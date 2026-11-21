#!/usr/bin/env python3
"""
One-command builder for the Zephyr RTOS vs. FreeRTOS presentation.

Usage:
  python build.py                  # Rebuild skeleton (if needed) + generate deck + apply animations
  python build.py --skeleton       # Force rebuild 34-slide XML skeleton from assets/template_base.pptx
  python build.py --no-anim        # Build PPTX without launching PowerPoint COM animations (Linux/macOS safe)
  python build.py --pdf --qa       # Also regenerate PDF backup and 1600x900 PNGs in system temp dir
  python build.py --out test.pptx  # Output to a custom path
"""
import argparse
import os
import re
import subprocess
import sys

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(REPO_ROOT, "src")
sys.path.insert(0, SRC_DIR)

import stage1_skeleton
import stage2_build_deck
from pptx import Presentation


def verify_deck(pptx_path):
    validate_py = os.path.join(REPO_ROOT, ".agents", "skills", "pptx", "scripts", "office", "validate.py")
    base_tpl = os.path.join(REPO_ROOT, "assets", "template_base.pptx")
    if os.path.exists(validate_py) and os.path.exists(base_tpl):
        subprocess.run([sys.executable, validate_py, pptx_path, "--original", base_tpl], check=True)

    prs = Presentation(pptx_path)
    pat = re.compile(r"(lorem|ipsum|jupiter|mercury|venus|neptune|placeholder|slidesgo|\bmars\b(?!\s+pathfinder))", re.I)
    notes_ok = 0
    hits = 0
    for si, s in enumerate(prs.slides, 1):
        if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip():
            notes_ok += 1
        for sh in s.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    t = p.text.strip()
                    if t and pat.search(t):
                        hits += 1
                        print(f"PLACEHOLDER HIT slide {si}: {t}")
    print(f"VERIFIED: {len(prs.slides)} slides, {notes_ok} speaker-notes slides, {hits} placeholder hits")
    assert hits == 0, "Unexpected placeholder text found!"


def main():
    parser = argparse.ArgumentParser(description="Build Zephyr RTOS vs. FreeRTOS presentation")
    parser.add_argument("--out", default=stage2_build_deck.OUT_PPTX, help="Output .pptx file path")
    parser.add_argument("--skeleton", action="store_true", help="Force rebuild skeleton from template_base.pptx")
    parser.add_argument("--no-anim", action="store_true", help="Skip PowerPoint COM animation pass")
    parser.add_argument("--pdf", action="store_true", help="Export Zephyr_RTOS_vs_FreeRTOS.pdf via PowerPoint")
    parser.add_argument("--qa", action="store_true", help="Export 1600x900 PNGs to build/qa-slides/ via PowerPoint")
    args = parser.parse_args()

    out_pptx = os.path.abspath(args.out)

    if args.skeleton or not os.path.exists(stage1_skeleton.SKEL):
        stage1_skeleton.build_skeleton()

    stage2_build_deck.build(out_pptx)

    if not args.no_anim:
        if os.name != "nt":
            print("Note: --no-anim skipped automatically (PowerPoint COM animations require Windows).")
        else:
            ps1 = os.path.join(SRC_DIR, "add_pro_animations.ps1")
            subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, "-TargetPptx", out_pptx],
                check=True,
            )

    verify_deck(out_pptx)

    if (args.pdf or args.qa) and os.name == "nt":
        qa_ps1 = os.path.join(SRC_DIR, "export_qa.ps1")
        cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", qa_ps1, "-TargetPptx", out_pptx]
        if args.pdf:
            cmd.append("-ExportPdf")
        if args.qa:
            cmd.append("-ExportPngs")
        subprocess.run(cmd, check=True)


if __name__ == "__main__":
    main()
