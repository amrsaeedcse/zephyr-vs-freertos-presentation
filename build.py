#!/usr/bin/env python3
"""
One-command builder for the Zephyr RTOS Embedded Systems presentation (Team 8 / CSE 421-422).

Usage:
  python build.py                  # Rebuild skeleton (if needed) + generate 32-slide deck + apply animations
  python build.py --skeleton       # Force rebuild 32-slide XML skeleton from assets/template_base.pptx
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
    assert len(prs.slides) == 32, f"Expected 32 slides, got {len(prs.slides)}"

    pat = re.compile(r"(lorem|ipsum|jupiter|mercury|venus|neptune|placeholder|slidesgo|\bmars\b(?!\s+pathfinder))", re.I)
    bad_terms_pat = re.compile(r"(zypher|gerrit)", re.I)
    notes_ok = 0
    hits = 0
    bad_hits = 0
    tanween_hits = 0

    for si, s in enumerate(prs.slides, 1):
        notes_txt = ""
        if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip():
            notes_txt = s.notes_slide.notes_text_frame.text.strip()
            notes_ok += 1
            if "\u0627\u064b" in notes_txt or any(c + "\u064b" in notes_txt for c in "\u0623\u0625\u0622"):
                tanween_hits += 1
                print(f"ARABIC TANWEEN HIT in speaker notes slide {si}")
            if bad_terms_pat.search(notes_txt):
                bad_hits += 1
                print(f"FORBIDDEN TERM HIT in notes slide {si}")

        for sh in s.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    t = p.text.strip()
                    if t and pat.search(t):
                        hits += 1
                        print(f"PLACEHOLDER HIT slide {si}: {t}")
                    if t and bad_terms_pat.search(t):
                        bad_hits += 1
                        print(f"FORBIDDEN TERM HIT slide {si}: {t}")
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        t = cell.text.strip()
                        if t and pat.search(t):
                            hits += 1
                            print(f"PLACEHOLDER TABLE HIT slide {si}: {t}")
                        if t and bad_terms_pat.search(t):
                            bad_hits += 1
                            print(f"FORBIDDEN TERM TABLE HIT slide {si}: {t}")

    print(f"VERIFIED: {len(prs.slides)} slides, {notes_ok} speaker-notes slides, {hits} placeholder hits, {bad_hits} bad-term hits, {tanween_hits} tanween hits")
    assert notes_ok == 32, f"Expected 32 slides with speaker notes, got {notes_ok}"
    assert hits == 0, "Unexpected placeholder text found!"
    assert bad_hits == 0, "Unexpected forbidden term (Zypher/Gerrit) found!"
    assert tanween_hits == 0, "Unexpected Arabic Tanween on Alef found in notes!"


def main():
    parser = argparse.ArgumentParser(description="Build Zephyr RTOS Embedded Systems presentation (32 slides)")
    parser.add_argument("--out", default=stage2_build_deck.OUT_PPTX, help="Output .pptx file path")
    parser.add_argument("--skeleton", action="store_true", help="Force rebuild skeleton from template_base.pptx")
    parser.add_argument("--no-anim", action="store_true", help="Skip PowerPoint COM animation pass")
    parser.add_argument("--pdf", action="store_true", help="Export Zephyr_RTOS_Embedded_Systems.pdf via PowerPoint")
    parser.add_argument("--qa", action="store_true", help="Export 1600x900 PNGs to temp dir via PowerPoint")
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
