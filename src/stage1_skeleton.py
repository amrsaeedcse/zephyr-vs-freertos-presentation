import os, shutil, subprocess, sys, tempfile, zipfile, re
from defusedxml import minidom

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SRC_DIR)
ASSETS_DIR = os.path.join(REPO_ROOT, "assets")

TPL = os.path.join(ASSETS_DIR, "template_base.pptx")
WORK = os.path.join(tempfile.gettempdir(), "zephyr_embedded_skel_work")
SKEL = os.path.join(ASSETS_DIR, "template_skeleton_32.pptx")
SCRIPTS = os.path.join(REPO_ROOT, ".agents", "skills", "pptx", "scripts")
ADD_SLIDE = os.path.join(SCRIPTS, "add_slide.py")
CLEAN = os.path.join(SCRIPTS, "clean.py")
VALIDATE = os.path.join(SCRIPTS, "office", "validate.py")

# Exact 32-slide sequence mapped to original template slide numbers (1-indexed)
# Split: S1 Shared Cover | S2-7 Speaker 1 | S8-13 Speaker 2 | S14-19 Speaker 3 | S20-25 Speaker 4 | S26-31 Speaker 5 | S32 Shared Verdict
PLAN = [
    1,   # 01: Hero Cover — Zephyr RTOS: Silicon-to-Cloud Embedded OS (Shared: Team 8 / CSE 421/422)
    # --- Speaker 1: Omar Abdelsalam Abdelfattah (Slides 2-7) ---
    3,   # 02: 6-Stage Presentation Roadmap (30-Min / 5-Speaker Plan)
    4,   # 03: Pure Embedded Systems: Direct Silicon Control + Grand Protocol Unifier (mcu_silicon.jpg)
    6,   # 04: Why RTOS Is Mandatory: CSE 421/422 Breaking Points (4 Course Traps)
    19,  # 05: Real-Time Determinism vs. Average Throughput (Hard/Firm/Soft Deadlines)
    6,   # 06: Firmware Concurrency Models: Super-Loop vs. ISR-Heavy (volatile) vs. RTOS
    12,  # 07: The Golden Course Map: ATmega32 (Lec 1-10) -> Zephyr RTOS (8-Row Table)
    # --- Speaker 2: Amr Abdelazim Saeed (Slides 8-13) ---
    6,   # 08: Single-Core Execution Illusion & 4 Thread States (Time-Slicing + Blocking Discipline)
    12,  # 09: Inside Every Thread: SRAM Layout (TCB, PSP/MSP, MPU Stack Guard & Static-First)
    8,   # 10: Hardware Context Switching: AVR RETI -> ARM Cortex-M SysTick, z_swap() & PendSV (Prio 15)
    6,   # 11: Interrupt Latency & Selective Critical Sections: BASEPRI vs. PRIMASK / cli()
    8,   # 12: Priority Inversion & The Mars Pathfinder (1997) Bug — Solved by k_mutex Inheritance
    6,   # 13: Day 1 vs. Day 90 Memory Reality: Heap Fragmentation & Static-First Allocation
    # --- Speaker 3: Omar Ahmed Abdelaziz Mashhour (Slides 14-19) ---
    5,   # 14: Section Pivot — Enter Zephyr RTOS v4.5 (Oct 2026 Reality & Real Hardware, zephyr_boards.jpg)
    15,  # 15: Zephyr RTOS: Full "Batteries-Included" Embedded OS Canvas (Trims DOWN vs Builds UP)
    18,  # 16: Devicetree: 3-Layer Hardware Decoupling (.dtsi -> .dts -> app.overlay -> 0 C Edits)
    11,  # 17: 0 Bytes RAM for Devicetree: .rodata Flash Compilation & Real Footprint Chart
    6,   # 18: Unified Driver Model: Solving Lecture 2 Slide 45 (Flash V-Tables + static inline)
    8,   # 19: Before main() Starts: 5 Linker-Sorted SYS_INIT Boot Levels
    # --- Speaker 4: Ali Mohamed Ali (Slides 20-25) ---
    17,  # 20: Kconfig (prj.conf), CMake Dead-Code Pruning & west (--sysbuild / west spdx)
    9,   # 21: Rich Scheduling & Tickless Idle (Meta-IRQ, Coop <0, Preempt >=0, EDF, Workqueues, CTC)
    19,  # 22: Smarter Interrupts: Why No *FromISR()? (_isr_wrapper, 1-cycle IPSR, Direct, Zero-Latency)
    6,   # 23: Deterministic Memory & Power: O(1) k_mem_slab vs. sys_mem_blocks SRAM Bank Power-Off
    17,  # 24: IPC Toolbox, The k_fifo Stack Trap & Zbus Pub/Sub
    6,   # 25: Hardware Memory Isolation: User Mode Threads, k_mem_domain & SVC Firewall
    # --- Speaker 5: Ahmed Hassan Mohamed Mostafa (Slides 26-31) ---
    7,   # 26: Embedded Protocol Unifier (All 16 Class Teams), TrustZone (TF-M), MCUboot & SIL 3 Safety
    19,  # 27: Automatic Power Management: Midterm Q2 Automated (PM Policy, Runtime PM, Domains)
    6,   # 28: Course Upgrade Showdown: Lec 7 HVAC (ATmega32 TCNT1/ADATE -> Portable Zephyr C)
    12,  # 29: Embedded RTOS Matrix: Bare-Metal (ATmega32) vs. FreeRTOS vs. Zephyr v4.5
    4,   # 30: Boot Firmware on Your PC: native_sim, west twister & Shell Demo (embedded_lab.jpg)
    6,   # 31: Honest Engineering Trade-Offs & Embedded Linux / AMP Boundary
    # --- Shared Closing Slide ---
    20,  # 32: Engineering Verdict — Decision Guide & Thank You / Q&A Ready
]


def build_skeleton(out_skel=SKEL, validate=True):
    assert len(PLAN) == 32, f"Expected 32 slides in PLAN, got {len(PLAN)}"
    os.makedirs(os.path.dirname(out_skel), exist_ok=True)
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    with zipfile.ZipFile(TPL, "r") as zf:
        zf.extractall(WORK)

    used_original = set()
    final_slide_files = []

    for src_num in PLAN:
        src_name = f"slide{src_num}.xml"
        if src_num not in used_original:
            used_original.add(src_num)
            final_slide_files.append(src_name)
        else:
            proc = subprocess.run(
                [sys.executable, ADD_SLIDE, WORK, src_name],
                capture_output=True, text=True, check=True
            )
            m = re.search(r"(slide\d+\.xml)", proc.stdout)
            if not m:
                raise RuntimeError(f"Could not parse add_slide output: {proc.stdout} {proc.stderr}")
            final_slide_files.append(m.group(1))

    pres_rels_path = os.path.join(WORK, "ppt", "_rels", "presentation.xml.rels")
    rels_doc = minidom.parse(pres_rels_path)
    target_to_rid = {}
    for rel in rels_doc.getElementsByTagName("Relationship"):
        tgt = rel.getAttribute("Target").replace("\\", "/")
        if tgt.startswith("slides/"):
            target_to_rid[os.path.basename(tgt)] = rel.getAttribute("Id")

    pres_path = os.path.join(WORK, "ppt", "presentation.xml")
    pres_doc = minidom.parse(pres_path)
    sld_id_lst = pres_doc.getElementsByTagName("p:sldIdLst")[0]
    rid_to_node = {}
    for node in list(sld_id_lst.childNodes):
        if node.nodeType == node.ELEMENT_NODE and node.tagName == "p:sldId":
            rid = node.getAttribute("r:id")
            rid_to_node[rid] = node
            sld_id_lst.removeChild(node)

    for sname in final_slide_files:
        sld_id_lst.appendChild(rid_to_node[target_to_rid[sname]])

    with open(pres_path, "wb") as f:
        f.write(pres_doc.toxml(encoding="utf-8"))

    subprocess.run([sys.executable, CLEAN, WORK], check=True)

    if os.path.exists(out_skel):
        os.remove(out_skel)
    with zipfile.ZipFile(out_skel, "w", zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(WORK):
            for fn in files:
                full = os.path.join(root_dir, fn)
                rel = os.path.relpath(full, WORK).replace("\\", "/")
                zf.write(full, rel)

    if os.path.exists(WORK):
        shutil.rmtree(WORK)

    print("BUILT SKELETON 32:", out_skel, os.path.getsize(out_skel))
    if validate and os.path.exists(VALIDATE):
        subprocess.run([sys.executable, VALIDATE, out_skel, "--original", TPL], check=True)
    return out_skel


if __name__ == "__main__":
    build_skeleton()
