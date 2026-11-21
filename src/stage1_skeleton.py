import os, shutil, subprocess, sys, tempfile, zipfile, re
from defusedxml import minidom

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SRC_DIR)
ASSETS_DIR = os.path.join(REPO_ROOT, "assets")

TPL = os.path.join(ASSETS_DIR, "template_base.pptx")
WORK = os.path.join(tempfile.gettempdir(), "zephyr_freertos_skel_work")
SKEL = os.path.join(ASSETS_DIR, "template_skeleton_34.pptx")
SCRIPTS = os.path.join(REPO_ROOT, ".agents", "skills", "pptx", "scripts")
ADD_SLIDE = os.path.join(SCRIPTS, "add_slide.py")
CLEAN = os.path.join(SCRIPTS, "clean.py")
VALIDATE = os.path.join(SCRIPTS, "office", "validate.py")

# Exact 34-slide sequence mapped to original template slide numbers (1-indexed)
PLAN = [
    1,   # 01: Cover (UX balanced, zero ugly white box)
    3,   # 02: Presentation Roadmap (6 coherent stages + clean arrows)
    4,   # 03: Why RTOS? Super-Loop Stall vs Event-Blocked Threads (with MCU photo)
    6,   # 04: Inside Every Thread (TCB + Private Stack + 4-State Lifecycle Diagram)
    8,   # 05: How Threads Share One CPU (Preemptive Priority & Round-Robin Lanes)
    8,   # 06: Inside Every 1 ms System Tick (4-Stage Left-to-Right Pipeline + Context Switch)
    19,  # 07: Blocking Discipline (xTicksToWait Concentric Rings + Narrative Story)
    6,   # 08: Thread Coordination (Queues vs Binary Semaphores vs Mutexes)
    8,   # 09: The Priority Inversion Trap & Priority Inheritance Fix (Visual Tracks)
    6,   # 10: Fast Shared-Data Protection (Critical Sections & BASEPRI Threshold Ladder)
    8,   # 11: Vector Table Reality (Why FreeRTOS Needs *FromISR() & Manual Exit Switch)
    12,  # 12: Memory Allocation & The Swiss-Cheese Fragmentation Trap (Day 1 vs Day 90 Visual)
    5,   # 13: Architectural Pivot 02 — Moving Beyond Micro-Kernels: Enter Zephyr RTOS (Boards Photo)
    15,  # 14: Zephyr Batteries-Included Embedded OS Canvas
    18,  # 15: Devicetree 3-Layer Hardware Decoupling Tree
    11,  # 16: 0 Bytes RAM (.rodata) & Real Footprint Chart
    6,   # 17: Unified Driver Model (gpio_dt_spec -> struct device -> api + static inline)
    8,   # 18: Ordered SYS_INIT Boot Levels Before main() (5-Stage Left-to-Right Sequence)
    17,  # 19: Kconfig (prj.conf) + CMake + west & Sysbuild
    9,   # 20: Rich Scheduling: Cooperative (<0), Preemptive (>=0), EDF, Workqueues, Tickless & SMP
    19,  # 21: Smarter Interrupts (_isr_wrapper, 1-cycle IPSR, Direct & Zero-Latency)
    6,   # 22: Deterministic Memory (k_mem_slab Intrusive List vs sys_mem_blocks Power Gating)
    17,  # 23: IPC Toolbox & The k_fifo Stack Lifetime Trap
    7,   # 24: Zbus Pub/Sub Software Bus & Zero-Copy Multi-Core AMP
    18,  # 25: MPU User Mode, Memory Domains & Hardware Stack Guards
    7,   # 26: ARM TrustZone (TF-M), MCUboot Atomic OTA & Safety Certifications
    19,  # 27: Automatic Power Management (PM Policy Manager & Device Runtime PM)
    7,   # 28: Built-In Connectivity & Storage Stacks (BLE, IPv6/Matter, LittleFS)
    4,   # 29: Boot Firmware on Your PC: native_sim, west twister & Observability (Lab Photo)
    6,   # 30: Side-by-Side Code Showdown (freertos_stm32.c vs zephyr_portable.c)
    16,  # 31: Where Zephyr Hurts: 4 Honest Engineering Trade-Offs
    12,  # 32: FreeRTOS vs Zephyr: Full Architecture Matrix
    12,  # 33: References & Official Technical Documentation
    20,  # 34: Engineering Verdict — Thank You & Q&A
]


def build_skeleton(out_skel=SKEL, validate=True):
    assert len(PLAN) == 34
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

    print("BUILT SKELETON 34:", out_skel, os.path.getsize(out_skel))
    if validate and os.path.exists(VALIDATE):
        subprocess.run([sys.executable, VALIDATE, out_skel, "--original", TPL], check=True)
    return out_skel


if __name__ == "__main__":
    build_skeleton()
