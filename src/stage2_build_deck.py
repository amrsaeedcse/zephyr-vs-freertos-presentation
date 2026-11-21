import os, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from stage2_populate_utils import (
    set_text, set_cell_text, remove_shape, replace_pic_with_cropped,
    add_code_panel, add_pill_banner, add_monochrome_footprint_chart,
    add_chain_box, add_verdict_card, add_right_arrow, add_down_arrow, add_up_arrow,
    add_left_right_arrow, add_lead_subtitle, add_card, add_lane_block,
    clear_fill_and_line, set_notes,
    COL_CHARCOAL, COL_MUTED, COL_SILVER, COL_PLATINUM, COL_WHITE, COL_SOFTBG,
    FONT_HEAD, FONT_BODY, FONT_CODE, A_NS
)

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SRC_DIR)
ASSETS_DIR = os.path.join(REPO_ROOT, "assets")
SKEL = os.path.join(ASSETS_DIR, "template_skeleton_34.pptx")
OUT_PPTX = os.path.join(REPO_ROOT, "Zephyr_RTOS_vs_FreeRTOS.pptx")

IMG_MCU = os.path.join(ASSETS_DIR, "images", "mcu_silicon.jpg")
IMG_BOARDS = os.path.join(ASSETS_DIR, "images", "zephyr_boards.jpg")
IMG_LAB = os.path.join(ASSETS_DIR, "images", "embedded_lab.jpg")


def style_table_borders(cell, color_hex="BABABA", width_emu="6350"):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        old = tcPr.find(f"{{{A_NS}}}{tag}")
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, f"{{{A_NS}}}{tag}")
        ln.set("w", width_emu)
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        sf = etree.SubElement(ln, f"{{{A_NS}}}solidFill")
        srgb = etree.SubElement(sf, f"{{{A_NS}}}srgbClr")
        srgb.set("val", color_hex)


def add_clean_table(slide, x, y, w, h, headers, rows, col_widths, header_size=9.0, body_size=8.3, col0_bold=True):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = shape.table
    for ci, cw in enumerate(col_widths):
        tbl.columns[ci].width = Inches(cw)

    for ci, htxt in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COL_CHARCOAL
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.04)
        cell.margin_bottom = Inches(0.04)
        style_table_borders(cell, "595959")
        set_cell_text(cell, htxt, size_pt=header_size, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="left")

    for ri, rdata in enumerate(rows, 1):
        bg = COL_WHITE if ri % 2 == 1 else COL_SOFTBG
        for ci, val in enumerate(rdata):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.03)
            cell.margin_bottom = Inches(0.03)
            style_table_borders(cell, "D0D0D0")
            is_b = col0_bold and ci == 0
            col_hex = "383838" if is_b else "595959"
            set_cell_text(cell, val, size_pt=body_size, bold=is_b, color_hex=col_hex, font_name=FONT_BODY, align="left")
    return shape


def find_pic(slide):
    for sh in slide.shapes:
        if sh._element.tag.endswith("}pic"):
            return sh
    return None


def prepare_clean_header_slide(slide, title_idx, title_text, size_pt=21.5):
    """Keep ONLY the main title placeholder at consistent top y=0.40" and remove extra template junk."""
    title_sh = slide.shapes[title_idx]
    for sh in list(slide.shapes):
        if sh._element is not title_sh._element:
            remove_shape(sh)
    title_sh.left = Inches(0.75)
    title_sh.top = Inches(0.38)
    title_sh.width = Inches(8.50)
    title_sh.height = Inches(0.54)
    set_text(title_sh, title_text, size_pt=size_pt, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")
    return title_sh


def build(out_path=OUT_PPTX):
    if not os.path.exists(SKEL):
        import stage1_skeleton
        stage1_skeleton.build_skeleton()
    prs = Presentation(SKEL)
    assert len(prs.slides) == 34, f"Expected 34 slides, got {len(prs.slides)}"

    # =========================================================================
    # SLIDE 1: HERO COVER (Clean Platinum Canvas — NO harsh white box behind subtitle!)
    # =========================================================================
    s = prs.slides[0]
    s.shapes[0].left = Inches(3.79)
    s.shapes[0].top = Inches(1.70)
    s.shapes[0].width = Inches(5.43)
    s.shapes[0].height = Inches(1.55)
    s.shapes[0].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[0], ["Zephyr RTOS", "vs. FreeRTOS"], size_pt=43, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")

    # Remove the opaque white slab box behind subtitle so typography floats naturally on canvas
    clear_fill_and_line(s.shapes[1])
    s.shapes[1].left = Inches(3.79)
    s.shapes[1].top = Inches(3.32)
    s.shapes[1].width = Inches(5.43)
    s.shapes[1].height = Inches(0.48)
    s.shapes[1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[1], "From Real-Time Kernel Internals to a Full Embedded OS", size_pt=13.5, bold=False, color_hex="595959", font_name=FONT_BODY, align="center")
    set_notes(s, (
        "أهلاً بكم يا دكتور ويا شباب. النهاردة مش هنعرض Zephyr RTOS ككلام نظري؛ "
        "هنبدأ الأول بسرد هندسي واضح وسريع يشرح إزاي الـ RTOS وFreeRTOS بيشتغلوا من جوه خطوة بخطوة "
        "(تقسيم المهام، حالات الـ Thread، الجدولة، الـ Blocking، المزامنة، المقاطعات، والذاكرة)، "
        "عشان لما ندخل في Zephyr تشوف كل ميزة كحل مباشر لمشكلة حقيقية في الإنتاج."
    ))

    # =========================================================================
    # SLIDE 2: EXECUTIVE PRESENTATION ROADMAP (Coherent Narrative + Clean Uniform Arrows)
    # =========================================================================
    s = prs.slides[1]
    set_text(s.shapes[0], "Presentation Roadmap", size_pt=25, bold=True, color_hex="383838", font_name=FONT_HEAD)
    # Remove uneven template diagonal stripe groups ([13], [14])
    for sh in [s.shapes[14], s.shapes[13]]:
        remove_shape(sh)

    # Widen label boxes evenly centered beneath boxes 01..06 (Proportional: 2 Part-I + 4 Zephyr Core!)
    roadmap_cfg = [
        (7,  0.89, 2.24, "RTOS & FreeRTOS Basics",       "Why RTOS, TCB/Stack, Preemption, Tick & Blocking"),
        (8,  3.77, 2.24, "Sync, Interrupts & Memory",    "Queues, Mutex/Priority Inversion, BASEPRI, *FromISR & Heap"),
        (9,  6.64, 2.24, "Zephyr Hardware & Build",      "Devicetree (.dts), .rodata zero-RAM, Drivers & Kconfig"),
        (10, 0.89, 4.14, "Zephyr Kernel Internals",      "Coop/EDF/SMP scheduling, unified ISRs & k_mem_slab"),
        (11, 3.77, 4.14, "IPC, Security & Power",        "k_fifo/Zbus, MPU Domains, TrustZone, OTA & PM Policy"),
        (12, 6.64, 4.14, "Simulation, Code & Verdict",   "Connectivity, native_sim, Code Showdown & Matrix"),
    ]
    for si, lx, ly, lhead, ldesc in roadmap_cfg:
        sh = s.shapes[si]
        sh.left = Inches(lx)
        sh.top = Inches(ly)
        sh.width = Inches(2.46)
        sh.height = Inches(0.84)
        set_text(
            sh,
            [[(lhead, True, "383838")], [(ldesc, False, "595959")]],
            size_pt=10.0, font_name=FONT_BODY, align="center", space_after_pt=3
        )
    # Add 4 uniform horizontal arrows cleanly connecting 01->02->03 and 04->05->06
    for ax in (3.44, 6.32):
        add_right_arrow(s, ax, 1.72, 0.24, 0.16)
        add_right_arrow(s, ax, 3.62, 0.24, 0.16)
    set_notes(s, (
        "دي خريطة العرض الكاملة في 6 محطات متوازنة: أول محطتين (01 و02) بنبني فيهم أساس الـ RTOS وFreeRTOS بسرعة ووضوح "
        "(ليه RTOS، الـ TCB والـ Stack، الجدولة، الـ Blocking، المزامنة وPriority Inversion، المقاطعات والذاكرة)، "
        "وباقي 4 محطات (03 لحد 06) هم صلب العرض عن Zephyr RTOS: الهاردوير والـ Devicetree، قلب الكيرنل، الأمان والباور، لحد المحاكاة ومقارنة الكود والقرار النهائي."
    ))

    # =========================================================================
    # SLIDE 3: WHY AN RTOS? THREADS + BLOCKING VS SEQUENTIAL SUPER-LOOP (MCU Photo!)
    # =========================================================================
    s = prs.slides[2]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_MCU)

    # Shift title below top template circuit lines (y >= 0.96") so ZERO chrome overlaps text!
    s.shapes[1].left = Inches(0.72)
    s.shapes[1].top = Inches(0.96)
    s.shapes[1].width = Inches(4.92)
    s.shapes[1].height = Inches(0.74)
    set_text(s.shapes[1], "Why an RTOS? Threads Over\nOne Sequential Loop", size_pt=20.0, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    # Remove old raw text box; replace with 2 crystal-clear contrast cards (no cryptic mini-boxes!)
    remove_shape(s.shapes[0])

    add_card(
        s, 0.72, 1.82, 4.92, 1.24,
        "1. The Single while(1) Loop Problem",
        [
            ("Sequential Execution: ", "Every task runs one after another inside a single loop."),
            ("Stall Trap: ", "Calling a delay, waiting on a slow sensor read, or polling a flag freezes the whole CPU—urgent tasks miss deadlines."),
        ],
        dark=False, title_pt=10.8, body_pt=8.9
    )
    add_card(
        s, 0.72, 3.18, 4.92, 1.30,
        "2. The RTOS Solution: Threads + Event Blocking",
        [
            ("Independent Threads: ", "Firmware splits into separate Threads, each handling one dedicated responsibility."),
            ("Event Blocking (0% CPU): ", "A waiting thread pauses cleanly until a Hardware Interrupt or Event wakes it—urgent threads run right away."),
        ],
        dark=True, title_pt=10.8, body_pt=8.9
    )

    add_pill_banner(s, 0.72, 4.62, 4.92, 0.42, "KEY IDEA", "Real-Time = Guaranteed deadlines where waiting never stalls urgent code.", font_size=8.3)
    set_notes(s, (
        "أول فكرة في القصة: في البرنامج العادي (while(1)) كل حاجة في مسار واحد؛ فلو عملت delay أو استنيت قراءة حساس أو polling، "
        "البرنامج كله بيقف مستني السطر ده والمهام الحرجة بتفوّت ميعادها. في الـ RTOS بنقسم البرنامج لـ Threads مستقلين، "
        "كل Thread مسئول عن مهمة واحدة؛ واللي مستني حاجة يتعمله Blocking ويركن بـ 0% CPU لحد ما ييجي Interrupt أو Event يصحيه!"
    ))

    # =========================================================================
    # SLIDE 4: INSIDE EVERY THREAD — TCB, THREAD STACK & 4 LIFECYCLE STATES
    # =========================================================================
    s = prs.slides[3]
    prepare_clean_header_slide(s, 1, "Inside Every Thread: TCB, Stack & Lifecycle States", size_pt=21.0)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Once firmware splits into Threads, what does the RTOS allocate per thread in RAM—and how do its 4 states transition?", size_pt=9.8)

    # Left Zone: Standard Embedded Engineering Terms (TCB & Thread Stack Frame)
    add_card(
        s, 0.75, 1.38, 3.75, 1.54,
        "1. TCB (Task Control Block)",
        [
            "Small kernel struct stored in RAM per thread:",
            ("• Priority & State: ", "Current priority level and lifecycle state."),
            ("• Stack Pointer (SP): ", "Points to the top of this thread's Stack Frame."),
            ("• Scheduler & Wait Links: ", "List pointers linking thread into Ready, Delayed, or Wait lists."),
        ],
        dark=False, title_pt=11.0, body_pt=8.7
    )
    add_card(
        s, 0.75, 3.04, 3.75, 1.54,
        "2. Thread Stack (Stack Frame in RAM)",
        [
            "Dedicated stack memory owned by each thread:",
            ("• Local C Frame: ", "Local variables & nested function return addresses."),
            ("• Saved Context Frame: ", "On Cortex-M switch, HW + PendSV push PC, LR, xPSR & R0–R12 onto this stack (TCB stores SP)."),
            ("• Isolation: ", "Separate stacks stop threads corrupting each other."),
        ],
        dark=False, soft=True, title_pt=10.8, body_pt=8.5
    )

    # Right Zone: 100% Accurate Connected 4-State Lifecycle Visual Diagram
    add_lane_block(s, 4.78, 1.38, 4.47, 0.32, "The 4 Thread Lifecycle States & Exact Transitions", dark=True, font_pt=9.2)

    # Top Row: READY <-> RUNNING
    add_card(
        s, 4.78, 1.78, 1.98, 0.90,
        "READY",
        "Prepared to run;\nwaiting by priority.",
        dark=False, align="center", title_pt=11.0, body_pt=8.3
    )
    add_left_right_arrow(s, 6.82, 2.16, 0.36, 0.14)
    add_card(
        s, 7.24, 1.78, 2.01, 0.90,
        "RUNNING",
        "Executing on the\nCPU core right now.",
        dark=True, align="center", title_pt=11.0, body_pt=8.3
    )

    # Transition Corridor: Strictly between READY/RUNNING and BLOCKED below them!
    add_up_arrow(s, 5.02, 2.75, 0.14, 0.32)
    add_lead_subtitle(s, 5.20, 2.74, 1.78, 0.34, "Event / IRQ / Timeout\nwakes to READY", size_pt=7.7, bold=True, color=COL_CHARCOAL)

    add_down_arrow(s, 7.50, 2.75, 0.14, 0.32)
    add_lead_subtitle(s, 7.68, 2.74, 1.57, 0.34, "Blocks on Delay,\nQueue, or Event", size_pt=7.7, bold=True, color=COL_CHARCOAL)

    # Middle Row: BLOCKED spanning directly underneath BOTH READY & RUNNING
    add_card(
        s, 4.78, 3.15, 4.47, 0.70,
        "BLOCKED (Waiting — 0% CPU)",
        "Sleeping until a Delay timeout, Queue message, Semaphore, or Interrupt wakes it back to READY.",
        dark=False, align="left", title_pt=10.2, body_pt=8.2
    )

    # Bottom Row: SUSPENDED completely separate (Manual API Only!)
    add_card(
        s, 4.78, 3.94, 4.47, 0.64,
        "SUSPENDED (Manual Freeze Only)",
        "Frozen explicitly via vTaskSuspend() — ignores all events & timeouts until manually woken by vTaskResume().",
        dark=False, soft=True, align="left", title_pt=10.0, body_pt=8.0
    )

    add_pill_banner(s, 0.75, 4.70, 8.50, 0.38, "CORE INSIGHT", "A Context Switch saves PC & CPU registers to the outgoing Thread Stack, updates SP in its TCB, and loads the next Thread Stack.", font_size=8.3)
    set_notes(s, (
        "طيب بعد ما قسمنا البرنامج لـ Threads، إيه اللي بيتحجز لكل Thread في الرام؟ حاجتين أساسيتين: "
        "1) الـ TCB (Task Control Block): فيه أولويته وحالته والـ Stack Pointer (SP) والروابط بتاعته في قوايم الكيرنل. "
        "2) الـ Thread Stack في الرام: بيتخزن فيه المتغيرات المحلية، والـ PC والـ Registers لما يتبدّل. "
        "وبالنسبة للـ 4 حالات: الـ RUNNING بيتبادل مع READY؛ ولما RUNNING يستنى Delay أو Event بينزل لـ BLOCKED، ولما الحدث ييجي بيطلع من BLOCKED لـ READY. "
        "أما SUSPENDED فده حالة مستقلة تماماً بتتجمد بإيدك (vTaskSuspend) ومبتصحاش بـ Events، مبتصحاش غير لما تصحّيها بإيدك بـ vTaskResume()!"
    ))

    # =========================================================================
    # SLIDE 5: HOW MULTIPLE THREADS SHARE ONE CPU (Preemption + Round-Robin Visual Lanes)
    # =========================================================================
    s = prs.slides[4]
    prepare_clean_header_slide(s, 1, "How Multiple Threads Share One CPU Core", size_pt=21.5)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "When several Ready threads want the single CPU core at once, two scheduling rules govern who executes:", size_pt=10.0)

    # Top Policy Row: Preemptive Priority
    add_card(
        s, 0.75, 1.38, 3.55, 1.54,
        "1. Preemptive Priority",
        [
            ("Urgency Wins Immediately: ", "High-priority threads NEVER wait behind low-priority work."),
            "As soon as a Hardware Interrupt (IRQ) or Event wakes a critical thread, it pauses the lower thread mid-run and resumes it afterward.",
            ("Priority Scale: ", "In FreeRTOS, higher number = higher priority (opposite of Zephyr!)."),
        ],
        dark=False, title_pt=11.5, body_pt=8.5
    )
    # Top Right Visual Lane: Priority Preemption in Action (Distinct Row Labels vs Timeline Blocks!)
    add_card(s, 4.46, 1.38, 4.79, 1.54, "Execution Timeline — Priority Preemption", [], dark=False, soft=True, title_pt=9.8)

    # Plain bold row labels on left (visually distinct from timeline blocks!)
    add_lead_subtitle(s, 4.58, 1.84, 1.05, 0.36, "High Priority\n(Motor Ctrl)", size_pt=8.0, bold=True, color=COL_CHARCOAL)
    add_lead_subtitle(s, 4.58, 2.38, 1.05, 0.36, "Low Priority\n(UI / Logging)", size_pt=8.0, bold=True, color=COL_MUTED)

    # Timeline execution blocks on right (t1 -> t2 -> t3)
    add_lane_block(s, 5.72, 1.80, 1.12, 0.40, "Blocked (Idle)", muted=True, font_pt=7.5)
    add_lane_block(s, 6.90, 1.80, 1.24, 0.40, "RUNS <50us!", dark=True, font_pt=8.0)
    add_lane_block(s, 8.20, 1.80, 0.92, 0.40, "Blocked", muted=True, font_pt=7.5)

    add_lane_block(s, 5.72, 2.34, 1.12, 0.40, "UI Running", dark=True, font_pt=7.8)
    add_lane_block(s, 6.90, 2.34, 1.24, 0.40, "Preempted", muted=False, font_pt=7.6)
    add_lane_block(s, 8.20, 2.34, 0.92, 0.40, "Resumes", dark=True, font_pt=7.8)

    # Bottom Policy Row: Round-Robin Time-Slicing
    add_card(
        s, 0.75, 3.06, 3.55, 1.48,
        "2. Round-Robin Time-Slicing",
        [
            ("Fair Sharing at Equal Priority: ", "Triggered when multiple Ready threads share the exact SAME priority."),
            "The scheduler grants each equal thread 1 System Tick (typically 1 ms), then rotates turns so none is starved.",
        ],
        dark=False, title_pt=11.5, body_pt=8.6
    )
    add_card(s, 4.46, 3.06, 4.79, 1.48, "Execution Timeline — Equal Priority (1 System Tick = 1 ms)", [], dark=False, soft=True, title_pt=9.6)
    slices = [
        ("Thread A\n1 Tick", True),
        ("Thread B\n1 Tick", False),
        ("Thread A\n1 Tick", True),
        ("Thread B\n1 Tick", False),
        ("Thread A\n1 Tick", True),
    ]
    for si_s, (slbl, is_dk) in enumerate(slices):
        sx = 4.62 + si_s * 0.90
        add_lane_block(s, sx, 3.54, 0.84, 0.78, slbl, dark=is_dk, muted=not is_dk, font_pt=8.0)

    add_pill_banner(s, 0.75, 4.68, 8.50, 0.38, "NEXT STEP", "What actually generates that 1 System Tick (1 ms) heartbeat and switches threads in hardware? SysTick & PendSV!", font_size=8.4)
    set_notes(s, (
        "طيب عندنا كذا Thread جاهزين، إزاي بيتشاركوا المعالج الواحد؟ بقانونين أساسيين: "
        "1) الأولوية القاطعة (Preemptive Priority): لو مهمة عالية زي التحكم في الموتور صحيت بسبب Hardware Interrupt (IRQ) أو Event، بتقطع مهمة الـ UI/Logging فوراً، ولما تخلص التانية تكمل (في FreeRTOS الرقم الأكبر = أولوية أعلى، عكس Zephyr). "
        "2) التناوب الزمني (Round-Robin): لو كذا Thread ليهم نفس الأولوية بالظبط، كل واحد ياخد 1 System Tick (غالباً 1ms) وبعدين يسلّم للتاني عشان محدش يجوع."
    ))

    # =========================================================================
    # SLIDE 6: INSIDE EVERY SYSTEM TICK (1 ms) — SYSTICK & PENDSV (Linked to Slide 5!)
    # =========================================================================
    s = prs.slides[5]
    prepare_clean_header_slide(s, 1, "Inside Every System Tick (1 ms): SysTick & PendSV", size_pt=20.8)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Every 1 System Tick (1 ms), the hardware SysTick timer updates OS time and enforces both scheduling rules:", size_pt=9.8)

    stages_6 = [
        ("1. SysTick IRQ",
         "Hardware timer fires every 1 System Tick (typically 1 ms = 1,000 Hz).", True),
        ("2. Tick++ & Wake",
         "Increments OS clock & moves expired timers from Blocked -> Ready.", False),
        ("3. Pick Winner",
         "Checks Ready list for Preemption (higher priority) or Round-Robin (equal).", False),
        ("4. portYIELD() -> PendSV",
         "Sets the PendSV flag and exits SysTick fast to switch safely.", True),
    ]
    for ci, (st_h, st_b, dk) in enumerate(stages_6):
        cx = 0.75 + ci * 2.18
        add_card(s, cx, 1.40, 1.96, 1.88, st_h, st_b, dark=dk, title_pt=10.0, body_pt=8.7)
        if ci < 3:
            add_right_arrow(s, cx + 1.98, 2.27, 0.18, 0.14)

    add_card(
        s, 0.75, 3.44, 6.22, 1.08,
        "How portYIELD() & PendSV (Software Interrupt, Priority 15) Work",
        [
            ("• From a Thread: ", "portYIELD() sets the PendSV flag—because it is an Interrupt (even at Priority 15), it switches threads immediately!"),
            ("• From SysTick / ISR: ", "portYIELD() sets the same flag, but because Priority 15 is the lowest, PendSV waits until all active Interrupts finish, then switches cleanly."),
        ],
        dark=False, soft=True, title_pt=9.8, body_pt=8.3
    )
    add_card(
        s, 7.11, 3.44, 2.14, 1.08,
        "NOTE: IRQ Order",
        "Interrupt priorities are REVERSED vs Threads!\n0 = Highest\n15 = Lowest (PendSV)",
        dark=True, align="center", title_pt=9.5, body_pt=7.9
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "BATTERY HINT", "Waking up 1,000 times per second even when ALL threads are asleep drains battery — Zephyr solves this with Tickless Idle!", font_size=8.4)
    set_notes(s, (
        "طيب إيه اللي بيحصل كل 1 System Tick (1ms)؟ "
        "1) تايمر الهاردوير SysTick يضرب مقاطعة، 2) الكيرنل يزود الـ System Tick counter ويصحي الـ Threads اللي وقت نومها خلص ويرجعها Ready، "
        "3) يفحص قايمة الـ Ready: لو فيه حد أعلى أولوية يعمل Preemption، ولو فيه حد نفس الأولوية يعمل Round-Robin، "
        "4) ينادي portYIELD() عشان يرفع علم PendSV (Pendable Service Call) ويخرج بسرعة. "
        "ولاحظ المعلومة الاحترافية دي: في هاردوير ARM NVIC الأولويات بالعكس (0 الأعلى و15 الأضعف، عكس أولويات الـ Threads في FreeRTOS!). "
        "فلما portYIELD() يرفع PendSV من جوه Thread عادي بيتنفذ فوراً، أما من جوه SysTick أو أي Interrupt فالهاردوير يأجّله تلقائياً لحد ما كل المقاطعات الحرجة تخلص وبعدين يدخله بـ Tail-Chaining!"
    ))

    # =========================================================================
    # SLIDE 7: BLOCKING DISCIPLINE — xTicksToWait (Rings + Narrative Lead-In!)
    # =========================================================================
    s = prs.slides[6]
    s.shapes[0].left = Inches(0.79)
    s.shapes[0].top = Inches(0.36)
    s.shapes[0].width = Inches(8.43)
    s.shapes[0].height = Inches(0.48)
    set_text(s.shapes[0], "Blocking Discipline: Choosing How Long a Thread Waits", size_pt=20.5, bold=True, color_hex="383838", font_name=FONT_HEAD)

    add_lead_subtitle(
        s, 0.79, 0.88, 8.43, 0.34,
        "Whenever a thread waits for a Queue message, Binary Semaphore signal, or Mutex lock, one wait parameter governs its sleep:",
        size_pt=9.6
    )

    # Slightly nudge circles down +0.16" so there's clean breathing room below lead subtitle
    for ci in [1, 2, 3, 4, 5, 6]:
        s.shapes[ci].top = int(s.shapes[ci].top + Inches(0.16))
    set_text(s.shapes[4], "MAX", size_pt=14, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "100 ms", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "0", size_pt=15, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    # Right side callouts shifted slightly (+0.14") for airy separation under lead subtitle
    for ri in [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]:
        s.shapes[ri].top = int(s.shapes[ri].top + Inches(0.14))

    set_text(s.shapes[7], "Wait Forever (portMAX_DELAY)", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[8], "Sleep at 0% CPU until the exact event arrives (a Queue item, Binary Semaphore alert, or Mutex unlock). Ideal for pure event-driven worker threads.", size_pt=8.9, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[9], "Bounded Timeout (e.g. 100 ms — Best!)", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[10], "Block cleanly at 0% CPU, waking immediately when the event arrives—or returning a timeout error after 100 ms so a stuck sensor never freezes your device.", size_pt=8.9, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[11], "Zero Wait (xTicksToWait = 0)", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[12], "Never block: check the Queue or Semaphore once right now and return immediately. Mandatory inside Hardware Interrupts (ISRs) or quick polls.", size_pt=8.9, color_hex="595959", font_name=FONT_BODY)
    set_notes(s, (
        "طيب لما Thread يطلب حاجة من Queue أو Mutex أو Binary Semaphore وملقاهاش جاهزة، هيستنى قد إيه؟ المعامل xTicksToWait عنده 3 اختيارات: "
        "1) portMAX_DELAY: نام بـ 0% CPU لحد ما الحدث اللي مستنيه ييجي أياً كان وقته (ممتاز للمهام الـ Event-Driven). "
        "2) وقت محدد زي 100ms: نام واستنى الحدث، ولو مجاش خلال 100ms اصحى بـ error وتصرف بدل ما السيستم يتعلق للأبد (أفضل اختيار عملي). "
        "3) صفر (0): متنامش خالص، افحص ورجّع النتيجة فوراً — وإجباري جوه الـ ISRs."
    ))

    # =========================================================================
    # SLIDE 8: COORDINATION PRIMITIVES — QUEUES, SEMAPHORES & MUTEXES (Clear Purpose!)
    # =========================================================================
    s = prs.slides[7]
    prepare_clean_header_slide(s, 1, "Coordinating Threads: Queues, Semaphores & Mutexes", size_pt=21.0)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Wait—what ARE those events and locks threads block on? Each tool solves a specific problem:", size_pt=10.0)

    coords = [
        ("1. Message Queue", False, [
            ("Why Not a Plain Buffer? ", "A raw C array risks race conditions, needs manual locks, and forces the reader to busy-poll for new data."),
            ("How a Queue Solves It: ", "Thread-safe FIFO copy-by-value buffer managed by the OS."),
            ("Auto Blocking: ", "Reader sleeps at 0% CPU when empty; Sender can block when full."),
        ]),
        ("2. Binary Semaphore", True, [
            ("What For: ", "Fast 1-bit wake-up signal when an external event happens (most commonly from a Hardware Interrupt / ISR)."),
            ("How It Works: ", "The ISR raises the flag; the sleeping worker thread wakes immediately."),
            ("Critical Rule: ", "Has NO ownership — use strictly for signaling events, NEVER to lock shared resources!"),
        ]),
        ("3. Mutex Lock", False, [
            ("What For: ", "Protecting shared data or shared peripherals (SPI, I2C, UART, global structs) used by multiple threads."),
            ("How It Works: ", "Strict ownership: ONLY the thread holding the lock can unlock it."),
            ("Priority Inheritance: ", "Prevents Priority Inversion (next slide) — NEVER lock inside an ISR!"),
        ]),
    ]
    for ci, (ctitle, is_soft, cbullets) in enumerate(coords):
        cx = 0.75 + ci * 2.90
        add_card(s, cx, 1.40, 2.70, 2.86, ctitle, cbullets, dark=False, soft=is_soft, title_pt=11.6, body_pt=8.5)

    # High-contrast Dark Pro Note contrasting Mutex (Thread<->Thread) vs Spinlock (Thread<->ISR / Multi-Core)!
    add_card(
        s, 0.75, 4.34, 8.50, 0.76,
        "MUTEX vs. SPINLOCK  (Thread <-> Thread   vs.   Thread <-> Interrupt)",
        [
            ("• Mutex: ", "Protects shared data between Threads (puts waiting thread to sleep at 0% CPU — illegal inside an ISR!)."),
            ("• Spinlock / Critical Section: ", "Protects shared data between Threads & Interrupts (or across CPU cores) — pauses local IRQs & spins briefly without sleeping!"),
        ],
        dark=True, align="left", title_pt=9.0, body_pt=7.9
    )
    set_notes(s, (
        "طيب ليه بنحتاج 3 أدوات تنسيق مختلفين؟ "
        "1) Message Queue: ليه منستخدمش Array/Buffer عادي؟ لأن الـ Buffer العادي فيه Race Conditions وبيخليك تعمل Polling طول الوقت؛ الـ Queue بيديك FIFO آمن مع Blocking تلقائي لو فاضي أو مليان. "
        "2) Binary Semaphore: علم/إشارة سريعة لتصحيّة Thread لما يحصل حدث خارجي (أشهرها من ISR) ومفهوش ملكية فممنوع نقفل بيه مورد مشترك. "
        "3) Mutex: قفل بملكية صارمة لحماية الداتا المشتركة بين الـ Threads (وبينيّم اللي مستني بـ 0% CPU وممنوع جوه ISR) — وبيحل مشكلة Priority Inversion! "
        "وملحوظة ذكية تحت: لو الـ Mutex بيحمي الداتا المشتركة بين الـ Threads، طب لو الداتا مشتركة بين Thread وInterrupt (أو بين كذا Core)؟ هنا بنستخدم Spinlock / Critical Section بيوقف المقاطعات محلياً لحظة من غير نوم!"
    ))

    # =========================================================================
    # SLIDE 9: PRIORITY INVERSION TRAP & INHERITANCE FIX (Crystal-Clear P1/P2/P3 Story!)
    # =========================================================================
    s = prs.slides[8]
    prepare_clean_header_slide(s, 1, "The Priority Inversion Trap & How Mutexes Fix It", size_pt=21.0)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Suppose 3 threads (Low P1, Medium P2, High P3) share one CPU—and P1 & P3 share a resource (shared data or SPI):", size_pt=9.6)

    # Top Panel: The Bug (Without Priority Inheritance)
    add_lane_block(s, 0.75, 1.36, 8.50, 0.32, "THE TRAP — Using a Plain Flag / Semaphore (Priority Inversion: P2 Delays P3!)", muted=True, font_pt=8.8)
    trap_steps = [
        ("1. Low (P1) Takes Lock", "Low thread (P1) locks the shared data / peripheral and starts working.", False),
        ("2. High (P3) Blocks", "Urgent High thread (P3) needs that same resource -> sleeps waiting for P1.", False),
        ("3. Medium (P2) Cuts In!", "Unrelated Medium thread (P2) wakes up & preempts P1 (since P2 > P1)!", True),
        ("4. Inverted! P3 Starved", "P1 can't run to unlock, so Medium (P2) indirectly delays High (P3)!", False),
    ]
    for ci, (sh_t, sh_b, dk) in enumerate(trap_steps):
        cx = 0.75 + ci * 2.18
        add_card(s, cx, 1.74, 1.96, 1.18, sh_t, sh_b, dark=dk, title_pt=9.6, body_pt=8.1)
        if ci < 3:
            add_right_arrow(s, cx + 1.98, 2.26, 0.18, 0.14)

    # Bottom Panel: The Mutex Fix (Priority Inheritance)
    add_lane_block(s, 0.75, 3.06, 8.50, 0.32, "THE FIX — Mutex Temporary Priority Inheritance", dark=True, font_pt=8.8)
    fix_steps = [
        ("1. High (P3) Waits", "High (P3) blocks on the Mutex currently owned by Low (P1).", False),
        ("2. Boost: P1 -> P3!", "Mutex temporarily raises Low's priority up to High's level (P3)!", True),
        ("3. Medium (P2) Blocked", "Because boosted P1 runs at P3, Medium (P2) CANNOT preempt it.", False),
        ("4. Unlock -> P3 Runs", "P1 finishes fast, unlocks Mutex, drops to P1, and P3 runs right away!", False),
    ]
    for ci, (sh_t, sh_b, dk) in enumerate(fix_steps):
        cx = 0.75 + ci * 2.18
        add_card(s, cx, 3.44, 1.96, 1.14, sh_t, sh_b, dark=dk, title_pt=9.6, body_pt=8.1)
        if ci < 3:
            add_right_arrow(s, cx + 1.98, 3.94, 0.18, 0.14)

    add_pill_banner(s, 0.75, 4.70, 8.50, 0.36, "REAL-WORLD BUG", "NASA's Mars Pathfinder suffered repeated resets from this exact trap until Priority Inheritance was enabled!", font_size=8.3)
    set_notes(s, (
        "طيب إيه هي الـ Priority Inversion؟ تخيل 3 مهام: Low (P1)، Medium (P2)، High (P3). "
        "1) الضعيفة P1 مسكت قفل عادي على داتا مشتركة (أو SPI)، 2) العالية P3 احتاجت نفس الداتا فنامت تستنى P1 تخلص، "
        "3) في اللحظة دي المتوسطة P2 ملهاش علاقة بالموضوع خالص صحيت وقطعت P1 عشان P2 > P1! النتيجة إن المتوسطة عطّلت العالية (انقلاب أولويات)! "
        "الحل في الـ Mutex اسمه Priority Inheritance: أول ما P3 تستنى، الكيرنل بيرفع P1 مؤقتاً لمستوى P3 عشان P2 متقدرش تقطعها، فتخلص بسرعة وتفك القفل وP3 تشتغل فوراً!"
    ))

    # =========================================================================
    # SLIDE 10: CRITICAL SECTIONS, SPINLOCK & BASEPRI (Thread <-> Interrupt Protection!)
    # =========================================================================
    s = prs.slides[9]
    prepare_clean_header_slide(s, 1, "Thread <-> Interrupt: Critical Section, Spinlock & BASEPRI", size_pt=19.2)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "A Mutex guards Thread <-> Thread. For Thread <-> Interrupt sharing, a Critical Section / Spinlock pauses IRQs—without blinding urgent HW:", size_pt=9.0)

    # Left: 3 super-clean cause-and-effect steps
    add_card(
        s, 0.75, 1.38, 4.15, 0.96,
        "1. Critical Section / Spinlock Concept",
        "Protects 2–3 lines that must finish together (e.g. sensor x & y) by pausing local IRQs (on multi-core SMP, a Spinlock adds a cross-core lock).",
        dark=False, title_pt=10.3, body_pt=8.3
    )
    add_card(
        s, 0.75, 2.46, 4.15, 0.96,
        "2. Why Not Disable ALL Interrupts?",
        "Turning off every interrupt (__disable_irq()) protects your variables—but delays urgent hardware (Emergency Stop, Motor PWM, DMA).",
        dark=False, soft=True, title_pt=10.3, body_pt=8.4
    )
    add_card(
        s, 0.75, 3.54, 4.15, 1.00,
        "3. Smart Fix: BASEPRI Cutoff",
        "BASEPRI = 5 pauses only OS interrupts (5..15), while urgent hardware (0..4) stays live! Setting BASEPRI = 0 opens all interrupts again.",
        dark=False, title_pt=10.3, body_pt=8.4
    )

    # Right: Visual Priority Cutoff Ladder + Clean 3-Line Code
    add_card(
        s, 5.10, 1.38, 4.15, 0.90,
        "Priorities 0–4  ·  ALWAYS LIVE",
        "Emergency Stop, Motor PWM & DMA — zero delay.\n(Rule: cannot call OS APIs.)",
        dark=True, title_pt=10.0, body_pt=8.6
    )
    add_lane_block(s, 5.10, 2.35, 4.15, 0.30, "━━━━━  BASEPRI Cutoff = Priority 5  ━━━━━", muted=True, font_pt=8.4)
    add_card(
        s, 5.10, 2.72, 4.15, 0.90,
        "Priorities 5–15  ·  PAUSED FOR 2 us",
        "SysTick, PendSV, UART & Sensors wait briefly.\n(Safe to call *FromISR() OS APIs.)",
        dark=False, title_pt=10.0, body_pt=8.6
    )
    add_code_panel(
        s, 5.10, 3.72, 4.15, 0.82,
        "Protecting Shared x & y Together (Critical Section)",
        "taskENTER_CRITICAL();         // Pause IRQs 5..15\nsensor_x = x; sensor_y = y;   // Finishes together!\ntaskEXIT_CRITICAL();          // BASEPRI = 0 (open all)",
        font_size=7.2, title_size=7.8
    )

    add_pill_banner(s, 0.75, 4.68, 8.50, 0.38, "1-CORE vs MULTI-CORE", "1-core MCU: BASEPRI masking alone protects the section. Multi-core SMP (Zephyr): k_spin_lock adds a Spinlock across cores!", font_size=8.1)
    set_notes(s, (
        "طيب افرض فيه Thread وInterrupt بيتشاركوا متغيرين صغيرين (speed وdir) في سطرين بس؛ هنا الـ Mutex تقيل ومينفعش أصلاً جوه الـ ISR! "
        "لو الـ Interrupt ضرب بين السطر الأول والتاني هيقرأ داتا مكسورة (نص قديم ونص جديد). "
        "طب ليه منعملش disable لكل الـ Interrupts؟ عشان لو قفلتهم كلهم ممكن تفوّت إنذار أمان حرج زي Emergency Stop أو حماية موتور! "
        "فعلى ARM Cortex-M، FreeRTOS بيستخدم ريجستر BASEPRI عشان يرسم خط فاصل (مثلاً أولوية 5): المقاطعات العادية من 5 لـ 15 بتقف لحظة (2 ميكروثانية)، أما المقاطعات الحرجة من 0 لـ 4 فبتضرب بصفر تأخير!"
    ))

    # =========================================================================
    # SLIDE 11: HARDWARE INTERRUPTS — WHY FREERTOS NEEDS *FromISR() (Flow + Real Code!)
    # =========================================================================
    s = prs.slides[10]
    prepare_clean_header_slide(s, 1, "Hardware Interrupts: Why FreeRTOS Needs *FromISR()", size_pt=20.8)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Why does calling normal xQueueSend() inside a Hardware Interrupt crash FreeRTOS—requiring separate *FromISR() APIs?", size_pt=9.5)

    # Top 4-step Concrete Hardware Trace (Explicit Vector Table + Clear *FromISR() Role!)
    isr_steps = [
        ("1. HW IRQ Fires", "Peripheral (e.g. UART) interrupts CPU & saves registers.", False),
        ("2. Vector Table Jump", "CPU reads Vector Table & jumps straight to your C ISR (NO OS wrapper!).", True),
        ("3. Call *FromISR()", "Does the same Queue push + safe BASEPRI save/restore (no blocking!), sets woken flag.", False),
        ("4. Yield on Exit", "Final line calls portYIELD_FROM_ISR(woken) to trigger PendSV right as ISR exits.", False),
    ]
    for ci, (ih, ib, dk) in enumerate(isr_steps):
        cx = 0.75 + ci * 2.18
        add_card(s, cx, 1.38, 1.96, 1.18, ih, ib, dark=dk, title_pt=9.6, body_pt=8.0)
        if ci < 3:
            add_right_arrow(s, cx + 1.98, 1.90, 0.18, 0.14)

    # Bottom Left: 3 Short, Smart Reasons | Bottom Right: 4-Line Real FreeRTOS C Pattern!
    add_card(
        s, 0.75, 2.70, 4.15, 1.84,
        "Why Normal xQueueSend() Fails in an ISR",
        [
            ("1. Deadlock + Wrong Thread: ", "If full, it loops & blocks the innocent interrupted Thread while PendSV cannot run!"),
            ("2. Breaks Nested IRQs: ", "taskEXIT_CRITICAL() resets BASEPRI = 0; *FromISR() safely saves & restores old_mask."),
            ("3. Manual Exit Yield: ", "No OS exit wrapper exists—you trigger PendSV once on the final line."),
        ],
        dark=False, title_pt=10.4, body_pt=8.4
    )
    add_code_panel(
        s, 5.10, 2.70, 4.15, 1.84,
        "Real FreeRTOS ISR Pattern (uart_isr.c)",
        "void UART_IRQHandler(void) {\n    BaseType_t woken = pdFALSE;\n\n    // Same Queue push, ISR-safe BASEPRI:\n    xQueueSendFromISR(rx_q, &byte, &woken);\n\n    // Triggers PendSV right on ISR exit:\n    portYIELD_FROM_ISR(woken);\n}",
        font_size=7.9, title_size=8.4
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "ZEPHYR PREVIEW", "Later we'll see how Zephyr adds a central _isr_wrapper so ONE unified API works in both Threads & ISRs!", font_size=8.3)
    set_notes(s, (
        "ليه في FreeRTOS مينفعش تنادي xQueueSend() العادية جوه Interrupt وعندنا *FromISR()؟ لـ 3 أسباب حقيقية من الكود: "
        "1) لو الـ Queue مليانة، xQueueSend() بتلف في while وتنادي portYIELD() عشان تنيّم الـ pxCurrentTCB (فتنيّم الـ Thread البريء اللي اتقطع)، وتعمل Deadlock لأن PendSV ممنوع يشتغل قبل ما الـ ISR يخرج! "
        "2) taskEXIT_CRITICAL() في الدالة العادية بتصفّر BASEPRI = 0 فبتكسر حماية أي Nested Interrupt قبلها؛ أما *FromISR() فبتحفظ الـ old_mask وترجّعه زي ما كان بالظبط. "
        "3) لأن FreeRTOS معندهاش OS wrapper على مخرج الـ ISR، بتملى فلاج woken وتنادي portYIELD_FROM_ISR(woken) في آخر سطر عشان يرفع PendSV ويبدّل أول ما الـ ISR يخرج!"
    ))

    # =========================================================================
    # SLIDE 12: MEMORY ALLOCATION & SWISS-CHEESE FRAGMENTATION TRAP (Day 1 vs Day 90 Visual!)
    # =========================================================================
    s = prs.slides[11]
    prepare_clean_header_slide(s, 0, "RAM Allocation & The Swiss-Cheese Fragmentation Trap", size_pt=20.8)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Where do Thread stacks & Queues get RAM? FreeRTOS offers Static buffers or Dynamic Heaps (heap_1–5)—here is the long-term trap:", size_pt=9.4)

    # Visual Diagram: Day 1 Clean Heap vs Day 90 Fragmented Heap
    add_card(s, 0.75, 1.34, 8.50, 1.84, "", [], dark=False, soft=True)
    # Row 1: Day 1
    add_lane_block(s, 0.88, 1.46, 1.48, 0.56, "Day 1\n40 KB Free", muted=False, font_pt=8.0)
    add_lane_block(s, 2.42, 1.46, 2.00, 0.56, "Used (24 KB)", dark=True, font_pt=8.2)
    add_lane_block(s, 4.48, 1.46, 3.24, 0.56, "40 KB Contiguous Free Space", muted=False, font_pt=8.5)
    add_lane_block(s, 7.78, 1.46, 1.34, 0.56, "Malloc(8 KB)\nOK!", muted=True, font_pt=8.0)

    # Row 2: Day 90 Swiss-Cheese Heap
    add_lane_block(s, 0.88, 2.14, 1.48, 0.58, "Day 90\nMax Hole = 4 KB", muted=False, font_pt=7.8)
    for k in range(10):
        bx = 2.42 + k * 0.532
        add_lane_block(s, bx, 2.14, 0.49, 0.58, "Use" if k % 2 == 0 else "4K", dark=(k % 2 == 0), font_pt=7.6)
    add_lane_block(s, 7.78, 2.14, 1.34, 0.58, "Malloc(8 KB)\nNULL!", dark=True, font_pt=8.2)

    add_lead_subtitle(s, 0.88, 2.82, 8.24, 0.28, "Why did 8 KB fail when 40 KB is free? Microcontrollers lack an MMU to stitch scattered physical 4 KB holes into one continuous block!", size_pt=8.5)

    # Bottom 2 Concise Strategy Cards
    add_card(
        s, 0.75, 3.28, 4.15, 1.26,
        "Option A · Static Allocation & heap_1",
        "Reserve fixed compile-time arrays (xTaskCreateStatic) or allocate-only heap_1 (never frees). Guarantees 0% fragmentation forever, with extra setup code.",
        dark=False, title_pt=10.5, body_pt=8.6
    )
    add_card(
        s, 5.10, 3.28, 4.15, 1.26,
        "Option B · Dynamic heap_4 / heap_5",
        "heap_4 merges adjacent free blocks (and heap_5 spans split RAM banks; heap_3 wraps libc). Still vulnerable to scattered holes over months of uptime.",
        dark=False, title_pt=10.5, body_pt=8.6
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "BRIDGE TO ZEPHYR", "To eliminate dynamic allocation traps, Zephyr defaults to Static RAM + O(1) fixed-size k_mem_slab pools!", font_size=8.4)
    set_notes(s, (
        "طيب كل الـ Threads والـ Stacks دي بتاخد مساحتها منين في الرام؟ قدامك يا تحجز Static وقت الكومبايل (آمن بس يدوي)، "
        "يا تستخدم Heap ديناميكي (أشهرهم heap_4 اللي بيدمج البلوكات المتجاورة). المشكلة الخطيرة إن بعد شهور من الشغل، الرام ممكن يبقى فيها 40KB فاضيين "
        "بس متقطعين حتت 4KB متفرقة (زي الجبنة السويسري)! لأن الميكروكنترولر معندوش MMU يضمهم، أول ما تطلب 8KB الـ malloc ترجع NULL والسيستم يقع!"
    ))

    # =========================================================================
    # SLIDE 13: SECTION PIVOT 02 — ENTER ZEPHYR RTOS (with Zephyr Boards Photo!)
    # =========================================================================
    s = prs.slides[12]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_BOARDS)
    set_text(s.shapes[1], "02", size_pt=32, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    s.shapes[1].top = Inches(0.82)
    s.shapes[1].left = Inches(4.45)
    s.shapes[1].width = Inches(0.85)
    s.shapes[1].height = Inches(0.85)

    s.shapes[0].top = Inches(0.85)
    s.shapes[0].left = Inches(5.45)
    s.shapes[0].width = Inches(3.85)
    s.shapes[0].height = Inches(0.95)
    set_text(s.shapes[0], "Enter Zephyr RTOS", size_pt=27, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    tb = s.shapes.add_textbox(Inches(4.45), Inches(1.88), Inches(4.80), Inches(3.15))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    items_13 = [
        ("The Architectural Wall in FreeRTOS", True, 11.5, "383838", 4),
        ("FreeRTOS nails core RTOS fundamentals—and stops there. Everything touching silicon drivers (STM32Cube HAL / ESP-IDF), *FromISR() glue, heap safety, lwIP, NimBLE & OTA is YOUR manual burden.", False, 9.4, "595959", 8),
        ("What Zephyr RTOS Rebuilds Completely:", True, 11.5, "383838", 4),
        ("• Unified Drivers + Devicetree — swap STM32 <-> nRF52 <-> ESP32 with untouched main.c", False, 9.3, "595959", 3),
        ("• Unified ISRs, O(1) Memory, Native BLE/Matter, MPU Domains, PM Policy & MCUboot in 1 repo.", False, 9.3, "595959", 6),
        ("Switch MCU in FreeRTOS? Rewrite HAL. Switch MCU in Zephyr? Swap -b <board> & overlay!", True, 9.4, "383838", 0),
    ]
    for i, (txt, bld, sz, col, sp) in enumerate(items_13):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        r.text = txt
        r.font.name = FONT_HEAD if bld and sz > 10 else FONT_BODY
        r.font.size = Pt(sz)
        r.font.bold = bld
        r.font.color.rgb = RGBColor.from_string(col)
    set_notes(s, (
        "هنا نقطة التحوّل في القصة كلها: إحنا دلوقتي فاهمين قلب الـ RTOS كويس. FreeRTOS بيحل المهام والـ Tick والمزامنة ببراعة — بس بيقف عند حدود الكيرنل! "
        "أي حاجة تانية (درايفرات الهاردوير HAL، دوال *FromISR، الـ Heap، الشبكات، البلوتوث، الـ OTA، إدارة الطاقة) إنت بتلحمهم بإيدك. "
        "دلوقتي ندخل في Zephyr RTOS ونشوف إزاي أعاد تصميم النظام المدمج بالكامل!"
    ))

    # =========================================================================
    # SLIDE 14: PILLARS 01 & 02 — BATTERIES-INCLUDED OS CANVAS
    # =========================================================================
    s = prs.slides[13]
    set_text(s.shapes[0], "Zephyr RTOS: Full \"Batteries-Included\" OS", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    tbl = s.shapes[1].table
    canvas_cells = [
        ((0, 0), "1,000+ Boards", True, 9.5),
        ((1, 0), "13+ CPU Archs:\nARM, RISC-V, Xtensa, x86, TriCore.\nST, Nordic, NXP, ESP, TI, Renesas.", False, 8.2),
        ((0, 1), "App & Zbus", True, 9.5),
        ((1, 1), "Portable C/C++, POSIX APIs & Pub/Sub Zbus IPC.", False, 8.0),
        ((2, 1), "HW Desc & Build", True, 9.5),
        ((3, 1), "Devicetree (.dts), Kconfig, CMake, west & native_sim.", False, 8.0),
        ((0, 2), "Linux Foundation · Apache 2.0", True, 8.7),
        ((1, 2), "1 Unified Upstream Repo with continuous CI testing.\nKernel + Drivers + Connectivity + Bootloader ship verified together with LTS branches & zero third-party version hell.", False, 8.2),
        ((0, 4), "Native Subsystems", True, 9.5),
        ((1, 4), "BLE 5.x, IPv4/6, MQTT, Matter, Thread, USB & LittleFS.", False, 8.0),
        ((2, 4), "Security & Boot", True, 9.5),
        ((3, 4), "TrustZone (TF-M), PSA Crypto, MPU Domains & MCUboot OTA.", False, 8.0),
        ((0, 5), "Kernel & Drivers", True, 9.5),
        ((1, 5), "Coop, Preempt, EDF, SMP, O(1) k_mem_slab, Tickless Idle & Unified gpio/uart/spi APIs.", False, 8.2),
        ((4, 0), "FreeRTOS Paradigm (Builds UP)", True, 9.2),
        ((5, 0), "Starts from a tiny kernel; your team integrates vendor HALs, networking, Bluetooth & OTA manually.", False, 8.2),
        ((4, 3), "Zephyr Paradigm (Trims DOWN)", True, 9.2),
        ((5, 3), "Starts from a cohesive Linux-inspired platform; Kconfig prunes every unused subsystem at compile time.", False, 8.2),
    ]
    for (r, c), txt, is_h, sz in canvas_cells:
        set_cell_text(
            tbl.cell(r, c), txt, size_pt=sz, bold=is_h,
            color_hex="383838" if is_h else "595959",
            font_name=FONT_HEAD if is_h else FONT_BODY,
            align="center" if is_h else "left"
        )
    set_notes(s, (
        "هنا ندخل في Zephyr RTOS: بدل ما تلحم 5 مكتبات من شركات مختلفة، Zephyr تحت مظلة Linux Foundation وبرخصة Apache 2.0 "
        "بيجمع الكيرنل والدرايفرات الموحدة وBLE والشبكات وأنظمة الملفات وMCUboot في مستودع واحد بيتعمله CI testing مع بعض، "
        "وبيدعم أكتر من 1,000 بوردة و13+ معمارية (ARM, RISC-V, Xtensa, وحتى TriCore وHexagon حديثاً)."
    ))

    # =========================================================================
    # SLIDE 15: PILLAR 03 — DEVICETREE 3-LAYER HARDWARE DECOUPLING
    # =========================================================================
    s = prs.slides[14]
    set_text(s.shapes[0], "Devicetree: Decoupling Silicon & Pins from C", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    tree_15 = {
        12: ("main.c (Zero HAL)", "383838", 9.0),
        13: ("3-Layer DT Cascade", "383838", 8.8),
        14: ("Swappable Targets", "383838", 8.8),
        18: ("1. .dtsi (SoC)", "FFFFFF", 8.5),
        17: ("2. .dts (Board)", "FFFFFF", 8.5),
        16: ("3. app.overlay", "FFFFFF", 8.5),
        15: ("west build -b ...", "FFFFFF", 8.5),
        22: ("Regs, IRQs & compatible", "383838", 7.8),
        21: ("Clocks & alias led0", "383838", 7.8),
        20: ("Sensors / bt-hci", "383838", 7.8),
        19: ("STM32F4 / nRF52840", "383838", 7.8),
        23: ("ESP32-S3 (0 C edits)", "383838", 7.8),
    }
    for si, (txt, col, sz) in tree_15.items():
        set_text(s.shapes[si], txt, size_pt=sz, bold=True, color_hex=col, font_name=FONT_BODY, align="center")
    add_pill_banner(
        s, 0.78, 4.78, 8.44, 0.36,
        "app.overlay",
        "&gpioa { compatible = \"st,stm32-gpio\"; status = \"okay\"; };  —  MCU shortage? Swap -b <board> & overlay; keep main.c untouched!",
        font_size=8.2
    )
    set_notes(s, (
        "في Zephyr، الهاردوير منفصل تماماً عن كود الـ C عبر 3 طبقات Devicetree: ملف .dtsi للسيليكون "
        "(عناوين الـ registers وهوية البريفرال compatible)، وملف .dts للبوردة (توصيلات البنات والـ aliases)، "
        "وملف app.overlay لتعديلات مشروعك. بكده نفس كود main.c يشتغل على STM32 أو nRF52 أو ESP32 بمجرد تغيير اسم البوردة!"
    ))

    # =========================================================================
    # SLIDE 16: PILLAR 04 — 0 BYTES RAM (.rodata) & FOOTPRINT CHART
    # =========================================================================
    s = prs.slides[15]
    set_text(s.shapes[0], "0 Bytes RAM for Devicetree: .rodata & Footprint", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    stats_16 = [
        (4,  13, "0 Bytes RAM",     "Runtime RAM parsing Devicetree"),
        (3,  12, "Python DTS",      "Generates devicetree_generated.h"),
        (2,  11, "Flash (.rodata)", "const struct device lives in ROM"),
        (1,  10, "~195 vs ~220 KB", "Full IoT* stack delta vanishes!"),
    ]
    for hi, bi, htxt, btxt in stats_16:
        set_text(s.shapes[hi], htxt, size_pt=13.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")
        set_text(s.shapes[bi], btxt, size_pt=8.8, bold=False, color_hex="595959", font_name=FONT_BODY, align="left")
    set_text(
        s.shapes[9],
        "*Full IoT Stack = BLE + MQTT + TLS + OTA. Minimal Zephyr kernel starts ~8–18 KB Flash / ~4 KB RAM (3–4 KB if multithreading off).",
        size_pt=7.8, color_hex="595959", font_name=FONT_BODY, align="center"
    )
    old_chart_pic = find_pic(s)
    if old_chart_pic:
        remove_shape(old_chart_pic)
    add_monochrome_footprint_chart(s, 4.02, 1.25, 5.25, 3.30)
    set_notes(s, (
        "السؤال اللي بيسأله أي مهندس: الـ Devicetree ده مش هياخد وقت وRAM عشان يتقري في الـ runtime؟ "
        "المفاجأة إنه بيتقري وقت الـ Build بس بسكربتات Python تطلع ملف devicetree_generated.h، "
        "والـ const struct بيتحط في الفلاش (.rodata) بـ 0 بايت RAM إضافي! وفي مشروع IoT متكامل (BLE+MQTT+TLS+OTA) الفرق مع FreeRTOS (~195 مقابل ~220 KB) يكاد يتلاشى."
    ))

    # =========================================================================
    # SLIDE 17: PILLAR 05A — UNIFIED DRIVER MODEL
    # =========================================================================
    s = prs.slides[16]
    prepare_clean_header_slide(s, 1, "Unified Driver Model: Flash Function Pointers", size_pt=21.5)

    chain_17 = [
        ("1. gpio_dt_spec", "const in .rodata\n*port, pin (5), flags", False),
        ("2. struct device", "const in .rodata\n*config, *api, *data", False),
        ("3. gpio_driver_api", "const V-Table (Flash)\nstm32_cfg, stm32_tog", True),
        ("4. gpio_stm32.c", "Silicon Registers\nWrites GPIOA->BSRR", False),
    ]
    for ci, (ch_h, ch_b, dk) in enumerate(chain_17):
        cx = 0.75 + ci * 2.18
        add_chain_box(s, cx, 1.26, 1.94, 1.16, ch_h, ch_b, dark=dk)
        if ci < 3:
            add_right_arrow(s, cx + 1.97, 1.77, 0.18, 0.14)

    tb17 = s.shapes.add_textbox(Inches(0.75), Inches(2.64), Inches(4.05), Inches(1.85))
    tf = tb17.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for idx_l, (pre, rest) in enumerate([
        ("Standard Driver APIs Across All Boards", ""),
        ("• Core Buses: ", "gpio.h, uart.h, i2c.h, spi.h, pwm.h, adc.h, can.h, dma.h."),
        ("• Sensor Framework (sensor.h): ", "Unified fetch/channel API across 100s of IMUs & sensors."),
        ("• Display & UI: ", "display.h binds directly into native LVGL graphics."),
    ]):
        p = tf.paragraphs[0] if idx_l == 0 else tf.add_paragraph()
        p.space_after = Pt(5 if idx_l == 0 else 4)
        r1 = p.add_run()
        r1.text = pre
        r1.font.name = FONT_HEAD if idx_l == 0 else FONT_BODY
        r1.font.size = Pt(11.5 if idx_l == 0 else 9.2)
        r1.font.bold = True
        r1.font.color.rgb = COL_CHARCOAL
        if rest:
            r2 = p.add_run()
            r2.text = rest
            r2.font.name = FONT_BODY
            r2.font.size = Pt(9.2)
            r2.font.color.rgb = COL_MUTED

    add_code_panel(
        s, 4.98, 2.62, 4.27, 1.84,
        "gpio.h — static inline eliminates call overhead!",
        "static inline int z_impl_gpio_pin_toggle_dt(\n"
        "    const struct gpio_dt_spec *spec) {\n"
        "  const struct gpio_driver_api *api = spec->port->api;\n"
        "  return api->port_toggle_bits(spec->port, BIT(spec->pin));\n"
        "} // Direct Flash v-table call (User Mode off)",
        font_size=7.4
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "LINUX VFS PATTERN", "C polymorphism via const v-tables in Flash — mutable driver state only in RAM!", font_size=8.6)
    set_notes(s, (
        "هنا سر توحيد الدرايفرات في Zephyr: إنت بتنادي دالة عامة زي gpio_pin_toggle_dt(&led). الدالة دي static inline "
        "(يعني الكومبايلر بيفكها مكانها من غير call overhead لو User Mode مقفول)، وبتمشي في سلسلة مؤشرات ثابتة في الفلاش: "
        "من gpio_dt_spec للـ struct device ومنه لجدول الـ Function Pointers (api) بتاع درايفر STM32 أو Nordic اللي يكتب في الـ Registers مباشرة!"
    ))

    # =========================================================================
    # SLIDE 18: PILLAR 05B — ORDERED SYS_INIT BOOT LEVELS (True Left->Right 5 Stages!)
    # =========================================================================
    s = prs.slides[17]
    prepare_clean_header_slide(s, 1, "Before main() Starts: Ordered SYS_INIT Boot Levels", size_pt=21.0)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "No more manual HAL_Init() ordering bugs—the linker sorts driver & subsystem pointers into 5 ordered Flash stages:", size_pt=9.6)

    boot_stages = [
        ("1. EARLY",        "Tiny arch & reset setup\nbefore drivers.",           "No Kernel",   False),
        ("2. PRE_KERNEL_1", "Clocks, IRQ ctrl,\nbasic UART & GPIO.",             "IRQ Stack",   False),
        ("3. PRE_KERNEL_2", "Drivers depending\non PRE_KERNEL_1.",               "IRQ Stack",   False),
        ("4. POST_KERNEL",  "Kernel alive! I2C/SPI\n& k_mutex OK!",              "Main Thread", True),
        ("5. APPLICATION",  "App subsystems right\nbefore main() runs.",         "Main Thread", False),
    ]
    for bi, (bt, bd, bctx, dk) in enumerate(boot_stages):
        bx = 0.75 + bi * 1.74
        add_card(s, bx, 1.40, 1.54, 1.70, bt, [bd, "", f"[{bctx}]"], dark=dk, align="center", title_pt=9.2, body_pt=8.0)
        if bi < 4:
            add_right_arrow(s, bx + 1.56, 2.18, 0.16, 0.14)

    add_card(
        s, 0.75, 3.28, 4.15, 1.24,
        "FreeRTOS / Bare HAL (Manual Order)",
        "You manually call HAL_Init(); SystemClock_Config(); MX_GPIO_Init(); MX_SPI1_Init(); in main()—wrong call order breaks peripherals silently!",
        dark=False, title_pt=10.5, body_pt=8.7
    )
    add_card(
        s, 5.10, 3.28, 4.15, 1.24,
        "Zephyr Linker-Sorted Init Table",
        "DEVICE_DT_DEFINE() & SYS_INIT(fn, POST_KERNEL, prio) register pointers in a sorted Flash section—booted automatically before main()!",
        dark=False, soft=True, title_pt=10.5, body_pt=8.7
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "RULE OF THUMB", "Never sleep or lock a k_mutex in PRE_KERNEL_1 — the scheduler isn't alive yet! Use POST_KERNEL for threaded drivers.", font_size=8.4)
    set_notes(s, (
        "في FreeRTOS إنت بتضطر تنادي MX_GPIO_Init وMX_SPI_Init بإيدك في main بالترتيب الصح. في Zephyr الدرايفرات بتتسجل في جدول "
        "في الفلاش مترتب وقت الـ Link بـ 5 مراحل متتالية من الشمال لليمين: EARLY ثم PRE_KERNEL_1 وPRE_KERNEL_2 على الـ Interrupt Stack "
        "(فممنوع فيها نوم أو Mutex)، وبعد ما الكيرنل يشتغل تبدأ POST_KERNEL وAPPLICATION وبعدهم main() تلاقي كل الأجهزة جاهزة!"
    ))

    # =========================================================================
    # SLIDE 19: PILLAR 06 — KCONFIG + CMAKE + WEST & SYSBUILD
    # =========================================================================
    s = prs.slides[18]
    set_text(s.shapes[0], "Kconfig + CMake + west: Dead-Code Pruning", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)

    add_clean_table(
        s, 0.75, 1.32, 4.45, 3.12,
        headers=["west Command", "What It Automates"],
        rows=[
            ["west init / update", "Syncs Zephyr, HALs, mbedTLS & MCUboot"],
            ["west build -b <board>", "Runs DT + Kconfig + CMake + Ninja"],
            ["west build -t menuconfig", "Interactive TUI for Kconfig knobs"],
            ["west flash / debug", "Unified J-Link, ST-Link & OpenOCD GDB"],
            ["west build --sysbuild", "Builds App + MCUboot + BLE Net core!"],
        ],
        col_widths=[1.85, 2.60], header_size=9.0, body_size=8.3
    )

    tb19 = s.shapes.add_textbox(Inches(5.38), Inches(1.30), Inches(3.87), Inches(1.38))
    tf = tb19.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for idx_l, (pre, rest) in enumerate([
        ("Why prj.conf Alongside Devicetree?", ""),
        ("• Devicetree = HW PRESENT: ", "Board has 3 UARTs, 2 SPI, BLE."),
        ("• prj.conf = SW WANTED: ", "CONFIG_GPIO=y compiles GPIO only."),
        ("• Auto-Deps: ", "CONFIG_BT=y pulls crypto & pools automatically."),
    ]):
        p = tf.paragraphs[0] if idx_l == 0 else tf.add_paragraph()
        p.space_after = Pt(4 if idx_l == 0 else 3)
        r1 = p.add_run()
        r1.text = pre
        r1.font.name = FONT_HEAD if idx_l == 0 else FONT_BODY
        r1.font.size = Pt(11 if idx_l == 0 else 8.8)
        r1.font.bold = True
        r1.font.color.rgb = COL_CHARCOAL
        if rest:
            r2 = p.add_run()
            r2.text = rest
            r2.font.name = FONT_BODY
            r2.font.size = Pt(8.8)
            r2.font.color.rgb = COL_MUTED

    add_code_panel(
        s, 5.38, 2.78, 3.87, 1.66,
        "drivers/gpio/CMakeLists.txt (Vendor Filter)",
        "# DT 'compatible' + Kconfig -> CONFIG_GPIO_STM32=y\n"
        "zephyr_library_sources_ifdef(\n"
        "    CONFIG_GPIO_STM32 gpio_stm32.c)\n"
        "zephyr_library_sources_ifdef(\n"
        "    CONFIG_GPIO_ESP32 gpio_esp32.c)\n"
        "# gpio_esp32.c NEVER enters STM32 Flash!",
        font_size=7.4
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "PRECISION", "Devicetree asks: 'What hardware is on the board?'  |  Kconfig asks: 'Which software enters Flash?'", font_size=8.6)
    set_notes(s, (
        "الفرق الذهبي اللي لازم نفرقه: الـ Devicetree بيقول إيه الهاردوير اللي موجود على البوردة، "
        "أما ملف prj.conf (Kconfig) فبيقول إيه السوفتوير اللي عايز تفعله في مشروعك النهاردة! "
        "Kconfig مع compatible بيطلعوا CONFIG_GPIO_STM32=y، وCMake بيصفي الملفات بحيث ملفات ESP32 متدخلش الـ Compiler أصلاً."
    ))

    # =========================================================================
    # SLIDE 20: PILLAR 07 — RICH SCHEDULING & SMP
    # =========================================================================
    s = prs.slides[19]
    set_text(s.shapes[0], "Rich Scheduling: Coop, Preempt, EDF & SMP", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    grid_20 = [
        (5,  1, "Cooperative (< 0)",      "Priorities -1, -2, -15: NEVER preempted by threads (only ISRs)! Runs until k_yield() or blocking — avoids many mutexes."),
        (6,  2, "Preemptive (>= 0)",      "Traditional preemption where 0 = highest (opposite of FreeRTOS!). Optional Round-Robin time-slicing for equal priorities."),
        (7,  8, "Deadline (EDF)",         "CONFIG_SCHED_DEADLINE dynamically prioritizes whichever thread's deadline (k_thread_deadline_set) expires soonest."),
        (10, 3, "Workqueues (k_work)",    "Offload heavy ISR or callback work onto sysworkq or custom delayable workqueues — saves extra thread stacks."),
        (11, 4, "Tickless Idle",          "Stops 1 ms tick IRQs when idle! Programs one low-power alarm for next wake-up (5,000 ms sleep = 1 wakeup)."),
        (12, 9, "SMP & CPU Masks",        "True Symmetric Multiprocessing on multi-core SoCs with ticket spinlocks & k_thread_cpu_mask_enable() pinning."),
    ]
    for hi, bi, htxt, btxt in grid_20:
        set_text(s.shapes[hi], htxt, size_pt=12, bold=True, color_hex="383838", font_name=FONT_HEAD)
        set_text(s.shapes[bi], btxt, size_pt=9.2, color_hex="595959", font_name=FONT_BODY)
    set_notes(s, (
        "في الجدولة، Zephyr بيجمع 4 أنواع مع بعض: أولويات سالبة Cooperative متتقطعش غير لما تسيب المعالج بمزاجها (فتوفر Mutexes كتير)، "
        "وأولويات موجبة Preemptive مع Round-Robin (وخدوا بالكم الرقم الأصغر = أولوية أعلى عكس FreeRTOS!)، "
        "ومعاهم EDF وWorkqueues وTickless Idle (اللي بيوقف مقاطعة الـ 1ms وقت النوم) ودعم كامل للـ SMP!"
    ))

    # =========================================================================
    # SLIDE 21: PILLAR 08 — SMARTER INTERRUPTS (_isr_wrapper & IPSR)
    # =========================================================================
    s = prs.slides[20]
    set_text(s.shapes[0], "Why No *FromISR(): _isr_wrapper & IPSR", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[4], "99%", size_pt=14, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "FAST", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "0-LAT", size_pt=12, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "IRQ_CONNECT (_isr_wrapper + IPSR)", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[8], "One unified k_msgq_put(..., K_NO_WAIT) uses k_spin_lock & checks HW IPSR in 1 cycle: Thread yields now; ISR defers PendSV to _isr_wrapper exit!", size_pt=8.9, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[9], "IRQ_DIRECT_CONNECT (Fast Path)", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[10], "HW Vector Table jumps straight to my_isr(), skipping _isr_wrapper & _sw_isr_table (~20–30 cycles saved). Call ISR_DIRECT_FOOTER(1) if waking.", size_pt=8.9, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[11], "IRQ_ZERO_LATENCY (Ultra-Critical)", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[12], "Direct IRQ firing ABOVE irq_lock() for sub-microsecond HW fault response. Strict rule: NEVER call any kernel API inside!", size_pt=8.9, color_hex="595959", font_name=FONT_BODY)
    set_notes(s, (
        "هنا بقى حل المشكلة اللي شفناها في FreeRTOS: في Zephyr الجدول الرئيسي بيشاور في الحالات العادية (IRQ_CONNECT) على دالة Assembly اسمها _isr_wrapper، "
        "ودي ماسكة الدخول والخروج. فالدالة الواحدة زي k_msgq_put(..., K_NO_WAIT) بتسأل ريجستر الهاردوير IPSR: لو في Thread تبدل فوراً، "
        "ولو في ISR تأجل التبديل لباب الخروج في _isr_wrapper فيضرب PendSV أوتوماتيك! وعندك IRQ_DIRECT_CONNECT وIRQ_ZERO_LATENCY للحالات الحرجة."
    ))

    # =========================================================================
    # SLIDE 22: PILLAR 09 — DETERMINISTIC MEMORY & RAM BANK POWER-OFF (Inline Visual Diagrams!)
    # =========================================================================
    s = prs.slides[21]
    prepare_clean_header_slide(s, 1, "O(1) Slabs & Powering Off Unused RAM Banks", size_pt=21.5)

    # Left Card: k_mem_slab + Intrusive Pointer Blocks
    add_card(
        s, 0.75, 1.18, 4.15, 3.00,
        "k_mem_slab: O(1) Intrusive List",
        [
            ("• Zero Fragmentation: ", "Equal-size blocks; O(1) alloc/free & thread pend."),
            ("• Catch: ", "4-byte next pointer sits INSIDE each free block — powering off SRAM Bank 1 wipes pointers!"),
        ],
        dark=False, title_pt=11.5, body_pt=8.9
    )
    for b in range(3):
        bx = 0.90 + b * 1.32
        add_lane_block(s, bx, 3.22, 1.18, 0.34, f"4B -> Blk {b+1}" if b < 2 else "4B -> NULL", dark=True, font_pt=7.6)
        add_lane_block(s, bx, 3.58, 1.18, 0.44, f"Block {b} (Free 60B)", muted=False, font_pt=7.6)
        if b < 2:
            add_right_arrow(s, bx + 1.19, 3.56, 0.11, 0.11)

    # Right Card: sys_mem_blocks + Bank 0 Bitmap Diagram
    add_card(
        s, 5.10, 1.18, 4.15, 3.00,
        "sys_mem_blocks: External Bitmap",
        [
            ("• Separated Metadata: ", "Free/used bitmap lives safely in Always-On SRAM Bank 0."),
            ("• Deep Sleep Savings: ", "When Bitmap == [0|0|0|0], Zephyr cuts power to SRAM Bank 1 safely!"),
        ],
        dark=False, soft=True, title_pt=11.5, body_pt=8.9
    )
    add_lane_block(s, 5.26, 3.22, 1.80, 0.80, "SRAM Bank 0 (ON)\nBitmap: [0 | 0 | 0 | 0]", muted=False, font_pt=7.8)
    add_right_arrow(s, 7.12, 3.55, 0.16, 0.12)
    add_lane_block(s, 7.34, 3.22, 1.76, 0.80, "SRAM Bank 1\n(POWER OFF!)", dark=True, font_pt=8.0)

    # Bottom 3 Memory Pillars
    add_chain_box(s, 0.75, 4.34, 2.70, 0.70, "Static-First (Default)", "K_THREAD_DEFINE reserves RAM at compile time", dark=False)
    add_chain_box(s, 3.65, 4.34, 2.70, 0.70, "k_heap & Multi-Heap", "Thread-safe dynamic pool across RAM banks", dark=False)
    add_chain_box(s, 6.55, 4.34, 2.70, 0.70, "net_buf (Zero-Copy)", "Ref-counted chained buffers for BLE / TCP", dark=False)
    set_notes(s, (
        "في الميموري، Zephyr بيحجز كل ما يمكن statically وقت الكومبايل (K_THREAD_DEFINE). وللبلوكات السريعة عنده k_mem_slab "
        "اللي بيشتغل في O(1) ومن غير fragmentation، بس بيكتب عنوان البلوك الفاضي اللي بعده جوه أول 4 بايت من البلوك نفسه. "
        "طب لو معاك شريحة بتطفي كهرباء SRAM Bank 1 وقت النوم؟ هنا sys_mem_blocks بيفصل الخريطة في Bitmap جوه Bank 0 الصاحية!"
    ))

    # =========================================================================
    # SLIDE 23: PILLAR 10A — IPC TOOLBOX & k_fifo STACK LIFETIME TRAP
    # =========================================================================
    s = prs.slides[22]
    set_text(s.shapes[0], "IPC Toolbox & The k_fifo Stack Lifetime Trap", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)

    add_clean_table(
        s, 0.75, 1.32, 4.55, 3.12,
        headers=["Primitive", "Mechanism & When to Use"],
        rows=[
            ["k_msgq", "Fixed ring buffer using memcpy. Safe with local Stack vars!"],
            ["k_fifo / k_lifo", "Zero-copy linked list! Embeds 4B reserved ptr at struct start."],
            ["k_pipe", "Byte-stream pipe for variable-length UART / audio chunks."],
            ["k_poll", "Wait on MULTIPLE semaphores, FIFOs & signals (POSIX poll)."],
            ["k_mutex / k_sem", "Mutex (Priority Inheritance + Ownership) vs Sem / k_event."],
        ],
        col_widths=[1.30, 3.25], header_size=9.0, body_size=8.2
    )
    add_code_panel(
        s, 5.48, 1.32, 3.77, 2.02,
        "dangling_fifo_bug.c  (CRASHES!)",
        "void send_sensor(void) {\n"
        "  struct pkt item = {\n"
        "    .fifo_res = 0, .val = 42\n"
        "  };\n"
        "  k_fifo_put(&q, &item); // Links Stack!\n"
        "} // <-- Function returns! Stack WIPED!\n"
        "  // Receiver reads GARBAGE!",
        font_size=7.4
    )
    add_chain_box(
        s, 5.48, 3.46, 3.77, 0.98,
        "The Golden k_fifo Lifetime Rule",
        "1. Reserve first 4 bytes (void *res) for linked-list ptr.\n2. Allocate ONLY from k_mem_slab, k_heap, or static pools — NEVER local Stack!",
        dark=False
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "RULE OF THUMB", "Value copy from Stack -> k_msgq.  Zero-Copy huge payload -> k_mem_slab + k_fifo!", font_size=8.6)
    set_notes(s, (
        "مقارنة مهمة جداً بتوقع مبرمجين كتير: k_msgq بتعمل memcpy فآمنة تماماً تبعتلها متغير محلي من على الـ Stack. "
        "أما k_fifo وk_lifo فدول Zero-Copy linked list، بياخدوا أول 4 بايت reserved في الـ struct بتاعك يشاوروا بيهم على العنصر اللي بعده! "
        "فلو عرفت الـ struct جوه دالة على الـ Stack وبعتّه لـ k_fifo_put، أول ما الدالة تتقفل الـ Stack هيتمسح وتضرب كارثة!"
    ))

    # =========================================================================
    # SLIDE 24: PILLAR 10B — ZBUS PUB/SUB & ZERO-COPY MULTI-CORE AMP
    # =========================================================================
    s = prs.slides[23]
    set_text(s.shapes[0], "Zbus (Pub/Sub Bus) & Zero-Copy Multi-Core AMP", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "Zbus Channel: sensor_chan", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "Decoupled many-to-many software bus:\nOwns typed message + mutex, optional priority boost & cross-core proxy agents — zero spaghetti queues!",
        size_pt=9.4, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[1], "Publisher -> Subscribers", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "Sensor calls zbus_chan_pub(&sensor_chan, &reading, K_MSEC(10)). BLE Observer, Flash Logger & UI Display listen without knowing each other!",
        size_pt=9.4, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "AMP: ipc_service + OpenAMP", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "Links Core A (Cortex-A Linux / App) and Core B (Cortex-M Zephyr Real-Time / BLE) via zero-copy Shared SRAM Ring Buffers (icmsg / RPMsg-lite).",
        size_pt=9.4, color_hex="595959", font_name=FONT_BODY
    )
    set_notes(s, (
        "بدل ما كل Thread ينادي Queue بتاع Thread تاني ويعملوا شبكة متشابكة، Zephyr جاب Zbus بمبدأ Publish/Subscribe: "
        "الحساس ينشر على Channel وأي عدد من الـ Observers (بلوتوث، شاشة، تسجيل) يسمعوه من غير ما يعرفوا بعض! "
        "وللشرائح متعددة الأنوية غير المتماثلة (مثلاً Linux على Cortex-A وZephyr على Cortex-M) فيه إطار ipc_service وOpenAMP/RPMsg."
    ))

    # =========================================================================
    # SLIDE 25: PILLAR 11 — MPU USER MODE & MEMORY DOMAINS
    # =========================================================================
    s = prs.slides[24]
    set_text(s.shapes[0], "MPU User Mode & Memory Domains: Isolation", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    tree_25 = {
        12: ("KERNEL MODE (Priv)", "383838", 8.8),
        13: ("Domain A (User)", "383838", 8.8),
        14: ("Domain B (User)", "383838", 8.8),
        18: ("Network Parser", "FFFFFF", 8.5),
        17: ("HW Stack Guard", "FFFFFF", 8.5),
        16: ("Sensor / Motor", "FFFFFF", 8.5),
        15: ("HW Stack Guard", "FFFFFF", 8.5),
        22: ("net_buf Partition ONLY", "383838", 7.8),
        21: ("Faults Thread A Only", "383838", 7.8),
        20: ("SVC Syscalls Only", "383838", 7.8),
        19: ("Sensor RAM Partition", "383838", 7.8),
        23: ("Walled Off from Kernel", "383838", 7.8),
    }
    for si, (txt, col, sz) in tree_25.items():
        set_text(s.shapes[si], txt, size_pt=sz, bold=True, color_hex=col, font_name=FONT_BODY, align="center")
    add_pill_banner(
        s, 0.78, 4.78, 8.44, 0.36,
        "CONFIG_USERSPACE=y",
        "k_mem_domain_add_thread(&dom, &tid);  k_object_access_grant(&msgq, &tid);  —  MPU traps rogue pointers instantly!",
        font_size=8.2
    )
    set_notes(s, (
        "في الأنظمة العادية كل الثريدات بصلاحيات كاملة، فلو pointer ضرب في مكان غلط بيبوظ بيانات الكيرنل. "
        "في Zephyr تقدر تفعل User Mode: الهاردوير MPU بيحرس كل Stack بـ Stack Guard، وبيحط كل ثريد في Memory Domain خاص بيها "
        "متشوفش غير الـ Partitions والتصاريح المصرح ليها بيها (k_object_access_grant)، وأي تعامل مع الكيرنل يعدي على SVC Syscall."
    ))

    # =========================================================================
    # SLIDE 26: PILLAR 12 — SECURITY (TF-M), MCUboot OTA & SAFETY CERTS
    # =========================================================================
    s = prs.slides[25]
    set_text(s.shapes[0], "TrustZone (TF-M), Signed OTA & Safety Certs", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "MCUboot & Atomic A/B OTA", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "ECDSA/RSA Verified Boot. Writes update to Slot 1 while Slot 0 stays runnable — auto-rollbacks if self-test fails! Built in 1 command via --sysbuild.",
        size_pt=9.2, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[1], "ARM TrustZone & TF-M", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "SPE vs NSPE isolates Secure keys & attestation from Non-Secure Zephyr app. PSA Secure Storage encrypts flash secrets; native mbedTLS & HW crypto.",
        size_pt=9.2, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "Functional Safety & SBOM", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "IEC 61508 SIL 3 core kernel subset (SEooC), ISO 26262 (ASIL D) fault-containment mapping, plus native west spdx supply-chain SBOM export.",
        size_pt=9.2, color_hex="595959", font_name=FONT_BODY
    )
    set_notes(s, (
        "في الأمان والتحديثات، Zephyr مدمج معاه ARM TrustZone عبر TF-M/PSA لعزل المفاتيح والتشفير، "
        "ومعاه MCUboot لتوقيع السوفتوير وعمل تحديثات OTA بنظام Slot A/B بحيث لو التحديث الجديد فشل في الاختبار يرجع أوتوماتيك للنسخة القديمة. "
        "وفوق كل ده فيه مسار رسمي لشهادات السلامة الصناعية والسيارات IEC 61508 SIL 3 وISO 26262."
    ))

    # =========================================================================
    # SLIDE 27: PILLAR 13 — AUTOMATIC POWER MANAGEMENT
    # =========================================================================
    s = prs.slides[26]
    set_text(s.shapes[0], "Automatic Power Management: PM Policy & Device PM", size_pt=20.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[4], "WFI", size_pt=13.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "STOP", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "DEEP", size_pt=12, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "System PM Policy Manager", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[8], "Matches Tickless Idle time (4,500 ms) + Devicetree min-residency & exit-latency + DMA locks -> picks pm_state: RUNTIME_IDLE, STANDBY, or SOFT_OFF (32.768 kHz LSE).", size_pt=8.8, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[9], "Device Runtime PM (Auto Clock Gating)", size_pt=11, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[10], "Drivers track ref-counts via pm_device_runtime_get/put(): when no thread uses I2C or UART, Zephyr suspends & clock-gates that peripheral automatically!", size_pt=8.7, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[11], "Nested Power Domains (Devicetree)", size_pt=11, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[12], "Shuts off entire SoC peripheral rails & oscillators cleanly when every child device enters suspend — zero manual WFI glue code.", size_pt=8.7, color_hex="595959", font_name=FONT_BODY)
    set_notes(s, (
        "في إدارة الطاقة، مش بس فيه Tickless Idle؛ Zephyr فيه PM Policy Manager بياخد وقت السكون الجاي ويقارنه بالـ min-residency-us "
        "والـ exit-latency-us المتوصفين في الـ Devicetree، وبناءً عليه يختار أنسب حالة من pm_state (WFI / Stop / Standby مع 32.768 kHz LSE "
        "وBackup RAM). وكمان Device Runtime PM بيطفي ساعات الـ peripherals أوتوماتيك لما محدش يستخدمها!"
    ))

    # =========================================================================
    # SLIDE 28: PILLAR 14 — CONNECTIVITY & STORAGE STACKS
    # =========================================================================
    s = prs.slides[27]
    set_text(s.shapes[0], "Built-In Connectivity & Storage Stacks", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "IPv4/IPv6, Matter & Buses", size_pt=12.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "TCP/UDP, BSD Sockets, TLS/DTLS, MQTT, CoAP, LwM2M, HTTP & Wi-Fi. Native Matter, OpenThread, USB Device/Host, CANbus & Modbus.",
        size_pt=9.0, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[1], "Bluetooth LE (Host + Ctrl)", size_pt=12.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "Full GAP, GATT, Mesh, LE Audio & Direction Finding. Internal radio (nRF52/ESP32) OR external UART chip via 1 DT line: zephyr,bt-hci = &uart2;",
        size_pt=9.0, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "Storage, Flash & DFU", size_pt=12.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "LittleFS (power-fail safe + wear leveling), FATFS, NVS & ZMS key-value. MCUmgr (SMP over BLE/Serial/UDP) runs remote DFU & diagnostics.",
        size_pt=9.0, color_hex="595959", font_name=FONT_BODY
    )
    add_pill_banner(
        s, 0.78, 4.90, 8.44, 0.36,
        "SAME FORMULA",
        "subsys/ + prj.conf (CONFIG_BT=y, CONFIG_FILE_SYSTEM_LITTLEFS=y) + app.overlay — swap nRF52 radio <-> STM32 UART HCI with identical bt_enable()!",
        font_size=7.9
    )
    set_notes(s, (
        "الفرق بين الـ Drivers (في drivers/) والـ Stacks (في subsys/ زي البلوتوث والشبكات وLittleFS) إن الـ Stack بيقعد فوق الـ Driver "
        "ومش بيكلم Registers بإيده. بتفعل CONFIG_BT=y مثلاً، ولو الشريحة فيها راديو داخلي بيبقى متعرف في ملف البوردة جاهز، "
        "ولو مفيهاش بتشاور في app.overlay على &uart2 — وكود bt_enable() في main.c بيفضل زي ما هو!"
    ))

    # =========================================================================
    # SLIDE 29: PILLAR 15 — NATIVE_SIM, TWISTER & OBSERVABILITY (Lab Photo!)
    # =========================================================================
    s = prs.slides[28]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_LAB)

    s.shapes[1].top = Inches(0.91)
    s.shapes[1].left = Inches(0.75)
    s.shapes[1].width = Inches(4.85)
    s.shapes[1].height = Inches(0.80)
    set_text(s.shapes[1], "Boot Firmware on Your PC:\nnative_sim & Observability", size_pt=20.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    s.shapes[0].top = Inches(1.77)
    s.shapes[0].left = Inches(0.75)
    s.shapes[0].width = Inches(4.85)
    s.shapes[0].height = Inches(2.05)
    set_text(
        s.shapes[0],
        [
            [("• native_sim & west twister: ", True, "383838"), ("Run RTOS threads, Zbus & TCP/BLE natively on PC with GDB/Valgrind + automated CI testing.", False, "595959")],
            [("• Zephyr Shell: ", True, "383838"), ("Live interactive CLI over UART/RTT/BLE ('kernel threads', 'device list').", False, "595959")],
            [("• Deferred Logging & CTF: ", True, "383838"), ("Zero-ISR-delay severity logs + native SEGGER SystemView / Tracealyzer export.", False, "595959")],
            [("• Thread Analyzer & Coredump: ", True, "383838"), ("Live stack high-water marks + full CPU/RAM crash dumps saved to flash.", False, "595959")],
        ],
        size_pt=9.4, font_name=FONT_BODY, align="left", space_after_pt=5, line_spacing=1.15
    )
    add_code_panel(
        s, 0.75, 3.88, 4.85, 1.15,
        "native_sim.sh — Debug before touching a soldering iron!",
        "$ west build -b native_sim app\n"
        "$ ./build/zephyr/zephyr.exe\n"
        "# Runs RTOS + MQTT/BLE natively on host!",
        font_size=7.8
    )
    set_notes(s, (
        "ميزة رهيبة في الشغل الحقيقي: قبل ما البوردة توصل أصلاً، تقدر تبني الكود بتارجت -b native_sim وتشغل تطبيق Zephyr كبرنامج عادي "
        "على اللاب توب بتاعك وتعمله اختبارات آلية في الـ CI بـ west twister! وكمان جوه الكيرنل نفسه فيه Zephyr Shell تفاعلي، "
        "وLogging متدرج، وSystemView Tracing، وThread Analyzer، وCoredump يسجل حالة السيستم لو وقع."
    ))

    # =========================================================================
    # SLIDE 30: PILLAR 16A — SIDE-BY-SIDE CODE SHOWDOWN
    # =========================================================================
    s = prs.slides[29]
    set_text(s.shapes[1], "Code Showdown: FreeRTOS HAL vs. Portable Zephyr", size_pt=21, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 1]:
        remove_shape(sh)

    add_code_panel(
        s, 0.75, 1.32, 4.15, 3.16,
        "freertos_stm32.c  (Locked to ST Silicon)",
        "#include \"stm32f4xx_hal.h\"\n"
        "#include \"FreeRTOS.h\"\n"
        "#include \"task.h\"\n\n"
        "void blink_task(void *arg) {\n"
        "  for (;;) {\n"
        "    HAL_GPIO_TogglePin(GPIOA, GPIO_PIN_5);\n"
        "    vTaskDelay(pdMS_TO_TICKS(500));\n"
        "  }\n"
        "}\n"
        "int main(void) {\n"
        "  HAL_Init(); SystemClock_Config();\n"
        "  MX_GPIO_Init(); // Manual STM32 setup!\n"
        "  xTaskCreate(blink_task,\"b\",128,0,2,0);\n"
        "  vTaskStartScheduler();\n"
        "}",
        font_size=7.4
    )
    add_code_panel(
        s, 5.10, 1.32, 4.15, 3.16,
        "zephyr_portable.c  (STM32 / nRF52 / ESP32)",
        "#include <zephyr/kernel.h>\n"
        "#include <zephyr/drivers/gpio.h>\n\n"
        "static const struct gpio_dt_spec led =\n"
        "  GPIO_DT_SPEC_GET(DT_ALIAS(led0), gpios);\n\n"
        "void blink_thread(void *a,void *b,void *c){\n"
        "  gpio_pin_configure_dt(&led,\n"
        "      GPIO_OUTPUT_ACTIVE);\n"
        "  while (1) {\n"
        "    gpio_pin_toggle_dt(&led);\n"
        "    k_msleep(500); // Tickless sleep!\n"
        "  }\n"
        "}\n"
        "K_THREAD_DEFINE(blk_id, 512, blink_thread,\n"
        "                NULL, NULL, NULL, 7, 0, 0);",
        font_size=7.4
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "SPOT THE DIFF", "Zero HAL_Init(), zero GPIOA->ODR, zero dynamic heap — and K_THREAD_DEFINE auto-starts at boot!", font_size=8.6)
    set_notes(s, (
        "شوفوا الفرق قدام عينكم في الكود: في FreeRTOS إنت مربوط بـ stm32f4xx_hal.h وHAL_Init() وMX_GPIO_Init() وGPIOA_PIN_5 "
        "وxTaskCreate ديناميكي. في Zephyr نفس التاسك مكتوب بـ GPIO_DT_SPEC_GET(DT_ALIAS(led0)) في الفلاش (.rodata) وK_THREAD_DEFINE "
        "اللي بيحجز الـ Stack statically ويشغل الثريد أوتوماتيك — ونفس الملف ده حرفياً يتكومبايل على STM32 أو nRF52 أو ESP32!"
    ))

    # =========================================================================
    # SLIDE 31: PILLAR 16B — 4 HONEST ENGINEERING TRADE-OFFS
    # =========================================================================
    s = prs.slides[30]
    set_text(s.shapes[0], "Where Zephyr Hurts: 4 Real Engineering Trade-Offs", size_pt=21, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "Developer Experience", size_pt=9.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[3], "Build & Low-Level Rigidity", size_pt=9.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[4], "High Onboarding Cost", size_pt=9.0, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "Daily Workflow Friction", size_pt=9.0, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")

    quads = [
        (6, 2.25, 1.82, 2.45, 1.15, "01. Steep Learning Curve", "FreeRTOS takes hours; mastering Devicetree, Kconfig, CMake & west takes weeks of onboarding."),
        (9, 5.30, 1.82, 2.45, 1.15, "02. Macro Debug Friction", "Deep preprocessor macros (DT_INST_*, devicetree_generated.h) emit cryptic errors on overlay typos."),
        (7, 2.25, 3.30, 2.45, 1.15, "03. Heavy Workspace", "Full west pulls GBs of HALs & Python tooling; clean builds are slower than 5 FreeRTOS .c files."),
        (8, 5.30, 3.30, 2.45, 1.15, "04. Strict Opinionation", "Bypassing SYS_INIT for quick bare-metal hacks fights the framework; prune prj.conf debug flags!"),
    ]
    for qi, qx, qy, qw, qh, qh_txt, qb_txt in quads:
        sh = s.shapes[qi]
        sh.left, sh.top, sh.width, sh.height = Inches(qx), Inches(qy), Inches(qw), Inches(qh)
        set_text(
            sh,
            [[(qh_txt, True, "383838")], [(qb_txt, False, "595959")]],
            size_pt=9.0, font_name=FONT_BODY, align="center", space_after_pt=3
        )
    set_notes(s, (
        "وعشان نكون صادقين هندسياً ومش بنعمل دعاية، لازم نقول عيوب وتحديات Zephyr بصراحة: "
        "1) منحنى تعلم تقيل (Devicetree + Kconfig + CMake + west محتاجين أسابيع عكس FreeRTOS)، "
        "2) أخطاء الكومبايل جوه الـ macros الكتير ممكن تتعب في الديباج، 3) مساحة الـ workspace وزمن البناء أكبر، "
        "4) صعب تتحايل على فريمورك الدرايفرات بتاعه — ودايماً نظّف prj.conf من خيارات الـ debug قبل الإنتاج!"
    ))

    # =========================================================================
    # SLIDE 32: PILLAR 16C — FULL ARCHITECTURE COMPARISON MATRIX
    # =========================================================================
    s = prs.slides[31]
    set_text(s.shapes[0], "FreeRTOS vs. Zephyr: Every Layer Compared", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    remove_shape(s.shapes[1])
    add_clean_table(
        s, 0.72, 1.24, 8.56, 3.32,
        headers=["Dimension", "FreeRTOS", "Zephyr RTOS"],
        rows=[
            ["Core Scope", "Minimalist Micro-Kernel (MIT)", "Full Embedded OS (Apache 2.0, Linux Found.)"],
            ["Kernel Footprint", "~9 KB Flash / ~2 KB RAM", "~8–18 KB Flash / ~4 KB RAM (3–4 KB minimal)"],
            ["Full IoT Stack", "~195 KB Flash / ~44 KB RAM (DIY)", "~220 KB Flash / ~48 KB RAM (BLE+MQTT+TLS+OTA)"],
            ["HW Abstraction", "Vendor HAL lock-in (CubeMX / ESP-IDF)", "Devicetree (.dts/.overlay) + Unified APIs (.rodata)"],
            ["Configuration", "Manual #defines in FreeRTOSConfig.h", "Dependency-aware Kconfig (prj.conf + menuconfig)"],
            ["Sched & ISR", "Preempt + RR; *FromISR() APIs", "Coop (<0)+Preempt+EDF+SMP; IPSR + Zero-Latency"],
            ["Memory & MPU", "heap_1..5 (frag risk); FreeRTOS-MPU", "Static-first, k_mem_slab, sys_mem_blocks, Domains"],
            ["Stacks, PM & CI", "DIY glue (lwIP, NimBLE, manual sleep)", "Native BLE/Matter/LittleFS, PM Policy, native_sim"],
        ],
        col_widths=[1.42, 3.22, 3.92], header_size=9.0, body_size=8.0
    )
    add_pill_banner(s, 0.72, 4.70, 8.56, 0.38, "BOTTOM LINE", "Same silicon physics (SysTick, NVIC, PendSV) — vastly different system integration & portability!", font_size=8.5)
    set_notes(s, (
        "الجدول ده بيلخص المحاضرة كلها في صفحة واحدة للمقارنة السريعة بين FreeRTOS وZephyr عبر الـ 8 أبعاد الأساسية: "
        "الفلسفة، حجم الكيرنل، حجم مشروع IoT كامل، فصل الهاردوير، الكونفيجريشن، الجدولة والـ ISR، الميموري والـ MPU، والـ Stacks والاختبار."
    ))

    # =========================================================================
    # SLIDE 33: TECHNICAL REFERENCES & OFFICIAL DOCUMENTATION (Updated Slide Citations!)
    # =========================================================================
    s = prs.slides[32]
    set_text(s.shapes[0], "References & Official Technical Documentation", size_pt=21.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    remove_shape(s.shapes[1])
    add_clean_table(
        s, 0.72, 1.24, 8.56, 3.36,
        headers=["Ref", "Primary Source & Covered Slides", "Official Documentation URL"],
        rows=[
            ["[1]", "FreeRTOS Kernel & Memory (heap_1..5, BASEPRI, *FromISR) — Slides 3–12", "freertos.org/Documentation/00-Overview"],
            ["[2]", "R. Barry, Mastering the FreeRTOS Real Time Kernel — Slides 3–12", "freertos.org (Official Reference Manual)"],
            ["[3]", "ARM Cortex-M4 Generic User Guide (SysTick, PendSV, NVIC, IPSR, MPU)", "developer.arm.com/documentation/dui0553"],
            ["[4]", "Zephyr Build System, Devicetree (.dts/.overlay), Kconfig & west — Slides 15–19", "docs.zephyrproject.org/latest/build/dts/"],
            ["[5]", "Zephyr Kernel Services (EDF, SMP, _isr_wrapper, Slabs, IPC, Domains) — Slides 20–25", "docs.zephyrproject.org/latest/kernel/"],
            ["[6]", "Zephyr Subsystems, Drivers, Zbus, PM Policy & native_sim — Slides 17, 24, 27–29", "docs.zephyrproject.org/latest/services/"],
            ["[7]", "Zephyr Hardware Support, Security & Safety (SIL 3 / ISO 26262) — Slides 14, 26", "docs.zephyrproject.org/latest/security/"],
            ["[8]", "MCUboot Signed OTA Rollback & Trusted Firmware-M (TF-M / PSA) — Slide 26", "docs.mcuboot.com  |  trustedfirmware.org"],
        ],
        col_widths=[0.48, 4.98, 3.10], header_size=8.8, body_size=7.8
    )
    add_pill_banner(
        s, 0.72, 4.72, 8.56, 0.36,
        "VERIFIED",
        "All architectural figures, APIs, footprint ranges & 1,000+ board metrics are sourced directly from official upstream docs.",
        font_size=8.2
    )
    set_notes(s, (
        "ودي قائمة المراجع الرسمية الأولية اللي بنينا عليها كل سلايد في العرض: التوثيق الرسمي لـ FreeRTOS وكتاب Richard Barry، "
        "ودليل معمارية ARM Cortex-M4 للـ SysTick والـ PendSV والـ BASEPRI والـ IPSR والـ MPU، "
        "والتوثيق الرسمي المحدث لمشروع Zephyr RTOS (Devicetree, Kernel, Services, Hardware, Security)، بالإضافة لتوثيق MCUboot وTrusted Firmware-M."
    ))

    # =========================================================================
    # SLIDE 34: ENGINEERING VERDICT — THANK YOU & Q&A
    # =========================================================================
    s = prs.slides[33]
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)
    for lsh in list(s.slide_layout.shapes):
        if lsh.has_text_frame and "CREDITS" in lsh.text_frame.text:
            remove_shape(lsh)

    s.shapes[0].left = Inches(0.75)
    s.shapes[0].top = Inches(0.52)
    s.shapes[0].width = Inches(8.50)
    s.shapes[0].height = Inches(0.62)
    set_text(s.shapes[0], "Engineering Verdict  —  Thank You & Q&A", size_pt=23.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    add_verdict_card(
        s, 0.75, 1.38, 4.10, 2.94,
        "Choose FreeRTOS When...",
        [
            ("Ultra-tight silicon", "Flash < 64 KB or tiny RAM where every byte counts."),
            ("Single fixed MCU", "Deeply invested in STM32Cube or ESP-IDF with no swap planned."),
            ("Standalone control", "Motor driver, PSU or appliance loop without complex IoT stacks."),
            ("Speed to first blink", "Team wants hours of onboarding & full vendor IDE simplicity."),
        ],
        dark=False
    )
    add_verdict_card(
        s, 5.15, 1.38, 4.10, 2.94,
        "Choose Zephyr RTOS When...",
        [
            ("Connected IoT", "Need BLE, Wi-Fi, Matter, Thread, MQTT, TLS & OTA tested together."),
            ("Multi-vendor portability", "Product must survive chip shortages across ST, Nordic & ESP."),
            ("Security & Safety", "Need MPU User Domains, TF-M, MCUboot rollback or SIL 3 / ASIL D."),
            ("Battery & CI scale", "Automatic PM Policy + PC testing via native_sim & Twister."),
        ],
        dark=True
    )
    add_pill_banner(
        s, 0.75, 4.56, 8.50, 0.48,
        "QUESTIONS?",
        "Master FreeRTOS to understand the RTOS heart — adopt Zephyr when your product needs an OS, not just a kernel.",
        font_size=8.8
    )
    set_notes(s, (
        "الخلاصة الهندسية: اختيار الـ RTOS مبيبقاش بالموضة، بيبقى بقيود المشروع. اختار FreeRTOS لو الشريحة محدودة (<64KB) "
        "ومشروعك تحكم بسيط مرتبط بـ HAL واحدة ومحتاج سرعة بداية. واختار Zephyr لو بتبني منتج IoT متصل (BLE/Matter/MQTT) "
        "أو معتمد على بطارية أو محتاج عزل بالـ MPU وتحديثات OTA آمنة وقابلية نقل بين شرائح مختلفة. شكراً لحضراتكم وجاهزين لأسئلتكم!"
    ))

    prs.save(out_path)
    print(f"SAVED: {out_path} ({os.path.getsize(out_path)} bytes, {len(prs.slides)} slides)")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else OUT_PPTX
    build(target)
