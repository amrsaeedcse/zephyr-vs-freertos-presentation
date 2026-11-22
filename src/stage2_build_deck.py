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
    add_left_right_arrow, add_lead_subtitle, add_card, add_lane_block, add_dark_badge,
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
    s.shapes[0].top = Inches(1.48)
    s.shapes[0].width = Inches(5.43)
    s.shapes[0].height = Inches(1.50)
    s.shapes[0].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[0], ["Zephyr RTOS", "vs. FreeRTOS"], size_pt=43, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")

    # Remove the opaque white slab box behind subtitle so typography floats naturally on canvas
    clear_fill_and_line(s.shapes[1])
    s.shapes[1].left = Inches(3.79)
    s.shapes[1].top = Inches(3.06)
    s.shapes[1].width = Inches(5.43)
    s.shapes[1].height = Inches(0.44)
    s.shapes[1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[1], "From Real-Time Kernel Internals to a Full Embedded OS", size_pt=13.5, bold=False, color_hex="595959", font_name=FONT_BODY, align="center")

    tb_meta = s.shapes.add_textbox(Inches(3.55), Inches(3.64), Inches(5.90), Inches(0.32))
    tf_m = tb_meta.text_frame
    tf_m.word_wrap = True
    tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = 0
    p = tf_m.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "CSE 421/422 Embedded Systems  ·  Zagazig University  ·  Dr. Ahmed Osman Mahmoud"
    r.font.name = FONT_BODY
    r.font.size = Pt(9.0)
    r.font.bold = True
    r.font.color.rgb = COL_CHARCOAL

    add_dark_badge(s, 5.20, 4.02, 2.60, 0.30, "Team 8  ·  Topic 8  ·  24-11", font_size=8.6)

    set_notes(s, (
        "أهلاً بكم يا دكتور أحمد ويا شباب — إحنا Team 8 (Topic 8 · CSE 421/422 Embedded Systems). النهاردة مش هنعرض Zephyr RTOS ككلام نظري؛ "
        "هنبدأ الأول بسرد هندسي واضح وسريع يشرح إزاي الـ RTOS وFreeRTOS بيشتغلوا من جوه خطوة بخطوة "
        "(تقسيم المهام، حالات الـ Thread، الجدولة، الـ Blocking، المزامنة، المقاطعات، والذاكرة)، "
        "عشان لما ندخل في Zephyr تشوف كل ميزة كحل مباشر لمشكلة حقيقية في الإنتاج ومربوطة بمنهجنا."
    ))

    # =========================================================================
    # SLIDE 2: EXECUTIVE PRESENTATION ROADMAP (Fair 5 × 6-Slide Team 8 Split!)
    # =========================================================================
    s = prs.slides[1]
    set_text(s.shapes[0], "Presentation Roadmap", size_pt=25, bold=True, color_hex="383838", font_name=FONT_HEAD)
    # Remove uneven template diagonal stripe groups ([13], [14])
    for sh in [s.shapes[14], s.shapes[13]]:
        remove_shape(sh)

    # Exact fair 5 × 6-slide core split (S3–S32 = 30 slides) + Shared Intro (S1–S2) & Closer (S33–S34)
    roadmap_cfg = [
        (7,  0.89, 2.22, "RTOS Basics & Sync",           "Why RTOS, TCB/Stack, Preemption, Tick, Blocking & Sync",     "Omar Abdelsalam · S3–S8",  2.90),
        (8,  3.77, 2.22, "RTOS Traps & Enter Zephyr",    "Priority Inversion, BASEPRI, *FromISR, Heap & Zephyr OS",    "Amr Abdelazim · S9–S14",   2.90),
        (9,  6.64, 2.22, "Devicetree, Drivers & Sched",  "Step 1 DTS/prj.conf, Step 2 APIs, .rodata, SYS_INIT & SMP",  "Omar Mashhour · S15–S20",  2.90),
        (10, 0.89, 4.10, "Kernel, IPC, MPU & Security",  "Zero *FromISR, k_mem_slab, k_fifo/Zbus, MPU & TrustZone",    "Ali Mohamed · S21–S26",    4.78),
        (11, 3.77, 4.10, "Power, Sim, Code & Matrix",    "Auto PM, Stacks, native_sim, Code Showdown, Limits & Matrix", "Ahmed Hassan · S27–S32",  4.78),
        (12, 6.64, 4.10, "References, Verdict & Q&A",    "Official Documentation, Engineering Verdict & Full Team Q&A", "Team 8 · S33–S34",        4.78),
    ]
    for si, lx, ly, lhead, ldesc, spk_badge, by in roadmap_cfg:
        sh = s.shapes[si]
        sh.left = Inches(lx)
        sh.top = Inches(ly)
        sh.width = Inches(2.46)
        sh.height = Inches(0.66)
        set_text(
            sh,
            [[(lhead, True, "383838")], [(ldesc, False, "595959")]],
            size_pt=9.2, font_name=FONT_BODY, align="center", space_after_pt=2
        )
        add_dark_badge(s, lx + 0.14, by, 2.18, 0.25, spk_badge, font_size=7.6)
    # Add 4 uniform horizontal arrows cleanly connecting 01->02->03 and 04->05->06
    for ax in (3.44, 6.32):
        add_right_arrow(s, ax, 1.72, 0.24, 0.16)
        add_right_arrow(s, ax, 3.62, 0.24, 0.16)
    set_notes(s, (
        "دي خريطة العرض الكاملة متقسمة بالعدل الكامل على أعضاء Team 8 (5 متكلمين × 6 سلايدات محتوى بالظبط = 30 سلايد + افتتاح وختام مشترك): "
        "01 عمر عبدالسلام (S3–S8): أساسيات الـ RTOS وFreeRTOS والمزامنة؛ "
        "02 عمرو عبدالعظيم (S9–S14): تحديات الـ RTOS ودخول عالم Zephyr؛ "
        "03 عمر مشهور (S15–S20): الـ Devicetree والدرايفرات الموحدة والبيلد والجدولة؛ "
        "04 علي محمد (S21–S26): المقاطعات الموحدة والذاكرة والـ IPC والـ MPU والأمان؛ "
        "05 أحمد حسن (S27–S32): الباور والاتصالات والـ native_sim ومقارنة الكود والماتريكس؛ "
        "06 ختام مشترك لكل التيم (S33–S34): المراجع والقرار الهندسي والأسئلة!"
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
    add_dark_badge(s, 3.24, 5.12, 2.40, 0.25, "CSE · Lec 3 PWM · Lab 3 · Lec 4/6", font_size=7.2)
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
    add_lead_subtitle(s, 0.75, 0.94, 6.40, 0.32, "Every 1 System Tick (1 ms), the hardware SysTick timer updates OS time and enforces both scheduling rules:", size_pt=8.9)
    add_dark_badge(s, 7.20, 0.92, 2.05, 0.26, "CSE · Extends Lec 6 RETI", font_size=7.4)

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
    add_lead_subtitle(s, 0.75, 0.94, 6.00, 0.32, "Wait—what ARE those events and locks threads block on? Each tool solves a specific problem:", size_pt=9.8)
    add_dark_badge(s, 6.95, 0.92, 2.30, 0.26, "CSE · Lec 9 volatile ≠ atomic", font_size=7.5)

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
    add_dark_badge(s, 0.75, 5.14, 3.15, 0.25, "CSE · Lec 4 S22 · AVR cli() / Lec 9 DOR", font_size=7.4)
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
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "By default, FreeRTOS xTaskCreate() allocates Thread stacks & Queues from a dynamic Heap (pvPortMalloc)—here is the trap:", size_pt=9.1)

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
        "Fix A · Static Allocation & heap_1",
        "Skip the Heap: reserve fixed compile-time arrays (xTaskCreateStatic) or allocate-only heap_1 (never frees). 0% fragmentation forever, with extra setup code.",
        dark=False, title_pt=10.5, body_pt=8.5
    )
    add_card(
        s, 5.10, 3.28, 4.15, 1.26,
        "Fix B · Dynamic heap_4 / heap_5",
        "Common xTaskCreate() heap: heap_4 merges adjacent free blocks (heap_5 spans split RAM banks). Helps, but still vulnerable to scattered holes over months!",
        dark=False, title_pt=10.5, body_pt=8.5
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "BRIDGE TO ZEPHYR", "To avoid Heap fragmentation by default, Zephyr puts Threads in Static RAM + uses O(1) fixed-size k_mem_slab pools!", font_size=8.3)
    set_notes(s, (
        "الـ Threads والـ Stacks في FreeRTOS بتاخد رام منين؟ بالديفولت، xTaskCreate() بتحجز الـ Stack والـ TCB من الـ Heap بـ pvPortMalloc(). "
        "المشكلة الخطيرة إن بعد شهور من الحجز والتفضية، الرام ممكن يبقى فيها 40KB فاضيين بس متقطعين حتت 4KB متفرقة (زي الجبنة السويسري)! "
        "ولأن الميكروكنترولر معندوش MMU يضمهم، أول ما تطلب 8KB الـ malloc ترجع NULL والسيستم يقع! الحلول في FreeRTOS: يا تحجز Static وقت الكومبايل (xTaskCreateStatic / heap_1)، يا تستخدم heap_4 اللي بيدمج المتجاور بس لسه معرض للتقطيع على المدى الطويل."
    ))

    # =========================================================================
    # SLIDE 13: SECTION PIVOT 02 — ENTER ZEPHYR RTOS (Clean & Direct Transition!)
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

    tb = s.shapes.add_textbox(Inches(4.45), Inches(1.88), Inches(4.60), Inches(3.15))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    items_13 = [
        ("Where FreeRTOS Stops (Kernel Only)", True, 11.5, "383838", 4),
        ("FreeRTOS handles core kernel primitives well—threads, scheduling, queues, and mutexes—and intentionally stops there. Hardware drivers, build setup, connectivity, security & updates are outside the kernel.", False, 9.5, "595959", 10),
        ("What Zephyr RTOS Rebuilds (Full Embedded OS)", True, 11.5, "383838", 4),
        ("Zephyr combines a real-time kernel with a unified Hardware/Driver layer, an automated build system, and built-in subsystems—all designed to work together in one place.", False, 9.5, "595959", 10),
        ("Switch MCUs in FreeRTOS? Rewrite your hardware drivers.\nSwitch MCUs in Zephyr? Swap the target board & keep main.c!", True, 9.5, "383838", 0),
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
        "هنا نقطة التحوّل في العرض: إحنا دلوقتي فهمنا أساسيات الـ RTOS كويس. FreeRTOS بيشتغل على مستوى الكيرنل فقط "
        "(Threads، جدولة، Queues، Mutexes) وبيقف عند حدود الكيرنل؛ أي حاجة بره الكيرنل من درايفرات هاردوير أو بناء أو شبكات أو أمان إنت بتجمعها بنفسك. "
        "دلوقتي ندخل في Zephyr RTOS ونشوف إزاي اتبنى كـ نظام تشغيل مدمج متكامل!"
    ))

    # =========================================================================
    # SLIDE 14: ZEPHYR OVERVIEW — 4-LAYER EMBEDDED OS & OPEN GOVERNANCE
    # =========================================================================
    s = prs.slides[13]
    prepare_clean_header_slide(s, 0, "Zephyr RTOS: What It Is & How It's Layered", size_pt=21.2)
    add_lead_subtitle(s, 0.75, 0.94, 6.20, 0.32, "Not just a standalone kernel—Zephyr is an open-source embedded OS organized into 4 clear layers:", size_pt=9.5)
    add_dark_badge(s, 7.05, 0.92, 2.20, 0.26, "CSE · Lec 2 (4-Layer OS)", font_size=7.5)

    # Left Zone: 4 Clean Stacked OS Architecture Layers
    layers_14 = [
        ("4. Application Code (main.c)",        "Your portable C/C++ logic — never hardcodes chip registers.", False, False),
        ("3. Built-In Subsystems",              "Bluetooth, Wi-Fi/Networking, File Systems & OTA (running on top of drivers).", False, True),
        ("2. Real-Time Kernel",                 "Scheduling, Unified Interrupts, Memory Pools & MPU Isolation.", False, False),
        ("1. Unified Driver APIs + Devicetree", "Standard .h interfaces (gpio.h, uart.h, spi.h) + .dts/.dtsi HW files (1,000+ boards).", True, False),
    ]
    for li, (ltitle, ldesc, ldark, lsoft) in enumerate(layers_14):
        ly = 1.38 + li * 0.80
        add_card(s, 0.75, ly, 4.30, 0.72, ltitle, ldesc, dark=ldark, soft=lsoft, align="left", title_pt=10.0, body_pt=8.1)

    # Right Zone: Neutral Open Governance + Build UP vs Trim DOWN
    add_card(
        s, 5.25, 1.38, 4.00, 1.44,
        "Open & Vendor-Neutral (Linux Foundation)",
        [
            ("• No Chip Lock-In: ", "Backed by the Linux Foundation—not owned by ST, Nordic, or Espressif."),
            ("• Apache 2.0 License: ", "Free for commercial products without forcing you to open-source your app."),
        ],
        dark=False, title_pt=10.3, body_pt=8.5
    )
    add_card(
        s, 5.25, 2.94, 4.00, 1.56,
        "FreeRTOS (Build UP) vs. Zephyr (Trim DOWN)",
        [
            ("• FreeRTOS (Build UP): ", "Starts with a tiny kernel; your team adds HAL drivers, Bluetooth, TCP & OTA manually."),
            ("• Zephyr (Trim DOWN): ", "Starts with one tested OS; you enable only what you need—unused code is never compiled!"),
        ],
        dark=False, soft=True, title_pt=10.3, body_pt=8.5
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "NEXT STEP", "Step 1: What files do you give Zephyr to describe your board & select drivers? Devicetree + prj.conf!", font_size=8.3)
    set_notes(s, (
        "عشان نفهم Zephyr ببساطة: هو نظام تشغيل مفتوح المصدر تحت مظلة Linux Foundation (مؤسسة محايدة غير ربحية، يعني مش مملوك لشركة شرائح واحدة زي ST أو Espressif)، "
        "وبرخصة Apache 2.0 (تقدر تبيع منتجك التجاري بيه ببلاش). بيتكون من 4 طبقات واضحة: تطبيقك، ثم الـ Subsystems (بلوتوث، شبكات، OTA) اللي قاعدة فوق الدرايفرات، "
        "ثم الكيرنل، وتحتهم الدرايفرات الموحدة (gpio.h, uart.h, spi.h بنفس الباترن) مع ملفات وصف الهاردوير (.dts/.dtsi) لأكتر من 1,000 بوردة."
    ))

    # =========================================================================
    # SLIDE 15: STEP 1 — DEVICETREE + prj.conf -> GENERATED .h & DRIVER FILTER
    # =========================================================================
    s = prs.slides[14]
    prepare_clean_header_slide(s, 0, "Step 1 · Devicetree & prj.conf: Build Inputs & Output", size_pt=20.4)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Before compiling, Zephyr reads your board's hardware files + app config, then builds ONLY what your board needs:", size_pt=9.4)

    # Top Row: Exact Build Inputs (.dtsi + .dts + vendor .c + app.overlay + prj.conf)
    add_card(
        s, 0.75, 1.36, 2.58, 1.40,
        "1. Chip, Board & Drivers",
        [
            ("Included with Zephyr:", ""),
            ("• .dtsi (Chip): ", "Physical ports & IRQs."),
            ("• .dts (Board): ", "Built-in wiring (e.g. LED)."),
            ("• Vendor .c: ", "Drivers (e.g. gpio_stm32.c)."),
        ],
        dark=False, title_pt=9.6, body_pt=7.5
    )
    add_lead_subtitle(s, 3.37, 1.90, 0.24, 0.30, "+", size_pt=16.0, bold=True, color=COL_CHARCOAL, align="center")
    add_card(
        s, 3.65, 1.36, 2.74, 1.40,
        "2. Your Overlay (app.overlay)",
        [
            ("Same format as .dts — YOU write it:", ""),
            "Sets your custom pins/sensors & 'compatible' tag (picks which vendor .c driver to use, e.g. \"st,stm32-gpio\").",
        ],
        dark=True, title_pt=9.8, body_pt=7.9
    )
    add_lead_subtitle(s, 6.43, 1.90, 0.24, 0.30, "+", size_pt=16.0, bold=True, color=COL_CHARCOAL, align="center")
    add_card(
        s, 6.71, 1.36, 2.54, 1.40,
        "3. App Config (prj.conf)",
        [
            ("Features YOU want enabled:", ""),
            "Turns subsystems & drivers ON/OFF (e.g. CONFIG_GPIO=y, CONFIG_BT=y).",
        ],
        dark=False, soft=True, title_pt=9.8, body_pt=8.1
    )

    # Middle Merge Step
    add_down_arrow(s, 4.92, 2.82, 0.16, 0.24)
    add_lead_subtitle(s, 5.16, 2.84, 4.05, 0.26, "west build -b <board>  (Pre-build: Python + CMake runs 2 actions)", size_pt=8.2, bold=True, color=COL_CHARCOAL)

    # Bottom Result Row: Exact 2 Outputs of west build!
    add_card(
        s, 0.75, 3.16, 4.10, 1.34,
        "Output 1: devicetree_generated.h",
        [
            ("Pre-build #define Macros: ", "Generated by Python before C compilation."),
            ("Points to Vendor .c Device: ", "Tells main.c which struct device + config to use (e.g. pin 5, IRQ, address)."),
        ],
        dark=False, soft=True, title_pt=10.2, body_pt=8.2
    )
    add_right_arrow(s, 4.91, 3.76, 0.22, 0.14)
    add_card(
        s, 5.19, 3.16, 4.06, 1.34,
        "Output 2: CMake Filters Vendor .c",
        [
            ("Drops Unused .c Drivers: ", "Uses prj.conf + 'compatible' to compile ONLY your chip's .c (e.g. gpio_stm32.c)."),
            ("Zero Dead Code: ", "ESP32 / Nordic .c drivers never even enter the compiler!"),
        ],
        dark=False, title_pt=10.2, body_pt=8.2
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "NEXT STEP", "Step 2: Now that devicetree_generated.h & gpio_stm32.c are ready—how does main.c actually talk to the chip?", font_size=8.3)
    set_notes(s, (
        "الخطوة الأولى: إيه المدخلات اللي بتديها لـ Zephyr عشان يبني مشروعك؟ "
        "1) ملفات .dtsi/.dts (للشريحة والبوردة) + ملفات الـ Vendor .c drivers الجاهزة مع Zephyr (زي gpio_stm32.c وgpio_esp32.c). "
        "2) ملف app.overlay بتكتبه إنت بنفس أسلوب الـ .dts (تحط فيه بنّاتك/حساساتك والـ compatible اللي بتشاور على أنهي ملف vendor .c درايفر هتتعامل معاه). "
        "3) ملف prj.conf بتقول فيه أنهي ميزة عايز تفعّلها (زي CONFIG_GPIO=y). "
        "أول ما تشغّل west build، في مرحلة الـ Pre-build قبل الكومبايل، بايثون وCMake بيعملوا حاجتين: يطلعوا devicetree_generated.h (فيه #defines تشاور على الـ struct device جوه ملف الـ .c + بياناته زي البنّاية كمثال)، ويفلتروا ملفات الـ .c بحيث درايفر شريحتك بس اللي يدخل الكومبايلر!"
    ))

    # =========================================================================
    # SLIDE 16: STEP 2 — UNIFIED DRIVERS (Directly Follows Step 1!)
    # =========================================================================
    s = prs.slides[15]
    prepare_clean_header_slide(s, 1, "Step 2 · Unified Drivers: How main.c Reaches Any Chip", size_pt=20.2)
    add_lead_subtitle(s, 0.75, 0.92, 8.50, 0.30, "Continuing right from Step 1: how main.c uses devicetree_generated.h + ONE unified API to drive any chip (LED example):", size_pt=9.0)

    chain_16 = [
        ("1. Build Spec from .h", "GPIO_DT_SPEC_GET() reads\ndevicetree_generated.h -> spec", False),
        ("2. Call Unified API", "main.c passes &spec to gpio.h\ne.g. gpio_pin_toggle_dt(&led)", False),
        ("3. Unified API -> Table", "Unified API (static inline) reads\nspec.port (device) -> Flash *api", True),
        ("4. Silicon Registers", "Vendor .c runs stm32_toggle()\n-> writes HW (GPIOA->BSRR)", False),
    ]
    for ci, (ch_h, ch_b, dk) in enumerate(chain_16):
        cx = 0.75 + ci * 2.18
        add_chain_box(s, cx, 1.30, 1.94, 1.12, ch_h, ch_b, dark=dk)
        if ci < 3:
            add_right_arrow(s, cx + 1.97, 1.79, 0.18, 0.14)

    add_card(
        s, 0.75, 2.54, 4.08, 1.00,
        "Same Pattern Across Every Peripheral",
        [
            ("• Unified Headers: ", "gpio.h, uart.h, i2c.h, spi.h, sensor.h — even Bluetooth radios follow this exact flow!"),
            ("• Zero Overhead: ", "CMake compiles only your chip's .c; static inline jumps straight to Flash *api."),
        ],
        dark=False, soft=True, title_pt=9.6, body_pt=7.8
    )

    add_card(
        s, 0.75, 3.62, 4.08, 0.92,
        "Note · Step 1: Spec vs. Direct Device",
        [
            ("• Spec (Extra Config): ", "In Step 1, bundle *device + config (e.g. GPIO pin/flags, I2C addr)."),
            ("• Direct Device: ", "No extra config? Pull device directly (e.g. uart_poll_out(dev, 'A'))."),
        ],
        dark=False, soft=False, title_pt=9.3, body_pt=7.5
    )

    add_code_panel(
        s, 4.98, 2.54, 4.27, 2.00,
        "Example: main.c & gpio.h (LED Toggle)",
        "// 1) Read devicetree_generated.h -> spec (*device + pin)\n"
        "static const struct gpio_dt_spec led =\n"
        "    GPIO_DT_SPEC_GET(DT_ALIAS(led0), gpios); // nickname\n\n"
        "// 2) main.c passes &spec to Unified API (same on any MCU):\n"
        "gpio_pin_toggle_dt(&led);\n\n"
        "// 3) Inside gpio.h (static inline -> device -> Flash *api):\n"
        "led.port->api->port_toggle_bits(led.port, BIT(led.pin));",
        font_size=6.8, title_size=8.1
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "NEXT QUESTION", "Wait—does all this Devicetree & Flash-table infrastructure waste precious microcontroller RAM? Let's check!", font_size=8.3)
    add_dark_badge(s, 0.75, 5.14, 2.75, 0.25, "CSE · Solves Lec 2 S45 HAL Overhead", font_size=7.4)
    set_notes(s, (
        "الخطوة التانية مباشرة بعد Step 1 — إزاي main.c بتوصل فعلياً للـ Registers بتاعة الشريحة (هناخد مثال الـ LED مع GPIO): "
        "1) ماكرو GPIO_DT_SPEC_GET بيقرأ من devicetree_generated.h اللي طلع في Step 1 ويطلعلك struct اسمه spec فيه بوينتر للـ struct device بتاع الدرايفر + بيانات التوصيل زي البنّاية والـ flags (وDT_ALIAS(led0) مجرد nickname للـ LED على البوردة). "
        "2) في main.c إنت بتباصي &spec لدالة الـ Unified API الموحدة زي gpio_pin_toggle_dt(&led). "
        "3) جوه gpio.h، الدالة الموحدة (static inline) بتاخد spec.port (اللي هو struct device في gpio_stm32.c) وتروح منه لجدول الـ API في الفلاش وتنادي دالة الشريحة وتكتب في الهاردوير! "
        "وملحوظة مهمة: مش كل الدرايفرات بتستعمل spec — الـ spec بيتعمل لما نحتاج config إضافي زي رقم البن أو عنوان I2C، إنما حاجات تانية زي UART بتباصي struct device * مباشرة زي uart_poll_out(uart_dev, 'A')."
    ))

    # =========================================================================
    # SLIDE 17: DOES THIS WASTE RAM? — 0 BYTES RAM (.rodata) & FOOTPRINT CHART
    # =========================================================================
    s = prs.slides[16]
    set_text(s.shapes[0], "Does Devicetree Waste RAM? Zero-Cost Flash & Size", size_pt=20.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    stats_17 = [
        (4,  13, "0 Bytes RAM",     "Devicetree is never parsed at runtime"),
        (3,  12, "Build-Time Only", "Python compiles .dts into C #defines"),
        (2,  11, "Flash (.rodata)", "const device & API tables live in Flash"),
        (1,  10, "~195 vs ~220 KB", "In a real IoT app, Flash delta vanishes"),
    ]
    for hi, bi, htxt, btxt in stats_17:
        set_text(s.shapes[hi], htxt, size_pt=13.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")
        set_text(s.shapes[bi], btxt, size_pt=8.8, bold=False, color_hex="595959", font_name=FONT_BODY, align="left")
    set_text(
        s.shapes[9],
        "*Full IoT App = Bluetooth + MQTT + TLS + OTA. Minimal Zephyr kernel starts ~8–18 KB Flash / ~4 KB RAM.",
        size_pt=8.0, color_hex="595959", font_name=FONT_BODY, align="center"
    )
    old_chart_pic = find_pic(s)
    if old_chart_pic:
        remove_shape(old_chart_pic)
    add_monochrome_footprint_chart(s, 4.02, 1.25, 5.25, 3.30)
    set_notes(s, (
        "السؤال اللي بييجي فوراً بعد ما فهمنا Step 1 وStep 2: هل كل الـ Devicetree وجداول الـ API دي بتاكل RAM من الميكروكنترولر؟ "
        "الإجابة: لأ، صفر بايت RAM! الـ Devicetree بيتحوّل وقت الـ Build لـ #defines، وجداول struct device وAPI بتتخزن const في الفلاش (.rodata). "
        "وفي مشروع IoT كامل (BLE+MQTT+TLS+OTA) حجم Zephyr (~220 KB) بيقرب جداً من FreeRTOS (~195 KB)!"
    ))

    # =========================================================================
    # SLIDE 18: PILLAR 05B — ORDERED SYS_INIT BOOT LEVELS (5 Clean Stages!)
    # =========================================================================
    s = prs.slides[17]
    prepare_clean_header_slide(s, 1, "Before main() Starts: Automatic Boot Order (SYS_INIT)", size_pt=20.4)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "No more ordering 20 manual HAL_Init() / MX_*_Init() calls in main()—Zephyr boots drivers automatically in 5 stages:", size_pt=9.2)

    boot_stages = [
        ("1. EARLY",        "Basic CPU & reset\nsetup before drivers.",          "No Kernel",   False),
        ("2. PRE_KERNEL_1", "Clocks, IRQ ctrl,\nbasic UART & GPIO.",             "IRQ Stack",   False),
        ("3. PRE_KERNEL_2", "Drivers depending\non PRE_KERNEL_1.",               "IRQ Stack",   False),
        ("4. POST_KERNEL",  "Kernel alive! I2C/SPI\n& Mutexes OK!",              "Main Thread", True),
        ("5. APPLICATION",  "App subsystems right\nbefore main() runs.",         "Main Thread", False),
    ]
    for bi, (bt, bd, bctx, dk) in enumerate(boot_stages):
        bx = 0.75 + bi * 1.74
        add_card(s, bx, 1.40, 1.54, 1.70, bt, [bd, "", f"[{bctx}]"], dark=dk, align="center", title_pt=9.2, body_pt=8.0)
        if bi < 4:
            add_right_arrow(s, bx + 1.56, 2.18, 0.16, 0.14)

    add_card(
        s, 0.75, 3.28, 4.15, 1.24,
        "The Pain in FreeRTOS / Bare HAL",
        "In main(), you manually call HAL_Init(); SystemClock_Config(); MX_GPIO_Init(); MX_SPI1_Init(); — swap two lines by mistake and peripherals fail silently!",
        dark=False, title_pt=10.4, body_pt=8.5
    )
    add_card(
        s, 5.10, 3.28, 4.15, 1.24,
        "Zephyr Automatic Boot Table (SYS_INIT)",
        "Each driver registers its init function & boot stage. Before entering main(), Zephyr runs the sorted Flash table so every peripheral is ready on line 1 of main()!",
        dark=False, soft=True, title_pt=10.4, body_pt=8.5
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "GOLDEN RULE", "Never sleep or lock a Mutex in PRE_KERNEL_1/2 — the scheduler isn't running yet! Use POST_KERNEL for threaded drivers.", font_size=8.3)
    set_notes(s, (
        "الميزة هنا واضحة جداً: في مشاريع STM32 وFreeRTOS العادية بتكتب أول 20 سطر في main() بإيدك: HAL_Init ثم SystemClock ثم MX_GPIO ثم MX_SPI — ولو لخبطت الترتيب الهاردوير بيقف في صمت! "
        "في Zephyr كل درايفر بيسجّل دالة تجهيزه في مرحلة من 5 مراحل (من الـ Clocks في PRE_KERNEL_1 لحد I2C/SPI بعد ما الكيرنل يصحى في POST_KERNEL). "
        "الـ Linker بيرتبهم أوتوماتيك في الفلاش وكود الـ Boot بيشغّلهم قبل main() — فتدخل main() تلاقي كل البريفرالز جاهزة على أول سطر!"
    ))

    # =========================================================================
    # SLIDE 19: PILLAR 06 — KCONFIG (prj.conf), CMAKE & WEST
    # =========================================================================
    s = prs.slides[18]
    set_text(s.shapes[0], "Kconfig, CMake & west: Compile Only What You Use", size_pt=20.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)

    add_clean_table(
        s, 0.75, 1.32, 4.45, 3.12,
        headers=["west Command", "What It Does for You"],
        rows=[
            ["west init / update", "Downloads Zephyr, chip HALs & bootloader"],
            ["west build -b <board>", "Runs Devicetree + Kconfig + CMake + compile"],
            ["west build -t menuconfig", "Visual menu to turn OS features ON / OFF"],
            ["west flash / debug", "One command for ST-Link, J-Link & OpenOCD"],
            ["west build --sysbuild", "Builds App + Bootloader together in 1 step!"],
        ],
        col_widths=[1.85, 2.60], header_size=9.0, body_size=8.2
    )

    tb19 = s.shapes.add_textbox(Inches(5.38), Inches(1.30), Inches(3.87), Inches(1.38))
    tf = tb19.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for idx_l, (pre, rest) in enumerate([
        ("Why prj.conf Alongside Devicetree?", ""),
        ("• Devicetree = Hardware Present: ", "Board has 3 UARTs, 2 SPI, BLE."),
        ("• prj.conf = Software Wanted: ", "CONFIG_GPIO=y compiles GPIO only."),
        ("• Auto-Dependencies: ", "CONFIG_BT=y enables crypto & buffers automatically."),
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
        "CMakeLists.txt — Filters Out Other Chips!",
        "# Devicetree 'compatible' + prj.conf -> STM32=y\n"
        "zephyr_library_sources_ifdef(\n"
        "    CONFIG_GPIO_STM32 gpio_stm32.c)\n"
        "zephyr_library_sources_ifdef(\n"
        "    CONFIG_GPIO_ESP32 gpio_esp32.c)\n"
        "# gpio_esp32.c NEVER enters an STM32 build!",
        font_size=7.4
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "CLEAR RULE", "Devicetree asks: 'What hardware is on the board?'  |  prj.conf (Kconfig) asks: 'Which software enters Flash?'", font_size=8.4)
    set_notes(s, (
        "عشان متتلخبطش بين Devicetree وprj.conf: الـ Devicetree بيقول إيه الهاردوير اللي موجود على البوردة، "
        "إنما ملف prj.conf (Kconfig) بيقول إيه السوفتوير اللي عايز تفعّله في مشروعك النهاردة! "
        "وCMake بياخد نتيجة الاتنين ويفلتر ملفات الـ .c بحيث درايفر ESP32 ما يدخلش الكومبايلر أصلاً لو إنت بتبني لـ STM32، وأداة west بتجمع البناء والحرق والـ Debug في أمر واحد."
    ))

    # =========================================================================
    # SLIDE 20: PILLAR 07 — 3 SCHEDULING POLICIES + 3 KERNEL EXECUTION UPGRADES
    # =========================================================================
    s = prs.slides[19]
    set_text(s.shapes[0], "Zephyr Kernel: 3 Scheduling Policies + 3 Execution Upgrades", size_pt=19.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    grid_20 = [
        (5,  1, "Policy · Cooperative (< 0)", "Negative priorities (-1, -2): never preempted by other threads (only ISRs). Runs until blocking or k_yield() (SW IRQ like portYIELD)."),
        (6,  2, "Policy · Preemptive (>= 0)", "Preempts lower threads immediately. Note: 0 = highest priority (opposite of FreeRTOS!). Supports Round-Robin time-slicing."),
        (7,  8, "Policy · Earliest Deadline", "EDF scheduler: instead of fixed priority, runs whichever thread's time deadline expires soonest (ideal for hard real-time)."),
        (10, 3, "Upgrade · Workqueues",       "Replaces manual ISR + Binary Semaphore + extra Thread! Pass a callback in struct k_work; a shared worker thread runs it."),
        (11, 4, "Upgrade · Tickless Idle",    "Fixes Slide 6's 1 ms battery drain! Stops periodic tick IRQs when idle & sets ONE wake-up timer alarm (5 s sleep = 1 wakeup)."),
        (12, 9, "Upgrade · Multi-Core (SMP)", "Runs threads simultaneously in parallel on 2+ CPU cores (e.g. ESP32, RP2040) using real Spinlocks (k_spin_lock) & core pinning."),
    ]
    for hi, bi, htxt, btxt in grid_20:
        set_text(s.shapes[hi], htxt, size_pt=11.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
        set_text(s.shapes[bi], btxt, size_pt=8.7, color_hex="595959", font_name=FONT_BODY)
    add_dark_badge(s, 3.55, 4.92, 2.90, 0.26, "CSE · Tickless = Lec 4 Timer CTC", font_size=7.6)
    set_notes(s, (
        "السلايد دي مقسّمة بوضوح لصفّين: الصف اللي فوق فيه 3 سياسات جدولة (Scheduling Policies): "
        "1) Cooperative (أولويات سالبة مبتتقطعش إلا بالـ block أو k_yield زي portYIELD)، "
        "2) Preemptive (أولويات موجبة و0 أعلى أولوية عكس FreeRTOS)، و3) Earliest Deadline (EDF بيشغّل اللي ميعاد تسليمه قرب يخلص أسرع). "
        "الصف اللي تحت فيه 3 تطويرات تنفيذ في الكيرنل: Workqueues (بديل حوار الـ ISR + Binary Semaphore، بتحط دالتك في struct k_work)، "
        "وTickless Idle (حل بطارية Slide 6)، وMulti-Core SMP (تشغيل فعلي بالتوازي على أكتر من Core زي ESP32 وRP2040 بـ k_spin_lock)."
    ))

    # =========================================================================
    # SLIDE 21: PILLAR 08 — WHY ZERO *FromISR() IN ZEPHYR (_isr_wrapper & IPSR)
    # =========================================================================
    s = prs.slides[20]
    set_text(s.shapes[0], "Why Zephyr Needs ZERO *FromISR() Functions", size_pt=21.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[4], "99%", size_pt=14, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "FAST", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "0-LAT", size_pt=12, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "Standard (IRQ_CONNECT — 99%)", size_pt=11.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[8], "Solves Slide 11! Zephyr wraps ISR entry/exit via _isr_wrapper. One unified API (e.g. k_msgq_put()) checks the CPU's IPSR register in 1 cycle: in a Thread it switches now; in an ISR it switches automatically at exit!", size_pt=8.6, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[9], "Direct (IRQ_DIRECT_CONNECT)", size_pt=11.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[10], "Need maximum speed? Skips the OS wrapper (~20–30 cycles saved) so the Vector Table jumps straight to your C function, just like FreeRTOS.", size_pt=8.7, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[11], "Zero-Latency (IRQ_ZERO_LATENCY)", size_pt=11.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[12], "Matches FreeRTOS unmasked priorities (0–4)! Never paused by OS critical sections—for emergency hardware faults (strict rule: never call OS APIs inside!).", size_pt=8.7, color_hex="595959", font_name=FONT_BODY)
    set_notes(s, (
        "هنا الحل المباشر لكابوس Slide 11 في FreeRTOS: في Zephyr بدل ما الجدول ينط على دالتك مباشرة، في 99% من الحالات (IRQ_CONNECT) بينط على _isr_wrapper "
        "ماسك الدخول والخروج. فالدالة الواحدة k_msgq_put() تسأل ريجستر الهاردوير IPSR في 1 Cycle: أنا في Thread ولا ISR؟ لو Thread تبدّل فوراً، ولو ISR تسيب _isr_wrapper يضرب PendSV أوتوماتيك عند الخروج! "
        "ولو محتاج سرعة قصوى عندك Direct أو Zero-Latency (زي أولويات 0–4 في BASEPRI)."
    ))

    # =========================================================================
    # SLIDE 22: PILLAR 09 — SAFE MEMORY POOLS & POWERING OFF UNUSED RAM BANKS
    # =========================================================================
    s = prs.slides[21]
    prepare_clean_header_slide(s, 1, "Zero-Fragmentation Pools & Turning Off Unused RAM", size_pt=20.2)
    add_lead_subtitle(s, 0.75, 0.92, 8.50, 0.30, "1) Zephyr puts Threads in Static RAM by default (no Heap!).  2) Fixed-size pools eliminate Slide 12's fragmentation & save power:", size_pt=9.0)

    # Left Card: k_mem_slab + Intrusive Pointer Blocks
    add_card(
        s, 0.75, 1.28, 4.15, 2.90,
        "k_mem_slab: Fixed-Size O(1) Pool",
        [
            ("• Why 0% Fragmentation: ", "Every block is identical (e.g. 64 B)—any freed block in the middle fits any new request!"),
            ("• O(1) Free-List & Catch: ", "Stores a 4B 'next free' ptr INSIDE each free block—cutting SRAM Bank 1 power wipes them!"),
        ],
        dark=False, title_pt=11.0, body_pt=8.4
    )
    for b in range(3):
        bx = 0.90 + b * 1.32
        add_lane_block(s, bx, 3.22, 1.18, 0.34, f"4B -> Blk {b+1}" if b < 2 else "4B -> NULL", dark=True, font_pt=7.6)
        add_lane_block(s, bx, 3.58, 1.18, 0.44, f"Block {b} (Free 60B)", muted=False, font_pt=7.6)
        if b < 2:
            add_right_arrow(s, bx + 1.19, 3.56, 0.11, 0.11)

    # Right Card: sys_mem_blocks + Bank 0 Bitmap Diagram
    add_card(
        s, 5.10, 1.28, 4.15, 2.90,
        "sys_mem_blocks: Power Off Unused RAM",
        [
            ("• Separate Bitmap: ", "Moves the free/used map (1 bit per block) safely into Always-On SRAM Bank 0."),
            ("• Deep Sleep Savings: ", "When Bank 1 is empty (Bitmap = [0|0|0|0]), Zephyr cuts power to SRAM Bank 1 safely!"),
        ],
        dark=False, soft=True, title_pt=11.0, body_pt=8.4
    )
    add_lane_block(s, 5.26, 3.22, 1.80, 0.80, "SRAM Bank 0 (ON)\nBitmap: [0 | 0 | 0 | 0]", muted=False, font_pt=7.8)
    add_right_arrow(s, 7.12, 3.55, 0.16, 0.12)
    add_lane_block(s, 7.34, 3.22, 1.76, 0.80, "SRAM Bank 1\n(POWER OFF!)", dark=True, font_pt=8.0)

    # Bottom 3 Memory Pillars
    add_chain_box(s, 0.75, 4.34, 2.70, 0.70, "Static-First (Default)", "K_THREAD_DEFINE puts stacks in Static RAM", dark=False)
    add_chain_box(s, 3.65, 4.34, 2.70, 0.70, "k_heap (When Dynamic)", "Thread-safe dynamic pool if sizes vary", dark=False)
    add_chain_box(s, 6.55, 4.34, 2.70, 0.70, "net_buf (Zero-Copy)", "Shared packet buffers for BLE & Wi-Fi", dark=False)
    set_notes(s, (
        "حل مشكلة الجبنة السويسري في Slide 12 على خطوتين: أولاً في Zephyr الـ Default هو Static RAM (K_THREAD_DEFINE بيحجز الـ Stacks في Static segment وقت الكومبايل مش في Heap). "
        "ثانياً للـ buffers اللي بتتحجز وتفضى وقت الشغل، k_mem_slab بيقسمها لبلوكات كلها مقاس واحد بالظبط (مثلاً 64 بايت)، فأي بلوك يفضى حتى في النص ينفع لأي طلب جديد = صفر fragmentation! "
        "وبيستخدم Free Linked List بـ O(1) بيكتب فيها عنوان البلوك الفاضي الجاي جوه أول 4 بايت من البلوك الفاضي نفسه. "
        "طب لو عايز تطفي كهرباء SRAM Bank 1 وقت النوم من غير ما تضيع الـ pointers؟ هنا sys_mem_blocks بيفصل الـ Bitmap (بتّاية لكل بلوك) في Bank 0 الصاحية عشان تطفي Bank 1 بأمان!"
    ))

    # =========================================================================
    # SLIDE 23: PILLAR 10A — IPC TOOLBOX & k_fifo STACK LIFETIME TRAP
    # =========================================================================
    s = prs.slides[22]
    set_text(s.shapes[0], "Thread Communication (IPC) & The k_fifo Stack Trap", size_pt=20.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)

    add_clean_table(
        s, 0.75, 1.32, 4.55, 3.12,
        headers=["Tool", "How It Works & When to Use"],
        rows=[
            ["k_msgq", "Copies bytes by value (like FreeRTOS Queue). Safe with Stack vars!"],
            ["k_fifo / k_lifo", "Zero-copy list (uses first 4B): FIFO = 1st-in 1st-out; LIFO = last-in 1st-out."],
            ["k_pipe", "Byte stream for variable-length UART or audio chunks."],
            ["k_poll", "Wait on MULTIPLE Queues, Semaphores & Events at the same time."],
            ["k_mutex / k_sem", "Mutex (Ownership + Priority Inheritance) vs. Semaphore / Event."],
        ],
        col_widths=[1.30, 3.25], header_size=9.0, body_size=8.0
    )
    add_code_panel(
        s, 5.48, 1.32, 3.77, 2.02,
        "dangling_fifo_bug.c  (CRASHES!)",
        "void send_sensor(void) {\n"
        "  // Local var lives on Stack (NOT copied!)\n"
        "  struct pkt item = { .res = 0, .val = 42 };\n"
        "  k_fifo_put(&q, &item); // Saves Stack addr!\n"
        "} // <-- Function returns! Stack WIPED!\n"
        "  // Receiver reads GARBAGE!",
        font_size=7.3
    )
    add_chain_box(
        s, 5.48, 3.46, 3.77, 0.98,
        "The Golden k_fifo Lifetime Rule",
        "1. Reserve first 4 bytes (void *res) for linked-list ptr.\n2. Allocate ONLY from k_mem_slab, k_heap, or static RAM — NEVER local Stack!",
        dark=False
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "RULE OF THUMB", "Small copy from Stack -> k_msgq.  Zero-copy large data -> k_mem_slab + k_fifo!", font_size=8.6)
    set_notes(s, (
        "في التواصل بين الـ Threads: k_msgq بتنسخ الداتا بالـ memcpy فآمنة مع المتغيرات المحلية على الـ Stack. "
        "أما k_fifo (اللي يدخل الأول يخرج الأول) وk_lifo (اللي يدخل الآخر يخرج الأول) فدول Zero-Copy linked lists سريعة للداتا الكبيرة وبتستخدم أول 4 بايت كـ pointer. "
        "الفخ اللي في الكود على اليمين: item متعرّف كمتغير محلي جوه الدالة (على الـ Stack)، فلو بعت عنوانه (&item) لـ k_fifo_put، أول ما الدالة ترجع الـ Stack بيتمسح واللي بيستقبل يقرأ Garbage! لازم ييجي من static أو k_mem_slab."
    ))

    # =========================================================================
    # SLIDE 24: PILLAR 10B — ZBUS PUB/SUB & MULTI-CORE COMMUNICATION
    # =========================================================================
    s = prs.slides[23]
    set_text(s.shapes[0], "Scaling Communication: Zbus (Pub/Sub) & Multi-Core", size_pt=20.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "Queues (1-to-1) vs. Zbus (1-to-Many)", size_pt=11.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "With k_msgq/k_fifo, sending one sensor reading to 3 threads needs 3 separate Queues—and the sensor must know all 3!\nZbus replaces that spaghetti with ONE shared, mutex-protected Channel (e.g. sensor_chan).",
        size_pt=9.0, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[1], "Publish / Subscribe (Radio Model)", size_pt=11.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "The Sensor thread publishes once to sensor_chan without knowing who listens.\nBluetooth, Display & Flash Logger subscribe automatically—add a 4th listener tomorrow without touching sensor code!",
        size_pt=9.0, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "Multi-Core AMP (Linux + Zephyr)", size_pt=11.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "On asymmetric multi-core chips (e.g. Cortex-A running Linux + Cortex-M running Zephyr real-time), built-in IPC shares messages via zero-copy Shared RAM ring buffers.",
        size_pt=9.0, color_hex="595959", font_name=FONT_BODY
    )
    set_notes(s, (
        "الفرق الجوهري بين Slide 23 وSlide 24: الـ Queues العادية (k_msgq/k_fifo) علاقة 1-to-1؛ فلو حساس الحرارة عايز يبعت لـ 3 Threads (شاشة، بلوتوث، تسجيل)، "
        "لازم تعمل 3 Queues وحساس الحرارة يعرفهم التلاتة! في Zephyr، Zbus بيشتغل بنظام 1-to-Many (Publish/Subscribe زي قناة راديو): "
        "الحساس ينشر مرة واحدة على قناة sensor_chan، والكل يشترك فيها أوتوماتيك من غير ما يعرفوا بعض! ولو عندك شريحة فيها Core Linux وCore Zephyr (AMP)، فيه تواصل جاهز عبر Shared RAM."
    ))

    # =========================================================================
    # SLIDE 25: PILLAR 11 — HARDWARE MEMORY ISOLATION (MPU USER MODE & DOMAINS)
    # =========================================================================
    s = prs.slides[24]
    prepare_clean_header_slide(s, 0, "Hardware Isolation: MPU User Mode & Memory Domains", size_pt=20.2)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.32, "Default FreeRTOS runs Flat RAM (no MPU)—every thread sees ALL memory, so one bad pointer can crash the whole chip!", size_pt=9.2)

    # Left Card: The Risk Without Isolation
    add_card(
        s, 0.75, 1.38, 3.25, 3.12,
        "Without Isolation (Flat RAM)",
        [
            ("• Full Access Everywhere: ", "Every thread runs Privileged & sees all RAM (FreeRTOS-MPU exists, but needs manual setup)."),
            ("• One Bug Crashes All: ", "A buffer overflow in Wi-Fi/BLE silently overwrites Kernel tables or Motor RAM."),
            ("• Hard to Debug: ", "The chip crashes far away from the real bug!"),
        ],
        dark=False, soft=True, title_pt=10.6, body_pt=8.4
    )

    # Right Visual: Zephyr MPU User Mode Isolation Diagram
    add_card(
        s, 4.22, 1.38, 5.03, 0.88,
        "KERNEL MODE (Privileged — Owns Hardware & OS)",
        "User threads see ONLY their own Stack + their assigned Memory Domain (approved global buffers). OS calls go via SVC.",
        dark=True, align="left", title_pt=10.1, body_pt=8.1
    )
    add_down_arrow(s, 5.40, 2.32, 0.14, 0.24)
    add_down_arrow(s, 7.92, 2.32, 0.14, 0.24)
    add_lead_subtitle(s, 5.60, 2.33, 2.25, 0.24, "Hardware MPU Walls & Stack Guards", size_pt=8.0, bold=True, color=COL_CHARCOAL, align="center")

    add_card(
        s, 4.22, 2.62, 2.42, 1.88,
        "User Domain A (BLE / Net)",
        [
            "Domain = its Stack + ble_rx_buf[] only.",
            ("• HW Stack Guard: ", "Traps stack overflow immediately."),
            ("• Contained Crash: ", "Only BLE faults—Motor stays safe!"),
        ],
        dark=False, title_pt=10.0, body_pt=8.2
    )
    add_card(
        s, 6.83, 2.62, 2.42, 1.88,
        "User Domain B (Motor)",
        [
            "Domain = its Stack + motor_pid RAM only.",
            ("• Walled Off: ", "Cannot touch Domain A or Kernel."),
            ("• Keeps Running: ", "Motor control stays safe!"),
        ],
        dark=False, title_pt=10.0, body_pt=8.2
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.40, "CONFIG_USERSPACE=y", "Powered by the CPU's hardware Memory Protection Unit (MPU) — catches bad pointers & stack overflows the instant they happen!", font_size=8.2)
    set_notes(s, (
        "في FreeRTOS العادي (الديفولت)، كل الـ Threads شغالة Flat RAM بصلاحيات كاملة من غير MPU (نسخة FreeRTOS-MPU موجودة بس إعدادها يدوي)، "
        "فلو Pointer خرف في كود البلوتوث ممكن يمسح بيانات الموتور أو الكيرنل. "
        "في Zephyr بسطر واحد (CONFIG_USERSPACE=y)، هاردوير الـ MPU بيدي كل Thread 'صندوق صلاحيات' اسمه Memory Domain: "
        "Domain A للبلوتوث (يشوف الـ Stack بتاعه + ble_rx_buf[] بس)، وDomain B للموتور (يشوف الـ Stack بتاعه + motor_pid بس) — فلو البلوتوث ضرب، الموتور والكيرنل يفضلوا في أمان تام!"
    ))

    # =========================================================================
    # SLIDE 26: PILLAR 12 — SECURITY, SAFE OTA UPDATES & SAFETY CERTS (Clean English!)
    # =========================================================================
    s = prs.slides[25]
    set_text(s.shapes[0], "Security, Safe OTA Updates & Safety Readiness", size_pt=21.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "MCUboot: Safe A/B OTA Updates", size_pt=12.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "Checks digital signatures before booting. Writes new updates into Slot B while Slot A stays intact—if the new firmware fails self-test, it auto-rolls back!",
        size_pt=9.1, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[1], "ARM TrustZone (Hardware Vault)", size_pt=12.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "Splits the chip into a Secure world (encryption keys & secrets) and a Non-Secure world (your Zephyr app)—so a network bug can never steal device keys.",
        size_pt=9.1, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "Industrial & Automotive Safety", size_pt=12.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "Ships with native TLS encryption and an official certification roadmap for safety-critical products (industrial IEC 61508 & automotive ISO 26262).",
        size_pt=9.1, color_hex="595959", font_name=FONT_BODY
    )
    set_notes(s, (
        "للمنتجات الحقيقية في السوق، Zephyr مدمج معاه 3 ركائز أمان: "
        "1) MCUboot للتحديث الهوائي الآمن (Slot A/B): يكتب التحديث في Slot B ويتأكد من التوقيع الرقمي، ولو فشل يرجع أوتوماتيك للنسخة القديمة. "
        "2) ARM TrustZone يعزل مفاتيح التشفير في منطقة آمنة منفصلة عن التطبيق. "
        "3) دعم رسمي لمسار شهادات السلامة الصناعية والسيارات (IEC 61508 وISO 26262)."
    ))

    # =========================================================================
    # SLIDE 27: PILLAR 13 — AUTOMATIC POWER MANAGEMENT (Plain, Crystal Clear!)
    # =========================================================================
    s = prs.slides[26]
    set_text(s.shapes[0], "Automatic Power Management: Deep Sleep Built In", size_pt=20.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[4], "WFI", size_pt=13.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "STOP", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "DEEP", size_pt=12, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "1. System Power Policy (Auto Sleep)", size_pt=11.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[8], "Checks how long all threads will stay idle (e.g. 4,500 ms) and automatically picks the deepest safe sleep state (Idle, Stop, or Deep Standby) defined in Devicetree.", size_pt=8.8, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[9], "2. Per-Peripheral Auto Clock Gating", size_pt=11.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[10], "Drivers track active usage automatically: the moment no thread uses I2C, SPI, or UART, Zephyr suspends & shuts off that peripheral's clock!", size_pt=8.8, color_hex="595959", font_name=FONT_BODY)

    set_text(s.shapes[11], "3. Automatic Power Domains", size_pt=11.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[12], "When every device inside a chip power block is suspended, Zephyr powers down the entire hardware rail cleanly—zero manual sleep glue code!", size_pt=8.8, color_hex="595959", font_name=FONT_BODY)
    add_dark_badge(s, 5.55, 4.94, 3.25, 0.26, "CSE · Automates Lec 2 & Midterm Q2", font_size=7.6)
    set_notes(s, (
        "في أجهزة البطارية، بدل ما تكتب كود النوم وتطفي الـ Clocks بإيدك زي FreeRTOS: "
        "1) مدير الطاقة في Zephyr بيشوف كل الـ Threads هتنام قد إيه ويختار أعمق وضع نوم مناسب أوتوماتيك (WFI / STOP / DEEP). "
        "2) أي Peripheral زي I2C أو UART محدش بيستخدمه دلوقتي بيتطفى عنه الـ Clock لوحده. "
        "3) لما كل الأجهزة في بلوك معين تنام، بيطفي كهرباء البلوك بالكامل أوتوماتيك!"
    ))

    # =========================================================================
    # SLIDE 28: PILLAR 14 — BUILT-IN CONNECTIVITY & STORAGE STACKS (Clean!)
    # =========================================================================
    s = prs.slides[27]
    set_text(s.shapes[0], "Built-In Connectivity & Storage Stacks", size_pt=22, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[2], "Wi-Fi, Networking & Smart Home", size_pt=12.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "Native TCP/IP, TLS encryption, MQTT, HTTP & Wi-Fi—plus built-in Matter, Thread, USB & CANbus ready out of the box.",
        size_pt=9.1, color_hex="595959", font_name=FONT_BODY
    )
    add_dark_badge(s, 2.75, 0.98, 4.50, 0.26, "CSE · Unifies Class Protocols (MQTT · CoAP · Modbus · LoRa)", font_size=7.5)
    set_text(s.shapes[1], "Complete Bluetooth LE Stack", size_pt=12.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "Full BLE Host & Controller included. Works with built-in chip radios (nRF52/ESP32) OR external UART radio chips using identical app code!",
        size_pt=9.1, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "Power-Fail Safe Storage & DFU", size_pt=12.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "Includes LittleFS (safe against sudden battery loss + flash wear leveling), FATFS, settings storage, and remote firmware updates over BLE or Serial.",
        size_pt=9.1, color_hex="595959", font_name=FONT_BODY
    )
    add_pill_banner(
        s, 0.78, 4.90, 8.44, 0.36,
        "SAME FORMULA",
        "Enable CONFIG_BT=y or CONFIG_FILE_SYSTEM_LITTLEFS=y in prj.conf + wire the radio in Devicetree — zero third-party glue!",
        font_size=8.1
    )
    set_notes(s, (
        "زي ما اتفقنا قبل كده: الـ Driver تحت بيكلم الهاردوير وبيتوصف في Devicetree، أما الـ Stack فوق (البلوتوث، Wi-Fi/MQTT، أنظمة الملفات زي LittleFS الآمنة لو الكهرباء قطعت) "
        "فمدمج جوه Zephyr وتفعله بسطر واحد في prj.conf من غير ما تجيب مكتبات خارجية وتلحمها بإيدك!"
    ))

    # =========================================================================
    # SLIDE 29: PILLAR 15 — RUN FIRMWARE ON YOUR PC (native_sim & Debugging!)
    # =========================================================================
    s = prs.slides[28]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_LAB)

    s.shapes[1].top = Inches(0.91)
    s.shapes[1].left = Inches(0.75)
    s.shapes[1].width = Inches(4.85)
    s.shapes[1].height = Inches(0.80)
    set_text(s.shapes[1], "Boot Firmware on Your PC:\nnative_sim & Built-In Debugging", size_pt=20.0, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    s.shapes[0].top = Inches(1.77)
    s.shapes[0].left = Inches(0.75)
    s.shapes[0].width = Inches(4.85)
    s.shapes[0].height = Inches(2.05)
    set_text(
        s.shapes[0],
        [
            [("• Run on PC (native_sim): ", True, "383838"), ("Compile & run your RTOS threads, networking & automated tests directly on your laptop before hardware arrives!", False, "595959")],
            [("• Interactive Zephyr Shell: ", True, "383838"), ("Live CLI over UART or BLE — type 'kernel threads' or 'device list' right on the chip.", False, "595959")],
            [("• Background Logging: ", True, "383838"), ("Prints debug logs without slowing down real-time Interrupts; exports to SystemView.", False, "595959")],
            [("• Stack Analyzer & Crash Dump: ", True, "383838"), ("Live stack usage per thread + saves CPU/RAM state to Flash if a crash occurs.", False, "595959")],
        ],
        size_pt=9.2, font_name=FONT_BODY, align="left", space_after_pt=5, line_spacing=1.14
    )
    add_code_panel(
        s, 0.75, 3.88, 4.85, 1.15,
        "native_sim — Test firmware before touching hardware!",
        "$ west build -b native_sim app\n"
        "$ ./build/zephyr/zephyr.exe\n"
        "# Runs RTOS threads + MQTT/BLE natively on PC!",
        font_size=7.8
    )
    set_notes(s, (
        "ميزة عملية رهيبة: قبل ما البوردة تيجي من المصنع أصلاً، تقدر تبني مشروعك بـ -b native_sim وتشغّل نظام الـ RTOS كبرنامج عادي على اللاب توب وتعمله Tests! "
        "وكمان جوه Zephyr فيه Shell تفاعلي على الـ UART تشوف بيه الـ Threads والـ Stacks لايف، وLogging مبيبطأش الـ Interrupts، وCrash Dump يحفظ سبب أي عطل في الفلاش."
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
        (6, 2.25, 1.82, 2.45, 1.15, "01. Steep Learning Curve", "FreeRTOS takes hours to learn; mastering Devicetree, Kconfig, CMake & west takes weeks."),
        (9, 5.30, 1.82, 2.45, 1.15, "02. Macro Debug Friction", "Generated Devicetree macros can emit cryptic compiler errors on small overlay typos."),
        (7, 2.25, 3.30, 2.45, 1.15, "03. Heavy Workspace", "A full west checkout pulls GBs of chip HALs & tools; clean builds are slower than 5 FreeRTOS files."),
        (8, 5.30, 3.30, 2.45, 1.15, "04. Strict Framework Rules", "Bypassing SYS_INIT for quick bare-metal hacks fights the framework; always prune debug flags in prj.conf!"),
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
            ["Core Scope", "Minimalist Kernel Only (MIT)", "Complete Embedded OS (Apache 2.0, Linux Found.)"],
            ["Kernel Footprint", "~9 KB Flash / ~2 KB RAM", "~8–18 KB Flash / ~4 KB RAM (minimal)"],
            ["Full IoT App", "~195 KB Flash / ~44 KB RAM (manual)", "~220 KB Flash / ~48 KB RAM (BLE+MQTT+TLS+OTA)"],
            ["HW & Drivers", "Vendor HAL lock-in (CubeMX / ESP-IDF)", "Devicetree (.dts/.overlay) + Unified Flash APIs"],
            ["Configuration", "Manual #defines in FreeRTOSConfig.h", "Dependency-aware Kconfig (prj.conf + menuconfig)"],
            ["Sched & ISR", "Preempt + RR; separate *FromISR() APIs", "Coop, Preempt, EDF, SMP; 1 unified ISR API"],
            ["Memory & MPU", "heap_1..5 (frag risk); FreeRTOS-MPU", "Static-first, k_mem_slab, RAM power-off, MPU Domains"],
            ["Stacks, PM & PC", "Manual integration & custom sleep code", "Built-in BLE/Matter/LittleFS, Auto PM, native_sim"],
        ],
        col_widths=[1.42, 3.22, 3.92], header_size=9.0, body_size=8.0
    )
    add_pill_banner(s, 0.72, 4.70, 8.56, 0.38, "BOTTOM LINE", "Same underlying CPU physics (SysTick, Interrupts, PendSV) — vastly different portability & built-in ecosystem!", font_size=8.4)
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
            ["[6]", "Zephyr Subsystems, Drivers, Zbus, PM Policy & native_sim — Slides 16, 24, 27–29", "docs.zephyrproject.org/latest/services/"],
            ["[7]", "Zephyr Hardware Support, Security & Safety (IEC 61508 / ISO 26262) — Slides 14, 26", "docs.zephyrproject.org/latest/security/"],
            ["[8]", "MCUboot Signed OTA Rollback & ARM TrustZone (TF-M) — Slide 26", "docs.mcuboot.com  |  trustedfirmware.org"],
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
            ("Security & Safety", "Need MPU isolation, TrustZone, safe A/B OTA, or safety certifications."),
            ("Battery & PC testing", "Automatic deep-sleep power policy + PC simulation via native_sim."),
        ],
        dark=True
    )
    add_pill_banner(
        s, 0.75, 4.56, 8.50, 0.48,
        "QUESTIONS?",
        "Master FreeRTOS to understand the RTOS heart — adopt Zephyr when your product needs an OS, not just a kernel.",
        font_size=8.8
    )
    add_dark_badge(s, 0.75, 5.14, 2.15, 0.25, "Team 8 · CSE 421/422", font_size=7.5)
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
