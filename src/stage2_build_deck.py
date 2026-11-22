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
SKEL = os.path.join(ASSETS_DIR, "template_skeleton_32.pptx")
OUT_PPTX = os.path.join(REPO_ROOT, "Zephyr_RTOS_Embedded_Systems.pptx")

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


def add_clean_table(slide, x, y, w, h, headers, rows, col_widths, header_size=9.0, body_size=8.2, col0_bold=True):
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


def prepare_clean_header_slide(slide, title_idx, title_text, size_pt=21.0):
    """Keep ONLY the main title placeholder at consistent top y=0.38" and remove extra template junk."""
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
    assert len(prs.slides) == 32, f"Expected 32 slides, got {len(prs.slides)}"

    # =========================================================================
    # SLIDE 1: HERO COVER (Shared Intro · Team 8 · CSE 421/422 · 24-11 · 32 Slides)
    # =========================================================================
    s = prs.slides[0]
    s.shapes[0].left = Inches(3.65)
    s.shapes[0].top = Inches(1.26)
    s.shapes[0].width = Inches(5.70)
    s.shapes[0].height = Inches(1.20)
    s.shapes[0].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[0], ["Zephyr RTOS"], size_pt=46, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")

    clear_fill_and_line(s.shapes[1])
    s.shapes[1].left = Inches(3.65)
    s.shapes[1].top = Inches(2.52)
    s.shapes[1].width = Inches(5.70)
    s.shapes[1].height = Inches(0.72)
    s.shapes[1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(
        s.shapes[1],
        [
            [("Silicon-to-Cloud Embedded Operating System", True, "383838")],
            [("From ATmega32 Bare-Metal Roots to Real-Time Edge Products", False, "595959")],
        ],
        size_pt=12.8, font_name=FONT_BODY, align="center", space_after_pt=3
    )

    tb_meta = s.shapes.add_textbox(Inches(3.65), Inches(3.54), Inches(5.70), Inches(1.36))
    tf_m = tb_meta.text_frame
    tf_m.word_wrap = True
    tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = 0
    meta_lines = [
        ("CSE 421/422 Embedded Systems  ·  Zagazig University  ·  Dr. Ahmed Osman Mahmoud", True, 8.9, "383838", 5),
        ("Topic 8  ·  Date: 24-11  ·  32 Slides (30 Min · 5 Speakers × 6 Core Slides)", True, 8.5, "595959", 4),
        ("Omar Abdelsalam Abdelfattah (S2–7)  ·  Amr Abdelazim Saeed (S8–13)", False, 8.2, "595959", 2),
        ("Omar Ahmed Mashhour (S14–19)  ·  Ali Mohamed Ali (S20–25)  ·  Ahmed Hassan (S26–31)", False, 8.2, "595959", 0),
    ]
    for mi, (mtxt, mbld, msz, mcol, msp) in enumerate(meta_lines):
        p = tf_m.paragraphs[0] if mi == 0 else tf_m.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(msp)
        r = p.add_run()
        r.text = mtxt
        r.font.name = FONT_BODY
        r.font.size = Pt(msz)
        r.font.bold = mbld
        r.font.color.rgb = RGBColor.from_string(mcol)

    set_notes(s, (
        "[Shared Opening — Speaker 1: Omar Abdelsalam] صباح الخير يا دكتور أحمد، وأهلًا بكم يا شباب. "
        "الأسبوع اللي فات زمايلنا في تيم 4 عرضوا فكرة الـ RTOS، والنهاردة تيم 6 شرحوا الـ MQTT وتيم 7 شرحوا الـ MIPI CSI-2. "
        "إحنا تيم 8، وعلى مدار 30 دقيقة (6 دقايق لكل مهندس فينا على 6 شرائح)، هنوريكم إزاي الأساس اللي درسناه على ATmega32 مع د. أحمد "
        "بيتحول لنظام تشغيل Embedded متكامل اسمه Zephyr RTOS: من قلب السيليكون والمقاطعات، مرورًا بالـ Devicetree وتوحيد كل بروتوكولات الدفعة، "
        "لحد تطبيق كود المحاضرة السابعة وتجربته لايف!"
    ))

    # =========================================================================
    # SLIDE 2: 6-STAGE PRESENTATION ROADMAP (Speaker 1 · Slides 2-7)
    # =========================================================================
    s = prs.slides[1]
    set_text(s.shapes[0], "Presentation Roadmap (30 Min · 5 Speakers)", size_pt=23.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [s.shapes[14], s.shapes[13]]:
        remove_shape(sh)

    roadmap_cfg = [
        (7,  0.89, 2.24, "1. Embedded Roots (S2–7)",      "Omar Abdelsalam: Silicon physics, CSE traps, determinism & Course Map"),
        (8,  3.77, 2.24, "2. Kernel & Hardware (S8–13)",  "Amr Abdelazim: TCB/PSP, PendSV (P15), BASEPRI, Mars Pathfinder & Frag"),
        (9,  6.64, 2.24, "3. Zephyr & Devicetree (S14–19)", "Omar Mashhour: v4.5, Trims DOWN, .dts/.rodata, Flash V-Tables & SYS_INIT"),
        (10, 0.89, 4.14, "4. Build, ISR & IPC (S20–25)",   "Ali Mohamed: Kconfig/west, Tickless/CTC, IPSR, Bank-off RAM, Zbus & MPU"),
        (11, 3.77, 4.14, "5. Systems & Code (S26–31)",    "Ahmed Hassan: Protocols/SIL 3, Q2 PM Policy, Lec 7 HVAC, native_sim & Limits"),
        (12, 6.64, 4.14, "6. Verdict & Q&A (S32)",        "Shared Team 8 Closer: Decision verdict & full Q&A readiness across curriculum"),
    ]
    for si, lx, ly, lhead, ldesc in roadmap_cfg:
        sh = s.shapes[si]
        sh.left = Inches(lx)
        sh.top = Inches(ly)
        sh.width = Inches(2.46)
        sh.height = Inches(0.88)
        set_text(
            sh,
            [[(lhead, True, "383838")], [(ldesc, False, "595959")]],
            size_pt=9.3, font_name=FONT_BODY, align="center", space_after_pt=3
        )
    for ax in (3.44, 6.32):
        add_right_arrow(s, ax, 1.72, 0.24, 0.16)
        add_right_arrow(s, ax, 3.62, 0.24, 0.16)

    set_notes(s, (
        "[Speaker 1: Omar Abdelsalam] العرض متقسم بدقة 6 محطات على 5 أعضاء: "
        "المحطة الأولى معايا: بنقفل جدل علاقة الـ RTOS بالـ Embedded وبنكشف حدود الكورس وبندرس الحتمية وخريطة المنهج. "
        "التانية مع عمرو: بنشوف الرام من جوه، التبديل عبر PendSV، وتجنب الـ Priority Inversion وتفتت الذاكرة. "
        "التالتة مع عمر مشهور: ندخل Zephyr v4.5 ومعجزة Devicetree وموديل الدرايفرات. "
        "الرابعة مع علي: Kconfig، التخلص من *FromISR، طاقة بنوك الرام والـ MPU. "
        "والخامسة مع أحمد: كل بروتوكولات الدفعة وSIL 3 وترقية كود المحاضرة 7 وديمو native_sim، ثم الخاتمة والأسئلة."
    ))

    # =========================================================================
    # SLIDE 3: PURE EMBEDDED SYSTEMS + PROTOCOL UNIFIER (Speaker 1 · mcu_silicon.jpg)
    # =========================================================================
    s = prs.slides[2]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_MCU)

    s.shapes[1].left = Inches(0.72)
    s.shapes[1].top = Inches(0.96)
    s.shapes[1].width = Inches(4.92)
    s.shapes[1].height = Inches(0.74)
    set_text(s.shapes[1], "Pure Embedded Systems:\nSilicon Control + Protocol Unifier", size_pt=18.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")
    remove_shape(s.shapes[0])

    add_card(
        s, 0.72, 1.82, 4.92, 1.28,
        "1. Direct Bare-Silicon Physics (Not Desktop OS)",
        [
            ("No MMU Required: ", "Runs inside 8–18 KB Flash & 4 KB RAM on Cortex-M / RISC-V MCUs."),
            ("Hardware Mastery: ", "Touches NVIC IRQs, GPIOA->BSRR, 32.768 kHz RTC, SRAM Bank power gates & microamp (µA) deep sleep."),
        ],
        dark=False, title_pt=10.6, body_pt=8.7
    )
    add_card(
        s, 0.72, 3.20, 4.92, 1.32,
        "2. The Grand Unifier for Every Class Protocol",
        [
            ("One Verified Embedded OS: ", "Every bus & IoT protocol presented in this course ships inside Zephyr's tree:"),
            ("Built-In: ", "I2C · SPI · CAN · LIN (v4.5) · I3C · MIPI CSI-2 · MQTT (Team 6) · CoAP · Modbus · LoRaWAN · Matter/Thread."),
        ],
        dark=True, title_pt=10.6, body_pt=8.6
    )
    add_pill_banner(s, 0.72, 4.64, 4.92, 0.40, "CORE ANSWER", "Protocols need deterministic firmware orchestration — Zephyr IS that silicon layer.", font_size=8.1)

    set_notes(s, (
        "[Speaker 1: Omar Abdelsalam] أول سؤال نقفله قبل ما نكمل: هل الـ RTOS وZephyr فعلًا صميم Embedded Systems ولا سوفتوير عام؟ "
        "الإجابة بالأرقام: Zephyr معندوش MMU؛ بيشتغل في 8–18 كيلوبايت فلاش و4 كيلوبايت رام، بيتعامل مباشرةً مع ريجسترات الهاردوير، "
        "الـ NVIC، الكريستالة 32.768 kHz، وفصل كهرباء بنوك الـ SRAM عشان ينزل للميكروأمبير! "
        "الأهم: الدكتور طلب من معظم الفرق تشرح بروتوكولات — زي الـ MQTT بتاع تيم 6 النهاردة، والـ MIPI CSI-2، والـ CAN، والـ LIN، والـ Modbus، والـ CoAP، والـ LoRa. "
        "البروتوكول لوحده مجرد باكتات؛ Zephyr هو الطبقة الحية اللي بتربط الحساس والدرايفر والبروتوكول مع بعض في نظام واحد مُختبر!"
    ))

    # =========================================================================
    # SLIDE 4: CSE 421/422 BREAKING POINTS (Speaker 1 · 4 Course Traps)
    # =========================================================================
    s = prs.slides[3]
    prepare_clean_header_slide(s, 1, "Why RTOS Is Mandatory: CSE 421/422 Breaking Points", size_pt=20.2)
    add_lead_subtitle(s, 0.75, 0.92, 8.50, 0.30, "ATmega32 Bare-Metal builds deep hardware mastery—until 4 structural limits from our own lectures break multi-task products:", size_pt=9.3)

    add_card(
        s, 0.75, 1.32, 4.15, 1.54,
        "1. Blocking Delays & Missed Deadlines",
        [
            ("Course Proof (Lec 3 GEN_PWM): ", "Software PWM deceleration uses _delay_ms() for 18 seconds straight!"),
            ("Production Failure: ", "During those 18s, buttons, comms & Emergency Stop stall completely."),
        ],
        dark=False, title_pt=10.3, body_pt=8.4
    )
    add_card(
        s, 5.10, 1.32, 4.15, 1.54,
        "2. Hardware Timer Exhaustion",
        [
            ("Course Proof (Lab 3): ", "Consumes ALL 3 hardware timers on ATmega328P at once!"),
            ("RTOS Solution: ", "One 16/32-bit HW timer (SysTick / CTC) multiplexes infinite software timers."),
        ],
        dark=False, title_pt=10.3, body_pt=8.4
    )
    add_card(
        s, 0.75, 2.98, 4.15, 1.54,
        "3. Multi-Rate Loops, DOR & volatile Bugs",
        [
            ("Course Proof (Lec 7 S62 & Lec 9): ", "0.5s Temp + 2s Anemometer + UART polling loses bytes (DOR overrun)!"),
            ("volatile Trap: ", "volatile stops compiler caching ONLY—it never fixes multi-byte torn reads!"),
        ],
        dark=False, title_pt=10.3, body_pt=8.4
    )
    add_card(
        s, 5.10, 2.98, 4.15, 1.54,
        "4. Static Power & Single-Silicon Lock-In",
        [
            ("Course Proof (Lec 2 & Midterm): ", "Hand-tuned ATmega32 sleep tables lock firmware to 1 chip."),
            ("Market Reality: ", "Connected IoT demands BLE + MQTT + OTA across swappable Cortex-M / RISC-V chips."),
        ],
        dark=False, title_pt=10.3, body_pt=8.4
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.38, "CORE INSIGHT", "Polling loop complexity grows O(N²) when adding sensors & radios; RTOS isolates each task's timing budget!", font_size=8.4)

    set_notes(s, (
        "[Speaker 1: Omar Abdelsalam] المنهج اللي درسناه على ATmega32 بنى عندنا فهمًا عميقًا للهاردوير، لكن في مشاريع الإنتاج الحقيقية بتظهر 4 نقاط انهيار شوفنا بوادرها في محاضراتنا نفسهَا: "
        "1) في المحاضرة 3 (مثال GEN_PWM)، الكود فضل محبوس في _delay_ms لمدة 18 ثانية كاملة — يعني لو عندك زرار طوارئ أو شاشة أو اتصال لاسلكي كله مات! "
        "2) في لاب 3 استخدمنا الـ 3 تايمرات كلهم في نفس اللحظة — طب لو احتجت توقيت رابع في السيليكون؟ مفيش! الـ RTOS بتايمر واحد يديك عدد لا نهائي من الـ Kernel Timers. "
        "3) واجب المحاضرة 7 شريحة 62 (قراءة حرارة كل 0.5 ثانية وسرعة رياح كل ثانيةين) ومعاهم شاشة وUART من المحاضرة 9: أي تأخير يعمل DOR Overrun، ومجرد كلمة volatile مش كافية تمنع التداخل في الرام! "
        "4) وفي المحاضرة 2 والميدتيرم كنا بنحسب أوضاع الطاقة يدويًا على شريحة واحدة — أما منتج الـ IoT الحديث فمحتاج مرونة تنقله بين ARM وRISC-V."
    ))

    # =========================================================================
    # SLIDE 5: REAL-TIME DETERMINISM VS. AVERAGE THROUGHPUT (Speaker 1 · TPL 19)
    # =========================================================================
    s = prs.slides[4]
    set_text(s.shapes[0], "Real-Time Determinism vs. Average Throughput", size_pt=21.0, bold=True, color_hex="383838", font_name=FONT_HEAD)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Fundamental RTOS Law: Predictable worst-case deadline compliance beats high average speed every time:", size_pt=9.4, align="center")

    set_text(s.shapes[4], "SOFT", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "FIRM", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "HARD", size_pt=12.5, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "Hard, Firm & Soft Real-Time Deadlines", size_pt=10.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[8],
        "Hard (Lec 4 Pacemaker / Lec 6 Ventilator INT0): missing 1 deadline causes physical damage or failure. Firm drops stale frame; Soft degrades gracefully.",
        size_pt=8.3, color_hex="595959", font_name=FONT_BODY
    )

    set_text(s.shapes[9], "Worst-Case Latency (WCET) > Average Speed", size_pt=10.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[10],
        "Linux optimizes average throughput (3 GHz UI can still stutter). RTOS bounds interrupt-to-thread dispatch strictly inside ~1–3 µs even on a 48 MHz MCU.",
        size_pt=8.3, color_hex="595959", font_name=FONT_BODY
    )

    set_text(s.shapes[11], "O(1) Preemption vs. Super-Loop O(N) Jitter", size_pt=10.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[12],
        "In a bare-metal Super-Loop, critical response time equals the worst-case sum of ALL modules. Zephyr swaps straight to the highest ready thread cleanly.",
        size_pt=8.3, color_hex="595959", font_name=FONT_BODY
    )

    set_notes(s, (
        "[Speaker 1: Omar Abdelsalam] نقطة أساسية دايمًا بيسأل فيها أي دكتور امبيديد: هل كلمة Real-Time معناها 'أسرع معالج'؟ "
        "لأ! الفرق بين الـ Embedded RTOS وأنظمة السيرفرات هو الحتمية (Determinism) مش مجرد متوسط السرعة (Average Throughput). "
        "عندنا 3 مستويات للمواعيد النهائية من أمثلة المنهج: "
        "Hard Real-Time زي منظم ضربات القلب في المحاضرة 4 أو مقاطعة INT0 لجهاز التنفس الصناعي في المحاضرة 6 — التأخير مللي ثانية واحدة معناها كارثة! "
        "Firm زي قراءة إطار كاميرا أو فريم من الـ HVAC لو اتأخر يترمي، وSoft زي تحديث شاشة OLED أو سجل إحصائيات. "
        "في الـ Super-Loop، زمن استجابتك محكوم بأطول فرع في الدالة الرئيسية O(N)؛ أما الـ RTOS فبيضمن لك إن الخيط الأعلى أولوية يشتغل في حدود ثابتة O(1) لا تتعدى ميكروثواني معدودة!"
    ))

    # =========================================================================
    # SLIDE 6: FIRMWARE CONCURRENCY MODELS (Speaker 1 · 3 Columns)
    # =========================================================================
    s = prs.slides[5]
    prepare_clean_header_slide(s, 1, "Concurrency Models: Super-Loop vs. ISR-Heavy vs. RTOS", size_pt=20.0)
    add_lead_subtitle(s, 0.75, 0.92, 8.50, 0.30, "How embedded firmware architecture evolves across CSE 421/422—from single while(1) polling to isolated multi-thread scheduling:", size_pt=9.3)

    add_lane_block(s, 0.75, 1.32, 2.70, 0.34, "1. Bare-Metal Super-Loop (Lec 1–3)", dark=False, font_pt=8.6)
    add_card(
        s, 0.75, 1.72, 2.70, 2.78,
        "Sequential while(1)",
        [
            ("Control Flow: ", "Single stack; polls buttons, ADC & UART serially."),
            ("Good For: ", "Tiny LEDs, simple motor speed, < 2 KB RAM MCUs."),
            ("Failure Mode: ", "One 18s GEN_PWM loop or slow I2C poll blocks every module!"),
        ],
        dark=False, soft=True, title_pt=9.8, body_pt=8.3
    )

    add_lane_block(s, 3.65, 1.32, 2.70, 0.34, "2. Super-Loop + ISRs (Lec 4–9)", dark=False, font_pt=8.6)
    add_card(
        s, 3.65, 1.72, 2.70, 2.78,
        "Flags + volatile Pointers",
        [
            ("Control Flow: ", "HW IRQs preempt main(), set volatile flags or ring buffers."),
            ("Good For: ", "Medium AVR control loops (Lec 7 HVAC / Lec 9 USART)."),
            ("Failure Mode: ", "Flag soup, torn multi-byte reads, UART DOR & no isolation!"),
        ],
        dark=False, title_pt=9.8, body_pt=8.3
    )

    add_lane_block(s, 6.55, 1.32, 2.70, 0.34, "3. Priority-Preemptive RTOS", dark=True, font_pt=8.6)
    add_card(
        s, 6.55, 1.72, 2.70, 2.78,
        "Isolated Stacks + Events",
        [
            ("Control Flow: ", "Separate thread + Private Stack per module; event-driven."),
            ("Good For: ", "Multi-rate sensors + BLE/MQTT + UI + Safety compliance."),
            ("Advantage: ", "0% CPU idle blocking, O(1) sync & portable across 1,000+ boards!"),
        ],
        dark=True, title_pt=9.8, body_pt=8.3
    )

    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "LECTURE 9 PROOF", "Remember 'volatile char *message_ptr'? volatile blocks compiler optimization ONLY—atomicity & queues require an RTOS!", font_size=8.3)

    set_notes(s, (
        "[Speaker 1: Omar Abdelsalam] لو بصينا لتطور معمارية السوفتوير اللي اتعلمناه في المحاضرات من 1 لـ 10 هنلاقي 3 نماذج واضحة: "
        "النموذج الأول Super-Loop البسيط: while(1) على مكدس واحد، مناسب للـ ATmega32 لو عندنا مهمة واحدة، لكنه ينهار قدام أي تأخير. "
        "النموذج التاني اللي استخدمناه في المحاضرات 4 لـ 9: Super-Loop مع Interrupts. "
        "وهنا بتيجي مصيدة volatile اللي شوفناها في المحاضرة 9 مع مؤشر السيريال message_ptr: volatile بس بتمنع الكومبايلر يعمل caching للمتغير في ريجستر، لكنها مبتحميش قراءة 32-بت متقسمة أو تضارب مهام! "
        "النموذج التالت هو الـ RTOS: كل وظيفة ليها خيط ومكدس خاص وأولوية صريحة، وبتنام 0% CPU لحد ما ييجي الحدث أو التايمر."
    ))

    # =========================================================================
    # SLIDE 7: THE GOLDEN COURSE MAP + HAND-OFF TO SPEAKER 2 (Speaker 1 · TPL 12)
    # =========================================================================
    s = prs.slides[6]
    prepare_clean_header_slide(s, 0, "The Golden Course Map: ATmega32 (Lec 1–10) → Zephyr", size_pt=19.8)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Every bare-metal primitive built in CSE 421/422 evolves directly into an automated industrial service inside Zephyr:", size_pt=9.2)

    add_clean_table(
        s, 0.75, 1.26, 8.50, 3.24,
        headers=["CSE 421/422 Course Topic", "ATmega32 Bare-Metal Mechanism", "Zephyr RTOS v4.5 Upgrade"],
        rows=[
            ["Lec 1–2: Architecture & GPIO",   "Direct registers (DDRB, PORTB, 32 KB Flash)", "Devicetree (.dts) -> 0B RAM .rodata + GPIO APIs"],
            ["Lec 2 S45: Layered HAL Cost",   "Wrapper function calls add cycle overhead",     "__syscall + static inline V-tables (zero extra cost!)"],
            ["Lec 2 & Midterm: Power Modes",  "Manual SM2..0 + PRR bits (Q2: 30 / 61 mW)",     "PM Policy Manager + Runtime PM + 32.768 kHz LSE"],
            ["Lec 3: SW Delays & 18s PWM",    "Blocking _delay_ms() freezes entire CPU",      "Non-blocking k_msleep() (0% CPU) & PWM DTS API"],
            ["Lec 4 & Lab 3: Timer CTC/ISR",  "OCR1A compare ticks & ISR overhead (S22)",      "Tickless Idle dynamically schedules next HW tick only"],
            ["Lec 6: IVT & INT0 Ventilator",  "Vector table, RETI & stack PC push",           "_isr_wrapper (1-cycle IPSR) + Workqueue offloading"],
            ["Lec 7: HVAC + S62 Assignment",  "Timer1 COMPA + ADATE (0.5s & 2s dual-rate)",   "Dedicated ADC thread + K_MSEC(500)/2000 + Zbus"],
            ["Lec 9: USART DOR & volatile",   "Manual volatile RX flags & buffer overruns",    "Interrupt/DMA Ring-Buffers, k_msgq & k_mutex"],
        ],
        col_widths=[2.20, 3.05, 3.25], header_size=8.8, body_size=7.95
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "HAND-OFF -> SPEAKER 2", "Amr Abdelazim now opens the kernel internals: Thread SRAM anatomy, PendSV (Prio 15), BASEPRI & Memory reality.", font_size=8.2)

    set_notes(s, (
        "[Speaker 1: Omar Abdelsalam] والجدول ده هو 'خريطة الكنز' اللي بتربط كل محاضرة خدناها بالـ Zephyr: "
        "شوف المحاضرة 2 شريحة 45 (تكلفة الـ HAL الإضافية) اتحلت بـ static inline؛ "
        "المحاضرة 2 والميدتيرم Q2 (الـ 30 والـ 61 ملي وات وPower-save vs Extended Standby) بقى بيشتغل لوحده بـ PM Policy؛ "
        "المحاضرة 4 ووضع الـ CTC بقى الأساس الفيزيائي لـ Tickless Idle؛ "
        "المحاضرة 6 ومقاطعة INT0 بقت Workqueue؛ ومثال المكيف والـ USART DOR اتحولوا لقنوات وخيوط متزامنة. "
        "وبعد ما أثبتنا من منهجنا إن الـ RTOS ضرورة بنيوية، الباشمهندس عمرو عبدالعظيم هياخدنا جوه الرام والهاردوير عشان يشوف النواة بتشتغل إزاي."
    ))

    # =========================================================================
    # SLIDE 8: SINGLE-CORE EXECUTION & 4 THREAD STATES (Speaker 2 · TPL 6)
    # =========================================================================
    s = prs.slides[7]
    prepare_clean_header_slide(s, 1, "Single-Core Concurrency & Thread Lifecycle (4 States)", size_pt=19.8)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "How 1 core creates concurrency using Zephyr priority numbers (smaller # = higher priority; IDLE sleeps at 0% CPU):", size_pt=9.2)

    # Top CPU Execution Timeline
    add_lane_block(s, 0.75, 1.24, 1.32, 0.42, "HW IRQ\nADC Ready", dark=True, font_pt=7.8)
    add_right_arrow(s, 2.09, 1.38, 0.14, 0.14)
    add_lane_block(s, 2.25, 1.24, 1.85, 0.42, "Sensor Thread\n(Prio 2 · Higher)", dark=True, font_pt=7.8)
    add_right_arrow(s, 4.12, 1.38, 0.14, 0.14)
    add_lane_block(s, 4.28, 1.24, 1.82, 0.42, "Logger Thread\n(Prio 5 · Lower)", dark=False, font_pt=7.8)
    add_right_arrow(s, 6.12, 1.38, 0.14, 0.14)
    add_lane_block(s, 6.28, 1.24, 1.45, 0.42, "Idle (WFI Sleep)\n(0% CPU)", dark=False, dashed=True, font_pt=7.8)
    add_right_arrow(s, 7.75, 1.38, 0.14, 0.14)
    add_lane_block(s, 7.91, 1.24, 1.34, 0.42, "IRQ Preempts!\nPrio 2 Runs", dark=True, font_pt=7.8)

    # Bottom Left: 4 Thread State Cards + arrows
    add_card(s, 0.75, 1.82, 2.20, 1.12, "READY", "Prepared to run; waiting in Run-Queue ordered strictly by numeric priority.", dark=False, align="center", title_pt=10.5, body_pt=8.1)
    add_left_right_arrow(s, 3.01, 2.31, 0.26, 0.14)
    add_card(s, 3.33, 1.82, 2.20, 1.12, "RUNNING", "Active on CPU core right now (executing on its HW PSP Stack).", dark=True, align="center", title_pt=10.5, body_pt=8.1)

    add_up_arrow(s, 1.00, 2.98, 0.14, 0.30)
    add_lead_subtitle(s, 1.18, 2.97, 1.78, 0.32, "IRQ / Msg / Timeout\nwakes to READY", size_pt=7.5, bold=True, color=COL_CHARCOAL)
    add_down_arrow(s, 3.60, 2.98, 0.14, 0.30)
    add_lead_subtitle(s, 3.78, 2.97, 1.75, 0.32, "Blocks: K_FOREVER,\nK_MSEC(100), NO_WAIT", size_pt=7.5, bold=True, color=COL_CHARCOAL)

    add_card(s, 0.75, 3.35, 2.32, 1.14, "BLOCKED (0% CPU)", "Suspended cleanly waiting for data, mutex, or k_msleep(). CPU enters sleep or serves others!", dark=False, align="left", title_pt=9.6, body_pt=7.9)
    add_card(s, 3.21, 3.35, 2.32, 1.14, "SUSPENDED (Frozen)", "Explicit k_thread_suspend() freeze—ignores all events until k_thread_resume()!", dark=False, soft=True, align="left", title_pt=9.6, body_pt=7.9)

    # Bottom Right: Blocking Timeout Discipline & Priority Convention
    add_card(
        s, 5.73, 1.82, 3.52, 2.67,
        "Timeout Discipline & Priority Sign",
        [
            ("• Never Default Blindly to K_FOREVER: ", "If a dead I2C sensor hangs, bounded K_MSEC(100) returns -EAGAIN so recovery logic runs!"),
            ("• K_NO_WAIT Option: ", "Returns immediately if resource busy (mandatory inside ISRs)."),
            ("• Zephyr Priority Scale: ", "Smaller numeric value = Higher priority (<0 Cooperative, >=0 Preemptive) — opposite FreeRTOS!"),
        ],
        dark=False, soft=True, title_pt=10.0, body_pt=8.2
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "EFFICIENCY RULE", "Unlike AVR _delay_ms() burning 16 MIPS spinning, RTOS BLOCKED threads surrender CPU cleanly for IDLE WFI deep sleep!", font_size=8.2)

    set_notes(s, (
        "[Speaker 2: Amr Abdelazim] شكرًا يا عمر. ندخل دلوقتي تحت الكبوت ونفهم إزاي نواة معالج واحد بتدير مهام متعددة بالتوازي الظاهري. "
        "فوق قدامنا مسار التنفيذ بأولويات Zephyr الصحيحة (الرقم الأصغر هو الأعلى!): "
        "المقاطعة بتصحّي خيط الحساس (أولوية 2) يخلّص في ميكروثواني وينام، بعدين الـ Logger (أولوية 5) يشتغل، ولما الكل ينام الـ Idle ينفذ تعليمة WFI بـ 0% CPU. "
        "وكل خيط بيتنقل بين 4 حالات فقط: READY مستني دوره، RUNNING شغال على الـ CPU بـ PSP، "
        "BLOCKED نايم مستني رسالة أو وقت — وهنا قاعدة هندسية مهمة: بلاش K_FOREVER مع هاردوير خارجي، استخدم K_MSEC(100) عشان لو الحساس فصل الكود يرجعلك -EAGAIN ويتعافى! "
        "أما SUSPENDED فحالة تجميد يدوية مبتصحاش إلا بـ k_thread_resume."
    ))

    # =========================================================================
    # SLIDE 9: INSIDE EVERY THREAD — SRAM ANATOMY (TCB, PSP/MSP, MPU GUARD)
    # =========================================================================
    s = prs.slides[8]
    prepare_clean_header_slide(s, 0, "Inside Every Thread: SRAM Layout, PSP/MSP & MPU Guard", size_pt=19.4)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Each Thread owns private SRAM structures—completely isolated from other tasks by ARM Cortex-M stack pointers & HW bounds:", size_pt=9.2)

    # Left SRAM Stack Diagram
    add_lane_block(s, 0.75, 1.26, 4.25, 0.30, "Per-Thread SRAM Footprint (Static Compile Allocation)", dark=True, font_pt=8.6)
    add_chain_box(s, 0.75, 1.62, 4.25, 0.76, "1. struct k_thread (TCB — ~120 Bytes)", "Stores Priority, Thread State, Link Nodes & Saved Stack Pointer (PSP).", dark=False)
    add_chain_box(s, 0.75, 2.44, 4.25, 0.62, "2. HW MPU Stack Guard (32 B Read-Only Zone)", "Armored hardware tripwire—any overflow fires immediate MemManage Fault!", dark=True)
    add_chain_box(s, 0.75, 3.12, 4.25, 0.90, "3. Private Thread Stack (e.g. K_THREAD_STACK_DEFINE 1024 B)", "Local variables + Callee Frame {R4–R11} + IRQ Hardware Frame {R0–R3, R12, LR, PC, xPSR}.", dark=False)
    add_lane_block(s, 0.75, 4.08, 4.25, 0.42, "Dual ARM Stack Pointers:  MSP = ISRs/Boot   |   PSP = User Threads", dark=False, dashed=True, font_pt=8.1)

    # Right Code Panel + Diagnostics Card
    add_code_panel(
        s, 5.18, 1.26, 4.07, 2.26,
        "Static Thread Creation (Exact 9-Arg K_THREAD_DEFINE)",
        "void sensor_task(void *p1, void *p2, void *p3) {\n"
        "  while (1) {\n"
        "    read_bme280();\n"
        "    k_msleep(500); // BLOCKED (0% CPU)\n"
        "  }\n"
        "}\n"
        "/* Statically allocated in .bss — 0 malloc! */\n"
        "K_THREAD_DEFINE(sens_tid, 1024, sensor_task,\n"
        "                NULL, NULL, NULL, 2, 0, 0);",
        font_size=7.35
    )
    add_card(
        s, 5.18, 3.60, 4.07, 0.90,
        "Overflow Protection & Linker Sizing",
        "Linker fails build if RAM overflows before flashing. Runtime Thread Analyzer logs high-water usage!",
        dark=False, soft=True, title_pt=9.6, body_pt=8.1
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "MEMORY REALITY", "Private stacks consume RAM cleanly: use Workqueues or tight Kconfig stacks on smaller 16–32 KB SRAM Cortex-M0+ MCUs!", font_size=8.1)

    set_notes(s, (
        "[Speaker 2: Amr Abdelazim] طب إيه اللي بيتحجز في الـ SRAM لكل Thread عشان كل واحد يبقى ليه حياته المستقلة؟ "
        "على اليمين شايفين التعريف القياسي في Zephyr بماكرو K_THREAD_DEFINE ذو الـ 9 معاملات — بيتحجز كل حاجة وقت الكومبايل في قطاع .bss بدون malloc! "
        "وعلى الشمال تشريح الرام بتاعه: "
        "1) الـ TCB أو struct k_thread (~120 بايت) شايل الأولوية والحالة وآخر عنوان Stack للخيط (PSP). "
        "2) حاجز HW MPU Stack Guard (منطقة Read-Only 32 بايت)؛ لو كودك وقع في Recursion وطفح المكدس يمسكه الهاردوير في اللحظة قبل ما يدمّر بيانات مهمة تانية! "
        "3) الـ Private Stack اللي شايل المتغيرات المحلية ومخزن المسجلات وقت التبديل. وARM Cortex-M بيفصل بين MSP للمقاطعات وPSP للخيوط!"
    ))

    # =========================================================================
    # SLIDE 10: HARDWARE CONTEXT SWITCHING — AVR RETI -> PENDSV (Speaker 2 · TPL 8)
    # =========================================================================
    s = prs.slides[9]
    prepare_clean_header_slide(s, 1, "Hardware Context Switching: SysTick, z_swap() & PendSV", size_pt=19.8)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.30, "Extending Lecture 6 (AVR saves PC on Stack & returns via RETI) to full multi-thread register swapping on ARM Cortex-M:", size_pt=9.4)

    stages_10 = [
        ("1. SysTick / IRQ",
         "HW stacks {R0–R3, R12, LR, PC, xPSR} onto Thread A's PSP stack.", True),
        ("2. z_swap() Pick",
         "Kernel scheduler selects Thread B; sets PENDSVSET in NVIC ICSR.", False),
        ("3. PendSV (Prio 15)",
         "Lowest-IRQ handler pushes {R4–R11}, swaps PSP in TCB & pops B's {R4–R11}.", False),
        ("4. EXC_RETURN",
         "HW auto-pops B's frame; Thread B resumes exact interrupted opcode!", True),
    ]
    for ci, (st_h, st_b, dk) in enumerate(stages_10):
        cx = 0.75 + ci * 2.18
        add_card(s, cx, 1.38, 1.96, 1.88, st_h, st_b, dark=dk, title_pt=9.8, body_pt=8.4)
        if ci < 3:
            add_right_arrow(s, cx + 1.98, 2.25, 0.18, 0.14)

    add_card(
        s, 0.75, 3.42, 6.08, 1.12,
        "Why PendSV Sits at Lowest Priority (15) + Tail-Chaining & Lazy FPU",
        [
            ("• Safe Boundary via Tail-Chaining: ", "PendSV waits until ALL active hardware IRQs finish—never leaving an unprivileged Thread stuck inside interrupt context!"),
            ("• Zero Waste Optimizations: ", "Uses CONFIG_USE_SWITCH fast-path & ARM Lazy FPU stacking (S16–S31) only for threads that actually execute float math."),
        ],
        dark=False, soft=True, title_pt=9.7, body_pt=8.2
    )
    add_card(
        s, 6.97, 3.42, 2.28, 1.12,
        "PRIORITY TRAP!",
        "ARM NVIC IRQs:\n0 = Highest · 15 = Lowest\nZephyr Threads:\nSmaller # = Higher!",
        dark=True, align="center", title_pt=9.5, body_pt=8.0
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.38, "LECTURE 6 LINK", "ATmega32 HW saves PC + I-bit & returns via RETI. Cortex-M adds software {R4–R11} + PSP swap between TCBs!", font_size=8.1)

    set_notes(s, (
        "[Speaker 2: Amr Abdelazim] في المحاضرة السادسة شفنا إزاي الـ AVR في المقاطعة بيحفظ الـ PC ويرجع بـ RETI. "
        "طيب على ARM Cortex-M إزاي النواة بتبدل من Thread A لـ Thread B بالكامل في 4 خطوات هاردوير؟ "
        "1) أول ما يحصل SysTick أو z_swap()، الهاردوير تلقائيًا بيدفع 8 مسجلات (R0-R3, R12, LR, PC, xPSR) على PSP الخاص بـ A. "
        "2) المجدول يختار الخيط B ويرفع علم مقاطعة سوفتوير اسمها PendSV. "
        "3) ليه PendSV محطوطة عند أقل أولوية في الـ NVIC (أولوية 15 — افتكر إن 0 في الـ NVIC هو الأعلى و15 هو الأقل!)؟ عشان تستنى كل المقاطعات العتادية الأعلى تخلص وتدخل بـ Tail-Chaining، تدفع R4–R11، تبدّل مؤشر PSP في الـ TCB، وتسحب R4–R11 للخيط B! "
        "4) مع تعليمة EXC_RETURN الهاردوير بيسحب باقي المسجلات ويكمل الخيط B. وبفضل Lazy FPU ما بيحفظش مسجلات الفلوت غير للخيط اللي بيستخدمها فعلًا!"
    ))

    # =========================================================================
    # SLIDE 11: INTERRUPT LATENCY & CRITICAL SECTIONS — BASEPRI VS PRIMASK
    # =========================================================================
    s = prs.slides[10]
    prepare_clean_header_slide(s, 1, "Interrupt Latency & Critical Sections: BASEPRI vs. PRIMASK", size_pt=19.4)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Applying Lecture 4 Slide 22 ('Minimize ISR frequency & overhead'): protecting shared data without blinding urgent pins:", size_pt=9.2)

    # Top Latency Chain
    add_lane_block(s, 0.75, 1.28, 1.72, 0.78, "1. HW Signal\nPin / Timer IRQ", dark=False, font_pt=8.2)
    add_right_arrow(s, 2.50, 1.60, 0.16, 0.14)
    add_lane_block(s, 2.69, 1.28, 2.02, 0.78, "2. NVIC Entry\n~12 Cycles (~150 ns)", dark=True, font_pt=8.2)
    add_right_arrow(s, 4.74, 1.60, 0.16, 0.14)
    add_lane_block(s, 4.93, 1.28, 2.02, 0.78, "3. Short ISR\nRead Reg + Signal (<2 µs)", dark=False, font_pt=8.2)
    add_right_arrow(s, 6.98, 1.60, 0.16, 0.14)
    add_lane_block(s, 7.17, 1.28, 2.08, 0.78, "4. PendSV Dispatch\nTarget Thread Runs!", dark=True, font_pt=8.2)

    add_card(
        s, 0.75, 2.26, 4.15, 2.22,
        "Global Lockout: cli() / ARM PRIMASK",
        [
            ("Mechanism: ", "AVR cli() and Cortex-M PRIMASK disable ALL maskable interrupts across the entire silicon chip."),
            ("Interrupt Jitter Risk: ", "If a task holds irq_lock() even 100 µs, high-speed USART RX (Lec 9 DOR) or Motor PWM faults sit deaf & blind!"),
            ("Course Warning: ", "Keep critical sections to < 10 instructions—never run printf(), loops, or I2C transfers inside lock!"),
        ],
        dark=False, soft=True, title_pt=10.1, body_pt=8.3
    )
    add_card(
        s, 5.10, 2.26, 4.15, 2.22,
        "Selective Masking: ARM BASEPRI (Zephyr Way)",
        [
            ("Mechanism: ", "Instead of shutting down the whole NVIC, BASEPRI blocks ONLY interrupts at priority >= threshold."),
            ("Hardware Guarantee: ", "Critical NVIC Priority 0 IRQs (Zero-Latency BLE / Motor Safety) ALWAYS preempt kernel locks cleanly!"),
            ("Thread-Only Lock: ", "Need to stop task switches without stopping IRQs? k_sched_lock() blocks Scheduler alone—HW stays live!"),
        ],
        dark=True, title_pt=10.1, body_pt=8.3
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "GOLDEN RULE", "ISR reads HW & signals event (< 5 µs); heavy parsing lives in Threads or Workqueues; ZERO_LATENCY never gets masked!", font_size=8.2)

    set_notes(s, (
        "[Speaker 2: Amr Abdelazim] في المحاضرة 4 شريحة 22، الدكتور ركز جدًا على تقليل زمن وتكرار الـ ISR. "
        "فوق شايفين مكونات الـ Interrupt Latency: الهاردوير ARM Cortex-M بيدخل الفيكتور في 12 سيكل فقط (~150 نانوثانية)، بعدين ISR خاطف يقرا الريجستر ويبعت Signal. "
        "طيب لما نحب نحمي متغير مشترك في Critical Section: هل نقفل المقاطعات كلها بـ cli() أو PRIMASK زي ATmega32؟ "
        "الخطر هنا إنك بتعمّي الشريحة كلها، ومقاطعة UART أو Motor Fault تفوت منك وتعمل DOR! "
        "الحل الأذكى على ARM هو مسجل BASEPRI: تمنع مقاطعات السوفتوير والكيرنل فقط، وتسيب مقاطعات الهاردوير الحرجة (Zero-Latency Priority 0) تدخل بلا أي تأخير. "
        "ولو عايز تمنع تبديل الخيوط بدون لمس المقاطعات خالص، عندك k_sched_lock()!"
    ))

    # =========================================================================
    # SLIDE 12: PRIORITY INVERSION & MARS PATHFINDER (1997) (Speaker 2 · TPL 8)
    # =========================================================================
    s = prs.slides[11]
    prepare_clean_header_slide(s, 1, "Priority Inversion & The Mars Pathfinder (1997) Bug", size_pt=19.8)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Real aerospace lesson: how a low-priority task starved a high-priority system watchdog—and how Zephyr k_mutex fixes it automatically:", size_pt=9.2)

    # Left: THE BUG
    add_lane_block(s, 0.75, 1.26, 4.15, 0.34, "UNPROTECTED: Binary Semaphore (Priority Inversion)", dark=False, font_pt=8.7)
    add_card(
        s, 0.75, 1.66, 4.15, 2.82,
        "1997 Mars Pathfinder Reset Loop",
        [
            ("1. Low (Telemetry): ", "Acquires shared Info-Bus binary semaphore cleanly."),
            ("2. High (Bus Mgmt): ", "Wakes, requests Bus Semaphore -> BLOCKS waiting on Low!"),
            ("3. Med (Comms): ", "Independent Med-priority task preempts Low for seconds!"),
            ("4. Systemic Failure: ", "Low cannot finish to release Bus -> High starves -> Watchdog fires total spacecraft reboot!"),
        ],
        dark=False, soft=True, title_pt=10.1, body_pt=8.35
    )

    # Right: THE FIX
    add_lane_block(s, 5.10, 1.26, 4.15, 0.34, "SOLVED IN ZEPHYR: k_mutex (Priority Inheritance PIP)", dark=True, font_pt=8.7)
    add_card(
        s, 5.10, 1.66, 4.15, 2.82,
        "Built-In Priority Inheritance Protocol",
        [
            ("1. Low (Telemetry): ", "Locks k_mutex (&bus_mtx) around shared peripheral."),
            ("2. High Blocks: ", "Zephyr kernel instantly boosts Low's priority UP to match High's priority!"),
            ("3. Med Blocked Out: ", "Med-priority Comms CANNOT preempt boosted Low now!"),
            ("4. Clean Release: ", "Low finishes immediately, unlocks k_mutex -> priority restores -> High runs on deadline!"),
        ],
        dark=True, title_pt=10.1, body_pt=8.35
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "SYNC RULE", "Binary k_sem = Event Signaling (ISR -> Thread, NO ownership).  k_mutex = Resource Ownership + Priority Inheritance (Thread <-> Thread)!", font_size=8.1)

    set_notes(s, (
        "[Speaker 2: Amr Abdelazim] واحدة من أشهر كوارث الهندسة المدمجة في التاريخ هي مشكلة Priority Inversion على مركبة Mars Pathfinder سنة 1997! "
        "اللي حصل على الشمال: مهمة Low (أولوية واطية) خدت Semaphore للـ Bus. مهمة High صحيت طلبت الـ Bus فوقفت تستنى Low. "
        "في اللحظة دي دخلت مهمة Med مستقلّة قطعت Low ومنعته يكمل! النتيجة: Med أخرت Low، وLow محبوس بالـ High، فالعالي استنى كتير والـ Watchdog رستر المركبة! "
        "الحل جوه نواة Zephyr مبني جاهز في كائن الـ k_mutex بخاصية Priority Inheritance (PIP): "
        "أول ما High يقف مستني الـ Mutex، النواة ترفع أولوية Low مؤقتًا لأولوية High، فمهمة Med متقدرش تقطعه، يخلص ويسلم الـ Mutex فورًا، وترجع أولويته طبيعية! "
        "وافتكر دايمًا: الـ k_sem للتنبيه (Signal)، والـ k_mutex لملكية الموارد عشان فيه Priority Inheritance!"
    ))

    # =========================================================================
    # SLIDE 13: DAY 1 VS DAY 90 MEMORY REALITY + HAND-OFF TO SPEAKER 3 (Speaker 2 · TPL 6)
    # =========================================================================
    s = prs.slides[12]
    prepare_clean_header_slide(s, 1, "Day 1 vs. Day 90 Memory: Fragmentation & Static-First", size_pt=19.6)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Why raw C malloc()/free() passes university demos on Day 1 yet silently locks up edge devices after months in the field:", size_pt=9.2)

    # Day 1 contiguous visual blocks
    add_lane_block(s, 0.75, 1.28, 8.50, 0.30, "DAY 1: Contiguous Heap After Startup  —  malloc(32 KB) SUCCEEDS!", dark=False, font_pt=8.4)
    add_lane_block(s, 0.75, 1.62, 1.85, 0.42, "Boot / App (18 KB)", dark=True, font_pt=8.0)
    add_lane_block(s, 2.63, 1.62, 1.55, 0.42, "Sensor (12 KB)", dark=True, font_pt=8.0)
    add_lane_block(s, 4.21, 1.62, 5.04, 0.42, "Contiguous Free Heap: 98 KB Available  ✅", dark=False, dashed=True, font_pt=8.2)

    # Day 90 fragmented visual blocks
    add_lane_block(s, 0.75, 2.18, 8.50, 0.30, "DAY 90: External Fragmentation After Alloc/Free Churn  —  malloc(32 KB) RETURNS NULL!", dark=True, font_pt=8.4)
    frag_segs = [
        (0.75, 1.25, "18 KB Used", True),
        (2.03, 1.05, "20 KB Free", False),
        (3.11, 1.15, "16 KB Used", True),
        (4.29, 1.12, "22 KB Free", False),
        (5.44, 1.25, "24 KB Used", True),
        (6.72, 1.18, "24 KB Free", False),
        (7.93, 1.32, "No 32K Gap ❌", True),
    ]
    for fx, fw, flbl, fdk in frag_segs:
        add_lane_block(s, fx, 2.52, fw, 0.44, flbl, dark=fdk, dashed=(not fdk), font_pt=7.5)

    add_card(
        s, 0.75, 3.14, 4.15, 1.36,
        "1. O(1) Fixed-Block Memory Pools",
        [
            ("No Walking Free Lists: ", "FreeRTOS heap_4 scans variable blocks dynamically (non-deterministic O(N) & locks scheduler!)."),
            ("Deterministic Fix: ", "RTOS divides buffers into identical fixed blocks via bitmaps/slabs -> O(1) constant time & ZERO fragmentation forever!"),
        ],
        dark=False, soft=True, title_pt=9.9, body_pt=8.2
    )
    add_card(
        s, 5.10, 3.14, 4.15, 1.36,
        "2. Static-First Embedded Philosophy",
        [
            ("Compile-Time Guarantee: ", "Define Threads, Queues, Mutexes & Buffers globally in .bss at build time—linker checks RAM 100% before deployment!"),
            ("Zero Surprise OOM: ", "Critical medical/industrial loops NEVER call heap malloc() in the field."),
        ],
        dark=True, title_pt=9.9, body_pt=8.2
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "HAND-OFF -> SPEAKER 3", "Omar Mashhour now reveals Zephyr RTOS v4.5: ecosystem scale, Zero-RAM Devicetree, Flash V-Tables & SYS_INIT boot!", font_size=8.1)

    set_notes(s, (
        "[Speaker 2: Amr Abdelazim] آخر مشكلة قاتلة في النواة هي الذاكرة: ليه أي مهندس امبيديد بيخاف من malloc وfree العادية؟ "
        "بص فوق على اليوم الأول: الرام حتة واحدة متصلة بمساحة 98 كيلوبايت، فتطلب 32 كيلوبايت تنجح. "
        "بعد 90 يوم تشغيل متواصل من حجز ومسح أحجام مختلفة، الرام بتتقطع 'جبنة سويسرية': عندك 66 كيلوبايت فاضيين مجموعين، بس مفيش ولا فراغ واحد متصل يشيل 32 كيلوبايت -> يرجعلك NULL والجهاز يقع! "
        "عشان كده الكيرنل القوي بيشتغل بمبدأين: Static-First (احجز كل خيوطك وطوابيرك وقت الـ Compile في .bss والـ Linker يضمن لك الرام من غير مفاجآت)، "
        "ومجمعات ذاكرة بكتل ثابتة O(1) من غير تفتت أبدًا! "
        "وشفنا كده كل أساسيات النواة؛ الباشمهندس عمر أحمد مشهور هيدخلنا دلوقتي عالم Zephyr RTOS v4.5 ويشرح معجزة الـ Devicetree بدون استهلاك رام!"
    ))

    # =========================================================================
    # SLIDE 14: SECTION PIVOT — ENTER ZEPHYR RTOS v4.5 (Speaker 3 · TPL 5)
    # =========================================================================
    s = prs.slides[13]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_BOARDS)

    s.shapes[1].left = Inches(4.45)
    s.shapes[1].top = Inches(0.78)
    s.shapes[1].width = Inches(0.95)
    s.shapes[1].height = Inches(0.82)
    s.shapes[1].fill.solid()
    s.shapes[1].fill.fore_color.rgb = COL_CHARCOAL
    s.shapes[1].line.fill.background()
    s.shapes[1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[1], "v4.5", size_pt=24, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    s.shapes[0].left = Inches(5.48)
    s.shapes[0].top = Inches(0.80)
    s.shapes[0].width = Inches(3.82)
    s.shapes[0].height = Inches(0.86)
    set_text(s.shapes[0], "Enter Zephyr\nRTOS v4.5", size_pt=25, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    tb = s.shapes.add_textbox(Inches(4.45), Inches(1.72), Inches(4.56), Inches(3.38))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    v45_items = [
        ("Oct 2026 Release Scale: ", "18,075 commits by 1,230 contributors inside v4.5 alone (+178 boards, now 1,000+ supported boards total)."),
        ("v4.5 Technical Highlights: ", "Infineon TriCore & Qualcomm Hexagon ports; native USB-C mpipe & LIN bus drivers; official west spdx (SPDX 3.0) SBOM compliance."),
        ("Industry Backing (Apache 2.0): ", "Linux Foundation neutral governance, v3.7 LTS (to 2029), planned v4.6 LTS (Apr 2027) & 21% Eclipse IoT market share (2024)."),
        ("Shipping Commercial Products: ", "Ray-Ban Meta smart glasses (with ExecuTorch), Nordic nRF Connect SDK, Garmin wearables & Oticon medical aids."),
    ]
    for idx_v, (vh, vb) in enumerate(v45_items):
        p = tf.paragraphs[0] if idx_v == 0 else tf.add_paragraph()
        p.space_after = Pt(7.5)
        p.line_spacing = 1.15
        r1 = p.add_run()
        r1.text = "▪ " + vh
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.7)
        r1.font.bold = True
        r1.font.color.rgb = COL_CHARCOAL
        r2 = p.add_run()
        r2.text = vb
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.7)
        r2.font.color.rgb = COL_MUTED

    set_notes(s, (
        "[Speaker 3: Omar Mashhour] شكرًا يا عمرو. ندخل دلوقتي على بطل عرضنا: Zephyr RTOS v4.5! "
        "إحنا قدام أحدث إصدار نزل 28 أكتوبر 2026: في v4.5 لوحده دخل 18,075 commit من 1,230 مساهم، "
        "واتضاف 178 بوردة جديدة عشان الإجمالي يتخطى +1,000 بوردة على أكتر من 13 معمارية "
        "(ومعانا في v4.5 معماريات ضخمة زي Infineon TriCore للسيارات وQualcomm Hexagon!). "
        "كمان نزل فيه درايفر LIN رسمي وUSB mpipe وأداة west spdx بـ SPDX 3.0. "
        "ومشروع محكوم من Linux Foundation برخصة Apache 2.0 مع خطة LTS حتى 2029 (وv4.6 LTS مخطط أبريل 2027) وحصته وصلت 21% في استطلاع Eclipse، "
        "وده مش نظام تجارب: ده النظام اللي مشغل نظارات Ray-Ban Meta (مع ExecuTorch)، ومنصة Nordic nRF Connect، وساعات Garmin!"
    ))

    # =========================================================================
    # SLIDE 15: FULL "BATTERIES-INCLUDED" OS CANVAS (Speaker 3 · TPL 15)
    # =========================================================================
    s = prs.slides[14]
    set_text(s.shapes[0], "Zephyr RTOS: Full \"Batteries-Included\" Embedded OS", size_pt=20.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    tbl = s.shapes[1].table
    canvas_cells = [
        ((0, 0), "1,000+ Boards", True, 9.5),
        ((1, 0), "13+ CPU Archs:\nARM, RISC-V, Xtensa, x86, TriCore, Hexagon.\nST, Nordic, NXP, ESP, TI.", False, 8.1),
        ((0, 1), "App & Zbus", True, 9.5),
        ((1, 1), "Portable C/C++, POSIX APIs & Pub/Sub Zbus IPC.", False, 8.0),
        ((2, 1), "HW Desc & Build", True, 9.5),
        ((3, 1), "Devicetree (.dts), Kconfig, CMake, west & native_sim.", False, 8.0),
        ((0, 2), "Linux Foundation · Apache 2.0", True, 8.7),
        ((1, 2), "1 Unified GitHub Repo (PRs + MAINTAINERS.yml + Twister CI).\nKernel + Drivers + Protocols + Bootloader verified together—zero third-party version hell.", False, 8.0),
        ((0, 4), "Native Protocols", True, 9.5),
        ((1, 4), "BLE 5.x, IPv4/6, MQTT, CoAP, Matter, CAN, LIN, Modbus.", False, 8.0),
        ((2, 4), "Security & Boot", True, 9.5),
        ((3, 4), "TrustZone (TF-M), PSA Crypto, MPU Domains & MCUboot.", False, 8.0),
        ((0, 5), "Kernel & Drivers", True, 9.5),
        ((1, 5), "Coop, Preempt, EDF, SMP, O(1) Slabs, Tickless (CTC) & Unified APIs.", False, 8.1),
        ((4, 0), "Micro-Kernel Paradigm (Builds UP)", True, 9.0),
        ((5, 0), "Starts from a tiny ~9 KB kernel; your team integrates vendor HALs, lwIP, NimBLE & OTA manually.", False, 8.1),
        ((4, 3), "Zephyr Paradigm (Trims DOWN — Lec 2 Layers!)", True, 9.0),
        ((5, 3), "Industrializes Lecture 2's 4-layer architecture; Kconfig & CMake prune unused code at compile time.", False, 8.1),
    ]
    for (r, c), txt, is_h, sz in canvas_cells:
        set_cell_text(
            tbl.cell(r, c), txt, size_pt=sz, bold=is_h,
            color_hex="383838" if is_h else "595959",
            font_name=FONT_HEAD if is_h else FONT_BODY,
            align="center" if is_h else "left"
        )

    set_notes(s, (
        "[Speaker 3: Omar Mashhour] السلايد ده بيوضح الفارق الفلسفي الأعمق بين Zephyr وأنظمة Micro-Kernel التقليدية زي FreeRTOS: "
        "الـ Micro-Kernel ماشي بمبدأ Builds UP: تاخد نواة ~9 كيلوبايت، وبعدها فريقك يقعد شهور يركب HAL المورد، ويدمّج lwIP وNimBLE وMCUboot ويحل تعارض نسخهم! "
        "أما Zephyr فماشي بمبدأ Trims DOWN، وده حرفيًا تطبيق صناعي للمعمارية الطبقية اللي خدناها في المحاضرة التانية: "
        "كل شيء — من الـ Devicetree والدرايفرات لحد الـ BLE والـ MQTT والـ CAN والـ LIN والـ MCUboot — في مستودع GitHub واحد محكوم بـ MAINTAINERS.yml ومُختبر آليًا عبر west twister مع كل PR، "
        "وKconfig بيقص أي جزء مش عايزه!"
    ))

    # =========================================================================
    # SLIDE 16: DEVICETREE — 3-LAYER HARDWARE DECOUPLING (Speaker 3 · TPL 18)
    # =========================================================================
    s = prs.slides[15]
    set_text(s.shapes[0], "Devicetree: Decoupling Silicon & Pins from C", size_pt=21.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    tree_16 = {
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
    for si, (txt, col, sz) in tree_16.items():
        set_text(s.shapes[si], txt, size_pt=sz, bold=True, color_hex=col, font_name=FONT_BODY, align="center")
    add_pill_banner(
        s, 0.78, 4.78, 8.44, 0.36,
        "app.overlay",
        "&i2c1 { status = \"okay\";  bme280@76 { compatible = \"bosch,bme280\"; reg = <0x76>; }; };   // Re-targets SoCs without touching C!",
        font_size=7.9
    )

    set_notes(s, (
        "[Speaker 3: Omar Mashhour] في المنهج لو حبيت تنقل الكود من ATmega32 لـ STM32، هتضطر تعدل أسماء ريجسترات ودبابيس جوه ملفات الـ C. "
        "في Zephyr، كود التطبيق في main.c معزول تمامًا عن السيليكون عبر 3 طبقات Devicetree وراثية: "
        "1) .dtsi بيوصف الشريحة نفسهَا من الشركة الصانعة (الريجسترات والـ IRQs). "
        "2) .dts بيوصف البوردة (الكريستالة والـ led0 والدبابيس المربوطة). "
        "3) app.overlay خاص بتطبيقك (إضافة حساس BME280 على I2C1 بعناوينه). "
        "لما الشركة تطلب تغير من STM32 لـ Nordic أو ESP32: تغير علم -b في سطر البناء والـ overlay فقط من غير ما تعدل سطر واحد في C!"
    ))

    # =========================================================================
    # SLIDE 17: 0 BYTES RAM FOR DEVICETREE (.rodata) & CHART (Speaker 3 · TPL 11)
    # =========================================================================
    s = prs.slides[16]
    set_text(s.shapes[0], "0 Bytes RAM for Devicetree: .rodata & Footprint", size_pt=21.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    stats_17 = [
        (4, 13, "0 Bytes RAM",     "Runtime RAM parsing Devicetree"),
        (3, 12, "Python DTS",      "Generates devicetree_generated.h"),
        (2, 11, "Flash (.rodata)", "const struct device lives in ROM"),
        (1, 10, "~195 vs ~220 KB", "Full IoT* stack delta vanishes!"),
    ]
    for hi, bi, htxt, btxt in stats_17:
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
        "[Speaker 3: Omar Mashhour] السؤال اللي بييجي فورًا في ذهن أي مهندس امبيديد لما نقول Devicetree: طب ما الملفات النصية دي هتستهلك رام؟ "
        "الإجابة: 0 بايت RAM وقت التشغيل! لأن الـ Devicetree بيتحلل بالكامل وقت الـ Build بسكربتات Python تطلع هيدر C ثابت، "
        "وهياكل الأجهزة const struct device بتتخزن جوه قطاع .rodata في الفلاش فقط! "
        "والرسم البياني بيكشف حقيقة الحجم: نواة FreeRTOS لوحدها أصغر (~9KB مقابل ~14KB)، لكن أول ما تبني منتج IoT حقيقي (بلوتوث + MQTT + TLS + OTA) الفجوة بتتلاشى تقريبًا (~195KB مقابل ~220KB)!"
    ))

    # =========================================================================
    # SLIDE 18: UNIFIED DRIVER MODEL — SOLVING LEC 2 SLIDE 45 (Speaker 3 · TPL 6)
    # =========================================================================
    s = prs.slides[17]
    prepare_clean_header_slide(s, 1, "Unified Driver Model: Solving Lecture 2 Slide 45", size_pt=20.8)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.30, "Lecture 2 warned that HAL layers add call overhead. Zephyr delivers universal portability at direct-register speed:", size_pt=9.4)

    drivers_18 = [
        ("1. App Calls Abstract API",
         "Application invokes device_get_binding() & i2c_write(dev, ...) — 100% chip-agnostic.", False),
        ("2. Flash V-Table Struct",
         "const struct i2c_driver_api lives purely in .rodata Flash (0 bytes copied to SRAM).", False),
        ("3. static inline -> HW Reg",
         "Compiler inlines api->transfer() directly—zero extra function-call penalty!", True),
    ]
    for ci, (dh, db, dk) in enumerate(drivers_18):
        cx = 0.75 + ci * 2.90
        add_card(s, cx, 1.38, 2.70, 1.54, dh, db, dark=dk, title_pt=10.1, body_pt=8.4)
        if ci < 2:
            add_right_arrow(s, cx + 2.72, 2.08, 0.16, 0.14)

    add_code_panel(
        s, 0.75, 3.06, 5.20, 1.48,
        "Zephyr Zero-Overhead Driver Dispatch (Flash .rodata + static inline)",
        "__syscall int gpio_pin_toggle(const struct device *port, gpio_pin_t pin);\n"
        "static inline int z_impl_gpio_pin_toggle(const struct device *port, gpio_pin_t pin) {\n"
        "    const struct gpio_driver_api *api = (const struct gpio_driver_api *)port->api;\n"
        "    return api->port_toggle_bits(port, BIT(pin)); // Direct Flash V-Table call!\n"
        "}",
        font_size=7.4
    )
    add_card(
        s, 6.10, 3.06, 3.15, 1.48,
        "Zero Wrapper Penalty in Production",
        "When CONFIG_USERSPACE=n, z_impl_ inlines away completely: portable STM32/nRF/ESP32 driver calls execute almost identically to direct REG=VAL!",
        dark=False, soft=True, title_pt=10.0, body_pt=8.4
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.38, "LEC 2 SLIDE 45 SOLVED", "Dual safety: full MPU syscall gating in User Mode, OR zero-overhead static inline compilation in Supervisor Mode!", font_size=8.2)

    set_notes(s, (
        "[Speaker 3: Omar Mashhour] في المحاضرة التانية شريحة 45، د. أحمد نبهنا لعيوب الطبقية الـ HAL وهي إنها ممكن تضيف تكلفة نداء دوال (Call Overhead) وتبطئ الهاردوير. "
        "بص Zephyr حل المعضلة دي إزاي بذكاء هندسي رائع: "
        "التطبيق بينادي دالة موحدة زي gpio_pin_toggle، وهي متعرفة static inline مع ماكرو __syscall. "
        "لو أنت شغال في Supervisor Mode، الكومبايلر بيفرد الدالة جوه كودك مباشرةً ويروح منادي مؤشر الدالة من جدول Flash V-Table في .rodata (0 بايت RAM!) بضربة واحدة، فبتجمع بين نظافة المحاضرة 2 وسرعة الكود المباشر!"
    ))

    # =========================================================================
    # SLIDE 19: BEFORE main() STARTS — 5 SYS_INIT LEVELS + HAND-OFF TO SPEAKER 4
    # =========================================================================
    s = prs.slides[18]
    prepare_clean_header_slide(s, 1, "Before main() Starts: 5 Linker-Sorted SYS_INIT Boot Levels", size_pt=19.4)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.30, "No manual HAL_Init() ordering bugs: Zephyr sorts Flash init pointers into 5 deterministic stages at link time:", size_pt=9.4)

    boot_19 = [
        ("1. EARLY",        "Clocks & CPU reset before device init.", False),
        ("2. PRE_KERNEL_1", "Interrupt stack; UART console comes up.", False),
        ("3. PRE_KERNEL_2", "Timer/SysTick prep; no threads yet.", False),
        ("4. POST_KERNEL",  "Kernel alive! Threaded sensors/buses.", True),
        ("5. APPLICATION",  "Services ready -> enters void main().", False),
    ]
    for ci, (bh, bb, dk) in enumerate(boot_19):
        cx = 0.75 + ci * 1.72
        add_card(s, cx, 1.38, 1.56, 1.72, bh, bb, dark=dk, align="center", title_pt=8.9, body_pt=8.0)
        if ci < 4:
            add_right_arrow(s, cx + 1.58, 2.16, 0.12, 0.14)

    add_code_panel(
        s, 0.75, 3.26, 4.20, 1.24,
        "How Drivers Auto-Register in Flash Linker Sections",
        "SYS_INIT(clock_early, PRE_KERNEL_1, 0);\n"
        "DEVICE_DT_DEFINE(node, sensor_init, NULL, &data, &cfg,\n"
        "                 POST_KERNEL, 50, &sensor_api);",
        font_size=7.4
    )
    add_card(
        s, 5.10, 3.26, 4.15, 1.24,
        "Zero Manual HAL_Init() Ordering Hell",
        "Linker script collects every init pointer sorted by (Level, Priority 0–99)—drivers wake automatically BEFORE main()!",
        dark=False, soft=True, title_pt=10.0, body_pt=8.4
    )
    add_pill_banner(s, 0.75, 4.66, 8.50, 0.38, "HAND-OFF -> SPEAKER 4", "Never call k_msleep() in PRE_KERNEL_1! Next: Ali Mohamed reveals Kconfig/west, Tickless/CTC, IPSR & SRAM Bank-off!", font_size=8.0)

    set_notes(s, (
        "[Speaker 3: Omar Mashhour] في البرمجة التقليدية، لو نسيت ترتيب HAL_Init أو SystemClock_Config أو MX_SPI_Init جوه main، الطرفية بتقف من غير صوت! "
        "في Zephyr، الـ Linker بيرتب مؤشرات تهيئة الهاردوير في الفلاش أوتوماتيك على 5 مستويات قبل main() ما تبدأ: "
        "EARLY ثم PRE_KERNEL_1 وPRE_KERNEL_2 ودول شغالين على الـ Interrupt Stack قبل ما الكيرنل يصحى (فعمرك ما تنادي k_msleep أو mutex جواهم!)، "
        "بعدين POST_KERNEL وAPPLICATION ولما main() تبدأ بتكون البوردة كلها مهيئة! "
        "ودلوقتي الباشمهندس علي محمد علي هيكمل رحلة البناء بالـ Kconfig والـ Tickless ومفاجأة إزاي Zephyr ألغى *FromISR!"
    ))

    # =========================================================================
    # SLIDE 20: KCONFIG (prj.conf) + CMAKE + WEST (Speaker 4 · TPL 17)
    # =========================================================================
    s = prs.slides[19]
    set_text(s.shapes[0], "Kconfig + CMake + west: Pruning & Sysbuild", size_pt=21.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)

    add_clean_table(
        s, 0.75, 1.30, 4.45, 3.16,
        headers=["west Command", "What It Automates"],
        rows=[
            ["west init / update", "Syncs Zephyr, HALs, mbedTLS & MCUboot"],
            ["west build -b <board>", "Runs DT + Kconfig + CMake + Ninja"],
            ["west build -t menuconfig", "Interactive TUI solving dependencies"],
            ["west flash / debug", "Unified J-Link, ST-Link & OpenOCD GDB"],
            ["west build --sysbuild", "Builds App + MCUboot + BLE Net core!"],
            ["west spdx", "Exports SPDX 3.0 SBOM (EU CRA ready)"],
        ],
        col_widths=[1.85, 2.60], header_size=8.8, body_size=8.0
    )

    tb20 = s.shapes.add_textbox(Inches(5.38), Inches(1.28), Inches(3.87), Inches(1.36))
    tf20 = tb20.text_frame
    tf20.word_wrap = True
    tf20.margin_left = tf20.margin_right = tf20.margin_top = tf20.margin_bottom = 0
    for idx_l, (pre, rest) in enumerate([
        ("Why prj.conf Alongside Devicetree?", ""),
        ("• Devicetree = HW PRESENT: ", "Board has UARTs, ADC, CAN, BLE."),
        ("• prj.conf = SW WANTED: ", "CONFIG_GPIO=y compiles GPIO only."),
        ("• Auto-Deps: ", "CONFIG_BT=y pulls crypto & pools automatically."),
    ]):
        p = tf20.paragraphs[0] if idx_l == 0 else tf20.add_paragraph()
        p.space_after = Pt(4 if idx_l == 0 else 3)
        r1 = p.add_run()
        r1.text = pre
        r1.font.name = FONT_HEAD if idx_l == 0 else FONT_BODY
        r1.font.size = Pt(10.8 if idx_l == 0 else 8.7)
        r1.font.bold = True
        r1.font.color.rgb = COL_CHARCOAL
        if rest:
            r2 = p.add_run()
            r2.text = rest
            r2.font.name = FONT_BODY
            r2.font.size = Pt(8.7)
            r2.font.color.rgb = COL_MUTED

    add_code_panel(
        s, 5.38, 2.74, 3.87, 1.72,
        "drivers/gpio/CMakeLists.txt (Dead-Code Pruning)",
        "# DT 'compatible' + Kconfig -> CONFIG_GPIO_STM32=y\n"
        "zephyr_library_sources_ifdef(\n"
        "    CONFIG_GPIO_STM32 gpio_stm32.c)\n"
        "zephyr_library_sources_ifdef(\n"
        "    CONFIG_GPIO_ESP32 gpio_esp32.c)\n"
        "# gpio_esp32.c NEVER enters STM32 Flash!",
        font_size=7.3
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "PRECISION", "Devicetree asks: 'What hardware is on the board?'  |  Kconfig asks: 'Which software enters Flash?'", font_size=8.5)

    set_notes(s, (
        "[Speaker 4: Ali Mohamed] شكرًا يا عمر. القاعدة الذهبية لنظام البناء بتلخص الفرق في سطر واحد: "
        "Devicetree بيجاوب 'إيه الهاردوير الموجود على البوردة؟'، أما ملف prj.conf في Kconfig فبيجاوب 'إيه السوفتوير اللي عايز تدخله الفلاش النهاردة؟'. "
        "لما تكتب CONFIG_BT=y، Kconfig يحل كل التبعيات لوحده، وCMake يعمل Dead-Code Pruning فيترجم gpio_stm32.c فقط. "
        "وكل ده بيتدار بأداة west: من سطر واحد تبني، تحرق، تفتح menuconfig، تجمع التطبيق مع البوت لودر بـ --sysbuild، وتصدر ملف SBOM عبر west spdx!"
    ))

    # =========================================================================
    # SLIDE 21: RICH SCHEDULING & TICKLESS IDLE — CTC BRIDGE (Speaker 4 · TPL 9)
    # =========================================================================
    s = prs.slides[20]
    set_text(s.shapes[0], "Scheduling & Tickless Idle (The CTC Bridge)", size_pt=21.2, bold=True, color_hex="383838", font_name=FONT_HEAD)
    grid_21 = [
        (5,  1, "Cooperative (< 0)",         "Priorities -1, -2: NEVER preempted by threads (only ISRs)! Runs until k_yield() or block—cuts mutex locks."),
        (6,  2, "Preemptive (>= 0)",         "Preempted immediately by smaller numeric priority (0 = highest, opposite of FreeRTOS!) + optional Round-Robin."),
        (7,  8, "Meta-IRQ & EDF",            "Meta-IRQ preempts even coop threads for microsecond BLE radio events; CONFIG_SCHED_DEADLINE adds EDF."),
        (10, 3, "Workqueues (Lec 6)",        "Offload heavy Ventilator INT0 logging/alerts via k_work_submit() onto sysworkq—zero extra thread stack!"),
        (11, 4, "Tickless Idle (Lec 4 CTC)", "Solves Lec 4 S22 ISR overhead! Stops 1 ms SysTick when idle; programs HW timer for next wakeup only."),
        (12, 9, "SMP & Multi-Core AMP",      "Symmetric multi-core spinlocks + CPU masks; AMP links Linux (Cortex-A) & Zephyr (Cortex-M) via OpenAMP."),
    ]
    for hi, bi, htxt, btxt in grid_21:
        set_text(s.shapes[hi], htxt, size_pt=11.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
        set_text(s.shapes[bi], btxt, size_pt=8.8, color_hex="595959", font_name=FONT_BODY)

    set_notes(s, (
        "[Speaker 4: Ali Mohamed] في الجدولة، Zephyr أوسع من مجرد Preemptive Round-Robin: "
        "عندنا أرقام سالبة Cooperative (-1, -2) متتقطعش أبدًا من أي Thread، وأرقام موجبة (>= 0) بتتقاطع فورًا، "
        "ولازم ننتبه: في Zephyr الرقم الأصغر هو الأولوية الأعلى! "
        "وفوقهم Meta-IRQ لأحداث الراديو بالمايكروثانية، وEDF بالمواعيد النهائية، وWorkqueues لتنفيذ شغل مقاطعة INT0 من محاضرة 6! "
        "والرابط العبقري مع المحاضرة الرابعة: فاكرين Timer CTC والمقارنة بـ OCR1A عشان نقلل عدد المقاطعات؟ ده بالضبط Tickless Idle: بدل ما تصحّي المعالج 1,000 مرة في الثانية بدون داعي، النواة تظبط الهاردوير يصحيه وقت المهمة الجاية وبس!"
    ))

    # =========================================================================
    # SLIDE 22: SMARTER INTERRUPTS — WHY NO *FromISR()? (Speaker 4 · TPL 19)
    # =========================================================================
    s = prs.slides[21]
    set_text(s.shapes[0], "Smarter Interrupts: Why Zephyr Has No *FromISR()", size_pt=19.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Unlike FreeRTOS's separate *FromISR() calls (which crash if missed), ARM's IPSR register makes Zephyr ISR-aware:", size_pt=9.2, align="center")

    set_text(s.shapes[4], "THREAD", size_pt=12.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "_isr_wrap", size_pt=11.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "1-Cycle\nIPSR", size_pt=10.5, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "Single Unified API: k_is_in_isr() via IPSR", size_pt=10.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[8],
        "FreeRTOS forces xQueueSend vs xQueueSendFromISR. Zephyr reads ARM IPSR in 1 cycle inside k_sem_give() / k_msgq_put()—one API works everywhere safely!",
        size_pt=8.2, color_hex="595959", font_name=FONT_BODY
    )

    set_text(s.shapes[9], "IRQ_CONNECT vs. Direct & Zero-Latency", size_pt=10.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[10],
        "IRQ_CONNECT uses _isr_wrapper for PendSV switches. Need pure C ISR? IRQ_DIRECT_CONNECT drops latency; IRQ_ZERO_LATENCY bypasses kernel locks!",
        size_pt=8.2, color_hex="595959", font_name=FONT_BODY
    )

    set_text(s.shapes[11], "Lec 4 & 6 Golden Rule: Defer via Workqueue", size_pt=10.5, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[12],
        "Following Lec 4 S22 & Lec 6 INT0: capture register fast, call k_work_submit(&w), and exit ISR cleanly to keep system jitter minimal.",
        size_pt=8.2, color_hex="595959", font_name=FONT_BODY
    )

    set_notes(s, (
        "[Speaker 4: Ali Mohamed] من أكبر أخطاء الكلاسيكيات في FreeRTOS إن لو مبرمج نسي يكتب xQueueSendFromISR جوه مقاطعة، النظام بيضرب Assertion Crash! "
        "ليه Zephyr معندوش *FromISR أصلًا؟ لأن معالجات ARM Cortex-M جواها مسجل هاردوير اسمه IPSR بيقولك في سيكل واحد أنت جوه مقاطعة ولا خيط عادي! "
        "فنفس الدالة k_sem_give وk_msgq_put بتشيك عبر k_is_in_isr وتتصرف صح! "
        "وغير IRQ_CONNECT العادي، لو عايز ISR عاري لـ Motor Control بدون تدخل النواة عندك IRQ_DIRECT_CONNECT وIRQ_ZERO_LATENCY، وطبعًا بنرحّل الشغل التقيل بـ Workqueues تطبيقا للمحاضرتين 4 و6!"
    ))

    # =========================================================================
    # SLIDE 23: DETERMINISTIC MEMORY & BANK POWER-OFF (Speaker 4 · TPL 6)
    # =========================================================================
    s = prs.slides[22]
    prepare_clean_header_slide(s, 1, "Deterministic Memory: k_mem_slab & Bank Power-Off", size_pt=19.8)
    add_lead_subtitle(s, 0.75, 0.94, 8.50, 0.30, "Turning Amr's Static-First & O(1) pool theory into Zephyr's hardware-saving allocators + physical SRAM power-gating:", size_pt=9.4)

    add_card(
        s, 0.75, 1.38, 4.15, 1.50,
        "1. O(1) Deterministic k_mem_slab",
        [
            ("Zero External Fragmentation: ", "Divides fixed SRAM chunks into identical aligned blocks (4–6 cycles alloc/free; ISR safe)."),
            ("Static-First Pools: ", "Compiler reserves .bss/.noinit footprint up front—build fails before field crash!"),
        ],
        dark=False, title_pt=10.1, body_pt=8.4
    )
    add_card(
        s, 5.10, 1.38, 4.15, 1.50,
        "2. sys_mem_blocks (Physical Bank Power-Off)",
        [
            ("SRAM Bank Contiguity: ", "Groups dynamic allocations inside lowest physical SRAM Bank 0 pages cleanly."),
            ("Hardware Energy Win: ", "Commands PM domains to cut leakage voltage from idle SRAM Bank 1 in sleep!"),
        ],
        dark=True, title_pt=10.1, body_pt=8.4
    )
    add_card(
        s, 0.75, 3.00, 4.15, 1.50,
        "3. net_buf Zero-Copy Headroom",
        [
            ("Protocol Stack Backbone: ", "Reserves leading headroom bytes so BLE, IPv6, CAN & MQTT prepend headers without memcpy()!"),
            ("OOM Immunity: ", "Reference-counted pools prevent a rogue packet flood from starving sensor loops."),
        ],
        dark=False, soft=True, title_pt=10.1, body_pt=8.4
    )
    add_card(
        s, 5.10, 3.00, 4.15, 1.50,
        "4. Flexible Heap Policy + k_heap",
        [
            ("Isolated Sub-Heaps: ", "Separate k_heap per subsystem (GUI vs. Network) stops memory leaks spreading across safety tasks."),
            ("Default Embedded Discipline: ", "Set CONFIG_HEAP_MEM_POOL_SIZE=0 on critical control loops for strict determinism."),
        ],
        dark=False, title_pt=10.1, body_pt=8.4
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "SILICON EDGE", "Zephyr treats SRAM as powered silicon banks (sys_mem_blocks power-off), not an endless desktop heap!", font_size=8.4)

    set_notes(s, (
        "[Speaker 4: Ali Mohamed] دلوقتي هنشوف إزاي Zephyr نفّذ نظرية الذاكرة اللي عمرو شرحها بس على مستوى الهاردوير: "
        "عندنا k_mem_slab لحجز كتل متساوية في زمن O(1) ثابت 4–6 سايكل بدون تفتت، وحتى net_buf بتسيب Headroom قدام الداتا عشان الـ MQTT والبلوتوث يضيفوا الترويسات من غير نسخ متعب! "
        "لكن قمة الذكاء السيليكوني هي أداة sys_mem_blocks: دي مبتحجزش عشوائي؛ بتلم البيانات كلها في أول بنك ذاكرة (SRAM Bank 0) عشان تسيب Bank 1 فاضي تمامًا، وتبعت للسيليكون يطفي الكهرباء عنه وقت النوم فيقلل تسريب الرام!"
    ))

    # =========================================================================
    # SLIDE 24: IPC TOOLBOX, k_fifo TRAP & ZBUS (Speaker 4 · TPL 17)
    # =========================================================================
    s = prs.slides[23]
    set_text(s.shapes[0], "IPC Toolbox, The k_fifo Stack Trap & Zbus", size_pt=20.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)

    add_clean_table(
        s, 0.75, 1.30, 4.55, 3.16,
        headers=["Zephyr IPC", "Behavior & Best Course Use"],
        rows=[
            ["k_sem",     "Counter signal (NO ownership); wakes Thread from HW ISR"],
            ["k_mutex",   "Ownership + Priority Inheritance (solves Mars Pathfinder!)"],
            ["k_msgq",    "Copies fixed data (e.g. ADC u16); Stack-safe!"],
            ["k_fifo",    "Zero-Copy linked pointers—needs 4B void* headroom!"],
            ["Zbus",      "Typed Pub/Sub Bus: decouples Temp sensor -> BLE/MQTT/UI"],
        ],
        col_widths=[1.08, 3.47], header_size=8.8, body_size=8.0
    )
    add_code_panel(
        s, 5.48, 1.30, 3.77, 2.04,
        "dangling_fifo_bug.c  (CRASHES!)",
        "struct pkt { void *fifo_reserved; int val; };\n"
        "void bad_send(struct k_fifo *q) {\n"
        "  struct pkt item = { .val = 42 };\n"
        "  k_fifo_put(q, &item); // Links Stack!\n"
        "} // <-- Function exits! Stack WIPED!\n"
        "  // Fix: allocate via k_mem_slab_alloc()!",
        font_size=7.25
    )
    add_chain_box(
        s, 5.48, 3.46, 3.77, 1.00,
        "Zbus Architecture Cleanliness",
        "Publishers push 1 message -> Zbus notifies BLE, MQTT & OLED automatically with 0 tight coupling!",
        dark=False
    )
    add_pill_banner(s, 0.75, 4.65, 8.50, 0.40, "RULE OF THUMB", "Small copy from Stack -> k_msgq.  Zero-Copy large frame -> k_mem_slab + k_fifo.  Many listeners -> Zbus!", font_size=8.3)

    set_notes(s, (
        "[Speaker 4: Ali Mohamed] في تواصل الخيوط: في غلطة مشهورة بتوقع مهندسين كتير لازم نحذر منها بالجدول والكود ده: "
        "k_msgq بتنسخ القيمة memcpy فآمنة جدًا تبعتلها متغير محلي من الـ Stack. "
        "أما k_fifo فـ Zero-Copy بياخد أول 4 بايت (void *fifo_reserved) يربط بيهم العقد؛ فلو عملت متغير محلي جوه الدالة وبعتّ عنوانه لـ k_fifo_put، أول ما الدالة ترجع الـ Stack بيتمسح والبرنامج ينهار! الحل: احجزه دايمًا من k_mem_slab. "
        "وفوقهم كلهم ناقل Zbus بنظام Publish/Subscribe اللي بيفصل الحساس عن الشاشة والبلوتوث والـ MQTT تمامًا!"
    ))

    # =========================================================================
    # SLIDE 25: HW MEMORY ISOLATION — USER MODE & k_mem_domain + HAND-OFF TO 5
    # =========================================================================
    s = prs.slides[24]
    prepare_clean_header_slide(s, 1, "Hardware Isolation: User Mode Threads & k_mem_domain", size_pt=19.4)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Stopping a buggy network parser from crashing safety loops—using ARM/RISC-V HW MPU firewalls:", size_pt=9.2)

    # Left: Kernel vs SVC vs User Mode diagram
    add_lane_block(s, 0.75, 1.28, 4.35, 0.34, "Privileged Kernel Space (MSP / Kernel Objects / NVIC)", dark=True, font_pt=8.4)
    add_card(
        s, 0.75, 1.68, 4.35, 0.86,
        "Kernel + Critical Motor / Safety Control",
        "Full Flash & SRAM hardware access; validates all user handles & pointers.",
        dark=False, soft=True, title_pt=9.6, body_pt=8.1
    )
    add_lane_block(s, 0.75, 2.60, 4.35, 0.36, "Hardware SVC Syscall Gate  (Z_OOPS Checks Args!)", dark=False, dashed=True, font_pt=8.2)
    add_card(
        s, 0.75, 3.02, 4.35, 1.46,
        "Unprivileged User Mode Thread (K_USER)",
        [
            ("• Assigned k_mem_domain: ", "MPU unlocks ONLY this thread's Partition & Stack—writing elsewhere causes instant MemManage HW fault!"),
            ("• Bug Containment: ", "Rogue MQTT / BLE buffer overflow kills ONLY that sandboxed thread!"),
        ],
        dark=True, title_pt=9.6, body_pt=8.15
    )

    # Right: Code + Containment Benefit
    add_code_panel(
        s, 5.28, 1.28, 3.97, 2.04,
        "Sandbox App Partition (CONFIG_USERSPACE=y)",
        "K_APPMEM_PARTITION_DEFINE(net_part);\n"
        "K_APP_BMEM(net_part) uint8_t pkt_buf[512];\n"
        "struct k_mem_domain net_dom;\n\n"
        "void enter_sandbox(void) {\n"
        "  k_mem_domain_init(&net_dom, 1, &parts);\n"
        "  k_thread_user_mode_enter(mqtt_fn, ...);\n"
        "} // HW MPU drops CPU privilege cleanly!",
        font_size=7.25
    )
    add_card(
        s, 5.28, 3.42, 3.97, 1.06,
        "Bare-Metal / Default FreeRTOS Contrast",
        "Classic monolithic loops leave EVERY library in Ring-0 privilege—a single-byte overflow corrupts PWM registers or kernel TCBs!",
        dark=False, soft=True, title_pt=9.7, body_pt=8.15
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "HAND-OFF -> SPEAKER 5", "Ahmed Hassan now closes the loop: All 16 Class Protocols & SIL 3, Midterm Q2 Auto-PM, Lec 7 Code, native_sim & Trade-Offs ->", font_size=8.0)

    set_notes(s, (
        "[Speaker 4: Ali Mohamed] آخر سلايد معايا بيورينا ميزة عتادية نادرة في عالم الـ RTOS: العزل العتادي بالـ MPU (CONFIG_USERSPACE)! "
        "في FreeRTOS العادي أو الـ Bare-Metal، كل الكود شغال بصلاحية كاملة على كل الرام؛ لو مكتبة بلوتوث أو MQTT حصل فيها Pointer Overflow، تكتب فوق متغيرات التحكم والمحركات وتوقع النظام! "
        "في Zephyr نقدر ننزل أي خيط شبكة لوضع Unprivileged User Mode جوه k_mem_domain: "
        "الـ MPU بيقفل على الخيط ده منطقته بس، وأي نداء للكيرنل يعدي من بوابة SVC Syscall بتشيك كل مؤشر. لو الخيط ده اتضرب، بيتعزل لوحده وبقية الجهاز يفضل شغال بأمان! "
        "ودلوقتي الباشمهندس أحمد حسن هيجمع لنا كل بروتوكولات الدفعة والـ SIL 3، أتمتة سؤال الميدتيرم Q2، وكود المحاضرة السابعة والديمو الحي!"
    ))

    # =========================================================================
    # SLIDE 26: PROTOCOL UNIFIER, TF-M TRUSTZONE, MCUBOOT & SIL 3 (Speaker 5 · TPL 7)
    # =========================================================================
    s = prs.slides[25]
    set_text(s.shapes[0], "Protocol Unifier, Security (TF-M) & SIL 3 Safety", size_pt=20.6, bold=True, color_hex="383838", font_name=FONT_HEAD)

    set_text(s.shapes[2], "Every Class Protocol Built-In", size_pt=11.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[4],
        "BLE 5.x (Lec 4 Pacemaker!), IPv4/6, MQTT (Team 6), CoAP (14), Matter (16), LoRa (15), CAN, LIN (v4.5), Modbus (13), MIPI (7) & LittleFS.",
        size_pt=8.8, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[1], "TrustZone (TF-M) & MCUboot OTA", size_pt=11.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[3],
        "ARM TrustZone TF-M/PSA Crypto isolates Root keys in Secure world. MCUboot adds ECDSA-signed Slot A/B OTA updates with instant atomic rollback.",
        size_pt=8.8, color_hex="595959", font_name=FONT_BODY
    )
    set_text(s.shapes[6], "IEC 61508 SIL 3 Safety & SPDX 3.0", size_pt=11.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[5],
        "IEC 61508 SIL 3 Written Concept Approval via Route 3s & SEooC (~15K LOC core subset, target 2027+), ISO 26262 follow-on & west spdx 3.0 compliance.",
        size_pt=8.8, color_hex="595959", font_name=FONT_BODY
    )

    set_notes(s, (
        "[Speaker 5: Ahmed Hassan] شكرًا يا علي. السلايد ده بيجسد قوة المنظومة المتكاملة فوق النواة في Zephyr: "
        "أول عمود فوق: كل البروتوكولات اللي باقي التيمات في دفعتنا شرحتها أو هتشرحها — MQTT بتاع تيم 6 النهاردة، بلوتوث مثال Pacemaker في محاضرة 4، CAN، وLIN الجديد في v4.5، وModbus، وCoAP، وLoRa، وMatter، وMIPI — كلها مدمجة رسميًا في نفس المستودع! "
        "تحت شمال: الأمان السيليكوني عبر TrustZone TF-M لعزل المفاتيح، وMCUboot للتحديثات الهوائية OTA بتوقيع ECDSA وتراجع تلقائي A/B لو التحديث فشل. "
        "تحت يمين: مسار اعتماد السلامة الوظيفية IEC 61508 SIL 3 بموافقة مفهوم رسمية عبر Route 3s بأسلوب SEooC (~15 ألف سطر من النواة) لـ 2027+ ومعاه west spdx!"
    ))

    # =========================================================================
    # SLIDE 27: AUTOMATIC POWER MANAGEMENT — MIDTERM Q2 (Speaker 5 · TPL 19)
    # =========================================================================
    s = prs.slides[26]
    set_text(s.shapes[0], "Automatic Power Management: Midterm Q2 Automated", size_pt=19.8, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(s.shapes[4], "WFI", size_pt=13.5, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[5], "STOP", size_pt=13, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    set_text(s.shapes[6], "DEEP", size_pt=12, bold=True, color_hex="FFFFFF", font_name=FONT_HEAD, align="center")

    set_text(s.shapes[7], "System PM Policy Manager (Lec 2 & Midterm)", size_pt=10.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[8],
        "Automates Midterm Q2 (Power-save vs Extended Standby)! Checks Tickless sleep + DTS min-residency-us to pick deepest safe pm_state with 32.768 kHz LSE.",
        size_pt=8.2, color_hex="595959", font_name=FONT_BODY
    )

    set_text(s.shapes[9], "Device Runtime PM (Auto Clock Gating)", size_pt=10.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[10],
        "Drivers track ref-counts via pm_device_runtime_get/put(): when no thread uses I2C, ADC or UART, Zephyr clock-gates that peripheral automatically!",
        size_pt=8.2, color_hex="595959", font_name=FONT_BODY
    )

    set_text(s.shapes[11], "Nested Power Domains & Wake Sources", size_pt=10.6, bold=True, color_hex="383838", font_name=FONT_HEAD)
    set_text(
        s.shapes[12],
        "Cuts SoC peripheral power rails & unused SRAM Bank 1 cleanly when child devices suspend, waking via unified pin/radio/RTC wake sources.",
        size_pt=8.2, color_hex="595959", font_name=FONT_BODY
    )

    set_notes(s, (
        "[Speaker 5: Ahmed Hassan] فاكرين سؤال الميدتيرم Q2 اللي حسبنا فيه الـ 30 والـ 61 ملي وات واختَرنا فيه بين Power-save وExtended Standby بإيدينا؟ "
        "في Zephyr، القرار ده بيتاخد أوتوماتيكيًا مئات المرات في الثانية بدون سطر كود يدوي في التطبيق! "
        "عند كل لحظة خمول في Tickless Idle، الـ System PM Policy بيشوف أقرب ميعاد صحيان وبيقارنه بـ exit-latency-us وmin-residency-us المكتوبين في Devicetree، ويختار أعمق وضع آمن مع الحفاظ على الكريستالة 32.768 kHz. "
        "وبالتوازي، Device Runtime PM بيطفي الساعة عن أي UART أو ADC أو I2C فاضي بـ reference counting، والـ Power Domains تفصل باور القطاعات — فالعقدة تعيش سنين على بطارية صغيرة!"
    ))

    # =========================================================================
    # SLIDE 28: COURSE UPGRADE SHOWDOWN — LEC 7 HVAC (Speaker 5 · TPL 6)
    # =========================================================================
    s = prs.slides[27]
    prepare_clean_header_slide(s, 1, "Course Upgrade: Lec 7 HVAC (ATmega32 -> Zephyr C)", size_pt=20.2)

    add_code_panel(
        s, 0.75, 1.12, 4.15, 3.40,
        "atmega32_hvac_lec7.c  (Single-Rate & Locked)",
        "// Manual: N=256 -> TCNT1=34285 for 0.5s\n"
        "ISR(ADC_vect) {\n"
        "  uint16_t raw = ADC;\n"
        "  if (raw > THRES) PORTB &= ~(1<<HEATER);\n"
        "  else             PORTB |=  (1<<HEATER);\n"
        "}\n"
        "int main(void) {\n"
        "  ADCSRA = (1<<ADEN)|(1<<ADATE)|(1<<ADIE);\n"
        "  SFIOR |= (5<<ADTS0); // Timer1 Compare B\n"
        "  TCCR1B |= (1<<CS12); // Fixed ATmega32 regs\n"
        "  sei(); while(1) { /* Stuck on 1 chip! */ }\n"
        "} // Adding Lec 7 S62 2s Anemometer breaks loop!",
        font_size=7.05
    )

    add_code_panel(
        s, 5.10, 1.12, 4.15, 3.40,
        "zephyr_hvac.c  (Portable Across 1,000+ Boards)",
        "static const struct adc_dt_spec temp =\n"
        "    ADC_DT_SPEC_GET(DT_PATH(zephyr_user));\n"
        "static const struct gpio_dt_spec heater =\n"
        "    GPIO_DT_SPEC_GET(DT_ALIAS(heater0), gpios);\n"
        "void hvac_task(void *a, void *b, void *c) {\n"
        "  int16_t buf; struct adc_sequence seq={...};\n"
        "  while (1) {\n"
        "    adc_read_dt(&temp, &seq); // V-Table\n"
        "    gpio_pin_set_dt(&heater, buf <= THRES);\n"
        "    k_msleep(500); // Auto Tickless + PM Policy!\n"
        "  } // Add 2nd K_THREAD_DEFINE(2000ms) -> 0 conflict!\n"
        "}\n"
        "K_THREAD_DEFINE(hvac, 1024, hvac_task, 0,0,0, 5,0,0);",
        font_size=6.85
    )

    add_pill_banner(s, 0.75, 4.66, 8.50, 0.38, "COURSE UPGRADE", "Same HVAC logic from Lec 7—now handles Slide 62's 2s Anemometer, sleep & STM32/nRF/ESP32 ports with 0 C changes!", font_size=8.2)

    set_notes(s, (
        "[Speaker 5: Ahmed Hassan] دي السلايد اللي بتترجم المنهج كله في مقارنة كود واحدة: مثال الـ HVAC من المحاضرة السابعة! "
        "على الشمال كود ATmega32 من المحاضرة: اضطرينا نحسب يدويًا N=256 وTCNT1=34285 عشان 0.5 ثانية ونربط SFIOR بـ ADATE؛ "
        "ده كود ممتاز لفهم السيليكون، بس مقفول على ATmega32 ولو ضفنا واجب شريحة 62 (حساس سرعة رياح كل ثانيةين) التايمرات والـ MUX بيتصادموا! "
        "على اليمين نفس المنطق في Zephyr: بيقرا من Devicetree بـ ADC_DT_SPEC_GET وبيشغل السخان بـ GPIO_DT_SPEC_GET، "
        "وينام k_msleep(500) فالسيليكون يدخل Tickless وPM Policy لوحده، ونضيف خيط تاني للرياح بـ K_THREAD_DEFINE من غير أي تعارض، ويتنقل على 1,000 بوردة بعلم -b!"
    ))

    # =========================================================================
    # SLIDE 29: EMBEDDED RTOS DECISION MATRIX (Speaker 5 · TPL 12)
    # =========================================================================
    s = prs.slides[28]
    prepare_clean_header_slide(s, 0, "Embedded RTOS Matrix: Bare-Metal vs. FreeRTOS vs. Zephyr", size_pt=19.6)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Objective engineering trade-off across compute footprint, hardware portability, driver ecosystem & safety assurance:", size_pt=9.2)

    add_clean_table(
        s, 0.75, 1.26, 8.50, 3.24,
        headers=["Engineering Axis", "Bare-Metal (ATmega32)", "FreeRTOS", "Zephyr RTOS v4.5"],
        rows=[
            ["Kernel Footprint",     "< 2 KB Flash / < 1 KB RAM", "~9 KB Flash / ~3 KB RAM", "~8–18 KB kernel (~195 vs ~220 KB full IoT)"],
            ["Hardware Portability", "Zero (locked to chip regs)", "Medium (manual vendor HAL)", "High (Devicetree .dts + 1,000+ boards)"],
            ["Driver & Config",      "Ad-hoc register C code",     "Vendor SDK per chip family", "Unified V-Tables + Kconfig pruning"],
            ["Interrupt APIs",       "Manual ISR & volatile",      "Dual Task / *FromISR() APIs", "Single unified API via 1-cycle IPSR"],
            ["Memory & Power",       "Manual sleep / 1 stack",     "heap_4 O(N) / basic tickless", "O(1) Slab + Bank-off + PM Policy"],
            ["Built-In Ecosystem",   "Write protocols manually",   "External lwIP/NimBLE stitch", "BLE, MQTT, CAN, LIN, Matter & TF-M"],
            ["Best Project Fit",     "Labs, < 8 KB / simple loops", "Single fixed chip & tight RAM", "Connected IoT, OTA, Multi-Board, SIL 3"],
        ],
        col_widths=[1.55, 2.05, 2.25, 2.65], header_size=8.8, body_size=7.9
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "ENGINEERING BALANCE", "FreeRTOS stays unbeatable for < 16 KB Flash static nodes; Zephyr dominates connected multi-protocol embedded products!", font_size=8.2)

    set_notes(s, (
        "[Speaker 5: Ahmed Hassan] دي مصفوفة المقارنة الموضوعية قدام أي لجنة تحكيم أو شركة: "
        "الـ Bare-Metal ممتاز للتعليم وللأنظمة الصغيرة جدًا زي مشاريع ATmega32 اللي فيها مهمة واحدة ورام 2 كيلوبايت. "
        "FreeRTOS خفيف في النواة الصافية (~9KB) ومثالي لما تكون مثبت على شريحة واحدة ومعاك HAL الشركة بس من غير تعقيد شبكات وOTA. "
        "أما Zephyr v4.5 فبيكسب لما تبني منتج متصل حقيقي: نقل بين +1,000 بوردة بـ Devicetree، درايفرات موحدة بـ 0 بايت رام، إلغاء *FromISR بـ IPSR، "
        "إطفاء بنوك الـ SRAM، وPM Policy وبروتوكولات مدمجة مع MCUboot وTF-M ومسار SIL 3!"
    ))

    # =========================================================================
    # SLIDE 30: BOOT FIRMWARE ON YOUR PC — native_sim & SHELL DEMO (Speaker 5 · TPL 4)
    # =========================================================================
    s = prs.slides[29]
    pic = find_pic(s)
    if pic:
        replace_pic_with_cropped(s, pic, IMG_LAB)

    s.shapes[1].left = Inches(0.72)
    s.shapes[1].top = Inches(0.91)
    s.shapes[1].width = Inches(4.85)
    s.shapes[1].height = Inches(0.74)
    set_text(s.shapes[1], "Boot Firmware on Your PC:\nnative_sim, Twister & Shell", size_pt=19.2, bold=True, color_hex="383838", font_name=FONT_HEAD, align="left")

    set_text(
        s.shapes[0],
        [
            [("Zero-Hardware Execution (native_sim): ", True, "383838"), ("Compiles your exact Zephyr C, Devicetree & Zbus as a native PC executable via ", False, "595959"), ("west build -b native_sim", True, "383838"), ("!", False, "595959")],
            [("Live Course Demo Commands: ", True, "383838"), ("Launch ", False, "595959"), ("kernel threads", True, "383838"), (", ", False, "595959"), ("device list", True, "383838"), (" & ", False, "595959"), ("pm stats", True, "383838"), (" live with zero MCU flashing!", False, "595959")],
            [("Industrial Telemetry & CI: ", True, "383838"), ("Non-blocking Deferred Logging + Flash Coredumps alongside ", False, "595959"), ("west twister", True, "383838"), (" automated PR suites.", False, "595959")],
        ],
        size_pt=8.7, font_name=FONT_BODY, align="left", space_after_pt=6, line_spacing=1.12
    )
    s.shapes[0].left = Inches(0.72)
    s.shapes[0].top = Inches(1.74)
    s.shapes[0].width = Inches(4.85)
    s.shapes[0].height = Inches(2.04)

    add_code_panel(
        s, 0.72, 3.86, 4.85, 1.18,
        "Live Laptop Demo Session (native_sim + Interactive Shell)",
        "$ west build -b native_sim samples/subsys/shell/shell_module && west build -t run\n"
        "uart:~$ kernel threads    # Shows ID, Stack Watermark & Priority!\n"
        "uart:~$ device list       # Proves Devicetree .rodata drivers are READY!",
        font_size=7.0
    )

    set_notes(s, (
        "[Speaker 5: Ahmed Hassan] ومش محتاجين نستنى البوردة توصل من الشحن عشان نجرب! "
        "بهدف اسمه native_sim، Zephyr بياخد كود الـ C والـ Devicetree والـ Zbus بتوعك ويترجمهم لملف تنفيذي يشتغل لايف على اللابتوب بالمايكروثانية والـ BLE والـ Shell التفاعلي! "
        "أمر واحد: west build -b native_sim ثم west build -t run، وتكتب قدام الدكتور أوامر زي kernel threads تشوف المهام واستهلاك المكدس الحقيقي، "
        "وdevice list تتأكد إن درايفرات الـ Devicetree شغالة READY، غير الـ Deferred Logging والـ Flash Coredumps اللي بتنقذ الـ debugging في الميدان!"
    ))

    # =========================================================================
    # SLIDE 31: HONEST TRADE-OFFS & EMBEDDED LINUX / AMP BOUNDARY (Speaker 5 · TPL 6)
    # =========================================================================
    s = prs.slides[30]
    prepare_clean_header_slide(s, 1, "Honest Trade-Offs & Embedded Linux Boundary", size_pt=20.2)
    add_lead_subtitle(s, 0.75, 0.90, 8.50, 0.28, "Senior embedded engineering means acknowledging platform costs cleanly—instead of blindly selling hype:", size_pt=9.2)

    add_card(
        s, 0.75, 1.30, 2.72, 1.56,
        "1. Steep Learning Curve",
        "Mastering DTS overlays, Kconfig symbols & west takes 2–4 weeks compared to FreeRTOS's 3-file start.",
        dark=False, title_pt=9.7, body_pt=8.2
    )
    add_card(
        s, 3.60, 1.30, 2.72, 1.56,
        "2. Minimal Kernel Size",
        "Stripped kernel (~8–18 KB) is heavier than minimal FreeRTOS (~9 KB) or tiny < 4 KB Bare-Metal loops.",
        dark=False, title_pt=9.7, body_pt=8.2
    )
    add_card(
        s, 0.75, 2.98, 2.72, 1.56,
        "3. AVR 2 KB RAM Ceiling",
        "ATmega32 stays too tight! Zephyr realistically shines on >= 16–32 KB SRAM Cortex-M or RISC-V nodes.",
        dark=False, soft=True, title_pt=9.7, body_pt=8.2
    )
    add_card(
        s, 3.60, 2.98, 2.72, 1.56,
        "4. Fast Ecosystem Cadence",
        "Brand-new silicon may arrive in vendor SDK weeks earlier; rapid 4-month releases make LTS vital.",
        dark=False, soft=True, title_pt=9.7, body_pt=8.2
    )

    add_card(
        s, 6.48, 1.30, 2.77, 3.24,
        "Embedded Linux + AMP Boundary",
        [
            ("Switch to Embedded Linux: ", "If the node has an MMU, > 32 MB RAM, USB cameras, OpenCV AI models, HDMI UIs, or Docker."),
            ("Power Duo (AMP / OpenAMP): ", "Modern SoCs run Embedded Linux on Cortex-A + Zephyr on companion Cortex-M for hard < 10 µs real-time I/O!"),
        ],
        dark=True, title_pt=10.1, body_pt=8.25
    )
    add_pill_banner(s, 0.75, 4.64, 8.50, 0.38, "OBJECTIVITY", "ATmega32 (2 KB RAM) -> Bare-Metal.  Single-chip static -> FreeRTOS.  Connected MCU -> Zephyr.  AI/MMU (>32 MB) -> Linux (+ Zephyr AMP)!", font_size=8.0)

    set_notes(s, (
        "[Speaker 5: Ahmed Hassan] وقبل الختام، المهندس الحقيقي لازم يكون أمين مع القيود والعيوب: Zephyr مش عصا سحرية لكل حالة! "
        "عندنا 4 تحديات صريحة: 1) منحنى التعلم في Devicetree وKconfig محتاج وقت أطول في البداية من مجرد ملفين FreeRTOS؛ "
        "2) حجم النواة الأساسية (~8–18 KB) أكبر من Bare-Metal؛ 3) شريحة ATmega32 برام 2 كيلوبايت مكانها Bare-Metal — Zephyr بيتألق من اول 16–32 KB رام؛ "
        "4) سرعة تحديثاته بتخلي الـ LTS ضروري للإنتاج. "
        "ولو المشروع محتاج MMU وكاميرات OpenCV ورام فوق 32 ميجا: روح لـ Embedded Linux على Cortex-A، وتقدر تجمع بينه وبين Zephyr على نفس الشريحة بـ OpenAMP!"
    ))

    # =========================================================================
    # SLIDE 32: SHARED CLOSER — ENGINEERING VERDICT & Q&A READY (TPL 20)
    # =========================================================================
    s = prs.slides[31]
    s.shapes[0].left = Inches(0.75)
    s.shapes[0].top = Inches(0.44)
    s.shapes[0].width = Inches(8.50)
    s.shapes[0].height = Inches(0.64)
    s.shapes[0].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(s.shapes[0], "Engineering Verdict — Thank You & Q&A", size_pt=22.0, bold=True, color_hex="383838", font_name=FONT_HEAD, align="center")
    for sh in [sh for i, sh in enumerate(s.shapes) if i != 0]:
        remove_shape(sh)
    for lsh in list(s.slide_layout.shapes):
        if lsh.has_text_frame and "CREDITS" in lsh.text_frame.text.upper():
            lsh._element.getparent().remove(lsh._element)

    add_verdict_card(
        s, 0.75, 1.32, 4.10, 2.98,
        "Bare-Metal / FreeRTOS When...",
        [
            ("8-bit / Tiny Flash", "ATmega32 (2 KB RAM) or < 64 KB Flash with 1 simple loop."),
            ("Single Fixed HAL", "Deeply tied to STM32Cube or vendor SDK with zero swap planned."),
            ("Standalone Control", "Simple motor/PSU loop without BLE, IP stacks or signed OTA."),
            ("Hours to First Blink", "Small team needing instant startup over long-term portability."),
        ],
        dark=False
    )
    add_verdict_card(
        s, 5.15, 1.32, 4.10, 2.98,
        "Choose Zephyr RTOS When...",
        [
            ("Connected Embedded IoT", "Need BLE, MQTT, CoAP, Matter, CAN, LIN & LittleFS tested together."),
            ("Chip Portability", "Swap ST, Nordic & ESP32 via Devicetree (-b board) without C rewrites."),
            ("Security & SIL 3", "Need MPU Domains, TF-M TrustZone, A/B OTA rollback or safety certs."),
            ("Ultra-Low Power & CI", "Automatic PM Policy, Bank-off RAM & PC testing via native_sim."),
        ],
        dark=True
    )
    add_pill_banner(
        s, 0.75, 4.52, 8.50, 0.50,
        "TEAM 8 · THANK YOU",
        "Bare-Metal (CSE 421/422) builds silicon mastery; Zephyr RTOS scales it globally. Ready for Dr. Ahmed & Class Q&A!",
        font_size=8.4
    )

    set_notes(s, (
        "[Shared Closing — Speaker 5: Ahmed Hassan on behalf of Team 8] الخلاصة الهندسية في جملة واحدة: "
        "دراسة الـ Bare-Metal مع الدكتور أحمد على ATmega32 هي الأساس اللي بيفهمك فيزياء السيليكون؛ "
        "استخدم Bare-Metal أو FreeRTOS للشرائح الضيقة الثابتة، واختار Zephyr RTOS أول ما تبني منتج Embedded متصل "
        "محتاج توحيد بروتوكولات، إدارة طاقة أوتوماتيكية، عزل عتادي بالـ MPU، تحديثات OTA، وقابلية نقل على +1,000 بوردة. "
        "شكرًا جزيلًا لحضراتكم — وتيم 8 بالكامل (عمر عبدالسلام، عمرو، عمر مشهور، علي، وأحمد) جاهز لأي سؤال من الدكتور والزملاء!"
    ))

    prs.save(out_path)
    print(f"SAVED: {out_path} ({os.path.getsize(out_path)} bytes, {len(prs.slides)} slides)")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else OUT_PPTX
    build(target)
