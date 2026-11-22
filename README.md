# Zephyr RTOS — Silicon-to-Cloud Embedded Operating System

A **32-slide (30-minute)** Embedded Systems presentation crafted for **Team 8 (Topic 8 · 24-11)** in **CSE 421/422 Embedded Systems** (Faculty of Engineering, Zagazig University — Dr. Ahmed Osman Mahmoud). Designed around **5 speakers × 6 core slides each (~1 min/slide = 6 min/speaker = 30 min)** plus a shared Hero Cover (`S1`) and Verdict/Q&A Closer (`S32`), anchored in the ATmega32 curriculum (`Lectures 1–10`, `Midterm`, `Sheets` & `Lab 3`), updated to **Zephyr v4.5 (Oct 2026)**, and complete with native PowerPoint animations/transitions and rich Egyptian Arabic + English Speaker Notes on all 32 slides.

## Team 8 Speakers (5 × 6-Slide Symmetric Split · 30 Min · 24-11)
- **Slide 1 (Shared Cover):** Hero Introduction (Omar Abdelsalam opens on behalf of Team 8)
1. **Slides 2–7 (6 Slides · ~6 min):** Omar Abdelsalam Abdelfattah (`عمر عبدالسلام عبدالفتاح`) — Embedded Roots, CSE 421/422 Breaking Points, Determinism, Concurrency Models & Golden Course Map
2. **Slides 8–13 (6 Slides · ~6 min):** Amr Abdelazim Saeed (`عمرو عبدالعظيم سعيد`) — Thread Execution & Lifecycle, SRAM Layout (`TCB/PSP/MSP/MPU`), Cortex-M `PendSV`, `BASEPRI` Latency, Priority Inversion (`Mars Pathfinder`) & Heap Fragmentation
3. **Slides 14–19 (6 Slides · ~6 min):** Omar Ahmed Abdelaziz Mashhour (`عمر أحمد عبدالعزيز مشهور`) — Zephyr `v4.5 (Oct 2026)`, OS Canvas, 3-Layer Devicetree, `0B RAM .rodata` Chart, `static inline` Flash V-Tables & 5-Stage `SYS_INIT`
4. **Slides 20–25 (6 Slides · ~6 min):** Ali Mohamed Ali (`علي محمد علي`) — `Kconfig`/`west` (`--sysbuild`/`spdx`), Tickless (`CTC`) Scheduling, `_isr_wrapper` (`IPSR`), `k_mem_slab`/`sys_mem_blocks` Bank-Off, `k_fifo` Trap/`Zbus` & User-Mode `k_mem_domain`
5. **Slides 26–31 (6 Slides · ~6 min):** Ahmed Hassan Mohamed Mostafa (`أحمد حسن محمد مصطفى`) — Protocol Unifier (`TF-M`/`MCUboot`/`SIL 3`), Midterm Q2 `PM Policy`, `Lec 7 HVAC` Code Upgrade, RTOS Matrix, `native_sim` Demo & Honest Trade-Offs
- **Slide 32 (Shared Closer):** Engineering Verdict — Decision Guide & Thank You / Q&A Ready

## Files Included

| Path | Description |
|---|---|
| [`Zephyr_RTOS_Embedded_Systems.pptx`](./Zephyr_RTOS_Embedded_Systems.pptx) | Main 32-slide animated PowerPoint deck (native transitions, choreography & speaker notes) |
| [`Zephyr_RTOS_Embedded_Systems.pdf`](./Zephyr_RTOS_Embedded_Systems.pdf) | Portable PDF backup for instant browser preview or projector fallback |
| [`build.py`](./build.py) | One-command builder (rebuilds skeleton, slides, animations, validation & optional PDF/PNG QA) |
| [`src/`](./src/) | Editable Python slide generator (`stage1_skeleton.py`, `stage2_build_deck.py`, `stage2_populate_utils.py`) & PowerShell animation/export scripts (`add_pro_animations.ps1`, `export_qa.ps1`) |
| [`assets/`](./assets/) | Clean base template ([`template_base.pptx`](./assets/template_base.pptx)), 32-slide skeleton ([`template_skeleton_32.pptx`](./assets/template_skeleton_32.pptx)) and images ([`assets/images/`](./assets/images/)) |
| [`fonts/`](./fonts/) | Typography pack (`Schibsted Grotesk`, `Figtree`, `JetBrains Mono`) for pixel-accurate rendering |
| [`requirements.txt`](./requirements.txt) | Python dependencies (`python-pptx`, `Pillow`, `lxml`, `defusedxml`) |
| [`.agents/`](./.agents/) | Installed `pptx` & `pdf` AI skills and OOXML schema/validation helpers |

## Editing & Rebuilding with Python

```bash
pip install -r requirements.txt
python build.py
```

Useful build flags:
- `python build.py --out preview.pptx --no-anim` — fast build without opening PowerPoint
- `python build.py --skeleton` — regenerate `assets/template_skeleton_32.pptx` from `assets/template_base.pptx`
- `python build.py --pdf --qa` — rebuild PPTX, apply animations, export PDF backup & render 1600×900 slide PNGs

## Team Font Setup (10 Seconds)
1. Open the [`fonts/`](./fonts/) folder.
2. Select all `.ttf` files → **Right-click → Install** (Windows) or double-click in Font Book (macOS).
3. Open [`Zephyr_RTOS_Embedded_Systems.pptx`](./Zephyr_RTOS_Embedded_Systems.pptx) and press **`F5`** to preview all animations & transitions.

## Presentation Outline (32 Slides)
1. **Cover:** Zephyr RTOS — Silicon-to-Cloud Embedded Operating System (`Team 8 · CSE 421/422`)
2. **Roadmap:** 6-stage Embedded Systems narrative (`5 Speakers × 6 Core Slides = 30 Min`)
3. **Pure Embedded Systems:** Bare-silicon physics (no MMU, `GPIOA->BSRR`, `µA` sleep) + Grand Protocol Unifier (`All 16 Class Protocols`)
4. **CSE 421/422 Breaking Points:** Lec 3 `18s GEN_PWM`, Lab 3 3-timer exhaustion, Lec 7 2-sensor HVAC / Lec 9 `DOR`, Midterm static power
5. **Determinism vs. Throughput:** Hard vs. Firm vs. Soft deadlines (Pacemaker / Ventilator) & worst-case bounded latency
6. **Concurrency Models:** Super-Loop polling vs. ISR-Heavy (`volatile` pitfalls, Lec 9 `message_ptr`) vs. RTOS isolated threads
7. **The Golden Course Map:** 8 direct bridges from ATmega32 (`Lec 1–10`) to Zephyr (`K_THREAD_DEFINE`, `Tickless/CTC`, `PM Policy`, `Workqueues`, `static inline`)
8. **Execution Illusion & Lifecycle:** Single-core Zephyr-priority scheduling (`0 = highest`) + `READY / RUNNING / BLOCKED / SUSPENDED` + `K_FOREVER` trap
9. **Inside Every Thread:** SRAM layout (`TCB`, `Private Stack`, `MPU Stack Guard`, `MSP/PSP`) & `K_THREAD_DEFINE` static-first allocation
10. **Hardware Context Switching:** From AVR `IVT/RETI` (`Lec 6`) to ARM Cortex-M `SysTick`, `z_swap()`, `PendSV (Priority 15)` & Lazy FPU
11. **Interrupt Latency & Critical Sections:** `BASEPRI` selective masking vs. global `cli()/PRIMASK` (`Lec 4 Slide 22` ISR discipline)
12. **Priority Inversion & Mars Pathfinder:** Live priority inversion timeline solved automatically via `k_mutex` Priority Inheritance (PIP)
13. **Day 1 vs. Day 90 Memory Reality:** Heap external fragmentation trap (`malloc(32 KB)` returns `NULL` despite `66 KB` total free RAM) vs. deterministic pools
14. **Enter Zephyr RTOS v4.5 (Oct 2026):** `18,075` commits, `1,230` contributors, `1,000+` boards, TriCore/Hexagon, `v3.7 LTS`, `21%` Eclipse, Ray-Ban Meta, Nordic, Garmin
15. **Full Embedded OS Canvas:** `Trims DOWN` vs `Builds UP`, `GitHub PR + Twister CI`, mapping Lec 2's 4-layer architecture
16. **Devicetree 3-Layer Decoupling:** `.dtsi` → `.dts` → `app.overlay` (zero C edits when swapping silicon)
17. **0 Bytes RAM for Devicetree:** Build-time `.rodata` compilation + Kernel & Full IoT Stack (`~195 vs ~220 KB`) footprint chart
18. **Unified Driver Model:** Solving `Lec 2 Slide 45` with Flash V-Tables & `static inline` zero-overhead polymorphism
19. **Before `main()` Starts (`SYS_INIT`):** 5 linker-sorted Flash boot stages (`EARLY → PRE_KERNEL_1/2 → POST_KERNEL → APPLICATION`)
20. **Kconfig (`prj.conf`), CMake & `west`:** Dead-code pruning + `--sysbuild` & `west spdx`
21. **Rich Scheduling & Tickless Idle:** `Meta-IRQ`, `Coop (< 0)`, `Preempt (>= 0)`, `EDF`, `Workqueues`, `Tickless (Lec 4 CTC)` & `SMP/AMP`
22. **Why No `*FromISR()`:** `_isr_wrapper`, 1-cycle `IPSR` (`k_is_in_isr()`), `IRQ_DIRECT_CONNECT` & `IRQ_ZERO_LATENCY`
23. **Deterministic Memory & Power:** `O(1)` `k_mem_slab` vs `sys_mem_blocks` SRAM Bank 1 power-off + `net_buf`
24. **IPC Toolbox, `k_fifo` Trap & `Zbus`:** `k_sem`, `k_mutex` (PIP), `k_msgq`, `k_fifo` dangling stack bug & `Zbus` Pub/Sub
25. **Hardware Memory Isolation:** Unprivileged User Mode threads, `k_mem_domain` & `SVC` syscall firewall
26. **Protocol Unifier, Security & Safety:** Every class protocol built-in + TrustZone (`TF-M`), MCUboot A/B OTA, `IEC 61508 SIL 3 (Route 3s / SEooC)` & `SPDX 3.0`
27. **Automatic Power Management:** Automating Midterm Q2 (`Power-save` / `Extended Standby`) via `PM Policy`, `Device Runtime PM` & Nested Power Domains
28. **Course Upgrade Showdown:** `Lec 7 HVAC` (`ATmega32 TCNT1=34285 / ADATE` → portable `Zephyr C` expandable to 2-sensor Slide 62 Assignment)
29. **Embedded RTOS Decision Matrix:** `Bare-Metal (ATmega32)` vs `FreeRTOS` vs `Zephyr v4.5`
30. **Boot Firmware on Your PC:** `native_sim`, `west twister`, Zephyr Shell (`kernel threads`, `device list`), Deferred Logging & Coredump
31. **Honest Trade-Offs & Linux Boundary:** 4 real costs (DTS/Kconfig curve, `~8–18 KB` kernel, RAM vs 2 KB AVR, vendor lag) + Embedded Linux / `AMP`
32. **Engineering Verdict — Thank You & Q&A:** Constraint-driven decision guide for Team 8 defense
