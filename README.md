# Zephyr RTOS vs. FreeRTOS — Embedded Systems Architecture Presentation

A 34-slide technical presentation comparing **FreeRTOS micro-kernel internals** against **Zephyr RTOS's full embedded OS architecture**, complete with clean conceptual diagrams, native PowerPoint animations/transitions, official upstream documentation references, and detailed bilingual (Egyptian Arabic + English) Speaker Notes on every slide.

## Files Included

| Path | Description |
|---|---|
| [`Zephyr_RTOS_vs_FreeRTOS.pptx`](./Zephyr_RTOS_vs_FreeRTOS.pptx) | Main 34-slide animated PowerPoint deck (native transitions, choreography & speaker notes) |
| [`Zephyr_RTOS_vs_FreeRTOS.pdf`](./Zephyr_RTOS_vs_FreeRTOS.pdf) | Portable PDF backup for instant browser preview or projector fallback |
| [`build.py`](./build.py) | One-command builder (rebuilds skeleton, slides, animations, validation & optional PDF/PNG QA) |
| [`src/`](./src/) | Editable Python slide generator (`stage1_skeleton.py`, `stage2_build_deck.py`, `stage2_populate_utils.py`) & PowerShell animation/export scripts (`add_pro_animations.ps1`, `export_qa.ps1`) |
| [`assets/`](./assets/) | Clean base template ([`template_base.pptx`](./assets/template_base.pptx)), 34-slide skeleton ([`template_skeleton_34.pptx`](./assets/template_skeleton_34.pptx)) and images ([`assets/images/`](./assets/images/)) |
| [`fonts/`](./fonts/) | Typography pack (`Schibsted Grotesk`, `Figtree`, `JetBrains Mono`) for pixel-accurate rendering |
| [`requirements.txt`](./requirements.txt) | Python dependencies (`python-pptx`, `Pillow`, `lxml`, `defusedxml`) |
| [`.agents/`](./.agents/) | Installed `pptx` & `pdf` AI skills and OOXML schema/validation helpers |

## Editing & Rebuilding with Python ("Vibe Coding" Workflow)

Any collaborator can edit slide content, layouts, diagrams, or speaker notes directly in Python and rebuild in one command:

```bash
pip install -r requirements.txt
python build.py
```

Useful build flags:
- `python build.py --out preview.pptx --no-anim` — fast build without opening PowerPoint (works on Windows/macOS/Linux)
- `python build.py --skeleton` — regenerate `assets/template_skeleton_34.pptx` from `assets/template_base.pptx`
- `python build.py --pdf --qa` — rebuild PPTX, apply animations, export PDF backup & render 1600×900 slide PNGs (Windows PowerPoint required for COM animations/PDF/PNG export)

## Team Setup (Recommended — 10 Seconds)

To make sure PowerPoint renders the exact layout and spacing on every teammate's laptop:
1. Open the [`fonts/`](./fonts/) folder.
2. Select all `.ttf` files → **Right-click → Install** (Windows) or double-click to install in Font Book (macOS).
3. Restart PowerPoint and open [`Zephyr_RTOS_vs_FreeRTOS.pptx`](./Zephyr_RTOS_vs_FreeRTOS.pptx) (press **`F5`** to preview all animations & transitions).

## Presentation Outline (34 Slides)

### Opening & Roadmap (Slides 1–2)
1. **Cover:** Zephyr RTOS vs. FreeRTOS — From Real-Time Kernel Internals to a Full Embedded OS
2. **Presentation Roadmap:** 6-stage connected narrative overview

### Part I — RTOS & FreeRTOS Under the Hood (Slides 3–12)
3. **Why an RTOS?** Super-Loop (`while(1)`) sequential stalls vs. independent Threads & non-blocking waits
4. **Inside Every Thread:** Task Control Block (Identity Card), Private Stack (Thread Workspace) & 4 Lifecycle Modes (`READY`, `RUNNING`, `BLOCKED`, `SUSPENDED`)
5. **How Multiple Threads Share One CPU Core:** Preemptive Priority vs. Round-Robin 1 ms time-slicing timelines
6. **Inside Every 1 ms Tick:** 4-step SysTick pipeline (`Tick IRQ` → `Timers` → `Unblock Ready` → `PendSV Context Switch`)
7. **Blocking Discipline:** Three core timeout strategies (`Wait Forever`, `Bounded Wait`, `No Wait`)
8. **Coordinating Threads:** Choosing between `Message Queue`, `Binary Semaphore`, and `Mutex Lock`
9. **The Priority Inversion Trap & How Mutexes Fix It:** Visual step-by-step trace of Priority Inversion vs. Mutex Priority Inheritance
10. **Thread <-> Interrupt Sharing (Critical Section, Spinlock & `BASEPRI`):** Selective ARM interrupt masking (`0–4` live vs. `5–15` paused) + 1-core vs. multi-core Spinlock link
11. **Hardware Interrupts & `*FromISR()`:** Why Vector Table direct jumps require IRQ-safe APIs & `portYIELD_FROM_ISR()`
12. **RAM Allocation & Fragmentation Trap:** Day 1 contiguous RAM vs. Day 90 Swiss-cheese 4 KB holes (`heap_1..5` vs. Static API)

### Architectural Pivot & Part II — Zephyr RTOS 16 Core Pillars (Slides 13–34)
13. **Section 02 — Enter Zephyr RTOS:** Breaking out of FreeRTOS's manual glue-code wall
14. **Full "Batteries-Included" OS:** 1,000+ boards, 13+ CPU architectures, Linux Foundation & Apache 2.0
15. **Step 1 · Devicetree & `prj.conf`:** `.dtsi` + `.dts` + `app.overlay` + `prj.conf` → `devicetree_generated.h` & CMake driver filtering
16. **Step 2 · Unified Drivers:** How `main.c` reaches any chip (`DT macro` → `struct device` → Flash API table → hardware registers)
17. **Does Devicetree Waste RAM? (`.rodata` & Footprint):** Zero runtime RAM overhead + Flash/RAM comparison chart
18. **Before `main()` Starts (`SYS_INIT`):** 5-stage linker-sorted boot sequence (`EARLY` → `PRE_KERNEL_1/2` → `POST_KERNEL` → `APPLICATION`)
19. **Kconfig (`prj.conf`), CMake & `west`:** Auto-dependency solving, dead-code pruning & `--sysbuild`
20. **Rich Scheduling & SMP:** Cooperative (`< 0`), Preemptive (`>= 0`), EDF, `k_work`, Tickless Idle & CPU masks
21. **Why No `*FromISR()`:** `_isr_wrapper`, 1-cycle `IPSR` check, `IRQ_DIRECT_CONNECT` & `IRQ_ZERO_LATENCY`
22. **Deterministic Memory & Power:** `k_mem_slab` (`O(1)` intrusive list) vs. `sys_mem_blocks` (SRAM Bank 1 power-off bitmap)
23. **IPC Toolbox & `k_fifo` Stack Trap:** `k_msgq` vs. `k_fifo`/`k_lifo` (4-byte reserved pointer & dangling Stack bug), `k_pipe`, `k_poll`
24. **Zbus & Multi-Core AMP:** Pub/Sub software bus + `ipc_service` / OpenAMP / RPMsg
25. **MPU User Mode & Memory Domains:** Hardware stack guards, `k_mem_domain`, `k_object_access_grant()` & SVC syscalls
26. **TrustZone (TF-M), Signed OTA & Safety Certs:** ARM TrustZone, MCUboot atomic A/B rollback, IEC 61508 SIL 3 / ISO 26262 & SBOM
27. **Automatic Power Management:** PM Policy Manager (`pm_state`), Device Runtime PM & nested Power Domains
28. **Built-In Connectivity & Storage Stacks:** BLE Host+Controller, IPv4/IPv6, Matter, Thread, LittleFS & MCUmgr
29. **Boot Firmware on Your PC:** `native_sim`, `west twister`, Zephyr Shell, Deferred Logging, CTF & Coredump
30. **Code Showdown:** `freertos_stm32.c` vs. portable `zephyr_portable.c`
31. **Where Zephyr Hurts:** Honest trade-offs (learning curve, macro debugging, workspace size & opinionation)
32. **FreeRTOS vs. Zephyr:** Full 8-layer architecture comparison matrix
33. **References & Official Technical Documentation:** Verified FreeRTOS, ARM Cortex-M4, Zephyr, MCUboot & TF-M sources
34. **Engineering Verdict — Thank You & Q&A:** Constraint-driven decision guide
