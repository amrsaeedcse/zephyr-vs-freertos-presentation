# Zephyr RTOS vs. FreeRTOS — Embedded Systems Architecture Presentation

A 32-slide technical presentation comparing **FreeRTOS micro-kernel internals** against **Zephyr RTOS's full embedded OS architecture**, complete with native PowerPoint animations/transitions, official upstream documentation references, and detailed bilingual (Egyptian Arabic + English) Speaker Notes on every slide.

## Files Included

| File | Description |
|---|---|
| [`Zephyr_RTOS_vs_FreeRTOS.pptx`](./Zephyr_RTOS_vs_FreeRTOS.pptx) | Main 32-slide animated PowerPoint deck (native transitions, choreography & speaker notes) |
| [`Zephyr_RTOS_vs_FreeRTOS.pdf`](./Zephyr_RTOS_vs_FreeRTOS.pdf) | Portable PDF backup for instant browser preview or projector fallback |
| [`fonts/`](./fonts/) | Typography pack (`Schibsted Grotesk`, `Figtree`, `JetBrains Mono`) for pixel-accurate rendering |

## Team Setup (Recommended — 10 Seconds)

To make sure PowerPoint renders the exact layout and spacing on every teammate's laptop:
1. Open the [`fonts/`](./fonts/) folder.
2. Select all `.ttf` files → **Right-click → Install** (Windows) or double-click to install in Font Book (macOS).
3. Restart PowerPoint and open [`Zephyr_RTOS_vs_FreeRTOS.pptx`](./Zephyr_RTOS_vs_FreeRTOS.pptx) (press **`F5`** to preview all animations & transitions).

## Presentation Outline (32 Slides)

### Opening & Roadmap (Slides 1–2)
1. **Cover:** Zephyr RTOS vs. FreeRTOS — From Micro-Kernel Internals to a Full Embedded OS
2. **Architecture Roadmap:** 6-stage overview

### Part I — RTOS & FreeRTOS Under the Hood (Slides 3–10)
3. **Why an RTOS?** Super-Loop (`while(1)`) stall trap vs. preemptive determinism
4. **Task Anatomy in RAM:** TCB (`pxTopOfStack`, priority, lists), Private Stack (HW/SW frames) & 4 lifecycle states
5. **Inside Every 1 ms SysTick:** `xTickCount++`, Delayed List wake-up, `PendSV`, Preemptive Priority & Round-Robin
6. **Blocking Discipline:** Three production rules for `xTicksToWait` (`0`, `pdMS_TO_TICKS(100)`, `portMAX_DELAY`)
7. **Sync Primitives & Priority Inversion:** Queues (`memcpy`), Binary Semaphores, Mutex ownership & Priority Inheritance
8. **Critical Sections (`BASEPRI`):** Selective ARM Cortex-M masking (`5–15` masked vs. `0–4` unmasked)
9. **Vector Table Reality:** Why FreeRTOS requires `*FromISR()` and manual `portYIELD_FROM_ISR()`
10. **FreeRTOS Memory Reality:** `heap_1..5` + Static API & external Swiss-cheese fragmentation without an MMU

### Architectural Pivot & Part II — Zephyr RTOS 16 Core Pillars (Slides 11–32)
11. **Section 02 — Enter Zephyr RTOS:** Breaking out of FreeRTOS's "Glue Code Wall"
12. **Batteries-Included Embedded OS:** 1,000+ boards, 13+ CPU architectures, Linux Foundation & Apache 2.0
13. **Devicetree Hardware Decoupling:** 3-layer cascade (`.dtsi` → `.dts` → `app.overlay`) & zero-edit MCU swapping
14. **Zero Bytes of RAM (`.rodata`) & Footprint:** Compile-time Python DTS binding + real Flash/RAM bar chart
15. **Unified Driver Model:** Flash function-pointer v-tables (`gpio_dt_spec → struct device → api`) & `static inline`
16. **Automatic Boot Bootstrap (`SYS_INIT`):** Ordered `EARLY` → `PRE_KERNEL_1/2` → `POST_KERNEL` → `APPLICATION`
17. **Kconfig (`prj.conf`), CMake & `west`:** Auto-dependency solving, dead-code pruning & `--sysbuild`
18. **Rich Scheduling & SMP:** Cooperative (`< 0`), Preemptive (`>= 0`), EDF, `k_work`, Tickless Idle & CPU masks
19. **Smarter Interrupts:** `_isr_wrapper`, 1-cycle `IPSR` check (no `*FromISR()`), `IRQ_DIRECT_CONNECT` & `IRQ_ZERO_LATENCY`
20. **Deterministic Memory:** `k_mem_slab` (`O(1)` intrusive list) vs. `sys_mem_blocks` (Bank 0 bitmap for SRAM Bank 1 power-off)
21. **IPC Toolbox & `k_fifo` Stack Trap:** `k_msgq` vs. `k_fifo`/`k_lifo` (4-byte reserved pointer & dangling Stack bug), `k_pipe`, `k_poll`
22. **Zbus & Multi-Core AMP:** Pub/Sub software bus + `ipc_service` / OpenAMP / RPMsg
23. **MPU User Mode & Memory Domains:** Hardware stack guards, `k_mem_domain`, `k_object_access_grant()` & SVC syscalls
24. **Security, OTA & Functional Safety:** ARM TrustZone (TF-M), MCUboot atomic A/B rollback, IEC 61508 SIL 3 / ISO 26262 & SBOM
25. **Automatic Power Management:** PM Policy Manager (`pm_state`), Device Runtime PM & nested Power Domains
26. **Production Connectivity & Storage Stacks:** BLE Host+Controller, IPv4/IPv6, Matter, Thread, LittleFS & MCUmgr
27. **Host Simulation & Observability:** `native_sim`, `west twister`, Zephyr Shell, Deferred Logging, CTF & Coredump
28. **Side-by-Side Code Showdown:** `freertos_stm32.c` vs. portable `zephyr_portable.c`
29. **Honest Engineering Trade-Offs:** Learning curve, macro debugging, workspace footprint & framework opinionation
30. **Full Architecture Comparison Matrix:** 8-layer side-by-side comparison
31. **References & Official Technical Documentation:** Verified FreeRTOS, ARM Cortex-M4, Zephyr, MCUboot & TF-M sources
32. **Engineering Verdict — Thank You & Q&A:** Constraint-driven decision guide
