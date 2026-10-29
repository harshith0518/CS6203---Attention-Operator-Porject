# Decision log

One line per decision, dated. Cheap to keep, and it is the raw material for the
report's "development velocity" and "debuggability" sections — plus it is the
only way to answer "why did you do it that way?" in a viva three weeks later.

Format: `YYYY-MM-DD | decision | why | evidence/where`

> **Scope note (2026-10-29): OP2 (attention) only. GEMM is not part of this project.**
> Any 2026-10-28 entry that implies GEMM was in scope has been superseded below.

---

## 2026-10-28 — planning day

| Date | Decision | Why |
|---|---|---|
| 2026-10-28 | Deep-dive on **OP2 Attention** for the written comparison | Brief allows choosing one operator; attention has the richer trade-off surface. |
| 2026-10-29 | **OP2 only — GEMM cut entirely.** Supersedes the 10-28 "implement all four kernels" decision. Not in the kernels, not in the benchmarks, not in the report. | Mentor clarification: the project is the attention operator. |
| 2026-10-29 | Reference switches from **cuBLAS → `torch.nn.functional.scaled_dot_product_attention`** (math / mem-efficient / flash backends) | With GEMM gone, cuBLAS has no role. SDPA is a *better* reference: its three backends are themselves the naive-vs-fused spectrum, on the same card, measured by someone else. |
| 2026-10-29 | ~~GQA promoted SHOULD → MUST~~ — **superseded**: GQA is **stretch-only, behind Gate B** | Fusion in both models is the graded deliverable. GQA is the best *slide*, not a requirement. |
| 2026-10-29 | **Toolchain = native Windows PowerShell 5.1, not WSL2 Ubuntu** | WSL uses the *same* Windows driver, so it does not fix the 566→580 blocker. NVIDIA's CUDA-on-WSL guide confirms persistence mode and other `nvidia-smi` state features are unsupported there — so WSL does not even deliver clock locking. CUDA dev tools in WSL are described as "preview", and profiling is 15% of the grade. MSVC 2026 is already installed; WSL would add a distro + second toolkit under a 9-day deadline. |
| 2026-10-29 | Repo scaffold + initial commit on `main` | The repo existed with zero commits; a baseline timestamp is the raw evidence for the "development velocity" axis. `.gitignore` excludes `.venv/` and `results/raw/`, but **keeps `profiles/` tracked** — Nsight reports are graded evidence. |
| 2026-10-29 | Python **3.12** for the venv (not 3.13/3.14) | cuTile supports 3.10–3.14, but `cupy`/`torch` binary wheels lag on the newest interpreters, and the whole point is to be installing, not debugging wheels. |
| 2026-10-29 | **Minimal surface: 1 fused kernel per model, `causal` as a flag. No naive kernel. 3 code files total.** | Owner must be able to work on this manually without an agent. Brief says OP2 is "any variant"; PyTorch's `MATH` backend supplies the materialised baseline for free, so writing our own would cost a day to reproduce a number we already have. |
| 2026-10-29 | **`AGENTS.md` introduced as the single context file** | So any agent (or the owner after a week away) resumes without being re-briefed. §3 State is the only section that changes; the file is capped at ~200 lines because it is loaded on every task. |
| 2026-10-29 | Docs consolidated: slide outline merged into `attention-variants.md` §10; `env-verified.md` folded into `AGENTS.md` §4 and deleted | Fewer files to keep in sync. Every doc is now either loaded by an agent or read for one specific job. |
| 2026-10-29 | **Docs condensed to bullets + in-line jargon glosses** (rule now in `AGENTS.md` §5) | Owner is new to the domain and reads these cold; a reader who hits *occupancy* or *TileIR* with no gloss stops and loses the thread. Words cut, facts kept — `PLAN.md` went 441 → 150 lines with no lost content. |
| 2026-10-29 | ~~AGENTS.md cap set to 175 lines, other docs 155~~ — **metric was wrong**: `Measure-Object -Line` skips blank lines (reported 160 for a 199-line file), so the caps were never actually enforced. **Now capped in words** (AGENTS ≤ 1900, others ≤ 2000) with the counting command written into `AGENTS.md` §5. | A rule you can't measure is not a rule. Words are also the real context cost. |
| 2026-10-29 | **Brief's three OP2 targets mapped to mechanism + evidence** (`PLAN.md` §3) | HBM traffic ↓ via fusion → `ncu` DRAM counters vs torch `MATH`; numerical stability → fp32 accumulate + fp32 running max + tail `ct.where`, evidenced by `max_abs_err` vs fp64; latency at target **N=2048** → p50/p90 plus the full curve. Without this mapping the three targets are slogans. |
| 2026-10-29 | **Gate B: standard MHA must be complete in *both* models before any other variant** (owner's rule) | Finishing one operator end-to-end beats two half-finished ones, and the graded comparison only works if both kernels exist at equal quality. Checkable criteria are in `PLAN.md` §1 (`max|out−ref| ≤ 5e-3·max|ref|` on every size, every config in the CSV, both profiled). Exploration is the reward for finishing, never a parallel track. |
| 2026-10-28 | Six-variant attention ladder, "one new problem each" | Directly answers Pravin's 14 Oct ask; doubles as the report structure. |
| 2026-10-28 | **One** `bench/harness.py` for both models; results only via CSV | 35% perf + 15% profiling. Hand-copied numbers into slides is the classic way to lose marks. |
| 2026-10-28 | Treat **"cuTile's arch-specific knobs are inert on sm_86"** as a headline finding, not a limitation | `allow_tma` / `num_ctas` / `mma_scaled` all require sm_90+. It is concrete evidence for the portability axis and no one else in the cohort will have it. |
| 2026-10-28 | Hardware gate on day 0: driver ≥580 before any code | cuTile needs r580+; machine is on 566.07. Failing fast here is worth more than a week of coding on the wrong box. |
| 2026-10-28 | Scripts target **PowerShell 5.1** | No `pwsh` on this machine; `#requires -Version 7` would have silently blocked setup. |

## Template for later

```
| YYYY-MM-DD | what changed | why | measurement that justified it |
```

Ideas worth logging as they happen:
- every tile-shape / block-size sweep result that changed a default
- every bug whose diagnosis took > 20 min (feeds the debuggability experiment)
- every cuTile API surprise (signature, supported dtype, silent no-op)
- the iteration-time log: `edit -> number in results.csv`, in seconds

---

## Open items needing an answer (blocking the schedule)

- [ ] Did the first meeting with Pravin already happen?
- [ ] Is 07 Nov still the deadline after the 14 Oct restart?
- [ ] Is there code elsewhere, or is this a cold start?
- [ ] Any access to a newer-than-Ampere GPU (lab / cloud / Colab)?
