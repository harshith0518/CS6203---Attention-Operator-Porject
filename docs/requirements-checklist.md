# What the assignment demands — condensed checklist

Our own summary of the CS-6023 course brief (source: `agrani-cutile-attention.pdf`, kept local
and **not redistributed** — it contains the mentors' email addresses). Use this when you need
the requirements without the original document.

## Grading

| Criterion | Weight |
|---|---|
| Correctness (vs. reference) | 25% |
| Performance & optimization | 35% |
| Comparative analysis | 25% |
| Code quality & profiling evidence | 15% |

## OP2 — attention (the operator we chose)

Brief allows **any variant**: scaled dot-product, causal, or fused/flash-style.

| Required in CUDA C++ | Required in cuTile |
|---|---|
| Fused QKᵀ → softmax → PV kernel, hand-tiled | The same fusion as chained tile operations |
| Online (streaming) softmax to bound memory traffic | Let the compiler place intermediate tiles |
| Manual shared-memory reuse across the fused stages | Compare tile granularity against the hand-fused kernel |

**Optimize for:** HBM traffic reduction from fusion · numerical stability · latency at the target
sequence length.

## The written deliverable

- **Choose either OP1 (GEMM) or OP2 (attention).** We chose OP2.
- A short report — **figures preferred over prose** — comparing CUDA C++ and cuTile across:
  1. **development velocity** — lines of code, iteration time
  2. **control & tunability** — what each model lets you tune, and what it hides
  3. **portability across GPU generations**
  4. **achieved performance** relative to a reference implementation
  5. **debuggability** — profiling and error visibility at each level
- Must **state a clear recommendation**: *when does the tile abstraction's productivity gain
  outweigh the control ceded to the compiler?*

## Expected artifacts

- CUDA C++ kernels **and** cuTile kernels
- Correctness tests against a reference (cuBLAS / PyTorch)
- Nsight Compute profiles
- The comparison report

> Note the tension: the comparison section says "choose either OP1 or OP2", but the artifact line
> says "kernels for both operators". We are building **OP2 only**, per mentor clarification —
> confirm this in writing.

## Milestones (as printed)

| Milestone | Date |
|---|---|
| Proposal & environment setup | 10/07 |
| One implementation working | 10/14 |
| Both implementations working | 10/21 |
| Benchmarks + written comparison | 10/28 |
| Final code, report | 11/07 |

Term 1/2026 · course CS-6023 · mentors: Pravin Jagtap and Sumanta Datta (Agrani Labs)
