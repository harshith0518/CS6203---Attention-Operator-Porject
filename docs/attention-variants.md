# The attention ladder — six variants, one new problem each

- **Use**: source for the mentor deck and report §2. **Not** a backlog (only fusion is built).
- **Claim**: these are not six features but a ladder — each rung adds exactly one new
  *kernel-engineering* problem, and each problem costs differently in CUDA C++ than in cuTile.
- Notation: batch `B`, query heads `H`, **KV heads** `H_kv` (heads sharing one key/value pair),
  query length `N`, key/value length `M`, head dim `D`, group `G = H/H_kv`, window `W`.
  Per-head FLOPs = `4·B·H·N·M·D` (halve if causal).

## 0. The table to show first

| # | Variant | The one new problem | Dominant resource | CUDA C++ cost | cuTile cost |
|---|---|---|---|---|---|
| 1 | Vanilla SDPA | why fusion is needed | **HBM O(N²d)** | low | low |
| 2 | Causal mask | mask **and skip whole tiles** | FLOPs halve, load imbalance | medium | low |
| 3 | **Fused flash + online softmax** | streaming softmax, live accumulator rescale | registers / shared memory | **high** | medium |
| 4 | Multi-head + GQA/MQA | head broadcast; decode is memory-bound | HBM (KV reuse) | medium | **very low** |
| 5 | Sliding-window / local | sparse tile traversal | FLOPs ∝ N·W | medium-high | low-medium |
| 6 | Cross-attention (N≠M) | asymmetric shapes, two tilings | FLOPs | low-medium | low |

## 1. Vanilla SDPA

- **SDPA** = scaled dot-product attention: `S = QKᵀ/√D`, `P = softmax(S)`, `O = PV`.
- Naive = 3 kernels, two materialised N×M tensors.

| | |
|---|---|
| Traffic | Q,K,V,O are `O(N·D)`; **S and P are `O(N·M)`** — written and re-read ≈ 4× |
| New problem | none mechanically — it is the reference that makes fusion *measurable*, not asserted |
| CUDA C++ | gemm → row-softmax kernel (block reduction, two passes: max then sum) → gemm |
| cuTile | 3 `@ct.kernel`s; softmax = load row tile → `ct.max(axis=1)` → `ct.exp2` → `ct.sum(axis=1)` |
| Reveals | **arithmetic intensity** = FLOPs per byte moved. At `D=64`: ~64 FLOP/B vs this card's **ridge point** (peak FLOPs ÷ peak BW ≈ 36 FLOP/B). Looks compute-bound on paper; the 4 round-trips put it on the memory side. Show both numbers, explain the gap. |

## 2. Causal

- `S_ij = -1e30 for j > i` (decoder-style: a token may not attend to the future).

| | |
|---|---|
| New problem | a mask is trivial; **not doing the masked work is not**. Separates "correct" from "fast" more cleanly than any other rung |
| Effect | FLOPs ≈ halve, but a rectangular grid over a triangular domain does all the work (~50% efficient); per-block early exit makes block runtimes wildly unequal |
| CUDA C++ | inner `if (col <= row)` for the diagonal tile + loop bound `k <= q_tile_end`. Cheap in code; occupancy and load balance become *yours* |
| cuTile | `ct.where(cols[None,:] <= rows[:,None], s, -1e30)` — broadcast gives the `[BQ,BK]` mask free — plus a dynamic `range` bound. The balance problem still exists, you just didn't write the arithmetic |
| Numerics trap | `-inf` is correct but `exp2(-inf − -inf) = NaN`. Use fp32 `-1e30`; in fp16 it overflows to `-inf` |

## 3. Fused flash + online softmax

- Recurrence and code: [`PLAN.md`](PLAN.md) §2.1–2.2.
- **flash attention** = fused exact attention whose N×M score matrix never reaches HBM.
- **online softmax** (Milakov & Gimelshein 2018) = streaming softmax with a running max/denominator;
  FlashAttention's contribution (Dao 2022) is fusing it into the PV product.

| | |
|---|---|
| New problem | the softmax is a *global* row reduction but you only hold a slice of the row ⇒ it must be streaming and idempotent under rescaling: every accumulated output is retroactively corrected by `α` when the max moves |
| Traffic | `O(N·D + M·D)` instead of `O(N·M)`. That is the whole result |
| CUDA C++ cost | becomes a memory-hierarchy project, not an arithmetic one: smem double-buffering + `cp.async`, `O` in ~64 registers/thread with `α` applied by hand, matched `mma.sync` fragment layouts, swizzled smem on the K transpose. Each is a place to be 3× off the roofline |
| cuTile cost | the same algorithm as chained tile ops — `acc` is a tile, not a register allocation. **Fusion is free; tile shapes are the work** |
| Reveals | the abstraction does not remove the work — it **moves it from the memory hierarchy to the type system** (report thesis) |

## 4. Multi-head, then GQA / MQA

- **MQA** = multi-query attention (`H_kv = 1`, Shazeer 2019); **GQA** = grouped-query (Shazeer's
  generalisation, Ainslie 2023). Standard in Llama, Mistral, Qwen, Gemma.

| | |
|---|---|
| New problem | one KV head serves `G` query heads ⇒ (a) index arithmetic, (b) **L2 reuse** — the `G` sharing heads should be scheduled together so K,V come from cache, not HBM `G` times, (c) the **decode regime**: `N=1` makes it a long skinny GEMV, ~0% of peak FLOPs |
| CUDA C++ | `int kv_head = q_head / G;` + hand-written strides + *your* decision whether to fuse `G` heads per block so the KV tile loads into smem once |
| cuTile | stack the group into the leading tile dim; `ct.mma` broadcasts the 2-D KV tile across a 3-D query tile (documented 3-D×2-D broadcast). Group sizes are powers of two ⇒ legal by construction |
| Verdict | **one index expression vs one broadcast** — the single most persuasive slide in the deck |
| Roofline flip | decode is memory-bound: report **GB/s**, not TFLOP/s. The one place naive and fused converge again (no N² tensor to avoid) |

## 5. Sliding-window / local

- `j ∈ [i−W, i]` (Longformer, Beltagy 2020; BigBird; Mistral's `sliding_window`). FLOPs → `O(N·W·D)`.

| | |
|---|---|
| New problem | the grid is a **band**, not a rectangle or triangle: bounds arithmetic + a two-sided mask; first/last blocks fall partly outside |
| CUDA C++ | `k_start = max(0,(q_begin−W)/BK)`, `k_end = min(n, (q_end+BK−1)/BK)`, two-sided `if`. Tuning question is real: smaller `BK` = finer skip vs larger `BK` = better mma utilisation |
| cuTile | the same two bounds as a dynamic `range()`, mask is `cols >= rows−W` ∧ `cols <= rows`. Barely longer than causal, and the two integers still feed `ct.tune.exhaustive_search` |
| Why include | tests the abstraction on *irregular* work — if the advantage evaporates here, that is an informative negative result |

## 6. Cross-attention

- Decoder queries attend to encoder keys: `Q: [N×D]`, `K,V: [M×D]`, `N ≠ M` (Whisper, T5).

| | |
|---|---|
| New problem | two different sequence dims ⇒ two tilings, and `N` is often 1 — which annihilates tiles (`BQ=64` computing one row wastes 63/64) |
| CUDA C++ | free: a 1-D block of 128 threads handles a row naturally. Asymmetry costs nothing at thread level |
| cuTile | honest cost — tile waste at `N=1`. Workaround: batch requests into the leading dim, which is what real serving kernels do |
| Why include | **the first rung where CUDA C++ wins.** A comparison with no such case is an advertisement, not an analysis |

## 7. Deliberately off the ladder

- **Paged / vLLM attention** — block-table indirection for the KV cache: a memory-management
  problem, not an abstraction one. Needs `ct.gather`.
- **MLA** (DeepSeek) — compressed latent KV; research-grade, doubles the algebra.
- **Linear / kernel attention** (Performer) — different asymptotics, not the same operator family.
- **Block-sparse** (SparTA) — rung 5 already tests irregular traversal; this adds a metadata
  format, not a concept.
- **Backward pass** — 5 more GEMMs, no new abstraction insight.

## 8. Sources

| Topic | Source |
|---|---|
| IO-aware exact attention; fusion; IO complexity Θ(N²d²/M) | Dao et al. — *FlashAttention* (2022) |
| Better work partitioning, non-matmul-FLOP argument | Dao — *FlashAttention-2* (2023) |
| **Online softmax** (the actual algorithmic idea, predates FA) | Milakov & Gimelshein (2018) |
| MQA | Shazeer — *Fast Transformer Decoding* (2019) |
| GQA | Ainslie et al. (2023) |
| Sliding window | Beltagy (2020) Longformer; Zaheer (2020) BigBird |
| Why plain FA is still not bandwidth-optimal | Dao et al. — *FlashAttention-3* (2024) |
| cuTile | NVIDIA cuTile Python docs · `NVIDIA/cutile-python` · **`NVIDIA/TileGym`** (reference transformer kernels) |

- `TileGym` is the closest existing analog. **Read its attention kernel before writing yours** —
  it is the intended reference and saves ~2 days. Then state in the report where you diverged.

## 9. Predictions — write these down *before* measuring

1. **Fusion**: ~3× at `N=1024–2048`; **crossover at N≈512**; above it HBM traffic, not
   arithmetic, sets latency.
2. **Abstraction cost**: fused cuTile trails fused CUDA C++ by <25% at `N=1024` for ~40% of the
   lines. If it trails by more, *which tile shape or hint the compiler gets wrong* is worth
   more than the speedup.
3. **Debuggability**: cuTile converts wrong-answer bugs into compile errors — diagnosis faster
   for masking bugs, slower for anything silently mis-lowered (no SASS story to tell).
4. **Portability**: on CC 8.6 two-thirds of cuTile's arch-specific tuning surface is inert, so
   the tile-shape sweep carries *all* of its tunability — and will be less shape-sensitive than
   a hand-tuned kernel is to block size, because the compiler re-derives the mapping.

## 10. Appendix — deck outline for the mentor meeting

Asked for on 14 Oct: *"pick 5-6 kernels, understand how they are expressed in CUDA C++ vs
cuTile, present slides."* Two jobs only: show you know the variants; show the same computation
written two ways. Nothing about your own optimisations. **14–16 slides, 12–15 min.**

| # | Slide | Content |
|---|---|---|
| 1 | Title | course, name, date |
| 2 | Framing | *"each variant adds exactly one new problem, and each costs differently"* + §0 table |
| 3 | What attention computes | the 3 equations, FLOPs `4BHMND` |
| 4 | Naive is memory-bound | N×M materialised twice; traffic vs ridge point |
| 5 | V1 vanilla | 3 kernels vs 3 `@ct.kernel`s |
| 6 | V2 causal | mask trivial, *skipping tiles* is not |
| 7 | V3 algorithm | online-softmax recurrence, `α` highlighted; cite Milakov + Dao |
| 8 | **V3 both implementations** | **centrepiece**: smem/`cp.async`/registers/hand-`α` vs the ~20-line cuTile kernel |
| 9 | V3 what fusion buys | HBM bytes `O(N²d)`→`O(Nd)`; predicted crossover N≈512 |
| 10 | **V4 GQA** | `q_head/G` by hand vs one `ct.mma` broadcast |
| 11 | V4 decode | `N=1` ⇒ ~0% FLOPs, ~100% bandwidth; roofline flips |
| 12 | V5–V6 | one line each; cross-attention = **where CUDA C++ wins** |
| 13 | Knob inventory | plan §4 axis 2 with the **"inert on sm_86"** column |
| 14 | Plan & scope | **"OP2 only, both models, no GEMM"** + 9-day schedule + what is cut. Ask: does 07 Nov hold? |
| 15 | Backup: API | `tiled_view`, `ct.mma`, `ct.max/sum/exp2/where`, `exhaustive_search`, `ByTarget` |
| 16 | Backup: constraints | needs driver ≥580 / CTK 13.2 / CC 8.x; you are on 4 GB Ampere; the fallback |

- **Say out loud**: "I predict crossover at N≈512 and will measure it" · "window + cross-attention
  are cut if the schedule slips; the fused kernel in both models is what must exist" · "on my
  card two-thirds of cuTile's tuning surface is inert".
- **Ask**: confirm OP2-only in writing · does 07 Nov hold?
- **Bring**: PDF (not a live notebook), one printed page of the plan, **nothing that runs**.
