# PLAN.md — technical spec

State, rules and traps live in [`AGENTS.md`](AGENTS.md). This file: what to build, how to
measure, what to write.

## 1. Scope — one operator, two kernels, three files

- Brief defines OP2 as *"any variant — e.g. scaled dot-product, causal, or fused/flash-style"*.
- ⇒ **one fused kernel per model**, `causal` as a runtime flag.

| File | What | ~Lines |
|---|---|---|
| `cutile/attn/flash.py` | fused attention, tile level | 60–80 |
| `cuda/attn/flash.cu` | same operator, hand-tiled | 180–250 |
| `bench/bench.py` | correctness + timing + CSV + figures | 130–180 |

- **No naive kernel.** PyTorch's `SDPBackend.MATH` *is* the materialised baseline — same card,
  same harness, zero work.
- **Not implemented** (written about, not coded): GQA beyond the stretch goal, sliding window,
  cross-attention, paged KV cache, backward pass. The survey is reading material, not a backlog.
- **Hard gate (Gate B)** — standard MHA is "done" only when: both kernels match fp64 PyTorch
  (`max|out−ref| ≤ 5e-3·max|ref|`, no NaN/Inf) on every size in §3, incl. `N=1000` and `N=1`;
  every config is a row in `results/results.csv`; both have an `ncu` profile. **Only then**:
  GQA, extra variants, or tuning beyond the §3 sweep.

**The operator — pin it down, do not drift:**

- **MHA** (multi-head attention, `H_kv = H`, no grouping), **self-attention** (`N = M`),
  **forward pass only**, and **exact** — FlashAttention is not an approximation.
- `Q,K,V : [B,H,N,D]` fp16 contiguous → `O : [B,H,N,D]` fp16. `D = 64`; `B=1, H=8`.
- `O = softmax(QKᵀ/√D + mask)·V`, fp32 accumulate, one fused launch, S never in HBM.
- `causal` flag: skip whole KV tiles, mask only the diagonal tile; handle `N % BK ≠ 0` (tail).
- **Out of scope**: dropout, RoPE/ALiBi/relative bias (positions are baked into Q,K *before* the
  kernel), KV cache, backward pass. Attention starts after the caller's linear projections.

## 2. Kernel design

### 2.1 The algorithm — identical in both models

Per query tile, streaming over KV tiles:

```
m = -1e30 (finite), l = 0, O = 0
for each KV tile k:
    S_k    = Q_tile · K_kᵀ · scale        # scale = log2(e)/√D
    m_new  = max(m, rowmax(S_k))          # running max
    P_k    = exp2(S_k − m_new)
    α      = exp2(m − m_new)              # rescale factor
    l      = l·α + rowsum(P_k)            # running denominator
    O      = O·α + P_k · V_k              # rescale BEFORE accumulate
    m      = m_new
O = O / l
```

- **online softmax** = softmax computed streaming: you hold only a slice of each row, so the
  running **max** `m` and denominator `l` are corrected as tiles arrive. A new larger max must
  retroactively rescale everything already accumulated — that is what `α` does.
- **Two numerics decisions** (report's numerics section): fold `log2(e)/√D` into the scale ⇒
  base-2 `exp2`, one instruction; start `m` at finite `-1e30`, since `-inf` makes the first
  `α = exp2(-inf − -inf) = NaN`. In fp16 `-1e30` overflows ⇒ keep `m`/masks **fp32**.
- **Build order**: match the fp64 reference on the **non-causal** path first, then enable
  `causal`. One moving part at a time — the mask is the second most common source of silent
  wrong answers after the tail.

### 2.2 cuTile — `cutile/attn/flash.py`

```python
@ct.kernel
def flash_fwd(Q, K, V, O, scale: ct.Constant[float], causal: ct.Constant[bool],
              BQ: ct.Constant[int], BK: ct.Constant[int], D: ct.Constant[int]):
    qi = ct.bid(0)                                    # this block's query tile index
    q  = Q.tiled_view((BQ, D)).load((qi, 0)).astype(ct.float16)
    m   = ct.full((BQ,), -1e30, ct.float32)
    l   = ct.zeros((BQ,), ct.float32)
    acc = ct.zeros((BQ, D), ct.float32)
    kv, vv = K.tiled_view((BK, D)), V.tiled_view((BK, D))
    n = kv.num_tiles(0) if not causal else (qi * BQ + BQ + BK - 1) // BK
    for k in range(n):
        kt = kv.load((k, 0))
        s  = ct.mma(q, ct.transpose(kt), ct.zeros((BQ, BK), ct.float32))
        if causal:
            rows = ct.arange(BQ) + qi * BQ
            cols = ct.arange(BK) + k * BK
            s = ct.where(cols[None, :] <= rows[:, None], s, -1e30)
        m_new = ct.maximum(m, ct.max(s, axis=1))
        p  = ct.exp2(s * scale - ct.expand_dims(m_new, 1))
        a  = ct.exp2(m - m_new)
        l  = l * a + ct.sum(p, axis=1)
        acc = ct.mma(p.astype(ct.float16), vv.load((k, 0)), acc * ct.expand_dims(a, 1))
        m  = m_new
    ct.store(O, (qi,), (acc / ct.expand_dims(l, 1)).astype(O.dtype))
```

- `ct.mma` = cuTile's tensor-core matrix-multiply-accumulate; `ct.transpose` flips K to Kᵀ.
- **Verify on first run, don't trust from memory**: `ct.transpose(x)` default axis pair, and
  that `ct.arange` + `[None, :]` broadcast yields a `[BQ, BK]` mask.

### 2.3 CUDA C++ — `cuda/attn/flash.cu`

128 threads, `BQ=BK=64`, `D=64`. All the difficulty is the memory hierarchy:

- `K`,`V` tiles in shared memory, **double-buffered with `cp.async`** (**`cp.async`** = sm_80+
  async global→shared copy; it overlaps load with compute).
- `O` accumulator in **registers** (`BQ/4 × D` floats/thread ≈ 64 regs), with **`α` applied by
  hand** to those accumulators before the second mma.
- `mma.sync.m16n8k16` (= warp-level tensor-core instruction, fp16 in / fp32 acc) for both GEMMs.
  Fragment layouts must match between QKᵀ and PV — the classic silent-wrong-answer bug.
- Shared memory padded to avoid **bank conflicts** (two threads hitting the same memory bank,
  which serialises the access).
- Causal: bound the k-loop, mask the diagonal tile.
- Host driver in the same file: allocate, verify once, then time.

```powershell
nvcc -arch=sm_86 -O3 --use_fast_math -lineinfo flash.cu -o flash.exe
```

## 3. Measurement

**The brief's three OP2 targets — mechanism → evidence:**

- **HBM traffic ↓ via fusion**: S never leaves registers/shared memory; evidence = `ncu` DRAM
  byte counters vs torch's `MATH` backend (the materialised baseline), figure 2.
- **Numerical stability**: fp32 accumulate, fp32 running max, `exp2`-folded scale, tail
  `ct.where`; evidence = `max_abs_err` vs fp64, zero NaN/Inf across the size matrix, **and a
  ×100 input-magnitude stress test** (a kernel with no max-subtraction dies here — that is the
  test proving the stabilisation is real, not decorative).
- **Latency at the target sequence length**: target **N = 2048** (prefill regime); evidence =
  p50/p90 at N=2048 plus the full latency-vs-N curve, figure 1.
- **Terminology**: the brief says HBM, but this laptop's 4 GB is **GDDR6** (6001 MHz DDR →
  12 Gbps, 128-bit → 192 GB/s). Same measurement, correct name: **device/DRAM bytes**. Metrics:
  `dram__bytes_read.sum`, `dram__bytes_write.sum`,
  `dram__throughput.avg.pct_of_peak_sustained_elapsed` — confirm locally with
  `ncu --query-metrics | findstr dram__bytes`.
- **cuTile-specific profiling**: `ncu --section Tile` reports launch config, execution info and
  GPU resource usage for **Tile IR** kernels — the official analogue of Launch/Occupancy for the
  cuTile side (Nsight Compute Profiling Guide, *Sets and Sections*).

**Protocol — identical for both models:**

1. Warm up 50 iters (pays cuTile's JIT compile). **Compile time reported separately.**
2. **Flush L2** between timed iters — write a 32 MB buffer. Without it you measure cache
   bandwidth and fusion looks free.
3. Time 200 iters; report **p50** (median) and **p90** (90th percentile — the tail matters on
   a throttling laptop).
4. Clock discipline: AC power, *Prefer maximum performance*, 10 min warmup, GPU idle. Report
   the spread — a stated caveat is a strength.
5. Reference through the **same harness**: `torch.nn.functional.scaled_dot_product_attention`
   under each backend the wheel provides (`MATH`, `EFFICIENT_ATTENTION`, `FLASH_ATTENTION`).

**CSV schema — append only, ASCII order fixed:**

```
model, causal, dtype, B, H, N, D, BQ, BK, p10_us, p50_us, p90_us,
tflops, gbytes_s, ref_backend, ref_us, pct_of_ref, max_abs_err, notes
```

- `tflops = 4·B·H·N²·D / t` (halve if causal). Correctness = `max|out − ref_fp64|`, also in CSV.

**Sizes:** `B=1, H=8, D=64`, fp16, `N ∈ {256, 512, 1024, 2048, 4096}` + `N=1000`
(non-divisible ⇒ exercises the tail mask) + `N=1` (decode, if cheap).

- Expect: fusion barely helps at `N=512`; ~3× at `N≥1024`; at `N=8192` a materialised N×N
  matrix **OOMs** on 4 GB while the fused kernel runs. Capture that error — it is a result.

**Figures — generated by `bench.py`, never hand-drawn:**

1. latency vs N (CUDA, cuTile, 3 SDPA backends);
2. **HBM bytes vs N** (from `ncu`) — the O(N²)→O(N) collapse, *the* money figure;
3. achieved TFLOP/s vs N, with measured peak as a horizontal line;
4. tile-shape sweep, best config marked.

## 4. The comparison — five axes → five experiments

| Axis | Experiment | Report |
|---|---|---|
| **Dev velocity** | `cloc` per kernel; wall-clock edit→CSV row, logged each iteration | LOC ratio; median iteration time; compile time (nvcc vs tileiras JIT, cold+warm) |
| **Control & tunability** | Knob inventory, each marked *active on sm_86* | Tile shape/`occupancy`/`num_worker_warps` vs block size/smem layout/`cp.async` depth. **`allow_tma`, `num_ctas`, `mma_scaled` inert here** |
| **Portability** | Count `__CUDA_ARCH__` guards + arch intrinsics; show `ByTarget(...)` | On sm_86 the modern CUDA path (`wgmma`/TMA) is unreachable. Concede: cuTile's ceiling = what `tileiras` implements, no `--ptx` escape |
| **Perf vs reference** | The CSV | % of each SDPA backend, % of measured peak, with variance |
| **Debuggability** | Inject 3 identical bugs in both, time the fix: (a) missing sync/race, (b) mask sign, (c) rescale placement | Minutes-to-fix + which tool found it: `compute-sanitizer` (NVIDIA's race/mem checker), `cuda-gdb`, `ncu` SASS view vs `TileTypeError` traceback, `ct.print`, `CUDA_TILE_LOGS=TILEIR` |

- **Prediction to test**: cuTile converts wrong-answer bugs into compile errors — earlier, but
  less searchable.
- **Required output**: one paragraph, unambiguous, on when the tile abstraction's productivity
  gain outweighs the ceded control. **Take a position.**

## 5. Schedule

| Day | Date | Work | Done when |
|---|---|---|---|
| 1 | Fri 30 Oct | driver ≥580 → packages → **gate green** | `smoke_cutile.py` passes |
| 2 | Sat 31 Oct | `bench.py` + torch baselines | one command → CSV with torch rows |
| 3 | Sun 1 Nov | `cutile/attn/flash.py` (fused, causal, tail) | matches torch fp64 in fp16 tolerance |
| 4 | Mon 2 Nov | `cuda/attn/flash.cu` | same tolerance; both rows in CSV |
| 5 | Tue 3 Nov | tile/block sweep; figures 1 + 3 | best config recorded both models |
| 6 | Wed 4 Nov | `ncu` both kernels; figure 2; debuggability | `profiles/` populated |
| 7 | Thu 5 Nov | report draft — all five axes | no placeholders |
| 8 | Fri 6 Nov | report final, cleanup, one-command reproduction | stranger regenerates every figure |
| 9 | Sat 7 Nov | **submit**, buffer | submitted |

- Lost a day ⇒ cut the debuggability experiment, then figure 4. **Never** cut day 6 (profiles)
  or day 7 (report) — 40% of the grade.

## 6. Risks

| Risk | Mitigation |
|---|---|
| Driver can't reach 580 | decide day 1; fall back to cloud/lab GPU, laptop = dev box |
| Laptop throttling ruins repro | §3 protocol 4 + p50/p90 + stated variance |
| 4 GB VRAM caps N | sizes fixed in §3; the OOM datapoint is a result |
| Tail-mask bug (`exp2(0)=1` phantom key) | `N=1000` in the test matrix from day 2 |
| cuTile detail profiling needs driver ≥582.16 | install ≥582; else profile CUDA fully and state the gap |
