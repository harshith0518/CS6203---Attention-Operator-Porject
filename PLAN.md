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

### 2.2 cuTile — `cutile/attn/flash.py` (pseudocode; ranks need checking on first run)

```python
# Host flattens [B,H,N,D] -> [B*H,N,D]; grid = (cdiv(N,BQ), B*H).
# Tiles are 3-D (1,BQ,D) so ct.mma broadcasts the 2-D KV tile — the same mechanism GQA uses.
@ct.kernel
def flash_fwd(Q, K, V, O, N, scale: ct.Constant[float], causal: ct.Constant[bool],
              BQ: ct.Constant[int], BK: ct.Constant[int], D: ct.Constant[int]):
    qi, bh = ct.bid(0), ct.bid(1)
    rows = ct.arange(BQ) + qi * BQ                      # absolute query indices
    q    = ct.load(Q, index=(bh, qi, 0), shape=(1, BQ, D)).astype(ct.float16)
    m    = ct.full((1, BQ), -1e30, ct.float32)           # fp32: -1e30 overflows fp16
    l    = ct.zeros((1, BQ), ct.float32)
    acc  = ct.zeros((1, BQ, D), ct.float32)
    nk   = ct.cdiv(N, BK)
    last = min(nk, ((qi + 1) * BQ - 1) // BK + 1) if causal else nk
    for k in range(last):                               # causal: skip tiles to the right
        kt = ct.load(K, index=(bh, k, 0), shape=(BK, D)).astype(ct.float16)
        s  = ct.mma(q, ct.transpose(kt), ct.zeros((1, BQ, BK), ct.float32)) * scale
        cols = ct.arange(BK) + k * BK
        keep = cols[None, :] < N                        # tail mask on BOTH paths
        if causal:
            keep = keep & (cols[None, :] <= rows[:, None])
        s = ct.where(keep, s, -1e30)                    # v1 masks every tile; split later if ncu says so
        m_new = ct.maximum(m, ct.max(s, axis=2))
        p  = ct.exp2(s - ct.expand_dims(m_new, 2))
        a  = ct.exp2(m - m_new)
        l  = l * a + ct.sum(p, axis=2)
        acc = ct.mma(p.astype(ct.float16), ct.load(V, index=(bh, k, 0), shape=(BK, D)),
                     acc * ct.expand_dims(a, 2))
        m  = m_new
    ct.store(O, index=(bh, qi, 0), shape=(1, BQ, D),
             tile=(acc / ct.expand_dims(l, 2)).astype(O.dtype))
```

- **Scale before the max.** `s * scale` happens inside the `mma` line, so `m`, `p` and `a` are all
  in log2 units. Scaling *after* the max gives `NaN` as soon as `p` underflows (caught at ×10
  inputs, not at unit scale).
- **The tail mask is unconditional.** Gating it on `causal` is the bug trap 3 warns about.
- **Verify on first run** (all unverified): `ct.transpose(x)` axis pair; rank of the `(1,BK)`
  predicate broadcasting against a `(1,BQ,BK)` tile; that `ct.load` accepts a 3-component index
  with a 2-component shape.
- `shortcut:` masking every tile costs a `where` per iteration. Split into an unmasked main loop
  plus a diagonal loop only if `ncu` shows it matters.

### 2.3 CUDA C++ — `cuda/attn/flash.cu`

128 threads (4 warps), `BQ=BK=64`, `D=64`. All the difficulty is the memory hierarchy:

- **v1 is single-buffered with plain loads.** Add `cp.async` only after Gate B, behind `-DPIPE=1`.
  The brief requires *manual shared-memory reuse*, not async pipelining — do not pay for both at once.
- `O` accumulator in registers: 4 warps × 16 rows/warp, so `16 × 64 / 32` = **32 fp32/thread** for
  `O` alone; budget **~100–128 total** with `S`/`P`, Q fragments and addressing. Confirm with
  `nvcc -Xptxas -v` and watch for spills.
- **`α` applied by hand** to those accumulator registers before the second mma.
- `mma.sync.m16n8k16` (warp-level tensor core, fp16 in / fp32 acc) for both GEMMs, with
  **`ldmatrix`** for fragment loads — `V` needs the `.trans` variant. Mismatched fragments between
  QKᵀ and PV are the classic silent-wrong-answer bug.
- Shared memory padded to avoid **bank conflicts** (two threads hitting one bank, which serialises).
- Causal: bound the k-loop; mask the diagonal tile.
- Realistically **350–450 lines** once CLI, JSON, host driver and timing are in — over the 250-line
  guideline, and that is the honest number for hand-written flash.
- **The exe cannot compute its own error** (no fp64 reference inside). Contract: `bench.py` writes
  `q/k/v.bin`, the exe writes `o.bin` + the JSON timings, `bench.py` computes `max_abs_err`.

```powershell
nvcc -arch=sm_86 -O3 --use_fast_math -lineinfo flash.cu -o flash.exe
```

### 2.4 Interfaces — how the three files talk

- `flash.exe --N 1024 --causal 0 --BQ 64 --BK 64` → writes `o.bin` + prints **one JSON line**:
  `{model,N,causal,BQ,BK,p50_us,p90_us}`. No `max_abs_err` — the exe has no reference.
- `cutile/attn/flash.py` → `attn(q,k,v,causal)` for import, plus `__main__` emitting the **same
  JSON** with the same flags.
- `bench/bench.py` → writes `q/k/v.bin`, shells out to the exe, imports the module, runs the fp64
  + SDPA references **and times the references itself**, then computes `max_abs_err` and **owns
  the CSV**.
- **Why the kernels self-time**: launch overhead would swamp a ~100 µs kernel, so each side times
  its own launch with CUDA events. The shared thing is the protocol (§3), not the code. The
  reference is timed in-process by `bench.py` because it is already in-process.

**Profiling commands** (both replay-safe — a bare `--set full` replays every one of your iterations):

```
ncu -k regex:flash --launch-skip 50 --launch-count 1 --set full -o profiles/flash_cuda flash.exe
ncu --section Tile --launch-count 1 -o profiles/flash_cutile python cutile/attn/flash.py
```

- On Windows, counters are usually blocked until **NVIDIA Control Panel → Desktop → Manage GPU
  Performance Counters** is enabled for your account (`ERR_NVGPUCTRPERM`). Test with a trivial
  kernel on day 0 — discovering this on day 5 costs the profiling marks.

## 3. Measurement

**The brief's three OP2 targets — mechanism → evidence:**

- **HBM traffic ↓ via fusion**: S never leaves registers/shared memory; evidence = `ncu` DRAM
  byte counters vs torch's `MATH` backend (the materialised baseline), figure 2.
- **Numerical stability**: fp32 accumulate, fp32 running max, `exp2`-folded scale, tail
  `ct.where`; evidence = `max_abs_err` vs fp64, zero NaN/Inf across the size matrix, **and a
  ×100 input-magnitude stress test** (a kernel with no max-subtraction dies here — that is the
  test proving the stabilisation is real, not decorative).
- **Latency at the target sequence length**: target **N = 2048** (prefill regime). The brief
  specifies **no** target — this is our assumption; state it in the report and ask Pravin.
  Evidence = p50/p90 at N=2048 plus the full latency-vs-N curve, figure 1.
- **Terminology**: the brief says HBM, but this laptop's 4 GB is **GDDR6** (6001 MHz DDR →
  12 Gbps, 128-bit → 192 GB/s). Same measurement, correct name: **device/DRAM bytes**. Metrics:
  `dram__bytes_read.sum`, `dram__bytes_write.sum`,
  `dram__throughput.avg.pct_of_peak_sustained_elapsed` — confirm locally with
  `ncu --query-metrics | findstr dram__bytes`.
- **cuTile-specific profiling**: `ncu --section Tile` reports launch config, execution info and
  GPU resource usage for **Tile IR** kernels — the official analogue of Launch/Occupancy for the
  cuTile side (Nsight Compute Profiling Guide, *Sets and Sections*).

**Protocol — identical for both models:**

1. Warm up 50 iters (pays cuTile's JIT compile). Compile time is reported separately, never inside
   a timing.
2. **Time to a duration, not a count**: repeat until ≥ 1 s of GPU work. 200 iterations of a 50 µs
   kernel is 10 ms total — that measures clock ramp, not steady state.
3. **Flush L2 outside the event window** — write 32 MB *before* `start.record()`. A dirty flush
   whose write-back lands inside the next kernel poisons the measurement.
4. Report p50/p90, and append `clocks.sm`, `power.draw`, `temperature.gpu` (via
   `nvidia-smi --query-gpu`) to **every CSV row**. No clock locking needed on Windows, and it
   turns "the laptop throttled" from an excuse into data.
5. **Cross-check** each config against `ncu gpu__time_duration` at N ∈ {256, 1024, 4096} — three
   timers (C++ events, Python events, torch) must agree, or you are measuring launch latency.
6. **Measure the peaks first**: fp16 4096³ torch matmul for compute, a large device copy for
   bandwidth. Every "% of peak" uses the measured number — not 7, not 20.
7. Reference through the same harness: torch SDPA, per backend.

**CSV schema — append only, ASCII order fixed:**

```
model, causal, dtype, B, H, N, D, BQ, BK, p50_us, p90_us, tflops, gbytes_s,
ref_backend, ref_us, pct_of_ref, max_abs_err, clocks_sm, power_w, temp_c, notes
```

- `tflops = 4·B·H·N²·D / t` (halve if causal). `max_abs_err` vs fp64 is written by `bench.py`.
- `pct_of_ref` = best available backend at that N — state which, per row. `gbytes_s` must say
  whether it is algorithmic or `ncu`-measured; say it in `notes`.

**Sizes:** `B=1, H=8, D=64`, fp16. Sweep `N ∈ {256, 512, 1024, 2048, 4096}`, plus:

- **Tail sizes `N ∈ {1, 65, 130, 1000}`.** A one-phantom-key tail error is ~1e-3 — *below* the
  5e-3 tolerance. `N=1023/2047/4095` therefore **do not catch a missing tail mask**; `N=1000`
  (24 phantom keys) does. Keep inputs ≈ randn: peaked inputs (`Q×4`) hide it at 1.3e-4.
- **`N=16384`** for the OOM demonstration. At `N=8192` one fp16 N×N tensor is 1.07 GB at H=8, so
  torch's `MATH` backend needs ≥ 2.15 GB against 3.78 GB free — it probably *runs*. `N=16384`
  needs 4.29 GB per tensor and cannot. Do not claim 8192 OOMs without measuring.
- The fp64 reference is computed **per head, chunked over query blocks** — a full `[8,4096,4096]`
  fp64 tensor is 1.07 GB, and impossible beyond that.
- **No `N=1` row.** Self-attention at `N=M=1` is trivially `O=V`. Real decode is `N_q=1` against a
  long `M_kv`, which breaks the `N=M` contract — that is cross-attention, out of scope.

**Figures — generated by `bench.py`, never hand-drawn. One per axis, which the brief asks for:**

1. latency vs N (CUDA, cuTile, 3 SDPA backends) — *performance*;
2. **DRAM bytes vs N** (from `ncu`) — the O(N²)→O(N) collapse, *the* money figure;
3. achieved TFLOP/s vs N against the measured peak — *performance*;
4. small multiples for the soft axes: LOC bar, iteration-time bar, knob active/inert grid —
   *velocity, tunability, portability* in one figure. Without it those three axes are prose.

## 4. The comparison — five axes → five experiments

| Axis | Experiment | Report |
|---|---|---|
| **Dev velocity** | `cloc` **kernel-only** (host driver and harness excluded from both sides); wall-clock edit→CSV row, split "algorithm" vs "model-specific" | LOC ratio; median iteration time; compile time (nvcc vs tileiras, cold+warm). **cuTile is written first and absorbs algorithm discovery — say so, or the comparison flatters CUDA.** |
| **Control & tunability** | Knob inventory, each marked *available on sm_86* for **both** columns | Tile shape/`occupancy`/`num_worker_warps` vs block size/smem layout/`cp.async` depth. TMA, clusters and FP8 are **unreachable on this GPU in either model** — not absent from either model |
| **Portability** | Count `__CUDA_ARCH__` guards + arch intrinsics; show `ByTarget(...)`; note `cp.async`/`mma.sync` ⇒ CUDA kernel needs sm_80+ | **This axis is argued, not measured** — there is one GPU. Label it that way. Concede: cuTile's ceiling = what `tileiras` implements, no `--ptx` escape |
| **Perf vs reference** | The CSV | % of each SDPA backend, % of measured peak, with variance |
| **Debuggability** | Inject 3 identical bugs in both — (a) missing sync/race, (b) mask sign, (c) rescale placement — and **record the first diagnostic each tool prints** | A 3×2 table of tool → first message. **No stopwatch**: n=1 and you already know where the bug is, so minutes-to-fix measures nothing |

- **Ablation ladder** (cheapest mark on the table). One kernel per model is an *endpoint*; a grader
  wants to see what each trick bought. Behind compile-time flags in the same files:
  single→double buffer, padding off/on, causal tile-skip off/on, tile shape. One CSV row each,
  tagged in `notes`. Still two kernels, still three files.
- **Prediction to test**: cuTile converts wrong-answer bugs into compile errors — earlier, but
  less searchable.
- **Required output**: one paragraph, unambiguous, on when the tile abstraction's productivity
  gain outweighs the ceded control. **Take a position.**

## 5. Schedule

Re-dated from **Fri 30 Oct** (gate slipped by the package install). Days 1–4 are the critical path.

| Day | Date | Work | Done when |
|---|---|---|---|
| 0 | **Fri 30 Oct** | **gate green** (pip → `.venv`) + **slides to Pravin**; start the Toolkit download in the background | `smoke_cutile.py` passes; deck sent |
| 1 | Sat 31 Oct | `bench.py` — correctness + CSV only; paste `ncu --csv` into `results/ncu.csv` later | one command → CSV with fp64 + SDPA rows |
| 2 | Sun 1 Nov | `cutile/attn/flash.py` — fused, `causal`, tail mask | matches fp64 within fp16 tolerance |
| 3 | Mon 2 Nov | `cuda/attn/flash.cu` — **single-buffered v1** (`nvcc` from day 0's download) | same tolerance; both rows in CSV |
| 4 | Tue 3 Nov | tile/block sweep (CUDA via templates); figures 1 + 3 | best config recorded, both models |
| 5 | Wed 4 Nov | `ncu` both kernels (`--section Tile` for cuTile); figure 2 | `profiles/` populated |
| 6 | Thu 5 Nov | report draft — **numbers only**; the static sections were written earlier | no placeholders |
| 7 | Fri 6 Nov | report final, 3-command README, save `compute-sanitizer` logs | a stranger reproduces in 3 commands |
| 8 | **Sat 7 Nov** | **submit**, buffer | submitted |

**Pre-cut now, before you need the time:**

1. `cp.async` double-buffering — v1 loads synchronously. Behind `-DPIPE=1`, after Gate B.
2. GQA, and the standalone sweep figure (keep the sweep as a table).
3. The timed bug experiment (use the diagnostic table) and "stranger regenerates every figure".
4. Figure 4 as a *fourth* performance figure — fold the soft axes into one small-multiples panel.

- **Day 0 is two jobs**: the gate *and* the deck — Pravin's price for the next conversation, so it
  does not queue behind code.
- The **CUDA Toolkit download (~8 GB, 20–40 min) is async** — start it before writing `bench.py`.
- Write the report's static sections (survey, algorithm, knob inventory) now; the deck already is
  §2. Day 6 should be numbers only.
- Lost a day ⇒ cut in the order listed above. **Never** cut day 5 (profiles) or day 6 (report) —
  those carry 40% of the grade.


## 6. Risks

| Risk | Mitigation |
|---|---|
| Driver can't reach 580 | decide day 1; fall back to cloud/lab GPU, laptop = dev box |
| Laptop throttling ruins repro | §3 protocol 4 + p50/p90 + stated variance |
| 4 GB VRAM caps N | sizes fixed in §3; the OOM datapoint is a result |
| Tail-mask bug (`exp2(0)=1` phantom key) | `N=1000` in the test matrix from day 2 |
| cuTile detail profiling needs driver ≥582.16 | install ≥582; else profile CUDA fully and state the gap |
