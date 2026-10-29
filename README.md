# CS-6023 — Attention under CUDA C++ vs cuTile (OP2)

- One merged attention kernel per programming model — thread level (CUDA C++) vs tile level
  (cuTile) — then a measured comparison. GEMM is out of scope.
- **Agent context is local only.** `AGENTS.md` (working state, rules, traps) is gitignored and
  therefore **not in this repository**. On a fresh clone it will be missing — keep your own copy.
- Detail: [`PLAN.md`](PLAN.md) · variants + deck: [`docs/attention-variants.md`](docs/attention-variants.md) · history: [`docs/decision-log.md`](docs/decision-log.md)

## Toolchain: native Windows, not WSL

- WSL2 uses the **same Windows driver** ⇒ it does not fix the driver blocker.
- NVIDIA's WSL guide: persistence mode and other `nvidia-smi` state features unsupported ⇒ no
  clock locking either. CUDA developer tools there are "preview".
- MSVC 2026 is already installed here; WSL = a distro + a second toolkit.

## Setup — four steps

### 1. NVIDIA driver ≥ 580 · **BLOCKER** · admin, then reboot

- Machine is on **566.07**. Nothing runs until this changes.
- <https://www.nvidia.com/Download/index.aspx> — Game Ready or Studio, either.
- Target **≥ 582.16**: that is what `ncu` needs for detailed counters on *cuTile* kernels.
- Verify: `nvidia-smi --query-gpu=driver_version --format=csv`

### 2. Python packages · no admin · ~10 min

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1
```

- Audits the machine, creates `.venv` (**venv** = isolated Python package set), installs,
  runs the smoke test. By hand:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install "cuda-tile[tileiras]" numpy cupy-cuda13x
```

- The `[tileiras]` extra vendors cuTile's compiler into the venv ⇒ **cuTile runs without a
  system-wide CUDA Toolkit**. That is why this comes before step 3.

### 3. CUDA Toolkit 13.2+ · admin · **CUDA C++ half only**

- Needed for `nvcc` (CUDA's compiler) and `compute-sanitizer`. cuTile does not need it.
- <https://developer.nvidia.com/cuda-downloads> → Windows x86_64; silent: `cuda_13.x_windows.exe -s`.
- ~8 GB, 20–40 min. **13.2 is the minimum** — the release that added Ampere (sm_80/86) to cuTile.
- If a build later fails on `cl.exe` (MSVC's compiler), add the **Desktop development with C++**
  workload to VS Build Tools.

### 4. Verify — the gate

```powershell
.\.venv\Scripts\python.exe .\scripts\smoke_cutile.py        # must print ALL THREE CHECKS PASSED
.\.venv\Scripts\python.exe -m pip freeze > docs\requirements.lock.txt
```

## Running

```powershell
.\.venv\Scripts\python.exe .\cutile\attn\flash.py      # tile-level kernel
.\.venv\Scripts\python.exe .\bench\bench.py            # correctness + benchmarks -> results\results.csv

cd cuda\attn
nvcc -arch=sm_86 -O3 --use_fast_math -lineinfo flash.cu -o flash.exe
ncu --set full -o ..\..\profiles\flash_cuda .\flash.exe
```

- Use `-lineinfo`, **never `-G`**: `-G` disables optimisation and makes every number meaningless.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| smoke test fails at device query | driver < 580 | step 1 |
| `ImportError: DLL load failed` | driver/toolkit mismatch | step 1, recreate venv |
| `pip`: no distribution for `cuda-tile` | wrong Python | `py -3.12 -m venv .venv` |
| `pip` fails on `cupy-cuda13x` | no wheel for this Python | 3.12, or `--only-binary :all:`, or use torch |
| SSL error from PowerShell 5.1 | TLS 1.0 default | `[Net.ServicePointManager]::SecurityProtocol = 'Tls12'` |
| `ensurepip` fails making the venv | temp-dir permissions | run it in your own terminal, not a sandboxed shell |
| `TileUnsupportedFeatureError` | FP8 on sm80 | stay fp16/bf16 |
| `ncu` profiles nothing | driver too old for the toolkit | step 1 (≥582.16 for cuTile detail) |
| `nvidia-smi -lgc` fails | clock locking is Linux-only | AC + *Prefer maximum performance* |

## Not installed, on purpose

- No WSL distro · no cuDNN/TensorRT/NCCL (unused) · no Triton (cuTile is the model under test —
  adding Triton muddies the comparison the brief asks for).
