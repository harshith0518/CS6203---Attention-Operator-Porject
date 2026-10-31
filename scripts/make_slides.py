"""Generate the dense 10-slide OP2 deck for the mentor review.

Covers, in order: problem and goal, the chosen operator, the implementation
approach, the stack, the measurement plan, the definition of done, the schedule,
references, and risks/asks.

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py

Output: docs/op2-attention-slides.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- design system
INK = RGBColor(0x13, 0x22, 0x29)
MUTED = RGBColor(0x5E, 0x6E, 0x76)
PANEL = RGBColor(0xEC, 0xF3, 0xF4)
RULE = RGBColor(0xD6, 0xE9, 0xEB)
TEAL = RGBColor(0x02, 0x76, 0x86)
MINT = RGBColor(0x5C, 0xD6, 0xB4)
ORANGE = RGBColor(0xC2, 0x5B, 0x22)
DEEP = RGBColor(0x0C, 0x33, 0x3B)
PAPER = RGBColor(0xFF, 0xFF, 0xFF)
CODEPANEL = RGBColor(0xF2, 0xF6, 0xF7)

HEAD = "Cambria"
BODY = "Calibri"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)
ML = Inches(0.55)
CW = W - 2 * ML
EYEBROW_Y = Inches(0.34)
TITLE_Y, TITLE_H = Inches(0.62), Inches(0.95)
CONTENT_Y = Inches(1.72)
FOOT_Y = Inches(7.06)
REPO = "github.com/harshith0518/CS6203---Attention-Operator-Porject"


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def alt(shape, text):
    shape._element._nvXxPr.cNvPr.set("descr", text)


def tb(slide, x, y, w, h, lines, size=14, bold=False, color=INK, font=BODY,
       align=PP_ALIGN.LEFT, space=1, line=1.1, alt_text=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, s in enumerate(lines if isinstance(lines, list) else [lines]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space)
        p.line_spacing = line
        r = p.add_run()
        r.text = s
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
    if alt_text:
        alt(box, alt_text)
    return box


def rect(slide, x, y, w, h, fill, alt_text=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    if alt_text:
        alt(s, alt_text)
    return s


def eyebrow(slide, s, dark=False):
    tb(slide, ML, EYEBROW_Y, Inches(11.0), Inches(0.24), s.upper(), size=11, bold=True,
       color=MINT if dark else TEAL)


def title(slide, s, dark=False, size=25):
    tb(slide, ML, TITLE_Y, CW, TITLE_H, s, size=size, bold=True,
       color=PAPER if dark else INK, font=HEAD, line=1.02)


def footer(slide, s, page, dark=False):
    c = RGBColor(0x9F, 0xBD, 0xC2) if dark else MUTED
    tb(slide, ML, FOOT_Y, Inches(11.2), Inches(0.24), s, size=10.5, color=c)
    tb(slide, W - ML - Inches(0.9), FOOT_Y, Inches(0.9), Inches(0.24), str(page),
       size=10.5, color=c, align=PP_ALIGN.RIGHT)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def table(slide, cols, rows, y=CONTENT_Y, widths=None, size=11.5, head_size=11.5,
          rowh=0.4, headh=0.3, tint_rows=(), x=ML, total=None):
    total = total or CW
    shape = slide.shapes.add_table(len(rows) + 1, len(cols), x, y, total,
                                   Inches(headh + rowh * len(rows)))
    t = shape.table
    if widths:
        for i, f in enumerate(widths):
            t.columns[i].width = int(total * f)
    t.rows[0].height = Inches(headh)
    for r in range(1, len(rows) + 1):
        t.rows[r].height = Inches(rowh)
    for c, hdr in enumerate(cols):
        cell = t.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DEEP
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.margin_top = cell.margin_bottom = Inches(0.01)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        r = cell.text_frame.paragraphs[0].add_run()
        r.text = hdr
        r.font.size = Pt(head_size)
        r.font.bold = True
        r.font.name = BODY
        r.font.color.rgb = PAPER
    for ri, row in enumerate(rows, start=1):
        tinted = (ri - 1) in tint_rows
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = (PANEL if tinted else (PAPER if ri % 2 else PANEL))
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.01)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            parts = val if isinstance(val, list) else (
                [val] if isinstance(val, tuple) else [(val, size, ci == 0, INK)])
            for k, (txt, sz, bold, col) in enumerate(parts):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.line_spacing = 0.98
                r = p.add_run()
                r.text = txt
                r.font.size = Pt(sz)
                r.font.bold = bold
                r.font.name = BODY
                r.font.color.rgb = col
    return t


def code(slide, x, y, w, h, header, lines, accent, size=9.5):
    rect(slide, x, y, w, h, CODEPANEL, "code panel: " + header)
    rect(slide, x, y, w, Inches(0.26), accent)
    tb(slide, x + Inches(0.14), y + Inches(0.02), w - Inches(0.24), Inches(0.22),
       header, size=11, bold=True, color=PAPER)
    box = slide.shapes.add_textbox(x + Inches(0.12), y + Inches(0.33),
                                   w - Inches(0.2), h - Inches(0.4))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.name = MONO
        r.font.color.rgb = MUTED if ln.strip().startswith(("#", "//")) else INK


def stat(slide, x, y, w, big, caption, colour):
    tb(slide, x, y, w, Inches(0.42), big, size=26, bold=True, color=colour, font=HEAD)
    tb(slide, x, y + Inches(0.44), w, Inches(0.4), caption, size=10.5, color=MUTED, line=1.15)


prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# ============================================================== 1 title (dark)
s = new(prs)
rect(s, 0, 0, W, H, DEEP, "dark background")
eyebrow(s, "CS-6023 GPU kernel programming  \u00b7  Project OP2", dark=True)
title(s, "Attention, written two ways", dark=True, size=50)
tb(s, ML, Inches(2.3), Inches(10.6), Inches(1.0),
   ["One fused attention kernel in CUDA C++ and in cuTile, compared on cost and payoff.",
    "Goal, approach, stack, evidence and definition of done."],
   size=17, color=RGBColor(0xC4, 0xDD, 0xDF), line=1.35)
tb(s, ML, Inches(5.55), Inches(11.6), Inches(1.0),
   ["Surya Harshith (cs23b010)  \u00b7  for Pravin Jagtap and Sumanta Datta  \u00b7  31 October 2026",
    REPO], size=11.5, color=RGBColor(0x8F, 0xB3, 0xB9), line=1.5)
notes(s, "Seven days to the deadline. This deck is the whole picture: what the operator is, why it "
         "is a memory problem, which variant I build, what I build it with, how I will know it is "
         "right, and what could stop me.")

# ========================================================== 2 problem and goal
s = new(prs)
eyebrow(s, "The problem")
title(s, "Attention is limited by moving an N\u00d7N matrix, not by arithmetic.")
stat(s, ML, CONTENT_Y, Inches(3.1), "268 MB",
     "one N\u00d7N fp16 tensor; N=4096, 8 heads, D=64  (computed)", ORANGE)
stat(s, ML + Inches(3.3), CONTENT_Y, Inches(3.1), "\u00d7 4 passes",
     "written out, read back, again: \u2248 1.07 GB  (computed)", ORANGE)
stat(s, ML + Inches(6.6), CONTENT_Y, Inches(3.1), "17 MB",
     "the same result, fused: S never leaves the chip  (computed)", TEAL)
stat(s, ML + Inches(9.9), CONTENT_Y, Inches(2.1), "64\u00d7",
     "less traffic  (computed)", TEAL)

rect(s, ML, Inches(2.86), CW, Pt(1.5), RULE)
tb(s, ML, Inches(3.0), Inches(6.3), Inches(0.9),
   ["32 FLOP per byte unfused, against a ridge point near 104 FLOP per byte.",
    "So the naive version is memory-bound, and fusion is the design, not a tuning trick."],
   size=12, color=INK, line=1.25)
tb(s, ML, Inches(3.95), Inches(6.3), Inches(0.7),
   "Prediction, not a result: 64\u00d7 less traffic, at most about 3\u00d7 less time, because compute "
   "becomes the limit once memory is not.", size=11, color=MUTED, line=1.25)

tb(s, ML, Inches(4.72), Inches(6.3), Inches(0.24), "WHAT FUSION BUYS AND COSTS",
   size=10.5, bold=True, color=TEAL)
tb(s, ML, Inches(4.98), Inches(6.3), Inches(1.6),
   ["Buys: one kernel instead of three, and no 268 MB tensor ever reaches memory.",
    "Costs: shared memory and register pressure, plus a softmax you only ever see one slice of.",
    "That trade is the whole project \u2014 and it is why the tile model is interesting here, since "
    "the tile compiler takes the second cost off your hands."],
   size=11, color=INK, line=1.28, space=3)

rect(s, Inches(7.15), Inches(2.92), Inches(5.63), Inches(1.95), PANEL, "goal statement")
tb(s, Inches(7.38), Inches(3.04), Inches(5.2), Inches(0.24), "GOAL", size=10.5, bold=True, color=TEAL)
tb(s, Inches(7.38), Inches(3.28), Inches(5.2), Inches(1.5),
   ["Build the same fused operator twice \u2014 hand-written CUDA C++ and tile-level cuTile \u2014 and "
    "measure which abstraction is worth the control it asks you to give up.",
    "Same maths, same sizes, same timing protocol, same reference."],
   size=11.5, color=INK, line=1.3)

tb(s, Inches(7.15), Inches(5.02), Inches(5.63), Inches(1.6),
   ["Constraint set: one developer, one laptop GPU (RTX 3050, 4 GB, compute capability 8.6),",
    "seven days, and a brief that grades correctness 25, performance 35, analysis 25, evidence 15."],
   size=10.5, color=MUTED, line=1.3)
footer(s, "Ridge assumes a 20 TFLOP/s tensor peak; the real peak is measured first and this figure "
          "recomputed. The card's memory is GDDR6, the brief says HBM.", 2)
notes(s, "Four passes over a 268 MB matrix is the 1.07 GB. The fused kernel never writes it. Two "
         "numbers matter: 32 FLOP per byte unfused, 2048 fused, against a ridge near 104. Do not "
         "read 64 times as a speedup.")

# ========================================================== 3 the operator
s = new(prs)
eyebrow(s, "The operator  \u00b7  the six common variants Pravin asked about, each in both models")
title(s, "One variant gets built. The other five decide why.", size=24)
tb(s, ML, Inches(1.58), CW, Inches(0.26),
   "O = softmax(Q K^T / \u221ad + mask) \u00b7 V      MHA, self-attention, forward only, exact, causal as "
   "a flag      fp16 in, fp32 accumulate      pass: max|out \u2212 ref| \u2264 5e-3 \u00b7 max|ref|",
   size=10, color=MUTED)
rows = [
    [[("1  Vanilla SDPA", 11, True, INK), ("baseline fusion replaces", 9, False, MUTED)],
     ("N\u00d7N matrix round-trips memory", 11, False, INK),
     ("3 launches: GEMM, softmax, GEMM", 11, False, INK),
     ("same 3 kernels, no help", 11, False, INK),
     ("baseline = PyTorch MATH", 11, False, MUTED)],
    [[("2  Causal", 11, True, INK), ("GPT-style decoders", 9, False, MUTED)],
     ("skip future tiles, mask the diagonal", 11, False, INK),
     ("loop bound + if, uneven block work", 11, False, INK),
     ("loop bound + ct.where", 11, False, INK),
     ("flag in my kernel", 11, True, TEAL)],
    [[("3  Fused + online softmax", 11, True, TEAL), ("FlashAttention", 9, False, MUTED)],
     ("softmax over a row seen in slices", 11, False, INK),
     ("rescale accumulators by hand", 11, False, INK),
     ("rescale is one tile multiply", 11, False, INK),
     ("BUILD", 11, True, TEAL)],
    [[("4  GQA / MQA", 11, True, INK), ("Llama 3", 9, False, MUTED)],
     ("query heads share one K/V head", 11, False, INK),
     ("index maths, choose who shares smem", 11, False, INK),
     ("group stacked into a tile dim (to verify)", 11, False, INK),
     ("stretch, after correctness", 11, False, MUTED)],
    [[("5  Sliding window", 11, True, INK), ("Mistral 7B", 9, False, MUTED)],
     ("band-shaped loop bounds", 11, False, INK),
     ("two bounds + edge masks", 11, False, INK),
     ("same two bounds as range()", 11, False, INK),
     ("read only", 11, False, MUTED)],
    [[("6  Cross-attention", 11, True, INK), ("T5, Whisper", 9, False, MUTED)],
     ("Nq \u2260 Nkv, two tilings", 11, False, INK),
     ("natural, one row per block", 11, False, INK),
     ("64-row tile on 1 query wastes 63/64", 11, False, INK),
     ("read only, CUDA may win", 11, False, MUTED)],
]
table(s, ["Variant  (used in)", "The new problem", "CUDA C++", "cuTile", "Status"], rows,
      y=Inches(1.94), widths=[0.22, 0.22, 0.21, 0.22, 0.13], size=11, rowh=0.5,
      headh=0.28, tint_rows=(2,))
tb(s, ML, Inches(5.35), CW, Inches(0.9),
   ["Why #3: it is the only variant where both models face the same hard problem \u2014 a softmax you "
    "can only see one slice of at a time. Causal then costs almost nothing, a loop bound and a mask "
    "rather than a second project.",
    "Rows 1\u20132 are the baseline and the flag; rows 4\u20136 are read and reported, not built."],
   size=10.5, color=MUTED, line=1.3)
footer(s, "CUDA and cuTile cells are hypotheses from reading the documentation, not measurements. "
          "Only row 3 and the causal flag produce numbers.", 3)
notes(s, "These are the six I would call common, and what each costs in each model. I build row 3 "
         "only. Everything in the two middle columns is my reading, not a measurement.")

# ========================================================== 4 approach
s = new(prs)
eyebrow(s, "The approach  \u00b7  three files, one loop, two owners of the machine")
title(s, "The same loop twice: registers and shuffles, or tiles.", size=24)
code(s, ML, Inches(1.66), Inches(6.05), Inches(4.5), "CUDA C++  \u00b7  sketch", [
    "// each thread owns 2 rows of the warp's 16-row tile",
    "float m[2] = {-1e30f,-1e30f}, l[2] = {0,0};",
    "float acc[D/8][4] = {};     // 32 fp32/thread",
    "for (int k = 0; k < nk; ++k) {",
    "  load_kv(k); __syncthreads();",
    "  mma_qk(s, q_frag, k);     // S = Q K^T",
    "  for (int r = 0; r < 2; ++r) {",
    "    float x = rowmax(s, r) * scale;",
    "    x = fmaxf(x, __shfl_xor_sync(~0u,x,1));",
    "    x = fmaxf(x, __shfl_xor_sync(~0u,x,2));",
    "    float mn = fmaxf(m[r], x);",
    "    float a  = exp2f(m[r] - mn);   // rescale",
    "    l[r] *= a; m[r] = mn;",
    "    for (int n=0;n<D/8;++n) acc[n][2*r]*=a;",
    "  }",
    "  // p = exp2(s*scale - m); l += rowsum(p)",
    "  // acc += p V (mma); tail masked with cols<N",
    "}",
], ORANGE, size=8.5)

code(s, Inches(6.83), Inches(1.66), Inches(5.95), Inches(4.5), "cuTile  \u00b7  sketch", [
    "bh, qi = ct.bid(1), ct.bid(0)",
    "q = ct.load(Q, (bh, qi, 0), (1, BQ, D))",
    "m = full(-1e30); l = zeros; acc = zeros",
    "for k in range(num_kv_tiles):",
    "    kt = ct.load(K, (bh, k, 0), (1, BK, D))",
    "    s = ct.mma(q, ct.transpose(kt)) * scale",
    "    s = ct.where(keep(qi, k), s, -1e30)",
    "    m_new = ct.maximum(m, ct.max(s, axis=2))",
    "    p = ct.exp2(s - ct.expand_dims(m_new, 2))",
    "    a = ct.exp2(m - m_new)   # rescale",
    "    l = l * a + ct.sum(p, axis=2)",
    "    acc = ct.mma(p.astype(f16), load_v(k),",
    "                 acc * ct.expand_dims(a, 2))",
    "    m = m_new",
    "ct.store(O, (bh, qi, 0), acc / ct.expand_dims(l, 2))",
], TEAL, size=8.5)

steps = [("1", "bench.py first", "fp64 reference, SDPA backends, one CSV. Defines the contract the "
                                  "kernels must meet."),
         ("2", "cuTile kernel", "Validates the algorithm cheaply; tile shapes are the tuning knob."),
         ("3", "CUDA kernel", "v1 single-buffered, then cp.async behind a flag, then a sweep.")]
for i, (n, name, what) in enumerate(steps):
    x = ML + i * Inches(4.0)
    tb(s, x, Inches(6.3), Inches(3.85), Inches(0.24), n + ".  " + name, size=11, bold=True, color=TEAL)
    tb(s, x, Inches(6.54), Inches(3.85), Inches(0.45), what, size=10, color=MUTED, line=1.15)
footer(s, "Sketches, not compiled yet. Both scale before the max, mask the tail, and rescale before "
          "accumulating. cuTile API details remain unverified until it runs.", 4)
notes(s, "Left is what I write per thread, right is what I write per tile. The rescale is 32 registers "
         "by hand on the left and one multiply on the right. Whether that one line costs speed is "
         "what the profiler will tell me.")

# ========================================================== 5 the stack
s = new(prs)
eyebrow(s, "The stack  \u00b7  installed and verified on this machine")
title(s, "What we build with, and what this GPU cannot do.", size=25)
table(s, ["Hardware / runtime", "Version", "Role"], [
    [("GPU", 11, True, INK), ("RTX 3050 Laptop, 4 GB, CC 8.6", 11, False, INK),
     ("the entire test bed", 11, False, MUTED)],
    [("Driver", 11, True, INK), ("617.42", 11, False, INK),
     ("needs \u2265580; \u2265582.16 for detailed ncu", 11, False, MUTED)],
    [("Memory", 11, True, INK), ("GDDR6, ~192 GB/s nominal", 11, False, INK),
     ("sets the roofline; 4 GB caps N", 11, False, MUTED)],
    [("OS / shell", 11, True, INK), ("Windows 11, PowerShell 5.1", 11, False, INK),
     ("no pwsh 7, no WSL needed", 11, False, MUTED)],
], y=CONTENT_Y, widths=[0.2, 0.44, 0.36], size=11, rowh=0.42, headh=0.28, total=Inches(6.0))

table(s, ["Software", "Version", "Why it is there"], [
    [("cuda-tile", 11, True, INK), ("1.7.0", 11, False, INK), ("the tile-level half", 11, False, MUTED)],
    [("tileiras, nvvm, ptxas", 11, True, INK), ("13.4.92", 11, False, INK),
     ("compiles TileIR to cubin in-process", 11, False, MUTED)],
    [("nvrtc", 11, True, INK), ("13.4.92", 11, False, INK),
     ("CuPy needs it for elementwise ops", 11, False, MUTED)],
    [("cupy-cuda13x", 11, True, INK), ("14.2.0", 11, False, INK),
     ("GPU arrays; DLPack into cuTile", 11, False, MUTED)],
    [("numpy", 11, True, INK), ("2.5.4", 11, False, INK), ("host-side checks", 11, False, MUTED)],
    [("nvcc", 11, True, INK), ("13.4.92", 11, False, INK), ("builds the CUDA C++ half", 11, False, MUTED)],
], y=CONTENT_Y, widths=[0.26, 0.16, 0.58], size=11, rowh=0.34, headh=0.28,
   x=Inches(6.85), total=Inches(5.93))

rect(s, ML, Inches(3.82), Inches(6.0), Inches(2.1), PANEL, "inert features panel")
tb(s, ML + Inches(0.2), Inches(3.93), Inches(5.6), Inches(0.24),
   "INERT ON THIS GPU, IN BOTH MODELS", size=10.5, bold=True, color=ORANGE)
tb(s, ML + Inches(0.2), Inches(4.19), Inches(5.6), Inches(1.6),
   ["TMA, the Tensor Memory Accelerator: sm_90 and later",
    "Thread-block clusters: num_ctas needs sm_90",
    "FP8 and block-scaled MMA: Ada or Hopper and later",
    "So three of cuTile's eight knobs do nothing on Ampere, and CUDA C++ cannot",
    "reach those instruction families either. That is a result for the portability",
    "axis, not a limitation to apologise for."],
   size=10.5, color=INK, line=1.25)

rect(s, Inches(6.85), Inches(4.32), Inches(5.93), Inches(1.05), PANEL, "status panel")
tb(s, Inches(7.05), Inches(4.42), Inches(5.5), Inches(0.24), "STATUS TODAY", size=10.5, bold=True, color=TEAL)
tb(s, Inches(7.05), Inches(4.66), Inches(5.5), Inches(0.7),
   ["Installed and importing: cuTile 1.7.0, cupy 14.2.0, numpy 2.5.4, the 13.4.92 toolchain.",
    "A cuTile kernel has compiled and launched on the GPU; the gate re-run is pending."],
   size=10.5, color=INK, line=1.25)
footer(s, "Verified with pip list, nvidia-smi and a smoke test on this machine, not inferred from "
          "documentation.", 5)
notes(s, "This is the whole toolchain and it is deliberately small: three Python libraries, one "
         "compiler, one driver. The important panel is the inert one: three features cuTile exposes "
         "do not exist on Ampere, which is itself part of the answer the project is after.")

# ========================================================== 6 measurement
s = new(prs)
eyebrow(s, "The measurement  \u00b7  one protocol, six metrics, four artefacts")
title(s, "How the numbers are produced, and why they can be believed.", size=25)
tb(s, ML, CONTENT_Y, Inches(6.0), Inches(0.24), "PROTOCOL, IDENTICAL FOR BOTH MODELS",
   size=10.5, bold=True, color=TEAL)
tb(s, ML, Inches(1.96), Inches(5.9), Inches(2.4),
   ["1.  Warm up 50 iterations; compile time reported separately, never inside a timing.",
    "2.  Time to a duration, not a count: at least 1 s of GPU work per configuration.",
    "3.  Flush L2 outside the event window, so write-back does not land in the next kernel.",
    "4.  Report p50 and p90; log clocks, power and throttle state on every row.",
    "5.  Cross-check against ncu gpu__time_duration at N = 256, 1024, 4096."],
   size=11, color=INK, line=1.3, space=3)

tb(s, ML, Inches(4.3), Inches(6.0), Inches(0.24), "WHY THESE CHOICES", size=10.5, bold=True, color=TEAL)
tb(s, ML, Inches(4.56), Inches(5.9), Inches(1.5),
   ["200 iterations of a 50 \u00b5s kernel is 10 ms of GPU time: that measures clock ramp, not steady "
    "state.",
    "Three timers have to agree, or the number is launch overhead wearing a lab coat."],
   size=10.5, color=MUTED, line=1.3)

table(s, ["Metric", "Produced by", "Answers"], [
    [("latency", 11, True, INK), ("p50/p90 from CUDA events", 11, False, INK), ("performance", 11, False, MUTED)],
    [("traffic", 11, True, INK), ("ncu DRAM bytes, fused vs unfused", 11, False, INK),
     ("the 64\u00d7 claim", 11, False, MUTED)],
    [("throughput", 11, True, INK), ("TFLOP/s vs measured peak", 11, False, INK),
     ("distance from roofline", 11, False, MUTED)],
    [("error", 11, True, INK), ("max abs vs the fp64 reference", 11, False, INK),
     ("correctness, 25%", 11, False, MUTED)],
    [("tunability", 11, True, INK), ("tile and block sweeps, ablations", 11, False, INK),
     ("what each trick bought", 11, False, MUTED)],
    [("velocity", 11, True, INK), ("kernel-only LOC, edit-to-number time", 11, False, INK),
     ("analysis, 25%", 11, False, MUTED)],
], y=Inches(1.94), widths=[0.22, 0.5, 0.28], size=11, rowh=0.4, headh=0.28,
   x=Inches(6.85), total=Inches(5.93))

rect(s, Inches(6.85), Inches(4.82), Inches(5.93), Inches(0.95), PANEL, "figures note")
tb(s, Inches(7.05), Inches(4.92), Inches(5.5), Inches(0.8),
   ["Four figures, one per axis: latency vs N, DRAM bytes vs N, TFLOP/s vs measured peak,",
    "and one small-multiples panel for velocity, tunability and portability."],
   size=10.5, color=INK, line=1.25)
footer(s, "Artefacts: results.csv, one row per configuration; profiles/*.ncu-rep; figs/*.pdf "
          "generated from the CSV; report.md.", 6)
notes(s, "The protocol exists to survive scrutiny. The third line is the one that matters: if the "
         "flush lands inside the timing, fusion looks free. Every row carries clock and temperature, "
         "because this is a laptop that throttles.")

# ========================================================== 7 definition of done
s = new(prs)
eyebrow(s, "Definition of done  \u00b7  what acceptable means, in numbers")
title(s, "Six checks decide whether this is finished. Each can fail.", size=25)
table(s, ["#", "Check", "Passes when", "Evidence"], [
    [("1", 11, True, INK), ("Correct on every size", 11, True, INK),
     ("max|out \u2212 ref_fp64| \u2264 5e-3 \u00b7 max|ref| at N \u2208 {1, 65, 130, 256, 512, 1024, 1000, 2048, 4096}",
      10.5, False, INK), ("results.csv", 10.5, False, MUTED)],
    [("2", 11, True, INK), ("Numerically stable", 11, True, INK),
     ("no NaN or Inf; inputs \u00d7100 stay finite; causal and non-causal both", 10.5, False, INK),
     ("results.csv", 10.5, False, MUTED)],
    [("3", 11, True, INK), ("Both models, one contract", 11, True, INK),
     ("cuTile and CUDA rows for identical shapes, produced by the same harness", 10.5, False, INK),
     ("results.csv", 10.5, False, MUTED)],
    [("4", 11, True, INK), ("Performance measured, not claimed", 11, True, INK),
     ("p50/p90 at N=2048; DRAM bytes from ncu; percent of the best SDPA backend", 10.5, False, INK),
     ("results.csv, profiles", 10.5, False, MUTED)],
    [("5", 11, True, INK), ("Profiling evidence exists", 11, True, INK),
     ("one ncu report per kernel, with section Tile for the cuTile side", 10.5, False, INK),
     ("profiles/*.ncu-rep", 10.5, False, MUTED)],
    [("6", 11, True, INK), ("The argument is made", 11, True, INK),
     ("five axes, one figure per axis, and a recommendation that takes a side", 10.5, False, INK),
     ("report.md", 10.5, False, MUTED)],
], y=CONTENT_Y, widths=[0.04, 0.2, 0.55, 0.21], size=10.5, rowh=0.54, headh=0.28)

rect(s, ML, Inches(5.4), CW, Inches(1.15), PANEL, "gate panel")
tb(s, ML + Inches(0.2), Inches(5.5), Inches(4.0), Inches(0.24), "THE HARD GATE",
   size=10.5, bold=True, color=ORANGE)
tb(s, ML + Inches(0.2), Inches(5.75), Inches(11.4), Inches(0.7),
   ["No variant work (GQA, sliding window, cross-attention) starts until checks 1 to 4 pass on both "
    "kernels. Exploration is the reward for finishing, not a parallel track.",
    "A correct kernel with a sharp argument scores. A fast kernel with no argument does not."],
   size=11, color=INK, line=1.25)
footer(s, "The tolerance is my choice: the brief prescribes none. It is stated here so it can be "
          "argued with.", 7)
notes(s, "Every row can fail, which is the point. Row 1 is what catches the tail bug, because those "
         "sizes make a missing mask produce an error above tolerance. Row 6 is where a quarter of "
         "the grade sits, and it is writing, not code.")

# ========================================================== 8 plan
s = new(prs)
eyebrow(s, "The plan  \u00b7  seven days, five gates")
title(s, "What happens on which day, and what gets cut first.", size=25)
table(s, ["Day", "Date", "Work", "Gate"], [
    [("0", 11, True, INK), ("Sat 1 Nov", 11, False, INK),
     ("gate green: smoke test passes; this deck sent", 11, False, INK),
     ("environment", 10.5, True, TEAL)],
    [("1", 11, True, INK), ("Sun 2 Nov", 11, False, INK),
     ("bench.py: fp64 reference, SDPA backends, CSV, measured peaks", 11, False, INK),
     ("contract", 10.5, True, TEAL)],
    [("2", 11, True, INK), ("Mon 3 Nov", 11, False, INK),
     ("cuTile fused kernel: causal flag, tail mask, correct on all sizes", 11, False, INK),
     ("kernels", 10.5, True, TEAL)],
    [("3", 11, True, INK), ("Tue 4 Nov", 11, False, INK),
     ("CUDA kernel v1 single-buffered, then cp.async behind a flag", 11, False, INK),
     ("kernels", 10.5, True, TEAL)],
    [("4", 11, True, INK), ("Wed 5 Nov", 11, False, INK),
     ("tile and block sweeps, ablation ladder, figures 1 and 3", 11, False, INK),
     ("numbers", 10.5, True, TEAL)],
    [("5", 11, True, INK), ("Thu 6 Nov", 11, False, INK),
     ("Nsight Compute on both kernels, DRAM figure, section Tile", 11, False, INK),
     ("evidence", 10.5, True, TEAL)],
    [("6", 11, True, INK), ("Fri 7 Nov", 11, False, INK),
     ("submit: report with five axes and the recommendation", 11, False, INK),
     ("argument", 10.5, True, TEAL)],
], y=CONTENT_Y, widths=[0.05, 0.13, 0.66, 0.16], size=10.5, rowh=0.46, headh=0.28)

rect(s, ML, Inches(5.42), Inches(6.0), Inches(1.2), PANEL, "cut list panel")
tb(s, ML + Inches(0.2), Inches(5.52), Inches(5.6), Inches(0.24), "CUT FIRST, IN THIS ORDER",
   size=10.5, bold=True, color=ORANGE)
tb(s, ML + Inches(0.2), Inches(5.77), Inches(5.6), Inches(0.8),
   ["1.  GQA, and the standalone sweep figure",
    "2.  Figure 4, the velocity and tunability small multiples",
    "3.  The debuggability table, last because it is cheap"],
   size=10.5, color=INK, line=1.25)
rect(s, Inches(6.85), Inches(5.42), Inches(5.93), Inches(1.2), PANEL, "never cut panel")
tb(s, Inches(7.05), Inches(5.52), Inches(5.5), Inches(0.24), "NEVER CUT", size=10.5, bold=True, color=TEAL)
tb(s, Inches(7.05), Inches(5.77), Inches(5.5), Inches(0.8),
   ["The fused kernel in both models, the correctness numbers, and the profiles.",
    "Together those carry three quarters of the grade."],
   size=10.5, color=INK, line=1.25)
footer(s, "The brief's own milestones, 10/07 to 10/28, have passed. This is the recovery plan; the "
          "questions on the last slide are what could change it.", 8)
notes(s, "I am behind the brief's milestones and this is the recovery plan. Two kernels, measured "
         "properly, with a report that takes a position. The cut list is written now so that "
         "decision is not made at midnight on the sixth.")

# ========================================================== 9 references
s = new(prs)
eyebrow(s, "References and resources")
title(s, "What we are standing on, and where to read it.", size=25)

groups = [
    ("cuTile and the toolchain", [
        ("cuTile Python documentation", "docs.nvidia.com/cuda/cutile-python"),
        ("NVIDIA/cutile-python: source, samples", "github.com/NVIDIA/cutile-python"),
        ("NVIDIA/TileGym: matmul + attention", "github.com/NVIDIA/TileGym"),
        ("NVIDIA/cuda-tile: the TileIR compiler", "github.com/NVIDIA/cuda-tile"),
        ("Nsight Compute profiling guide", "docs.nvidia.com/nsight-compute/ProfilingGuide"),
    ]),
    ("The algorithm", [
        ("Dao et al., FlashAttention, 2022", "arxiv.org/abs/2205.14135"),
        ("Milakov and Gimelshein, online softmax", "arxiv.org/abs/1801.05824"),
        ("Dao, FlashAttention-2, 2023", "arxiv.org/abs/2307.08691"),
        ("Vaswani et al., Attention Is All You Need", "arxiv.org/abs/1706.03762"),
    ]),
    ("Learning attention visually", [
        ("3Blue1Brown, Deep Learning ch.6", "3blue1brown.com/lessons/attention"),
        ("Alammar, The Illustrated Transformer", "jalammar.github.io/illustrated-transformer"),
        ("Georgia Tech, Transformer Explainer", "poloclub.github.io/transformer-explainer"),
        ("This repository: plan, decision log", REPO),
    ]),
]
col_x = [ML, Inches(4.65), Inches(8.75)]
col_w = [Inches(3.85), Inches(3.85), Inches(4.03)]
for gi, (gname, items) in enumerate(groups):
    x, w = col_x[gi], col_w[gi]
    tb(s, x, CONTENT_Y, w, Inches(0.24), gname.upper(), size=9.5, bold=True, color=TEAL)
    rect(s, x, Inches(1.94), w - Inches(0.12), Pt(1.2), RULE)
    y = 2.04
    for name, url in items:
        tb(s, x, Inches(y), w - Inches(0.14), Inches(0.27), name, size=10, bold=True,
           color=INK, line=1.05)
        tb(s, x, Inches(y + 0.25), w - Inches(0.14), Inches(0.24), url, size=8.5,
           color=TEAL, font=MONO)
        y += 0.62

rect(s, ML, Inches(5.32), CW, Inches(1.28), PANEL, "local resources panel")
tb(s, ML + Inches(0.2), Inches(5.42), Inches(11.4), Inches(0.24),
   "RESOURCES ON THIS MACHINE, AND WHAT IS STILL MISSING", size=10, bold=True, color=TEAL)
tb(s, ML + Inches(0.2), Inches(5.66), Inches(11.6), Inches(0.85),
   ["Hardware: RTX 3050 Laptop, 4 GB GDDR6, compute capability 8.6, driver 617.42.",
    "Software: a venv with cuda-tile 1.7.0, the 13.4.92 toolchain (tileiras, nvcc, nvvm, nvrtc), "
    "cupy 14.2.0, numpy 2.5.4.",
    "Missing, needed only for the CUDA half and profiling: ncu and compute-sanitizer from the CUDA "
    "Toolkit 13.2 or later. Neither is needed for the cuTile gate."],
   size=10, color=INK, line=1.22)
footer(s, "Every link here was opened and read while preparing this deck. Claims that are not "
          "verified are marked as such in the plan.", 9)
notes(s, "Row one is where the API truth lives: the cuTile documentation and TileGym, which already "
         "contains attention kernels. Row two is the algorithm. Row three is how I learned it. The "
         "fourth column is this repository, which carries the plan and the decision log.")

# ========================================================== 10 risks and asks
s = new(prs)
rect(s, 0, 0, W, H, DEEP, "dark background")
eyebrow(s, "Risks and asks", dark=True)
title(s, "What could stop this, and what I need from you.", dark=True, size=25)

tb(s, ML, Inches(1.6), Inches(6.4), Inches(0.24), "TOP RISKS  \u00b7  AND HOW EACH IS HANDLED",
   size=10.5, bold=True, color=MINT)
risks = [("Laptop throttling", "benchmarks drift between runs", "clocks and power logged per row; p50/p90, not means"),
         ("4 GB of VRAM", "caps N and forbids some sizes", "sizes fixed now; the OOM case is a result"),
         ("cuTile on sm_86", "some paths are still unproven", "kernel before harness; every call verified"),
         ("Seven days, one person", "no slack for a bad surprise", "cut list written; two kernels is the floor")]
y = 1.9
for name, why, mit in risks:
    tb(s, ML, Inches(y), Inches(1.6), Inches(0.42), name, size=11, bold=True, color=PAPER, line=1.08)
    tb(s, ML + Inches(1.66), Inches(y), Inches(1.85), Inches(0.42), why, size=9.5,
       color=RGBColor(0x9F, 0xBD, 0xC2), line=1.08)
    tb(s, ML + Inches(3.58), Inches(y), Inches(3.0), Inches(0.42), mit, size=9.5,
       color=PAPER, line=1.08)
    y += 0.7

tb(s, Inches(7.2), Inches(1.6), Inches(5.6), Inches(0.24), "THREE ANSWERS I NEED",
   size=10.5, bold=True, color=MINT)
tb(s, Inches(7.2), Inches(1.9), Inches(5.6), Inches(2.9),
   ["1.  Attention only, no GEMM. The brief says choose either but lists artefacts for both; a line "
    "in writing would settle it.",
    "2.  Does 07 November hold, given the restart on 14 October?",
    "3.  Target sequence length: I assume N = 2048 unless you say otherwise."],
   size=11, color=PAPER, line=1.35, space=9)
rect(s, Inches(7.2), Inches(4.72), Inches(5.58), Inches(1.05), RGBColor(0x11, 0x40, 0x49),
     "defaults panel")
tb(s, Inches(7.4), Inches(4.82), Inches(5.2), Inches(0.9),
   ["Default if I hear nothing by Monday 3 November: OP2 only, N = 2048, GQA stays stretch,",
    "and the schedule on the previous slide runs as written."],
   size=10, color=RGBColor(0xD3, 0xE5, 0xE8), line=1.3)

tb(s, ML, Inches(5.0), Inches(6.4), Inches(0.24), "WHAT ARRIVES ON 07 NOVEMBER",
   size=10.5, bold=True, color=MINT)
tb(s, ML, Inches(5.26), Inches(6.4), Inches(1.6),
   ["Two kernels, a harness that re-runs in three commands, profiles for both, a CSV behind every",
    "number, and a report that names which abstraction wins and under what conditions.",
    "If something is cut, it is named in the report rather than quietly missing."],
   size=10.5, color=PAPER, line=1.3)
footer(s, REPO, 10, dark=True)
notes(s, "Three questions, and the defaults I will follow if I hear nothing back. The risks are the "
         "ones I can actually mitigate: throttling and VRAM get measured rather than hidden. What "
         "arrives on 7 November is written down so it can be held against me.")

# --------------------------------------------------------------------- save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
