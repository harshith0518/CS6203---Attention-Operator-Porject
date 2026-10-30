"""Generate the 7-slide OP2 deck.

Built to the reviewed spec: fixed zones so titles cannot collide with content,
provenance tag on every number, speaker notes on every slide, no shadows, and
no grading weights in front of the mentors.

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
CODEPANEL = RGBColor(0xF3, 0xF7, 0xF8)

HEAD = "Cambria"
BODY = "Calibri"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)
ML, MR = Inches(0.7), Inches(0.7)
CW = W - ML - MR                    # 11.933
EYEBROW_Y = Inches(0.5)
TITLE_Y, TITLE_H = Inches(0.85), Inches(1.15)
CONTENT_Y = Inches(2.25)
FOOTER_Y = Inches(6.95)


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def alt(shape, text):
    """Alt text, required by the spec."""
    cNvPr = shape._element._nvXxPr.cNvPr
    cNvPr.set("descr", text)


def tb(slide, x, y, w, h, lines, size=16, bold=False, color=INK, font=BODY,
       align=PP_ALIGN.LEFT, space=2, line=1.14, alt_text=None):
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
    s.shadow.inherit = False        # writes an empty effectLst: no shadow at all
    if alt_text:
        alt(s, alt_text)
    return s


def eyebrow(slide, s, dark=False):
    tb(slide, ML, EYEBROW_Y, Inches(10.5), Inches(0.28), s.upper(), size=12, bold=True,
       color=MINT if dark else TEAL)


def title(slide, s, dark=False, size=28):
    tb(slide, ML, TITLE_Y, CW, TITLE_H, s, size=size, bold=True,
       color=PAPER if dark else INK, font=HEAD, line=1.05)


def footer(slide, s, page, dark=False):
    tb(slide, ML, FOOTER_Y, Inches(10.0), Inches(0.3), s, size=12,
       color=RGBColor(0x9F, 0xBD, 0xC2) if dark else MUTED)
    tb(slide, W - MR - Inches(1.2), FOOTER_Y, Inches(1.2), Inches(0.3), str(page),
       size=11, color=RGBColor(0x9F, 0xBD, 0xC2) if dark else MUTED,
       align=PP_ALIGN.RIGHT)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def shade(run, colour=TEAL):
    """Emphasise a code line: bold + accent colour.

    An a:highlight band was the first choice, but it made PowerPoint refuse to
    open the file while every schema check passed. Cosmetic XML surgery is not
    worth an unopenable deck, so the emphasis is structural instead.
    """
    run.font.bold = True
    run.font.color.rgb = colour


def code(slide, x, y, w, h, header, lines, accent, mark=(), size=12, alt_text=None):
    """Flat code panel: fill + 3 pt top rule + monospace body. No shadow."""
    panel = rect(slide, x, y, w, h, CODEPANEL, alt_text or ("code panel: " + header))
    head_bar = rect(slide, x, y, w, Inches(0.3), accent)
    tb(slide, x + Inches(0.16), y + Inches(0.04), w - Inches(0.3), Inches(0.24),
       header, size=12, bold=True, color=PAPER)
    box = slide.shapes.add_textbox(x + Inches(0.14), y + Inches(0.4),
                                   w - Inches(0.24), h - Inches(0.5))
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
        if i in mark:
            try:
                shade(r)
            except Exception:
                r.font.bold = True
    return panel


def table(slide, cols, rows, y=CONTENT_Y, widths=None, size=14, head_size=14,
          rowh=0.6, headh=0.44, accent_cols=None, alt_text=None):
    shape = slide.shapes.add_table(len(rows) + 1, len(cols), ML, y, CW,
                                   Inches(headh + rowh * len(rows)))
    t = shape.table
    if widths:
        for i, f in enumerate(widths):
            t.columns[i].width = int(CW * f)
    t.rows[0].height = Inches(headh)
    for r in range(1, len(rows) + 1):
        t.rows[r].height = Inches(rowh)
    for c, s in enumerate(cols):
        cell = t.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PAPER
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.vertical_anchor = MSO_ANCHOR.BOTTOM
        r = cell.text_frame.paragraphs[0].add_run()
        r.text = s
        r.font.size = Pt(head_size)
        r.font.bold = True
        r.font.name = BODY
        r.font.color.rgb = accent_cols[c] if accent_cols and c < len(accent_cols) and accent_cols[c] else INK
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = PANEL if ri % 2 == 0 else PAPER
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            parts = val if isinstance(val, list) else (
                [val] if isinstance(val, tuple) else [(val, size, ci == 0, INK)])
            for k, (txt, sz, bold, col) in enumerate(parts):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.line_spacing = 1.05
                r = p.add_run()
                r.text = txt
                r.font.size = Pt(sz)
                r.font.bold = bold
                r.font.name = BODY
                r.font.color.rgb = col
    if alt_text:
        alt(shape, alt_text)
    return t


def tile_motif(slide, x, y, cell=0.26, gap=0.06, rows=4, cols=4,
               lit=((1, 1), (2, 2)), dark=True):
    for r in range(rows):
        for c in range(cols):
            on = (r, c) in lit
            fill = TEAL if on else RGBColor(0x1B, 0x4B, 0x54)
            rect(slide, Inches(x + c * (cell + gap)), Inches(y + r * (cell + gap)),
                 Inches(cell), Inches(cell), fill)


prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# ================================================================== 1 title
s = new(prs)
rect(s, 0, 0, W, H, DEEP, "dark background")
tile_motif(s, 10.6, 5.4, cell=0.3, gap=0.07, lit=((0, 0), (2, 3)))
eyebrow(s, "CS-6023 GPU kernel programming  \u00b7  Project OP2", dark=True)
title(s, "Attention, written two ways", dark=True, size=54)
tb(s, ML, Inches(2.6), Inches(9.2), Inches(0.9),
   "One fused attention kernel in CUDA C++ and in cuTile, compared on cost and payoff.",
   size=20, color=RGBColor(0xC4, 0xDD, 0xDF), line=1.3)
tb(s, ML, Inches(6.05), Inches(11.2), Inches(0.8),
   ["Surya Harshith (cs23b010)  \u00b7  for Pravin Jagtap and Sumanta Datta  \u00b7  30 October 2026",
    "github.com/harshith0518/CS6203---Attention-Operator-Porject"],
   size=12, color=RGBColor(0x8F, 0xB3, 0xB9), line=1.5)
notes(s, "This is the plan for the attention project: one fused kernel in each model, and a "
         "measured comparison. I am sorry this is late; slide 7 says where I actually am.")

# ================================================================== 2 why fuse
s = new(prs)
eyebrow(s, "Why fuse")
title(s, "Unfused attention is limited by moving an N\u00d7N matrix, not by arithmetic.")
tb(s, ML, CONTENT_Y, CW, Inches(0.3),
   "1.07 GB  \u00b7  S and P written out and read back  \u00b7  computed", size=14, color=ORANGE)
rect(s, ML, Inches(2.62), CW, Inches(0.4), ORANGE, "unfused traffic bar, full width")
tb(s, ML, Inches(3.22), CW, Inches(0.3),
   "17 MB  \u00b7  only Q, K, V and O move; S stays on-chip  \u00b7  computed", size=14, color=TEAL)
rect(s, ML, Inches(3.6), int(CW / 64), Inches(0.4), TEAL, "fused traffic bar, 1/64 width")
tb(s, ML, Inches(4.16), CW, Inches(0.3),
   "bars drawn to scale  \u00b7  N=4096, 8 heads, D=64, fp16  \u00b7  64\u00d7 less traffic",
   size=12, color=MUTED)

stats = [("32 FLOP/byte", "unfused  (computed)", ORANGE),
         ("2,048 FLOP/byte", "fused  (computed)", TEAL),
         ("ridge \u2248 104", "above this, compute-bound  (assumed)", INK)]
for i, (big, sub, col) in enumerate(stats):
    x = ML + i * Inches(4.0)
    tb(s, x, Inches(4.65), Inches(3.7), Inches(0.5), big, size=34, bold=True, color=col, font=HEAD)
    tb(s, x, Inches(5.2), Inches(3.7), Inches(0.3), sub, size=12, color=MUTED)

tb(s, ML, Inches(5.7), CW, Inches(0.3),
   "* peak FLOP/s \u00f7 192 GB/s, assuming a 20 TFLOP/s tensor peak. I will measure the real peak.",
   size=12, color=MUTED)
tb(s, ML, Inches(6.12), CW, Inches(0.6),
   "Prediction, not a result: traffic falls 64\u00d7, but time falls by at most about 3\u00d7, because "
   "compute becomes the limit once memory is.", size=15, color=INK, line=1.25)
footer(s, "The card's memory is GDDR6. The brief says HBM.", 2)
notes(s, "Four passes over a 268 MB matrix is the 1.07 GB. The fused kernel never writes it. "
         "Do not read 64 times as a speedup: both time figures are roofline bounds, and I will "
         "measure the real ones.")

# ============================================================== 3 who owns what
s = new(prs)
eyebrow(s, "The two models")
title(s, "Same mathematics, different division of labour.")
table(s, ["", "CUDA C++  \u00b7  thread level", "cuTile  \u00b7  tile level"],
      [[("I write", 14, True, INK), ("the program for one thread", 14, False, INK),
        ("operations on whole tiles", 14, False, INK)],
       [("I decide", 14, True, INK),
        ("block shape, shared-memory layout, async copies, syncs", 14, False, INK),
        ("tile shape (BQ, BK, D) and which loops run", 14, False, INK)],
       [("Compiler decides", 14, True, INK),
        ("register allocation, instruction scheduling", 14, False, INK),
        ("thread mapping, staging, shared-memory layout, pipelining", 14, False, INK)]],
      y=CONTENT_Y, widths=[0.17, 0.42, 0.41], size=14, rowh=0.82,
      accent_cols=[None, ORANGE, TEAL],
      alt_text="table comparing what the programmer and compiler decide in each model")
tb(s, ML, Inches(5.55), CW, Inches(0.9),
   ["The middle row is what I can tune. The bottom row is what I can only inspect: ncu, and "
    "section Tile for cuTile.",
    "A tile is a 2-D block of values, with power-of-two sides, that the kernel treats as one value."],
   size=15, color=INK, line=1.3)
tb(s, ML, Inches(6.55), CW, Inches(0.3),
   "The project asks what the bottom-right cell costs and buys.", size=15, bold=True, color=TEAL)
notes(s, "In CUDA I choose everything, down to the shared-memory layout. In cuTile the compiler "
         "chooses it, and I choose tile shapes. The project is about what that bottom-right cell "
         "costs and what it buys.")

# ================================================================== 4 variants
s = new(prs)
eyebrow(s, "Six common variants")
title(s, "One new problem per variant. I build 3, with 2 as a flag.")
rows = [
    [[("1  Vanilla SDPA", 14, True, INK), ("the baseline fusion replaces", 12, False, MUTED)],
     ("N\u00d7N matrix round-trips memory", 14, False, INK),
     ("3 launches: GEMM, softmax, GEMM", 14, False, INK),
     ("same 3 kernels; no help", 14, False, INK),
     ("baseline: PyTorch MATH", 14, False, MUTED)],
    [[("2  Causal", 14, True, INK), ("GPT-style decoders", 12, False, MUTED)],
     ("skip future tiles, mask the diagonal", 14, False, INK),
     ("loop bound + if; uneven block work", 14, False, INK),
     ("loop bound + ct.where", 14, False, INK),
     ("flag in my kernel", 14, True, TEAL)],
    [[("3  Fused + online softmax", 14, True, TEAL), ("FlashAttention", 12, False, MUTED)],
     ("softmax over a row seen in slices", 14, False, INK),
     ("rescale accumulators by hand", 14, False, INK),
     ("rescale is one tile multiply", 14, False, INK),
     ("BUILD", 14, True, TEAL)],
    [[("4  GQA / MQA", 14, True, INK), ("Llama 3", 12, False, MUTED)],
     ("query heads share one K/V head", 14, False, INK),
     ("index math; choose who shares the smem tile", 14, False, INK),
     ("group stacked into a tile dimension, to verify", 14, False, INK),
     ("stretch, after correctness", 14, False, MUTED)],
    [[("5  Sliding window", 14, True, INK), ("Mistral 7B", 12, False, MUTED)],
     ("band-shaped loop bounds", 14, False, INK),
     ("two bounds + edge masks", 14, False, INK),
     ("same two bounds as range()", 14, False, INK),
     ("read only", 14, False, MUTED)],
    [[("6  Cross-attention", 14, True, INK), ("T5, Whisper", 12, False, MUTED)],
     ("Nq \u2260 Nkv: two tilings", 14, False, INK),
     ("natural: one row per block", 14, False, INK),
     ("a 64-row tile on 1 query wastes 63/64", 14, False, INK),
     ("read only; CUDA may win", 14, False, MUTED)],
]
table(s, ["Variant", "New problem", "CUDA C++", "cuTile", "Status"], rows,
      y=Inches(2.15), widths=[0.23, 0.22, 0.21, 0.21, 0.13], size=14, rowh=0.63,
      headh=0.36, alt_text="six attention variants with their expression in each model")
footer(s, "Cells are hypotheses from reading the docs. Only row 3 and the causal flag get numbers.", 4)
notes(s, "These are the six I would call common. I only build row 3, with causal as a flag. "
         "Everything in the CUDA and cuTile columns is my reading, not yet a measurement. "
         "Row 6 is where I expect the hand-written version to win.")

# ==================================================================== 5 kernel
s = new(prs)
eyebrow(s, "The kernel")
title(s, "The rescale is a register loop in CUDA and one multiply in cuTile.", size=26)

code(s, ML, CONTENT_Y, Inches(5.85), Inches(3.7), "CUDA C++  \u00b7  sketch", [
    "// each thread owns 2 rows of the warp's 16-row tile",
    "float m[2]={-1e30f,-1e30f}, l[2]={0,0};",
    "float acc[D/8][4] = {};      // 32 fp32",
    "for (int k = 0; k < nk; ++k) {",
    "  load_kv(k); __syncthreads();",
    "  mma_qk(s, q_frag, k);      // S = Q K^T",
    "  for (int r = 0; r < 2; ++r) {",
    "    float x = rowmax(s, r) * scale;",
    "    x = fmaxf(x, __shfl_xor_sync(~0u,x,1));",
    "    x = fmaxf(x, __shfl_xor_sync(~0u,x,2));",
    "    float mn = fmaxf(m[r], x);",
    "    float a  = exp2f(m[r] - mn);   // rescale",
    "    l[r] *= a; m[r] = mn;",
    "    for (int n=0;n<D/8;++n) acc[n][2*r]*=a;",
    "  }",
    "}",
], ORANGE, mark=(11, 13))

code(s, Inches(6.78), CONTENT_Y, Inches(5.85), Inches(3.7), "cuTile  \u00b7  sketch", [
    "bh, qi = ct.bid(1), ct.bid(0)",
    "q = ct.load(Q, (bh, qi, 0), (1, BQ, D))",
    "m = full(-1e30); l = zeros; acc = zeros",
    "for k in range(num_kv_tiles):",
    "    kt = ct.load(K, (bh, k, 0), (1, BK, D))",
    "    s = ct.mma(q, ct.transpose(kt)) * scale",
    "    s = ct.where(keep(qi, k), s, -1e30)",
    "    m_new = ct.maximum(m, ct.max(s, axis=2))",
    "    p = ct.exp2(s - ct.expand_dims(m_new, 2))",
    "    a = ct.exp2(m - m_new)     # rescale",
    "    l = l * a + ct.sum(p, axis=2)",
    "    acc = ct.mma(p.astype(f16), load_v(k),",
    "                 acc * ct.expand_dims(a, 2))",
    "    m = m_new",
    "ct.store(O, (bh, qi, 0), acc / ct.expand_dims(l, 2))",
], TEAL, mark=(9, 12))

tb(s, ML, Inches(6.08), CW, Inches(0.3),
   "Both sides scale before the max, mask the tail, and rescale before accumulating. "
   "Left: 32 registers per thread by hand. Right: acc * a.", size=15, color=INK)
footer(s, "Sketches, not compiled. cuTile names are abbreviated and unverified until it runs.", 5)
notes(s, "Left is what I write per thread, right is what I write per tile. Both scale before the "
         "max; skip that and the denominator underflows. The tail mask is in both. Whether the "
         "one-line version costs speed is what the profiler will tell me.")

# =================================================================== 6 evidence
s = new(prs)
eyebrow(s, "Evidence")
title(s, "The tests are built to fail: tail-heavy sizes, \u00d7100 inputs, one timing protocol.")
cols = [("CORRECTNESS", [
    "fp64 PyTorch reference. Pass: max|out \u2212 ref| \u2264 5e-3 \u00b7 max|ref|, no NaN or Inf.",
    "Sizes 1, 65, 130, 1000 are tail-heavy: a missing tail mask gives 1e-2 error at N=1000, "
    "but 1e-4 at N=4095, under tolerance.",
    "Inputs \u00d7100: the scaling-order bug turns into NaN."]),
    ("TIMING", [
        "One protocol for both models: at least 1 s per config, L2 flushed, p50 and p90.",
        "Clocks, power and throttle state logged per row, cross-checked against Nsight kernel time.",
        "Reference: PyTorch SDPA, every available backend. Target N=2048, assumed."]),
    ("PROFILING", [
        "DRAM bytes, fused against unfused: that tests the 64\u00d7 traffic prediction.",
        "Register spills, plus section Tile for cuTile.",
        "N=16384: unfused needs 8.6 GB for S and P; the card has about 3.8 GB."])]
for i, (hdr, items) in enumerate(cols):
    x = ML + i * Inches(4.0)
    tb(s, x, CONTENT_Y, Inches(3.7), Inches(0.3), hdr, size=12, bold=True, color=TEAL)
    rect(s, x, Inches(2.58), Inches(3.7), Pt(2), TEAL)
    box = s.shapes.add_textbox(x, Inches(2.78), Inches(3.6), Inches(3.4))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for j, t in enumerate(items):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.space_after = Pt(12)
        p.line_spacing = 1.2
        r = p.add_run()
        r.text = t
        r.font.size = Pt(15)
        r.font.name = BODY
        r.font.color.rgb = INK
footer(s, "Error figures from a numpy fp64 check, N(0,1) inputs, D=64. "
          "The only measured number in the deck.", 6)
notes(s, "I want every test to be able to fail. The surprise is that the obvious sizes miss a "
         "one-key tail bug, because the error sits under my tolerance, so I use tail-heavy "
         "sizes. Timing uses one protocol for both models, and I log the GPU clocks because this "
         "is a laptop.")

# ==================================================================== 7 status
s = new(prs)
rect(s, 0, 0, W, H, DEEP, "dark background")
eyebrow(s, "Where I am", dark=True)
title(s, "Plan fixed, kernels not started, eight days left.", dark=True, size=28)

rows = [["Sat 31 Oct", "bench.py: fp64 + SDPA reference rows", "one command writes the CSV"],
        ["Sun 1 Nov", "cuTile fused kernel, causal flag", "matches fp64 on all test sizes"],
        ["Mon 2 Nov", "CUDA kernel v1, single-buffered", "same tolerance"],
        ["Tue 3 Nov", "tile/block sweep, ablations", "best config per model"],
        ["Wed 4 Nov", "Nsight Compute on both", "profiles + DRAM figure"],
        ["Thu 5 \u2013 Fri 6 Nov", "report: five axes + recommendation", "no placeholders"],
        ["Sat 7 Nov", "submit", "\u2014"]]
shape = s.shapes.add_table(len(rows) + 1, 3, ML, CONTENT_Y, Inches(7.7),
                           Inches(0.34 + 0.5 * len(rows)))
t = shape.table
for i, f in enumerate([0.22, 0.44, 0.34]):
    t.columns[i].width = int(Inches(7.7) * f)
t.rows[0].height = Inches(0.34)
for r in range(1, len(rows) + 1):
    t.rows[r].height = Inches(0.5)
for c, hdr in enumerate(["Date", "Work", "Done when"]):
    cell = t.cell(0, c)
    cell.fill.solid()
    cell.fill.fore_color.rgb = DEEP
    cell.margin_left = Inches(0.07)
    r0 = cell.text_frame.paragraphs[0].add_run()
    r0.text = hdr
    r0.font.size = Pt(12)
    r0.font.bold = True
    r0.font.name = BODY
    r0.font.color.rgb = MINT
for ri, row in enumerate(rows, start=1):
    for ci, val in enumerate(row):
        cell = t.cell(ri, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x12, 0x44, 0x4D) if ri % 2 else RGBColor(0x0F, 0x3C, 0x45)
        cell.margin_left = Inches(0.07)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        r0 = cell.text_frame.paragraphs[0].add_run()
        r0.text = val
        r0.font.size = Pt(13)
        r0.font.name = BODY
        r0.font.bold = ci == 0
        r0.font.color.rgb = PAPER if ci == 0 else RGBColor(0xD3, 0xE5, 0xE8)
alt(shape, "seven-row schedule from 31 October to 7 November")

tb(s, ML, Inches(6.2), Inches(7.7), Inches(0.6),
   "Open today, all mine: finish the .venv install, install the CUDA Toolkit, "
   "enable GPU counters for ncu.", size=13, color=RGBColor(0xF0, 0xC0, 0xA8), line=1.25)

px, pw = Inches(8.75), Inches(3.9)
rect(s, px, CONTENT_Y, pw, Inches(4.4), RGBColor(0x11, 0x40, 0x49), "asks panel")
tb(s, px + Inches(0.28), CONTENT_Y + Inches(0.22), pw - Inches(0.5), Inches(0.3),
   "Three answers I need", size=15, bold=True, color=MINT)
tb(s, px + Inches(0.28), CONTENT_Y + Inches(0.64), pw - Inches(0.5), Inches(2.6),
   ["1.  Attention only, no GEMM. The brief says \u201cchoose either\u201d but lists artifacts for "
    "both. Can you confirm in writing?",
    "2.  Does 07 November hold?",
    "3.  Target sequence length: I will assume N=2048 unless you say otherwise."],
   size=13, color=PAPER, line=1.35, space=10)
tb(s, px + Inches(0.28), CONTENT_Y + Inches(3.35), pw - Inches(0.5), Inches(0.9),
   ["GQA stays stretch unless you want it earlier. With no reply by Mon 2 Nov, I proceed with "
    "these defaults."], size=12, color=RGBColor(0x9F, 0xBD, 0xC2), line=1.3)

footer(s, "github.com/harshith0518/CS6203---Attention-Operator-Porject", 7, dark=True)
notes(s, "I am behind the brief's milestones. This is the recovery plan: one kernel per model, "
         "measured properly. The three questions on the right are what could change that plan, "
         "and the defaults apply if I do not hear back.")

# --------------------------------------------------------------------- save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
