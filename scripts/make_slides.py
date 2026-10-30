"""Generate the OP2 survey deck for the mentor meeting.

Twelve slides. One idea each. Every slide carries a visual element (table, diagram,
chart, code panel or stat block) — no bare bullet lists — and the comparison slide
shows *real* CUDA C++ against cuTile, which is what Pravin actually asked for.

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py

Output: docs/op2-attention-slides.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ------------------------------------------------------------------ palette
# Teal Trust, with the dark tone carrying the title and question slides.
TEAL = RGBColor(0x02, 0x80, 0x90)
DEEP = RGBColor(0x0B, 0x3B, 0x43)      # dark ground for sandwich slides
SEA = RGBColor(0x00, 0xA8, 0x96)
MINT = RGBColor(0x02, 0xC3, 0x9A)
INK = RGBColor(0x14, 0x24, 0x2C)
MUTED = RGBColor(0x5A, 0x6B, 0x74)
PAPER = RGBColor(0xFF, 0xFF, 0xFF)
PANEL = RGBColor(0xEF, 0xF6, 0xF7)
CODEBG = RGBColor(0x11, 0x24, 0x2B)
CODEFG = RGBColor(0xD8, 0xEC, 0xEE)
WARN = RGBColor(0xB4, 0x4A, 0x0B)

HEAD = "Cambria"
BODY = "Calibri"
MONO = "Courier New"

W, H = Inches(13.333), Inches(7.5)


# ------------------------------------------------------------------ helpers
def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def tb(slide, x, y, w, h, lines, size=16, bold=False, color=INK, font=BODY,
       align=PP_ALIGN.LEFT, space=8, line=1.15):
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
    return box


def bullet(p, char="\u2013", indent=0.26, color=None):
    """Real bullet via a:buChar — never a literal bullet character in the text."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent))))
    pPr.set("indent", str(-int(Inches(indent))))
    if color is not None:
        buClr = pPr.makeelement(qn("a:buClr"), {})
        srgb = pPr.makeelement(qn("a:srgbClr"), {"val": str(color)})
        buClr.append(srgb)
        pPr.append(buClr)
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": char}))


def listy(slide, x, y, w, items, size=16, gap=12):
    """items: (level, text) or (level, text, colour)."""
    box = slide.shapes.add_textbox(x, y, w, Inches(4.4))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        lvl, s = it[0], it[1]
        col = it[2] if len(it) > 2 else INK
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap if lvl == 0 else max(4, gap - 5))
        p.line_spacing = 1.12
        if lvl == 0:
            bullet(p, color="028090")
        else:
            bullet(p, char="\u00b7", indent=0.5, color="5A6B74")
        r = p.add_run()
        r.text = s
        r.font.size = Pt(size - (1 if lvl else 0))
        r.font.bold = lvl == 0
        r.font.color.rgb = col
        r.font.name = BODY
    return box


def tile_motif(slide, x, y, cell=0.26, gap=0.06, rows=3, cols=3,
               filled=((0, 0), (1, 1), (2, 2)), color=TEAL):
    """The repeated visual motif: a grid of tiles, a few of them 'active'."""
    for r in range(rows):
        for c in range(cols):
            s = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(x + c * (cell + gap)), Inches(y + r * (cell + gap)),
                Inches(cell), Inches(cell))
            s.fill.solid()
            s.fill.fore_color.rgb = color if (r, c) in filled else PANEL
            s.line.fill.background()
            s.shadow.inherit = False


def head(slide, kicker, title, dark=False):
    """Title block. No rule, no stripe — the tile motif is the marker."""
    fg = PAPER if dark else INK
    tb(slide, Inches(0.75), Inches(0.55), Inches(10.6), Inches(0.3),
       kicker.upper(), size=11.5, bold=True, color=SEA if dark else TEAL)
    tb(slide, Inches(0.75), Inches(0.9), Inches(11.2), Inches(1.3),
       title, size=38, bold=True, color=fg, font=HEAD, line=1.02)
    tile_motif(slide, 11.95, 0.62, cell=0.17, gap=0.045,
               color=SEA if dark else TEAL)


def table(slide, cols, rows, y=Inches(2.35), widths=None, size=12, rowh=0.5,
          headh=0.42, x=Inches(0.75), total=Inches(11.85)):
    shape = slide.shapes.add_table(len(rows) + 1, len(cols), x, y, total,
                                   Inches(headh + rowh * len(rows)))
    t = shape.table
    if widths:
        for i, f in enumerate(widths):
            t.columns[i].width = int(total * f)
    t.rows[0].height = Inches(headh)
    for r in range(1, len(rows) + 1):
        t.rows[r].height = Inches(rowh)

    for c, s in enumerate(cols):
        cell = t.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DEEP
        cell.margin_left = cell.margin_right = Inches(0.09)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = s
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = PAPER
        r.font.name = BODY

    for ri, row in enumerate(rows, start=1):
        for ci, s in enumerate(row):
            cell = t.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = PANEL if ri % 2 else PAPER
            cell.margin_left = cell.margin_right = Inches(0.09)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = str(s)
            r.font.size = Pt(size - 0.5)
            r.font.name = BODY
            r.font.bold = ci == 0
            r.font.color.rgb = INK
    return t


def code(slide, x, y, w, h, lines, size=10.5, title=None):
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    panel.fill.solid()
    panel.fill.fore_color.rgb = CODEBG
    panel.line.fill.background()
    panel.shadow.inherit = False
    panel.adjustments[0] = 0.03
    ty = y
    if title:
        tb(slide, x + Inches(0.22), y + Inches(0.14), w - Inches(0.4), Inches(0.28),
           title, size=11.5, bold=True, color=MINT)
        ty = y + Inches(0.42)
    box = slide.shapes.add_textbox(x + Inches(0.22), ty, w - Inches(0.42),
                                   h - (ty - y) - Inches(0.1))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.line_spacing = 1.04
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.name = MONO
        r.font.color.rgb = MUTED if ln.strip().startswith("//") else CODEFG


def note(slide, s, y=Inches(6.62), color=MUTED, size=12.5):
    tb(slide, Inches(0.75), y, Inches(11.4), Inches(0.6), s,
       size=size, color=color, line=1.15)


def arrow(slide, x, y, w, h, color=SEA):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.fill.background()
    a.shadow.inherit = False
    return a


def chip(slide, x, y, w, h, text, fill=TEAL, fg=PAPER, size=13, bold=True):
    c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    c.fill.solid()
    c.fill.fore_color.rgb = fill
    c.line.fill.background()
    c.shadow.inherit = False
    c.adjustments[0] = 0.18
    tf = c.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.06)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = fg
    r.font.name = BODY
    return c


# ------------------------------------------------------------------ deck
prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# 1 ---------------------------------------------------------------- title
s = new(prs)
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
bg.fill.solid()
bg.fill.fore_color.rgb = DEEP
bg.line.fill.background()
bg.shadow.inherit = False

tile_motif(s, 9.85, 1.15, cell=0.42, gap=0.09, rows=5, cols=5,
           filled=((0, 0), (1, 1), (2, 2), (3, 3), (4, 4), (1, 3), (3, 1)),
           color=SEA)

tb(s, Inches(0.85), Inches(1.75), Inches(9.4), Inches(0.35),
   "CS-6023 GPU PROGRAMMING  \u00b7  PROJECT OP2", size=13, bold=True, color=MINT)
tb(s, Inches(0.85), Inches(2.25), Inches(9.6), Inches(2.0),
   ["Attention, written", "two ways"], size=54, bold=True, color=PAPER,
   font=HEAD, line=1.0)
tb(s, Inches(0.85), Inches(4.15), Inches(9.2), Inches(1.1),
   ["Six attention variants, one new kernel-engineering problem each \u2014",
    "and what each problem costs in CUDA C++ versus cuTile."],
   size=18, color=RGBColor(0xC8, 0xDE, 0xE2), line=1.28)
tb(s, Inches(0.85), Inches(5.85), Inches(9.4), Inches(0.9),
   ["Surya Harshith (cs23b010)  \u00b7  for Pravin Jagtap and Sumanta Datta",
    "30 October 2026"], size=12.5, color=RGBColor(0x93, 0xB4, 0xBB), line=1.35)

# 2 ---------------------------------------------------------------- argument
s = new(prs)
head(s, "the argument", "Six variants, one new problem each")
table(s, ["#", "Variant", "The one new problem", "CUDA C++", "cuTile"],
      [["1", "Vanilla SDPA", "why fusion is needed at all", "low", "low"],
       ["2", "Causal", "masking \u2014 and skipping whole tiles", "medium", "low"],
       ["3", "Fused + online softmax", "streaming softmax, live accumulator rescale", "high", "medium"],
       ["4", "GQA / MQA", "head broadcast; decode is memory-bound", "medium", "very low"],
       ["5", "Sliding window", "sparse tile traversal, ragged grids", "med-high", "low-med"],
       ["6", "Cross-attention", "asymmetric shapes \u2014 the one rung CUDA wins", "low-med", "low"]],
      widths=[0.05, 0.22, 0.41, 0.16, 0.13], size=12.5, rowh=0.54)
note(s, "These cost ratings are hypotheses, not measurements. Only #3 is implemented before "
        "the deadline; the rest are the framing.")

# 3 ---------------------------------------------------------------- dataflow
s = new(prs)
head(s, "the operator", "What attention computes")
bw, bh = Inches(1.9), Inches(0.82)
y0 = Inches(2.6)
step, gap_a = Inches(2.25), Inches(0.32)
chips = ["Q, K   [N \u00d7 D]", "S = Q K^T / \u221ad", "softmax, row-wise",
         "P  [N \u00d7 N]", "O = PV   [N \u00d7 D]"]
fills = [(PANEL, INK), (TEAL, PAPER), (SEA, PAPER), (PANEL, INK), (MINT, DEEP)]
for i, (t, (fill, fg)) in enumerate(zip(chips, fills)):
    chip(s, Inches(0.75) + i * step, y0, bw, bh, t, fill=fill, fg=fg, size=13.5)
    if i < 4:
        arrow(s, Inches(0.75) + i * step + bw + Inches(0.03), y0 + Inches(0.22),
              gap_a, Inches(0.38))

chip(s, Inches(0.75), Inches(4.15), Inches(7.4), Inches(0.72),
     "S and P are N \u00d7 N \u2192  fused: never written to memory",
     fill=WARN, size=15)
chip(s, Inches(8.45), Inches(4.15), Inches(4.15), Inches(0.72),
     "one kernel, not three", fill=PANEL, fg=INK, size=14)

tb(s, Inches(0.75), Inches(5.55), Inches(11.4), Inches(1.1),
   ["FLOPs = 4 \u00b7 B \u00b7 H \u00b7 N\u00b2 \u00b7 D  \u2014 two matmuls, each 2N\u00b2D per head.",
    "Softmax is a row reduction. Causal masks j > i; the mask goes before the softmax."],
   size=15, color=INK, line=1.3)

# 4 ---------------------------------------------------------------- memory bound
s = new(prs)
head(s, "why fusion is not optional", "The naive version is memory-bound")
stats = [("\u2248 1.07 GB", "intermediate traffic\nat N=4096, H=8, fp16", WARN),
         ("\u2248 17 MB", "the same attention, fused", MINT),
         ("32 FLOP/B", "intensity vs a ridge of \u2248104", TEAL)]
for i, (big, sub, col) in enumerate(stats):
    x = Inches(0.75 + i * 4.0)
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.35),
                              Inches(3.55), Inches(2.15))
    card.fill.solid()
    card.fill.fore_color.rgb = PANEL
    card.line.fill.background()
    card.shadow.inherit = False
    card.adjustments[0] = 0.06
    tb(s, x + Inches(0.3), Inches(2.62), Inches(3.0), Inches(0.8), big,
       size=40, bold=True, color=col, font=HEAD)
    tb(s, x + Inches(0.3), Inches(3.52), Inches(3.0), Inches(0.8),
       sub.split("\n"), size=12.5, color=MUTED, line=1.2)

listy(s, Inches(0.75), Inches(4.85), Inches(11.4), [
    (0, "Naive = three kernels: compute S, write it, read it for softmax, write P, read P for PV."),
    (0, "The brief says HBM; this laptop has GDDR6. Same measurement \u2014 device/DRAM bytes."),
    (1, "Peaks here are nominal. Measured peaks are step 1 of the benchmark plan, because "
        "the ridge point moves with them."),
], size=15, gap=11)
note(s, "Traffic is O(N\u00b2\u00b7D). No compiler trick saves you \u2014 only not writing the matrix does.")

# 5 ---------------------------------------------------------------- variants 1-2
s = new(prs)
head(s, "variants 1 and 2", "The baseline, and the mask that must also skip work")
table(s, ["", "Vanilla SDPA", "Causal"],
      [["New problem", "none \u2014 it exists to be beaten",
        "a mask is trivial; not doing the masked work is not"],
       ["CUDA C++", "3 launches, row-softmax with two passes",
        "bound the k-loop at the diagonal + an if on the diagonal tile \u2192 unequal block runtimes"],
       ["cuTile", "3 @ct.kernel; ct.max/ct.exp2/ct.sum over a tile",
        "ct.where(cols <= rows) \u2014 the [BQ\u00d7BK] mask is a broadcast; loop bound is a dynamic range"],
       ["Cost", "two N\u00d7N tensors round-trip through memory",
        "FLOPs halve \u2014 only if you skip the tiles"]],
      widths=[0.14, 0.42, 0.44], size=12.5, rowh=0.92)
note(s, "Neither model is saved by its abstraction here. Vanilla is the reference point, "
        "not something anyone ships.")

# 6 ---------------------------------------------------------------- algorithm
s = new(prs)
head(s, "variant 3 \u2014 the algorithm",
     "Online softmax: a softmax you can compute in slices")
code(s, Inches(0.75), Inches(2.3), Inches(7.55), Inches(3.75), [
    "m = -1e30 (finite!), l = 0, O = 0     // fp32",
    "for each KV tile k:",
    "    S_k   = Q_tile @ K_kT * scale     // scale first",
    "    m_new = max(m, rowmax(S_k))",
    "    P_k   = exp2(S_k - m_new)",
    "    a     = exp2(m - m_new)           // rescale",
    "    l     = l*a + rowsum(P_k)",
    "    O     = O*a + P_k @ V_k           // a BEFORE",
    "    m     = m_new",
    "O = O / l",
], size=11.5, title="the whole algorithm")
listy(s, Inches(8.6), Inches(2.4), Inches(4.0), [
    (0, "Softmax is a global row reduction \u2014 but a tile holds one slice of the row."),
    (0, "So the running max and denominator must be correct under rescaling. That is what a does."),
    (1, "Milakov 2018 for the recurrence; Dao 2022 for fusing it into PV."),
], size=14, gap=13)

# 7 ---------------------------------------------------------------- comparison
s = new(prs)
head(s, "variant 3 \u2014 two implementations", "The same algorithm, two costs")
code(s, Inches(0.75), Inches(2.25), Inches(6.05), Inches(3.55), [
    "__shared__ half sk[2][BK][D+8];   // padded",
    "float m[16], l[16], acc[16][VP];  // 32/thread",
    "for (int k = 0; k < nk; ++k) {",
    "  cp_async(sk[k&1], Kg + k*BK*D);",
    "  __syncthreads();",
    "  mma_m16n8k16(s, q_frag,",
    "               ldmatrix_x4(sk[k&1]), 0.f);",
    "  for (int t = 0; t < 16; ++t) {   // rows",
    "    float mx = rowmax(s[t]) * scale;",
    "    mx = __shfl_xor_sync(~0u, mx, 1);",
    "    mx = __shfl_xor_sync(~0u, mx, 2);",
    "    float mn = fmaxf(m[t], mx);",
    "    float a  = exp2f(m[t] - mn);",
    "    m[t] = mn; l[t] *= a;",
    "    for (int v = 0; v < VP; ++v)",
    "      acc[t][v] *= a;   // a by hand",
    "  }",
    "}",], size=9.5, title="CUDA C++ \u2014 what you write by hand")

code(s, Inches(7.0), Inches(2.25), Inches(5.6), Inches(3.55), [
    "q = ct.load(Q, (bh, qi, 0),",
    "            (1, BQ, D))",
    "for k in range(last):",
    "  kt = ct.load(K, (bh, k, 0),",
    "               (BK, D))",
    "  v  = ct.load(V, (bh, k, 0),",
    "               (BK, D))",
    "  s = ct.mma(q, ct.transpose(kt),",
    "             zeros) * scale",
    "  if causal:",
    "    s = ct.where(keep, s, -1e30)",
    "  m_new = ct.maximum(m, ct.max(s, 2))",
    "  p = ct.exp2(s - expand(m_new, 2))",
    "  a = ct.exp2(m - m_new)",
    "  l = l*a + ct.sum(p, 2)",
    "  acc = ct.mma(p.half(), v,",
    "               acc * expand(a, 2))",
    "  m = m_new",
], size=9.5, title="cuTile \u2014 what the compiler owns")
note(s, "CUDA C++: you pay in shared memory, registers and warp shuffles. cuTile: you pay in "
        "tile shapes. The abstraction moves the work \u2014 it does not remove it.")

# 8 ---------------------------------------------------------------- payoff
s = new(prs)
head(s, "variant 3 \u2014 the payoff", "Traffic drops 60\u00d7. Time drops about 3\u00d7.")
data = CategoryChartData()
data.categories = ["DRAM traffic", "Latency"]
data.add_series("Fused vs naive, N = 4096", (60.0, 3.2))
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.75), Inches(2.3),
                        Inches(6.6), Inches(3.9), data).chart
ch.has_title = True
ch.chart_title.text_frame.text = "reduction factor (roofline bounds, not measurements)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
ch.has_legend = False
plot = ch.plots[0]
plot.has_data_labels = True
plot.data_labels.show_value = True
plot.data_labels.number_format = '0.0"\u00d7"'
plot.data_labels.number_format_is_linked = False
plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
plot.data_labels.font.size = Pt(14)
plot.data_labels.font.bold = True
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = TEAL
ch.value_axis.visible = False
ch.value_axis.has_major_gridlines = False
ch.category_axis.tick_labels.font.size = Pt(12)
ch.category_axis.format.line.color.rgb = RGBColor(0xD5, 0xE3, 0xE5)

listy(s, Inches(7.6), Inches(2.45), Inches(5.0), [
    (0, "Once the traffic wall is gone you hit the compute wall: 34 GFLOP at ~20 TFLOP/s \u2248 1.7 ms."),
    (0, "These are roofline bounds \u2014 the honest framing is a factor, not a stopwatch result."),
    (0, "Prediction: crossover at N \u2248 512. Below it, launch overhead wins."),
], size=14.5, gap=13)
note(s, "Anyone reporting 60\u00d7 speedup is quoting the traffic number as if it were the time.")

# 9 ---------------------------------------------------------------- GQA
s = new(prs)
head(s, "variant 4", "GQA: one line of index maths vs one broadcast")
tb(s, Inches(0.75), Inches(2.25), Inches(11.4), Inches(0.5),
   "H query heads share H_kv key/value heads.   group G = H / H_kv.   "
   "MQA is H_kv = 1, i.e. G = H;  G = 1 is plain MHA.",
   size=15, color=MUTED)
table(s, ["", "CUDA C++", "cuTile"],
      [["Head map", "int kv_head = q_head / G;  plus hand-written strides",
        "stack the group into the leading tile dimension"],
       ["Reuse", "you decide whether to fuse G query heads into one block so the KV tile "
        "loads into shared memory once \u2014 and you reason about L2 reuse",
        "ct.mma broadcasts a 2-D KV tile across a 3-D query tile"],
       ["Legality", "any G", "G must be a power of two \u2014 unverified"]],
      widths=[0.13, 0.44, 0.43], size=12.5, rowh=1.02, y=Inches(3.0))
note(s, "The most persuasive slide in the deck \u2014 and the first to be cut. Decode is a different "
        "shape (N_q=1 against long M_kv), so it breaks the N=M contract and stays out of scope.")

# 10 --------------------------------------------------------------- knobs
s = new(prs)
head(s, "the tunability axis, on my hardware",
     "Three of the eight knobs do nothing on this GPU")
table(s, ["Knob", "CUDA C++", "cuTile", "Available on sm_86?"],
      [["tile / block shape", "free (any 3-D block)", "BQ, BK, D \u2014 powers of two", "yes"],
       ["threads / warps", "free", "num_worker_warps, occupancy", "yes"],
       ["smem layout, swizzle", "hand-written", "hidden", "yes (CUDA only)"],
       ["async pipelining", "cp.async, double buffer", "hidden (compiler)", "yes (CUDA only)"],
       ["prefetch hint", "manual", "latency = 1..10", "yes"],
       ["TMA", "sm_90+ only", "allow_tma", "no"],
       ["thread-block clusters", "sm_90+ only", "num_ctas", "no"],
       ["FP8 / block-scaled MMA", "Ada+ / Blackwell", "mma_scaled", "no"]],
      widths=[0.26, 0.22, 0.26, 0.26], size=12, rowh=0.46)
note(s, "RTX 3050 Laptop \u2014 compute capability 8.6, Ampere. Three of eight knobs are unreachable "
        "in *either* model; cuTile simply exposes ones CUDA does not name. That is the portability finding.")

# 11 --------------------------------------------------------------- plan
s = new(prs)
head(s, "the plan", "One operator, two kernels, three files, eight days")
table(s, ["Day", "Date", "Work", "Done when"],
      [["0", "Fri 30 Oct", "environment gate + this deck", "smoke test green; deck sent"],
       ["1", "Sat 31 Oct", "bench.py: fp64 + SDPA rows", "one command \u2192 CSV"],
       ["2", "Sun 1 Nov", "fused kernel in cuTile", "matches fp64 within fp16 tolerance"],
       ["3", "Mon 2 Nov", "fused kernel in CUDA C++ (v1, no cp.async)", "both rows in CSV"],
       ["4", "Tue 3 Nov", "tile / block sweep; ablation toggles", "best config recorded"],
       ["5", "Wed 4 Nov", "Nsight Compute, incl. --section Tile", "profiles/ populated"],
       ["6", "Thu 5 Nov", "report: numbers only", "no placeholders"],
       ["7", "Fri 6 Nov", "report final, 3-command README", "reproducible"],
       ["8", "Sat 7 Nov", "submit", "submitted"]],
      widths=[0.06, 0.15, 0.45, 0.34], size=11.5, rowh=0.42, y=Inches(2.25))
tb(s, Inches(0.75), Inches(6.45), Inches(11.4), Inches(0.7),
   "Pre-cut if a day slips: cp.async (v1 is synchronous) \u2192 GQA \u2192 the timed bug hunt. "
   "The fused kernel in both models is the only thing that must exist.",
   size=13, color=MUTED)

# 12 --------------------------------------------------------------- questions
s = new(prs)
bg2 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
bg2.fill.solid()
bg2.fill.fore_color.rgb = DEEP
bg2.line.fill.background()
bg2.shadow.inherit = False
tile_motif(s, 11.1, 4.7, cell=0.34, gap=0.08, rows=4, cols=4,
           filled=((0, 0), (2, 2)), color=TEAL)

tb(s, Inches(0.85), Inches(0.6), Inches(10.0), Inches(0.35),
   "WHAT I NEED FROM THIS MEETING", size=11.5, bold=True, color=MINT)
tb(s, Inches(0.85), Inches(1.0), Inches(10.4), Inches(0.9), "Three questions",
   size=38, bold=True, color=PAPER, font=HEAD)

qs = [("1", "OP2 only \u2014 can you confirm it in writing?",
       "The brief's comparison section says \u201cchoose either OP1 or OP2\u201d; its artifact "
       "list still says \u201cboth operators\u201d."),
      ("2", "Does 07 November hold?",
       "The brief's own 10/21 and 10/28 milestones have passed; we restarted from this "
       "survey on 14 October."),
      ("3", "Is fused MHA the right variant to finish?",
       "Or should GQA come before the deadline? GQA is the cleanest demonstration of the "
       "two models' difference.")]
y = 2.15
for n, q, sub in qs:
    chip(s, Inches(0.85), Inches(y - 0.02), Inches(0.42), Inches(0.42), n,
         fill=SEA, size=14)
    tb(s, Inches(1.5), Inches(y), Inches(10.0), Inches(0.4), q, size=19,
       bold=True, color=PAPER)
    tb(s, Inches(1.5), Inches(y + 0.42), Inches(9.6), Inches(0.5), sub, size=13,
       color=RGBColor(0xA9, 0xC6, 0xCC), line=1.2)
    y += 1.25

tb(s, Inches(0.85), Inches(6.15), Inches(10.6), Inches(0.9),
   ["Not doing: GEMM \u00b7 paged attention \u00b7 MLA \u00b7 backward pass \u00b7 a hand-written naive "
    "baseline", "(PyTorch's MATH backend is that baseline, on my card, for free)."],
   size=12.5, color=RGBColor(0x86, 0xA6, 0xAD), line=1.3)

# ------------------------------------------------------------------ save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
