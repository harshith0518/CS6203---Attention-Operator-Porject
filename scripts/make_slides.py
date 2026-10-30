"""Generate the OP2 progress deck for the mentor conversation.

Narrative, in this order:
  1  title
  2-3  what we have understood (two levels; why fusion decides everything)
  4  what attention is, explained visually, with verified reading links
  5  which attention we chose (contract + the variant ladder)
  6  the plan: the same maths at two levels, real code side by side
  7  what makes a good submission
  8  how we prove it is correct
  9  trajectory, step by step
  10 the repo, and what we need

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py
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
TEAL = RGBColor(0x02, 0x80, 0x90)
DEEP = RGBColor(0x0B, 0x3B, 0x43)
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
REPO = "github.com/harshith0518/CS6203---Attention-Operator-Porject"


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


def bullet(p, char="\u2013", indent=0.26, color="028090"):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent))))
    pPr.set("indent", str(-int(Inches(indent))))
    buClr = pPr.makeelement(qn("a:buClr"), {})
    buClr.append(pPr.makeelement(qn("a:srgbClr"), {"val": color}))
    pPr.append(buClr)
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": char}))


def listy(slide, x, y, w, items, size=16, gap=11, h=4.3, bcolor="028090"):
    box = slide.shapes.add_textbox(x, y, w, Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        lvl, s = it[0], it[1]
        col = it[2] if len(it) > 2 else INK
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap if lvl == 0 else max(3, gap - 5))
        p.line_spacing = 1.12
        if lvl == 0:
            bullet(p, color=bcolor)
        else:
            bullet(p, "\u00b7", 0.5, bcolor)
        r = p.add_run()
        r.text = s
        r.font.size = Pt(size - (1 if lvl else 0))
        r.font.bold = lvl == 0
        r.font.color.rgb = col
        r.font.name = BODY
    return box


def tile_motif(slide, x, y, cell=0.22, gap=0.05, rows=3, cols=3,
               filled=((0, 0), (1, 1), (2, 2)), color=TEAL):
    for r in range(rows):
        for c in range(cols):
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                       Inches(x + c * (cell + gap)),
                                       Inches(y + r * (cell + gap)),
                                       Inches(cell), Inches(cell))
            s.fill.solid()
            s.fill.fore_color.rgb = color if (r, c) in filled else PANEL
            s.line.fill.background()
            s.shadow.inherit = False


def head(slide, kicker, title, dark=False, size=36):
    tb(slide, Inches(0.75), Inches(0.52), Inches(10.4), Inches(0.3),
       kicker.upper(), size=11.5, bold=True, color=MINT if dark else TEAL)
    tb(slide, Inches(0.75), Inches(0.85), Inches(11.3), Inches(1.3), title,
       size=size, bold=True, color=PAPER if dark else INK, font=HEAD, line=1.03)
    tile_motif(slide, 12.0, 0.58, color=SEA if dark else TEAL)


def table(slide, cols, rows, y=Inches(2.3), widths=None, size=12, rowh=0.5,
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
        r = cell.text_frame.paragraphs[0].add_run()
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
            r = tf.paragraphs[0].add_run()
            r.text = str(s)
            r.font.size = Pt(size - 0.5)
            r.font.name = BODY
            r.font.bold = ci == 0
            r.font.color.rgb = INK
    return t


def code(slide, x, y, w, h, lines, size=9.5, title=None):
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    panel.fill.solid()
    panel.fill.fore_color.rgb = CODEBG
    panel.line.fill.background()
    panel.shadow.inherit = False
    panel.adjustments[0] = 0.03
    ty = y + Inches(0.42) if title else y
    if title:
        tb(slide, x + Inches(0.22), y + Inches(0.14), w - Inches(0.4), Inches(0.28),
           title, size=11.5, bold=True, color=MINT)
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
        r.font.color.rgb = MUTED if ln.strip().startswith(("#", "//")) else CODEFG


def note(slide, s, y=Inches(6.62), size=12.5, color=MUTED):
    tb(slide, Inches(0.75), y, Inches(11.4), Inches(0.6), s, size=size, color=color)


def arrow(slide, x, y, w, h, color=SEA):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.fill.background()
    a.shadow.inherit = False


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

# 1 ---- title
s = new(prs)
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
bg.fill.solid(); bg.fill.fore_color.rgb = DEEP
bg.line.fill.background(); bg.shadow.inherit = False
tile_motif(s, 9.9, 1.1, cell=0.42, gap=0.09, rows=5, cols=5,
           filled=((0, 0), (1, 1), (2, 2), (3, 3), (4, 4), (1, 3), (3, 1)), color=SEA)

tb(s, Inches(0.85), Inches(1.65), Inches(9.0), Inches(0.35),
   "CS-6023 GPU PROGRAMMING  \u00b7  PROJECT OP2", size=13, bold=True, color=MINT)
tb(s, Inches(0.85), Inches(2.15), Inches(9.4), Inches(1.9),
   ["Attention, written", "two ways"], size=54, bold=True, color=PAPER, font=HEAD, line=1.0)
tb(s, Inches(0.85), Inches(4.0), Inches(8.9), Inches(1.0),
   ["What we understood  \u00b7  which operator we chose  \u00b7  how we will build it",
    "in CUDA C++ and cuTile  \u00b7  how we will know it is right"],
   size=17, color=RGBColor(0xC8, 0xDE, 0xE2), line=1.3)
chip(s, Inches(0.85), Inches(5.35), Inches(6.9), Inches(0.44), REPO,
     fill=RGBColor(0x12, 0x4E, 0x59), fg=MINT, size=12.5, bold=False)
tb(s, Inches(0.85), Inches(6.05), Inches(9.0), Inches(0.9),
   ["Surya Harshith (cs23b010)  \u00b7  for Pravin Jagtap and Sumanta Datta",
    "30 October 2026"], size=12.5, color=RGBColor(0x93, 0xB4, 0xBB), line=1.35)

# 2 ---- understood 1/2
s = new(prs)
head(s, "what we have understood  \u00b7  1 of 2",
     "The same maths, two very different jobs for the programmer")
table(s, ["", "CUDA C++  \u2014  thread level", "cuTile  \u2014  tile level"],
      [["You write", "code for one thread; you decide which element it owns",
        "operations on whole **tiles**"],
       ["You own", "shared memory, register budget, sync points, warp shuffles",
        "tile shape, grid, and which loop to skip"],
       ["The compiler owns", "nothing \u2014 the mapping is already yours",
        "layout, staging, thread mapping, prefetch depth"],
       ["Cost of a mistake", "silent wrong numbers or a bank-conflict slowdown",
        "a compile error, or a shape that quietly does not fit"]],
      widths=[0.17, 0.42, 0.41], size=12.5, rowh=0.78)
listy(s, Inches(0.75), Inches(5.55), Inches(11.4), [
    (0, "A tile is a chunk of data \u2014 not a thread, and not a kind of memory. A block may hold many."),
    (1, "The question this project answers: what does the compiler hide, and what does that hiding cost?"),
], size=14, gap=6, h=1.0)

# 3 ---- understood 2/2
s = new(prs)
head(s, "what we have understood  \u00b7  2 of 2",
     "Fusion decides everything; the rest is tuning")
cards = [("\u2248 1.07 GB", "intermediate traffic\nif S and P are written out", WARN),
         ("\u2248 17 MB", "the same attention,\nfused into one kernel", MINT),
         ("32 FLOP/B", "arithmetic intensity vs\na ridge of \u2248 104", TEAL)]
for i, (big, sub, col) in enumerate(cards):
    x = Inches(0.75 + i * 4.0)
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.3),
                              Inches(3.55), Inches(2.05))
    card.fill.solid(); card.fill.fore_color.rgb = PANEL
    card.line.fill.background(); card.shadow.inherit = False
    card.adjustments[0] = 0.06
    tb(s, x + Inches(0.3), Inches(2.55), Inches(3.0), Inches(0.8), big,
       size=38, bold=True, color=col, font=HEAD)
    tb(s, x + Inches(0.3), Inches(3.4), Inches(3.0), Inches(0.8),
       sub.split("\n"), size=12.5, color=MUTED, line=1.2)

listy(s, Inches(0.75), Inches(4.75), Inches(11.4), [
    (0, "Naive attention = three kernels, with S and P round-tripping through device memory."),
    (0, "Intensity 32 is below the ridge point, so the naive version is memory-bound \u2014 not compute-bound."),
    (0, "Only fusion removes those tensors. It turns a memory-bound kernel into a compute-bound one."),
    (1, "Which is why the win is ~60\u00d7 less traffic but only ~3\u00d7 less time. We will measure both."),
], size=15, gap=9)

# 4 ---- attention, creatively
s = new(prs)
head(s, "attention, in one picture",
     "\u201cMole\u201d means three things. Attention decides which.", size=34)
tb(s, Inches(0.75), Inches(1.72), Inches(11.4), Inches(0.7),
   ["\u201cAmerican shrew mole\u201d        \u00b7        \u201cone mole of carbon dioxide\u201d        "
    "\u00b7        \u201ca biopsy of the mole\u201d"], size=16, color=INK, line=1.2)
tb(s, Inches(0.75), Inches(2.08), Inches(11.4), Inches(0.5),
   "After embedding, all three are the same vector. Attention is what makes them different again.",
   size=13.5, color=MUTED)

y = Inches(2.75)
bw, bh, step = Inches(2.13), Inches(0.72), Inches(2.43)
chips = [("Q  \u2014 what I seek", TEAL), ("K  \u2014 what might match", SEA),
         ("Q\u00b7K  \u2014 the score", TEAL), ("softmax", SEA),
         ("V  \u2014 what I take", MINT)]
for i, (t, f) in enumerate(chips):
    chip(s, Inches(0.75) + i * step, y, bw, bh, t, fill=f,
         fg=DEEP if f == MINT else PAPER, size=12.5)
    if i < 4:
        arrow(s, Inches(0.75) + i * step + bw + Inches(0.02), y + Inches(0.18),
              Inches(0.26), Inches(0.36))
tb(s, Inches(0.75), Inches(3.62), Inches(11.4), Inches(0.4),
   "Q is the text you type in a search box. K is the page titles. V is the page contents you get back.",
   size=13.5, color=MUTED)

tb(s, Inches(0.75), Inches(4.62), Inches(11.4), Inches(0.3),
   "GOOD PLACES TO LEARN IT", size=11, bold=True, color=TEAL)
links = [["3Blue1Brown \u2014 Deep Learning Ch.6 (the mole example)", "3blue1brown.com/lessons/attention"],
         ["Jay Alammar \u2014 The Illustrated Transformer", "jalammar.github.io/illustrated-transformer"],
         ["Georgia Tech \u2014 Transformer Explainer (live, interactive)", "poloclub.github.io/transformer-explainer"],
         ["Vaswani et al. 2017 \u2014 Attention Is All You Need", "arxiv.org/abs/1706.03762"]]
for i, (name, url) in enumerate(links):
    col, row = i % 2, i // 2
    x = Inches(0.75 + col * 5.95)
    yy = Inches(5.0 + row * 0.9)
    tb(s, x, yy, Inches(5.7), Inches(0.26), name, size=12.5, bold=True, color=INK)
    tb(s, x, yy + Inches(0.24), Inches(5.7), Inches(0.24), url, size=11, color=TEAL)

# 5 ---- which attention
s = new(prs)
head(s, "which attention we chose", "Plain multi-head, fused, exact \u2014 and why that one")
tb(s, Inches(0.75), Inches(1.95), Inches(11.4), Inches(0.4),
   "O = softmax(QK\u1d40/\u221ad + mask) \u00b7 V        B=1, H=8, D=64, fp16 in, fp32 accumulate",
   size=15, color=INK)
table(s, ["#", "Variant", "Take it?", "Why"],
      [["1", "Vanilla SDPA", "as the baseline", "PyTorch's MATH backend already is it \u2014 no code to write"],
       ["2", "Causal", "yes \u2014 as a flag", "free once fused; forces us to skip tiles, not just mask"],
       ["3", "Fused + online softmax", "THIS ONE", "the only variant that tests the two abstractions\u2019 real difference"],
       ["4", "GQA / MQA", "stretch", "best single demonstration, but only after the fused kernel works"],
       ["5", "Sliding window", "no", "irregular traversal; adds a format, not a concept"],
       ["6", "Cross-attention", "no", "the one place CUDA wins, but N\u2260M breaks our contract"]],
      widths=[0.05, 0.26, 0.16, 0.53], size=12, rowh=0.52)
note(s, "Tolerance for \u201ccorrect\u201d: max|out \u2212 ref_fp64| \u2264 5e-3 \u00b7 max|ref|, no NaN, no Inf \u2014 on every size.")

# 6 ---- the plan, two implementations
s = new(prs)
head(s, "the plan", "Same algorithm, two implementations, side by side")
code(s, Inches(0.75), Inches(2.15), Inches(6.05), Inches(3.45), [
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
    "}"], size=9, title="CUDA C++ \u2014 shared memory, registers, shuffles")
code(s, Inches(7.0), Inches(2.15), Inches(5.6), Inches(3.45), [
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
], size=9, title="cuTile \u2014 the compiler owns the mapping")
note(s, "Three files total: cuda/attn/flash.cu  \u00b7  cutile/attn/flash.py  \u00b7  bench/bench.py")

# 7 ---- good submission
s = new(prs)
head(s, "what makes a good submission",
     "The marks are in the measurement, not in the cleverness")
table(s, ["Criterion", "Wt", "What actually earns it", "What we do"],
      [["Correctness", "25%", "right numbers against a reference, on shapes that can expose the bugs",
        "fp64 reference, chunked; tail sizes chosen to be catchable"],
       ["Performance", "35%", "a credible fraction of the reference, and the reason for the gap",
        "fusion, tile sweep, and an ablation ladder per trick"],
       ["Comparative analysis", "25%", "the five axes, each with evidence \u2014 not opinion",
        "one figure per axis; portability labelled argued, not measured"],
       ["Code + profiling", "15%", "Nsight Compute evidence and a reproduction path",
        "ncu DRAM bytes + Tile section; 3-command README"]],
      widths=[0.17, 0.07, 0.42, 0.34], size=12, rowh=0.95)
note(s, "A perfect kernel with no argument caps at ~75%. A plain kernel with a sharp argument does not.")

# 8 ---- correctness
s = new(prs)
head(s, "how we prove it is correct", "The test size matters as much as the test")
table(s, ["Sizes", "What they catch", "Why those numbers"],
      [["1, 65, 130, 1000", "a missing tail mask", "phantom keys are a large fraction \u2192 error well above 5e-3"],
       ["1023, 2047, 4095", "almost nothing", "one phantom key \u2248 1e-3 \u2014 below tolerance, so it passes with the bug"],
       ["256 \u2026 4096", "the latency curve", "fusion barely helps at 512; crossover predicted there"],
       ["16384", "the OOM case", "4.29 GB per tensor at H=8; the fused kernel still runs"]],
      widths=[0.19, 0.24, 0.57], size=12, rowh=0.62)
listy(s, Inches(0.75), Inches(5.05), Inches(11.4), [
    (0, "Every size runs causal and non-causal, against a chunked fp64 reference \u2014 NaN/Inf fails the run."),
    (0, "A \u00d7100 input-magnitude stress is its own test: it is what exposes scaling order and overflow."),
    (0, "Three timers must agree (C++ events, Python events, ncu gpu__time_duration), or we are timing launches."),
], size=13.5, gap=8, h=1.5)

# 9 ---- trajectory
s = new(prs)
head(s, "trajectory", "Ten steps, each with something to show for it")
steps = [("\u2713", "Brief decoded; scope fixed to OP2 only", "done"),
         ("\u2713", "Environment: driver 617.42, repo, plan, this deck", "done"),
         ("\u2713", "Preflight tooling: audit + smoke test written", "done"),
         ("\u25b6", "Gate: cuTile smoke test passes on this GPU", "Fri 30"),
         ("\u25cb", "bench.py: fp64 + SDPA reference rows in the CSV", "Sat 31"),
         ("\u25cb", "Fused kernel in cuTile, causal + tail mask", "Sun 1"),
         ("\u25cb", "Fused kernel in CUDA C++, single-buffered first", "Mon 2"),
         ("\u25cb", "Tile/block sweep + ablation toggles", "Tue 3"),
         ("\u25cb", "Nsight Compute both kernels, incl. --section Tile", "Wed 4"),
         ("\u25cb", "Report, reproduction path, submit", "5\u20137 Nov")]
for i, (mark, text, when) in enumerate(steps):
    y = Inches(2.2 + i * 0.44)
    done = mark == "\u2713"
    now = mark == "\u25b6"
    col = MINT if done else (WARN if now else RGBColor(0x9F, 0xB3, 0xBA))
    tb(s, Inches(0.75), y, Inches(0.32), Inches(0.34), mark, size=15, bold=True, color=col)
    tb(s, Inches(1.15), y + Inches(0.02), Inches(8.6), Inches(0.34), text, size=13.5,
       bold=now, color=INK if not done else MUTED)
    tb(s, Inches(10.0), y + Inches(0.02), Inches(2.5), Inches(0.34), when, size=12,
       bold=now, color=col, align=PP_ALIGN.RIGHT)

# 10 ---- repo + asks
s = new(prs)
bgr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
bgr.fill.solid(); bgr.fill.fore_color.rgb = DEEP
bgr.line.fill.background(); bgr.shadow.inherit = False
tile_motif(s, 11.2, 5.0, cell=0.3, gap=0.07, rows=4, cols=4,
           filled=((0, 0), (2, 2)), color=TEAL)
head(s, "the repo, and what we need", "It runs in the open", dark=True, size=34)

tb(s, Inches(0.85), Inches(1.95), Inches(10.6), Inches(0.3), "PUBLIC REPO",
   size=11, bold=True, color=MINT)
tb(s, Inches(0.85), Inches(2.25), Inches(10.8), Inches(0.4), REPO, size=17,
   color=PAPER, font=MONO)

PALE = RGBColor(0xD5, 0xE6, 0xEA)
tb(s, Inches(0.85), Inches(2.95), Inches(5.4), Inches(0.3), "IN IT NOW",
   size=11, bold=True, color=MINT)
listy(s, Inches(0.85), Inches(3.25), Inches(5.6), [
    (1, "PLAN.md \u2014 kernel design, measurement protocol, schedule", PALE),
    (1, "docs/attention-variants.md \u2014 the six-variant survey", PALE),
    (1, "docs/decision-log.md \u2014 what changed and why", PALE),
    (1, "scripts/ \u2014 preflight audit, smoke test, this deck", PALE),
], size=12.5, gap=5, h=2.2, bcolor="02C39A")

tb(s, Inches(6.9), Inches(2.95), Inches(5.4), Inches(0.3), "BLOCKED ON, NOT ON CODE",
   size=11, bold=True, color=RGBColor(0xE8, 0xA5, 0x7C))
listy(s, Inches(6.9), Inches(3.25), Inches(5.6), [
    (1, "pip install into .venv \u2014 the environment gate is still red", RGBColor(0xF0, 0xC0, 0xA8)),
    (1, "CUDA Toolkit 13.2+ for nvcc and ncu", RGBColor(0xF0, 0xC0, 0xA8)),
    (1, "ncu counters are blocked by default on Windows", RGBColor(0xF0, 0xC0, 0xA8)),
], size=12.5, gap=5, h=1.8, bcolor="C4784A")

tb(s, Inches(0.85), Inches(5.3), Inches(10.6), Inches(0.3),
   "THREE THINGS I WOULD LIKE FROM YOU", size=11, bold=True, color=MINT)
tb(s, Inches(0.85), Inches(5.62), Inches(10.9), Inches(1.2),
   ["1.  Confirm OP2 only in writing \u2014 the brief\u2019s comparison section allows one operator,",
    "     but its artifact list still says \u201cboth\u201d.",
    "2.  Does 07 November hold, given we restarted from the survey on 14 October?",
    "3.  Should GQA come before the deadline, or stay stretch?"],
   size=13, color=RGBColor(0xD5, 0xE6, 0xEA), line=1.4)

# ------------------------------------------------------------------ save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
