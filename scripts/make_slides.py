"""Generate the OP2 deck in the "Silk Yellow Magazine" editorial style.

Devices taken from the reference template:
  - warm grey paper, black type, one saturated accent per chapter
  - a black capsule tag at top left that names the chapter
  - rotated page navigation on the right edge
  - oversized display headlines with one phrase in the accent colour
  - a 3x2 numbered card grid (01..06) with flat colour blocks
  - big-statistic slide with a hairline rule
  - label-pill rows (NAME  value) as used for SCALE / TIME / SPEND / DWELL
  - a black insight bar for the one sentence that must not be missed
  - a horizontal dot timeline
  - a dark closing slide carrying a pull-quote with a highlighted clause
  - a black footer bar with mono metadata

No fabricated data and no empty axes: every mark on every slide is either a
measured, computed or labelled-as-assumed number.

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ------------------------------------------------------------------- palette
PAPER = RGBColor(0xED, 0xEB, 0xE6)     # warm grey page
PAPER2 = RGBColor(0xE2, 0xDF, 0xD6)    # recessed block
INK = RGBColor(0x11, 0x11, 0x11)
MUTE = RGBColor(0x6B, 0x66, 0x60)
CORAL = RGBColor(0xE8, 0x48, 0x2B)     # CUDA, and chapter 1 accent
BLUE = RGBColor(0x2F, 0x6B, 0xE0)      # cuTile
YELLOW = RGBColor(0xF5, 0xC4, 0x00)    # highlight / insight
MINT = RGBColor(0xC9, 0xF0, 0xA8)
PINK = RGBColor(0xF7, 0xB6, 0xAE)
BLUEGRAY = RGBColor(0x93, 0xA9, 0xC4)
SAGE = RGBColor(0xBF, 0xD8, 0xCF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARKGREY = RGBColor(0x9A, 0x96, 0x90)

SANS = "Segoe UI"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.62)
CW = W - 2 * M
META_Y = Inches(6.98)
BODY_BOTTOM = Inches(6.82)
REPO = "github.com/harshith0518/CS6203---Attention-Operator-Porject"


# ------------------------------------------------------------------ helpers
def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def alt(shape, text):
    shape._element._nvXxPr.cNvPr.set("descr", text)


def _dim(v):
    """A shape extent must be positive; computed heights can go negative."""
    return max(int(abs(v)), 9144)


def rect(slide, x, y, w, h, fill, alt_text=None, shape=MSO_SHAPE.RECTANGLE, adj=None):
    s = slide.shapes.add_shape(shape, int(x), int(y), _dim(w), _dim(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    if adj is not None:
        try:
            s.adjustments[0] = adj
        except (IndexError, ValueError):
            pass
    if alt_text:
        alt(s, alt_text)
    return s


def bg(slide, colour=PAPER):
    rect(slide, 0, 0, W, H, colour, "page background")


def rule(slide, x, y, w, colour=INK, thick=Pt(1.0)):
    rect(slide, x, y, w, thick, colour)


def vrule(slide, x, y, h, colour=INK, thick=Pt(1.0)):
    rect(slide, x, y, thick, h, colour)


def txt(slide, x, y, w, h=Inches(0.3), lines=None, size=14, bold=False, colour=INK, font=SANS,
        align=PP_ALIGN.LEFT, space=2, line=1.2, caps=False):
    # callers may pass the text where the height goes; accept either order
    if isinstance(h, (str, list)):
        lines, h = h, Inches(0.3)
    if lines is None:
        lines = ""
    box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, s in enumerate(lines if isinstance(lines, list) else [lines]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space)
        p.line_spacing = line
        r = p.add_run()
        r.text = s.upper() if caps else s
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = colour
        r.font.name = font
    return box


def head(slide, x, y, w, lines, size=40, line=1.04, colour=INK):
    """Display headline. `lines` is a list of paragraphs; each paragraph is a
    list of (text, colour) runs so one phrase can carry the accent."""
    box = slide.shapes.add_textbox(int(x), int(y), int(w), Inches(0.2))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, runs in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line
        p.space_after = Pt(0)
        for text, col in runs:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = True
            r.font.color.rgb = col if col else colour
            r.font.name = SANS
    return box


def pill(slide, x, y, text, fill=INK, colour=WHITE, size=10.5, h=Inches(0.32),
         pad=Inches(0.2), font=MONO):
    cw = Inches(size * 0.60 / 72.0)
    w = pad * 2 + cw * len(text)
    rect(slide, x, y, w, h, fill, "capsule tag: " + text,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.5)
    txt(slide, x, y + (h - Inches(size * 1.35 / 72.0)) / 2 - Inches(0.01), w,
        Inches(0.24), text, size=size, colour=colour, font=font, align=PP_ALIGN.CENTER)
    return w


def chapter(slide, tag, fill=INK, colour=WHITE):
    pill(slide, M, Inches(0.44), tag, fill=fill, colour=colour)


def vnav(slide, page, total, dark=False):
    """Rotated page marker on the right edge, as in the reference template."""
    box = slide.shapes.add_textbox(int(W - Inches(0.5)), int(Inches(0.5)),
                                   int(Inches(0.3)), int(Inches(1.0)))
    tf = box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = f"{page:02d} / {total:02d}"
    r.font.size = Pt(9)
    r.font.name = MONO
    r.font.color.rgb = DARKGREY if dark else MUTE
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    bodyPr.set("vert", "vert270")


def metabar(slide, left, right):
    rect(slide, 0, META_Y, W, H - META_Y, INK, "footer bar")
    txt(slide, M, META_Y + Inches(0.15), Inches(9.0), Inches(0.24), left, size=10,
        colour=WHITE, font=MONO)
    txt(slide, W - M - Inches(2.2), META_Y + Inches(0.15), Inches(2.2), Inches(0.24), right,
        size=10, colour=DARKGREY, font=MONO, align=PP_ALIGN.RIGHT)


def insight(slide, x, y, w, h, label, body, size=15):
    rect(slide, x, y, w, h, INK, "insight bar")
    pill(slide, x + Inches(0.22), y + Inches(0.14), label, fill=YELLOW, colour=INK, size=9.5)
    body_h = max(h - Inches(0.62), Inches(0.22))
    txt(slide, x + Inches(0.22), y + Inches(0.56), w - Inches(0.44), body_h,
        body, size=size, colour=WHITE, line=1.25)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


prs = Presentation()
prs.slide_width, prs.slide_height = W, H
TOTAL = 12

# ==================================================================== 1 title
s = new(prs)
bg(s)
chapter(s, "CS-6023  \u00b7  PROJECT OP2  \u00b7  2026")
head(s, M, Inches(1.85), Inches(8.6),
     [[("Attention,", INK)], [("written ", INK), ("two ways.", CORAL)]], size=56)
rule(s, M, Inches(4.28), Inches(2.0), CORAL, Pt(3.0))
txt(s, M, Inches(4.55), Inches(6.9),
    "One fused attention kernel, written twice: by hand in CUDA C++ and by tile in "
    "cuTile. Same maths, different owner of the machine.", size=15, colour=MUTE, line=1.35)
txt(s, M, Inches(5.55), Inches(8.6),
    ["SURYA HARSHITH  \u00b7  CS23B010",
     "FOR PRAVIN JAGTAP AND SUMANTA DATTA  \u00b7  31 OCT 2026"], size=10.5, colour=INK,
    font=MONO, line=1.45)
# two tall swatches stand in for the photographic half of the reference layout
for i, (col, label) in enumerate([(CORAL, "CUDA C++"), (BLUE, "cuTile")]):
    x = Inches(9.5) + i * Inches(1.5)
    rect(s, x, Inches(1.35), Inches(1.32), Inches(4.85), col, label + " swatch")
    box = s.shapes.add_textbox(int(x + Inches(0.5)), int(Inches(1.6)),
                               int(Inches(0.32)), int(Inches(1.6)))
    tf = box.text_frame
    tf.word_wrap = False
    r = tf.paragraphs[0].add_run()
    r.text = label.upper()
    r.font.size = Pt(10)
    r.font.name = MONO
    r.font.color.rgb = WHITE
    tf._txBody.find(qn("a:bodyPr")).set("vert", "vert270")
metabar(s, "CS-6023 GPU KERNEL PROGRAMMING  \u00b7  IIT MADRAS", "01 / 12")
vnav(s, 1, TOTAL)
notes(s, "Seven days to the deadline. The deck is the whole picture: the operator, why it is a "
         "memory problem, which variant I build, what I build it with, how I will know it is "
         "right, and what could stop me.")

# ================================================================= 2 problem
s = new(prs)
bg(s)
chapter(s, "01  \u00b7  THE PROBLEM")
head(s, M, Inches(1.0), Inches(9.4),
     [[("Attention is a ", INK), ("memory problem.", CORAL)],
      [("The arithmetic was never the issue.", INK)]], size=34)

txt(s, M, Inches(2.72), Inches(3.4), "64\u00d7", size=104, bold=True, colour=CORAL, line=0.9)
txt(s, M, Inches(4.38), Inches(3.3), "less traffic once the N\u00d7N matrix never reaches memory",
    size=12.5, colour=INK, line=1.3)
txt(s, M, Inches(5.15), Inches(3.4),
    ["computed  \u00b7  N=4096, 8 heads, D=64, fp16"], size=10, colour=MUTE, font=MONO)

bx = Inches(4.55)
txt(s, bx, Inches(2.62), Inches(5.0), "UNFUSED  \u00b7  FOUR PASSES OVER 268 MB", size=10, bold=True,
    colour=MUTE, font=MONO)
for i, name in enumerate(["write S", "read S", "write P", "read P"]):
    y = Inches(2.92) + i * Inches(0.4)
    rect(s, bx, y, Inches(3.1), Inches(0.28), CORAL, "traffic block: " + name)
    txt(s, bx + Inches(3.22), y - Inches(0.01), Inches(1.6), name, size=10.5, colour=MUTE, font=MONO)
txt(s, bx, Inches(4.62), Inches(5.2), "1.07 GB", size=22, bold=True, colour=INK, font=MONO)
txt(s, bx + Inches(1.55), Inches(4.68), Inches(3.6), "in, and back out, four times", size=11.5,
    colour=MUTE)

txt(s, bx, Inches(5.35), Inches(5.0), "FUSED  \u00b7  17 MB", size=10, bold=True, colour=MUTE, font=MONO)
rect(s, bx, Inches(5.63), Inches(0.28), Inches(0.28), BLUE, "fused traffic block")
txt(s, bx + Inches(0.42), Inches(5.66), Inches(4.6), "Q, K, V, O  \u2014  the matrix stays on chip",
    size=11.5, colour=INK, font=MONO)

insight(s, Inches(9.55), Inches(2.62), Inches(3.16), Inches(3.05), "THE CATCH",
        ["At most about 3\u00d7 faster, not 64\u00d7. Once memory stops being the limit, "
         "compute is."], size=13)
txt(s, Inches(9.77), Inches(4.72), Inches(2.72),
    "Measured next: DRAM bytes at N = 2048, fused against unfused.", size=10.5,
    colour=DARKGREY, font=MONO, line=1.3)
txt(s, Inches(9.55), Inches(5.8), Inches(3.2),
    "Ridge assumes 20 TFLOP/s, which I measure before quoting.", size=10, colour=MUTE, line=1.25)
metabar(s, "THE BILL IS THE MATRIX, NOT THE MULTIPLY", "02 / 12")
vnav(s, 2, TOTAL)
notes(s, "Four passes over a 268 MB matrix is the 1.07 GB. The fused kernel never writes it. Do not "
         "read 64 times as a speedup: both time figures are roofline bounds and I measure the real "
         "ones.")

# =================================================================== 3 trade
s = new(prs)
bg(s)
chapter(s, "02  \u00b7  THE TRADE")
head(s, M, Inches(1.0), Inches(10.0),
     [[("Fusion buys traffic and ", INK), ("spends control.", CORAL)]], size=34)

pill(s, M, Inches(2.35), "BUYS", fill=BLUE)
txt(s, M, Inches(2.82), Inches(5.4),
    ["One kernel instead of three launches.",
     "No 268 MB tensor in device memory.",
     "A softmax that lives in registers."], size=15, colour=INK, line=1.35, space=9)

pill(s, Inches(6.85), Inches(2.35), "COSTS", fill=CORAL)
txt(s, Inches(6.85), Inches(2.82), Inches(5.6),
    ["Shared memory and registers are yours to manage.",
     "The row maximum only exists one slice at a time.",
     "An unmasked tail is a silent error, not a crash."], size=15, colour=INK, line=1.35, space=9)

rule(s, M, Inches(4.95), CW, INK, Pt(1.0))
txt(s, M, Inches(5.1), Inches(1.6), "THE GOAL", size=10, bold=True, colour=MUTE, font=MONO)
head(s, M, Inches(5.38), Inches(11.9),
     [[("Build it twice, then judge which abstraction ", INK),
       ("earns its keep.", CORAL)]], size=25, line=1.15)
txt(s, M, Inches(6.08), Inches(11.9),
    "Same mathematics, same sizes, same protocol, same reference. A fair comparison is the "
    "deliverable, not a novelty.", size=12, colour=MUTE)
metabar(s, "THE TILE MODEL TAKES THE SECOND COST OFF YOUR HANDS \u2014 WHETHER THAT PAYS IS THE QUESTION",
        "03 / 12")
vnav(s, 3, TOTAL)
notes(s, "This is the trade in one line. Fusion is what makes attention fast and the price is "
         "control over memory. The two models divide that price differently, which is the project.")

# ================================================================ 4 variants
s = new(prs)
bg(s)
chapter(s, "03  \u00b7  SIX VARIANTS, EACH IN BOTH MODELS")
head(s, M, Inches(0.98), Inches(11.0),
     [[("One gets built. ", INK), ("Five decide why.", CORAL)]], size=32)

cards = [
    ("01", "Vanilla SDPA", MINT, INK, "N\u00d7N on the bus, three launches",
     "cuda 3 kernels \u00b7 cutile no help"),
    ("02", "Causal", PINK, INK, "skip future tiles, mask the diagonal",
     "cuda bound + if \u00b7 cutile where"),
    ("03", "Fused, online softmax", INK, WHITE, "softmax seen one slice at a time",
     "rescale by hand \u00b7 one multiply"),
    ("04", "GQA / MQA", BLUEGRAY, INK, "query heads share one K/V head",
     "cuda index maths \u00b7 cutile stacked dim"),
    ("05", "Sliding window", SAGE, INK, "band-shaped loop bounds",
     "cuda two bounds \u00b7 cutile two bounds"),
    ("06", "Cross-attention", YELLOW, INK, "Nq \u2260 Nkv, so two tilings",
     "cuda one row/block \u00b7 64-row tile wastes 63/64"),
]
cw, ch, gap = Inches(3.86), Inches(1.92), Inches(0.27)
for i, (num, name, fill, fg, problem, cost) in enumerate(cards):
    col, row = i % 3, i // 3
    x = M + col * (cw + gap)
    y = Inches(1.95) + row * (ch + Inches(0.24))
    rect(s, x, y, cw, ch, fill, "variant card: " + name)
    txt(s, x + Inches(0.22), y + Inches(0.14), Inches(0.9), num, size=22, bold=True,
        colour=fg, font=MONO)
    txt(s, x + Inches(0.22), y + Inches(0.62), cw - Inches(0.4), name, size=15, bold=True,
        colour=fg)
    txt(s, x + Inches(0.22), y + Inches(0.94), cw - Inches(0.4), problem, size=11.5,
        colour=fg if fill is INK else MUTE, line=1.2)
    txt(s, x + Inches(0.22), y + ch - Inches(0.42), cw - Inches(0.4), cost, size=9.5,
        colour=fg if fill is INK else MUTE, font=MONO)

hero_x = M + 2 * (cw + gap)
pill(s, hero_x + cw - Inches(1.02), Inches(1.95) + Inches(0.18), "BUILD", fill=YELLOW,
     colour=INK, size=9.5, h=Inches(0.26), pad=Inches(0.14))
txt(s, M, Inches(6.42), Inches(9.5),
    "Only 03 runs end to end. Causal is a flag on it; 04 to 06 are read and reported.",
    size=11.5, colour=MUTE)
metabar(s, "CUDA AND CUTILE COSTS ARE MY READING OF THE DOCS, NOT MEASUREMENTS", "04 / 12")
vnav(s, 4, TOTAL)
notes(s, "These are the six I would call common, and what each costs in each model. I build row "
         "three only, with causal as a flag. The middle columns are reading, not measurement.")

# ==================================================================== 5 code
s = new(prs)
bg(s)
chapter(s, "04  \u00b7  THE KERNEL")
head(s, M, Inches(0.98), Inches(11.4),
     [[("The rescale is ", INK), ("32 registers", CORAL), (", or ", INK), ("one multiply.", BLUE)]],
     size=32)

cuda = [
    "float m[2]={-1e30f,-1e30f}, l[2]={0,0};",
    "float acc[D/8][4] = {};       // 32 fp32",
    "for (int k = 0; k < nk; ++k) {",
    "  load_kv(k); __syncthreads();",
    "  mma_qk(s, q_frag, k);       // S = Q K^T",
    "  for (int r = 0; r < 2; ++r) {",
    "    float x = rowmax(s, r) * scale;",
    "    x = fmaxf(x, __shfl_xor_sync(~0u,x,1));",
    "    x = fmaxf(x, __shfl_xor_sync(~0u,x,2));",
    "    float mn = fmaxf(m[r], x);",
    "    float a  = exp2f(m[r] - mn);  // rescale",
    "    l[r] *= a; m[r] = mn;",
    "    for (int n=0;n<D/8;++n) acc[n][2*r]*=a;",
    "  }",
]
tile = [
    "q = ct.load(Q, (bh, qi, 0), (1, BQ, D))",
    "m = full(-1e30); l = zeros; acc = zeros",
    "for k in range(num_kv_tiles):",
    "    kt = ct.load(K, (bh, k, 0), (1, BK, D))",
    "    s = ct.mma(q, ct.transpose(kt)) * scale",
    "    s = ct.where(keep(qi, k), s, -1e30)",
    "    m_new = ct.maximum(m, ct.max(s, axis=2))",
    "    p = ct.exp2(s - ct.expand_dims(m_new, 2))",
    "    a = ct.exp2(m - m_new)      # rescale",
    "    l = l * a + ct.sum(p, axis=2)",
    "    acc = ct.mma(p.astype(f16), load_v(k),",
    "                 acc * ct.expand_dims(a, 2))",
]


def codecol(slide, x, y, w, label, lines, colour, band):
    pill(slide, x, y, label, fill=colour)
    top = y + Inches(0.46)
    for r0, r1 in band:
        rect(slide, x - Inches(0.08), top + Inches(0.205) * r0 - Inches(0.03),
             w + Inches(0.16), Inches(0.205) * (r1 - r0 + 1) + Inches(0.02), PAPER2)
    box = slide.shapes.add_textbox(int(x), int(top), int(w), int(Inches(3.2)))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(11.5)
        r.font.name = MONO
        r.font.color.rgb = MUTE if ln.strip().startswith(("#", "//")) else INK


codecol(s, M, Inches(1.72), Inches(5.35), "CUDA C++", cuda, CORAL, [(10, 12)])
codecol(s, Inches(7.0), Inches(1.72), Inches(5.75), "CUTILE", tile, BLUE, [(8, 10)])

rule(s, M, Inches(5.28), CW, INK, Pt(1.0))
txt(s, M, Inches(5.42), Inches(5.3), "32 registers, rescaled by hand", size=14, bold=True, colour=CORAL)
txt(s, Inches(7.0), Inches(5.42), Inches(5.3), "one multiply on the accumulator", size=14, bold=True,
    colour=BLUE)
txt(s, M, Inches(5.82), Inches(11.8),
    "Both scale before the max, mask the tail, and rescale before accumulating. That order is what "
    "stops the denominator underflowing.", size=12, colour=MUTE, line=1.25)
metabar(s, "SKETCHES, NOT COMPILED \u00b7 CUTILE API DETAILS UNVERIFIED UNTIL IT RUNS", "05 / 12")
vnav(s, 5, TOTAL)
notes(s, "Left is per thread, right is per tile. The rescale is 32 registers by hand on the left and "
         "one multiply on the right. Whether that one line costs speed is what the profiler tells "
         "me.")

# =================================================================== 6 stack
s = new(prs)
bg(s)
chapter(s, "05  \u00b7  THE STACK")
head(s, M, Inches(0.98), Inches(11.4),
     [[("Two lanes, one ", INK), ("cubin", BLUE), (".", INK)]], size=32)

rows = [
    ("LANES", "attention.cu \u2192 nvcc \u2192 ptx \u2192 ptxas \u2192 cubin    |    "
              "attention.py \u2192 cuda.tile \u2192 tileir \u2192 tileiras \u2192 cubin"),
    ("HARDWARE", "RTX 3050 laptop  \u00b7  4 GB GDDR6  \u00b7  ~192 GB/s  \u00b7  sm_86, Ampere"),
    ("INSTALLED", "cuda-tile 1.7.0  \u00b7  tileiras, nvvm, nvrtc, nvcc 13.4.92  \u00b7  "
                  "cupy 14.2.0  \u00b7  numpy 2.5.4"),
    ("INERT", "TMA, thread-block clusters, FP8  \u2014  all need sm_90 or newer, in both models"),
    ("MISSING", "ncu and compute-sanitizer  \u2014  needed only for the CUDA half and profiling"),
]
for i, (label, body) in enumerate(rows):
    y = Inches(1.95) + i * Inches(0.82)
    pill(s, M, y + Inches(0.06), label, fill=INK if i < 3 else CORAL, size=9.5)
    txt(s, M + Inches(1.62), y, Inches(10.4),
        body, size=13 if i == 0 else 14, colour=INK if i < 3 else MUTE,
        font=MONO if i in (0, 2) else SANS, line=1.3)
    if i < len(rows) - 1:
        rule(s, M, y + Inches(0.62), CW, PAPER2, Pt(1.0))

insight(s, M, Inches(5.95), Inches(7.55), Inches(0.82), "WORTH SAYING",
        "Three cuTile knobs do nothing on Ampere. That is a portability result, not an apology.",
        size=12)
txt(s, Inches(8.45), Inches(5.98), Inches(4.3),
    ["Verified with pip list, nvidia-smi and a smoke test on this machine.",
     "One probe kernel compiles and runs; the matmul check has not run yet."],
    size=10.5, colour=MUTE, line=1.3)
metabar(s, "INSTALLED AND MEASURED HERE, NOT INTENDED", "06 / 12")
vnav(s, 6, TOTAL)
notes(s, "Two lanes, one destination. nvcc turns C++ into PTX and ptxas assembles it; cuTile goes "
         "through TileIR and tileiras. The versions are what is actually installed, and three "
         "features simply do not exist on this chip.")

# =================================================================== 7 status
s = new(prs)
bg(s)
chapter(s, "06  \u00b7  WHERE THIS STANDS")
head(s, M, Inches(1.0), Inches(11.4),
     [[("Built, running, ", INK), ("and next.", CORAL)]], size=34)

cols = [
    ("DONE", BLUE, WHITE, [
        "Brief decoded, scope frozen to OP2",
        "Six-variant survey written",
        "Driver 617.42, compute capability 8.6",
        "cuda-tile 1.7.0, toolchain 13.4.92",
        "Probe harness written and debugged"]),
    ("RUNNING NOW", YELLOW, INK, [
        "Probe kernel compiles and runs",
        "Three-check gate, re-run",
        "Peak FLOP/s and DRAM bandwidth, measured",
        "So the ridge figure is mine, not assumed"]),
    ("NEXT 48 HOURS", CORAL, WHITE, [
        "bench.py: fp64 reference, SDPA, CSV",
        "cuTile kernel, causal as a flag",
        "CUDA kernel v1, single-buffered",
        "Then sweeps, profiles, report"]),
]
colw = Inches(3.86)
for i, (label, fill, _fg, items) in enumerate(cols):
    x = M + i * (colw + Inches(0.27))
    pill(s, x, Inches(2.02), label, fill=fill, colour=WHITE if fill is not YELLOW else INK)
    for j, item in enumerate(items):
        y = Inches(2.6) + j * Inches(0.6)
        rect(s, x, y + Inches(0.08), Inches(0.07), Inches(0.07), fill, "list marker")
        txt(s, x + Inches(0.26), y, colw - Inches(0.32), item, size=13, colour=INK, line=1.25)

rect(s, M, Inches(5.92), CW, Inches(0.8), INK, "deliverables bar")
pill(s, M + Inches(0.22), Inches(6.14), "ON 07 NOV", fill=YELLOW, colour=INK, size=9.5,
     h=Inches(0.26), pad=Inches(0.14))
txt(s, M + Inches(1.95), Inches(6.15), Inches(9.9),
    "Two kernels  \u00b7  one CSV behind every number  \u00b7  profiles for both  \u00b7  "
    "a report that takes a side", size=12.5, colour=WHITE)
metabar(s, "NO CLAIM IN THIS DECK OUTRUNS THE EVIDENCE", "07 / 12")
vnav(s, 7, TOTAL)
notes(s, "This is the honest position: the environment is finished and one probe kernel runs. The "
         "kernels themselves start tomorrow. What you get on the seventh is written at the bottom, "
         "so it can be held against me.")

# ============================================================= 7 measurement
s = new(prs)
bg(s)
chapter(s, "07  \u00b7  THE MEASUREMENT")
head(s, M, Inches(0.98), Inches(11.4),
     [[("Four numbers, one protocol, ", INK), ("no stopwatch.", CORAL)]], size=32)

rows = [
    ("WARM-UP", "50 iterations before anything is timed. Compile time reported separately."),
    ("TIMING", "At least one second of GPU work per configuration, not a fixed iteration count."),
    ("FLUSH", "L2 cleared outside the event window, so write-back does not land in the next kernel."),
    ("REPORT", "p50 and p90, with clocks, power and throttle state logged on every row."),
]
for i, (label, body) in enumerate(rows):
    y = Inches(1.95) + i * Inches(0.78)
    pill(s, M, y + Inches(0.02), label, fill=PINK, colour=INK, size=9.5)
    txt(s, M + Inches(1.7), y, Inches(10.2), body, size=15, colour=INK, line=1.3)
    rule(s, M, y + Inches(0.6), CW, PAPER2, Pt(1.0))

insight(s, M, Inches(5.28), Inches(11.9), Inches(1.35), "WHY THIS, AND NOT A SIMPLE LOOP",
        ["Three timers have to agree, or the number is launch overhead wearing a lab coat. "
         "Every row carries clock and temperature, because this is a laptop that throttles."],
        size=14)
metabar(s, "ARTEFACTS: results.csv  \u00b7  profiles/*.ncu-rep  \u00b7  figs/*.pdf  \u00b7  report.md",
        "08 / 12")
vnav(s, 8, TOTAL)
notes(s, "The protocol exists to survive scrutiny. The flush sits outside the timed window; if it "
         "drifted inside, fusion would look free. Clock and temperature travel with every row "
         "because this machine throttles.")

# ============================================================ 8 definition
s = new(prs)
bg(s)
chapter(s, "08  \u00b7  DEFINITION OF DONE")
head(s, M, Inches(0.98), Inches(11.4),
     [[("Six checks. ", INK), ("Every one can fail.", CORAL)]], size=32)

checks = [
    ("01", "Correct on every size", "max|out \u2212 ref_fp64| \u2264 5e-3 \u00b7 max|ref|"),
    ("02", "Numerically stable", "no NaN or Inf; inputs \u00d7100 stay finite"),
    ("03", "Both models, one contract", "same shapes, same harness, same reference"),
    ("04", "Performance is measured", "p50/p90 at N=2048, DRAM bytes from ncu"),
    ("05", "Profiling evidence exists", "one ncu report per kernel, section Tile for cuTile"),
    ("06", "The argument is made", "five axes, one figure each, a recommendation"),
]
for i, (num, name, cond) in enumerate(checks):
    y = Inches(1.85) + i * Inches(0.55)
    txt(s, M, y, Inches(0.6), num, size=11, bold=True, colour=MUTE, font=MONO)
    txt(s, M + Inches(0.62), y - Inches(0.02), Inches(3.6), name, size=15, bold=True, colour=INK)
    txt(s, M + Inches(4.35), y, Inches(5.4), cond, size=12, colour=MUTE)
    txt(s, W - M - Inches(1.0), y, Inches(1.0), "pending", size=10, colour=MUTE, font=MONO,
        align=PP_ALIGN.RIGHT)
    rule(s, M, y + Inches(0.44), CW, PAPER2, Pt(1.0))
vrule(s, W - M - Inches(1.0), Inches(1.85), Inches(0.55 * 4) - Inches(0.1), CORAL, Pt(2.5))
txt(s, W - M - Inches(1.0), Inches(1.55), Inches(1.0), "GATE", size=9, bold=True, colour=CORAL,
    font=MONO, align=PP_ALIGN.RIGHT)

pill(s, M, Inches(5.22), "SIZES", fill=INK, size=9.5)
chips = ["1", "65", "130", "1000", "256", "512", "1024", "2048", "4096"]
x = M + Inches(1.28)
for c in chips:
    hot = c in ("1", "65", "130", "1000")
    w = Inches(0.62)
    rect(s, x, Inches(5.22), w, Inches(0.32), CORAL if hot else PAPER2, "size chip " + c)
    txt(s, x, Inches(5.28), w, c, size=10.5, colour=WHITE if hot else INK, font=MONO,
        align=PP_ALIGN.CENTER)
    x += w + Inches(0.1)
txt(s, x + Inches(0.15), Inches(5.3), Inches(4.2), "red = sizes that expose a missing tail mask",
    size=10, colour=MUTE, font=MONO)

# hard gate: one row, pill and sentence on the same baseline
rect(s, M, Inches(5.92), CW, Inches(0.62), INK, "hard gate bar")
pill(s, M + Inches(0.22), Inches(6.1), "HARD GATE", fill=YELLOW, colour=INK, size=9.5,
     h=Inches(0.26), pad=Inches(0.14))
txt(s, M + Inches(1.62), Inches(6.12), Inches(10.1),
    "No variant work starts until 01 to 04 pass on both kernels.", size=13, colour=WHITE)
metabar(s, "TOLERANCE IS MY CHOICE \u2014 THE BRIEF PRESCRIBES NONE, SO IT IS STATED TO BE ARGUED WITH",
        "09 / 12")
vnav(s, 9, TOTAL)
notes(s, "Every row can fail, which is the point. The red sizes are the ones where a missing mask "
         "produces an error above tolerance. Row six is where a quarter of the grade sits, and it "
         "is writing, not code.")

# ==================================================================== 9 plan
s = new(prs)
bg(s)
chapter(s, "09  \u00b7  THE PLAN")
head(s, M, Inches(0.98), Inches(10.4),
     [[("Seven days, five gates, ", INK), ("one cut list.", CORAL)]], size=32)

days = [("SAT 31", "gate green"), ("SUN 1", "bench.py"), ("MON 2", "cuTile kernel"),
        ("TUE 3", "CUDA kernel"), ("WED 4", "sweeps"), ("THU 5", "ncu profiles"),
        ("FRI 6", "report"), ("SAT 7", "submit")]
ty = Inches(3.55)
lx, rx = M + Inches(0.3), W - M - Inches(0.5)
rule(s, lx, ty, rx - lx, INK, Pt(2.0))
step = (rx - lx) / (len(days) - 1)
for i, (day, work) in enumerate(days):
    cx = lx + i * step
    gate = i in (0, 1, 2, 3, 7)
    rect(s, cx - Inches(0.055), ty - Inches(0.055), Inches(0.11), Inches(0.11), INK,
         "day marker")
    if gate:
        rect(s, cx - Inches(0.07), ty - Inches(0.52), Inches(0.14), Inches(0.14), CORAL,
             "gate marker", shape=MSO_SHAPE.DIAMOND)
        txt(s, cx - Inches(0.55), ty - Inches(0.82), Inches(1.1), "gate", size=8, colour=CORAL,
            font=MONO, align=PP_ALIGN.CENTER)
    txt(s, cx - Inches(0.62), ty - Inches(0.36), Inches(1.24), day, size=10, bold=True,
        colour=INK if i else CORAL, font=MONO, align=PP_ALIGN.CENTER)
    txt(s, cx - Inches(0.72), ty + Inches(0.2), Inches(1.44), work, size=11,
        colour=INK if i else CORAL, align=PP_ALIGN.CENTER, line=1.2)

rule(s, M, Inches(4.95), CW, INK, Pt(1.0))
pill(s, M, Inches(5.1), "CUT FIRST", fill=CORAL, size=9.5)
txt(s, M + Inches(1.6), Inches(5.08), Inches(4.6), "GQA  \u2192  Fig 4  \u2192  the debuggability table",
    size=12, colour=INK, font=MONO)
pill(s, Inches(7.0), Inches(5.1), "NEVER CUT", fill=INK, size=9.5)
txt(s, Inches(8.35), Inches(5.08), Inches(4.3), "both kernels, the correctness numbers, the profiles",
    size=12, colour=INK, font=MONO)
txt(s, M, Inches(5.85), Inches(11.9),
    ["The brief's own milestones, 10/07 to 10/28, have passed. This is the recovery plan: one "
     "kernel per model, measured properly, with a report that takes a position.",
     "The cut list is written now so that decision is not made at midnight on the sixth."],
    size=12, colour=MUTE, line=1.3, space=5)
metabar(s, "EIGHT COLUMNS BECAUSE DAY ZERO IS TODAY", "10 / 12")
vnav(s, 10, TOTAL)
notes(s, "I am behind the brief's milestones and this is the recovery plan. Two kernels, measured "
         "properly, a report that takes a position. The cut list exists now so that choice is not "
         "made at midnight.")

# ============================================================== 10 references
s = new(prs)
bg(s)
chapter(s, "10  \u00b7  SOURCES")
head(s, M, Inches(0.98), Inches(9.0),
     [[("Where this ", INK), ("comes from.", CORAL)]], size=32)

primary = [("cuTile Python documentation", "docs.nvidia.com/cuda/cutile-python"),
           ("NVIDIA/TileGym \u2014 attention and matmul kernels", "github.com/NVIDIA/TileGym"),
           ("Dao et al., FlashAttention, 2022", "arxiv.org/abs/2205.14135")]
for i, (name, url) in enumerate(primary):
    y = Inches(1.95) + i * Inches(0.95)
    txt(s, M, y, Inches(7.2), name, size=17, bold=True, colour=INK)
    txt(s, M, y + Inches(0.34), Inches(7.2), url, size=11.5, colour=BLUE, font=MONO)
    rule(s, M, y + Inches(0.72), Inches(7.2), PAPER2, Pt(1.0))

vrule(s, Inches(8.15), Inches(1.95), Inches(3.4), PAPER2, Pt(1.0))
pill(s, Inches(8.45), Inches(1.9), "ALSO USED", size=9.5)
secondary = [("Milakov and Gimelshein, online softmax", "arxiv.org/abs/1801.05824"),
             ("Dao, FlashAttention-2", "arxiv.org/abs/2307.08691"),
             ("Vaswani et al., Attention Is All You Need", "arxiv.org/abs/1706.03762"),
             ("Nsight Compute profiling guide", "docs.nvidia.com/nsight-compute/ProfilingGuide"),
             ("cutile-python samples", "github.com/NVIDIA/cutile-python")]
for i, (name, url) in enumerate(secondary):
    y = Inches(2.4) + i * Inches(0.56)
    txt(s, Inches(8.45), y, Inches(4.3), name, size=11.5, colour=INK)
    txt(s, Inches(8.45), y + Inches(0.22), Inches(4.3), url, size=9.5, colour=MUTE, font=MONO)

rule(s, M, Inches(5.35), CW, INK, Pt(1.0))
pill(s, M, Inches(5.5), "LEARNED FROM", size=9.5)
txt(s, M, Inches(5.92), Inches(11.9),
    "3Blue1Brown, Deep Learning ch.6   \u00b7   Alammar, The Illustrated Transformer   \u00b7   "
    "Georgia Tech, Transformer Explainer", size=11.5, colour=INK)
txt(s, M, Inches(6.3), Inches(11.9),
    "Repository, plan and decision log: " + REPO, size=10, colour=MUTE, font=MONO)
metabar(s, "EVERY LINK OPENED WHILE PREPARING THIS DECK", "11 / 12")
vnav(s, 11, TOTAL)
notes(s, "Three sources carry the technical weight: the cuTile docs, TileGym's attention kernels and "
         "the FlashAttention paper. The rest is support and background, and the repo carries the "
         "plan with its decision log.")

# ================================================================ 11 asks
s = new(prs)
bg(s, INK)
pill(s, M, Inches(0.44), "11  \u00b7  RISKS AND ASKS", fill=YELLOW, colour=INK)
head(s, M, Inches(1.85), Inches(7.4),
     [[("Three answers", WHITE)], [("change the plan.", YELLOW)]], size=44)
txt(s, M, Inches(4.05), Inches(6.6),
    "Everything else I can absorb. These three decisions are not mine to make, and each one "
    "moves the schedule.", size=14, colour=DARKGREY, line=1.35)

rows = [
    ("01", "Attention only, no GEMM?", "default: OP2 only"),
    ("02", "Does 07 November hold?", "default: it does"),
    ("03", "Is N = 2048 the target length?", "default: N = 2048"),
]
for i, (num, q, dflt) in enumerate(rows):
    y = Inches(1.85) + i * Inches(1.36)
    rect(s, Inches(8.4), y + Inches(0.1), Inches(0.13), Inches(0.13), YELLOW, "bullet")
    txt(s, Inches(8.72), y, Inches(0.5), num, size=10, colour=DARKGREY, font=MONO)
    txt(s, Inches(8.72), y + Inches(0.22), Inches(4.0), q, size=17, bold=True, colour=WHITE,
        line=1.15)
    txt(s, Inches(8.72), y + Inches(0.78), Inches(4.0), dflt, size=10, colour=YELLOW, font=MONO)
txt(s, Inches(8.4), Inches(6.05), Inches(4.4),
    "Silence by Tue 3 Nov and the defaults run.", size=11, colour=DARKGREY, font=MONO)
rect(s, 0, META_Y, W, H - META_Y, RGBColor(0x1C, 0x1C, 0x1C), "footer bar")
txt(s, M, META_Y + Inches(0.15), Inches(9.4), Inches(0.24),
    "ARRIVES 07 NOV:  TWO KERNELS  \u00b7  THREE COMMANDS PER RERUN  \u00b7  ONE REPORT", size=10,
    colour=WHITE, font=MONO)
txt(s, W - M - Inches(2.2), META_Y + Inches(0.15), Inches(2.2), Inches(0.24), "12 / 12", size=10,
    colour=DARKGREY, font=MONO, align=PP_ALIGN.RIGHT)
vnav(s, 12, TOTAL, dark=True)
notes(s, "Three questions and what I assume if I hear nothing. The risks are handled, not hidden: "
         "clocks are logged, sizes are fixed now, and the kernel comes before the harness.")

# ---------------------------------------------------------------------- save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention, written two ways"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
