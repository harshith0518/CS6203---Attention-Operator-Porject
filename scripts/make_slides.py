"""Generate the 11-slide OP2 deck: engineering-notebook layout.

Substance is unchanged from the reviewed version; this rewrites the packaging.
No tables (everything with shape is drawn), no shadows, no rounded cards, one
colour per concept, hairline rules instead of filled panels, and larger type.

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py

Output: docs/op2-attention-slides.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ------------------------------------------------------------------- palette
PAPER = RGBColor(0xF8, 0xF7, 0xF3)
INK = RGBColor(0x16, 0x18, 0x1D)
GREY = RGBColor(0x5B, 0x61, 0x6B)
HAIR = RGBColor(0xD9, 0xD6, 0xCE)
CUDA = RGBColor(0xD5, 0x5E, 0x00)      # Okabe-Ito vermillion: hand-written CUDA
TILE = RGBColor(0x00, 0x72, 0xB2)      # Okabe-Ito blue: cuTile
PASS = RGBColor(0x00, 0x9E, 0x73)      # Okabe-Ito green: verified / pass
DARK = RGBColor(0x0F, 0x14, 0x19)
DARKTX = RGBColor(0xE8, 0xE6, 0xE1)
CUDA_L = RGBColor(0xE6, 0x9F, 0x00)    # lighter on dark
TILE_L = RGBColor(0x56, 0xB4, 0xE9)
DARKGREY = RGBColor(0x9A, 0xA1, 0xAA)

SANS = "Segoe UI"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.6)
CW = W - 2 * M
REPO = "github.com/harshith0518/CS6203---Attention-Operator-Porject"


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def alt(shape, text):
    shape._element._nvXxPr.cNvPr.set("descr", text)


def bg(slide, colour):
    rect(slide, 0, 0, W, H, colour, "page background")


def rect(slide, x, y, w, h, fill, alt_text=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), int(y), int(w), int(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    if alt_text:
        alt(s, alt_text)
    return s


def hline(slide, x, y, w, colour=HAIR, thick=Pt(0.75)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), int(y), int(w), int(thick))
    s.fill.solid()
    s.fill.fore_color.rgb = colour
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def vline(slide, x, y, h, colour=HAIR, thick=Pt(0.75)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), int(y), int(thick), int(h))
    s.fill.solid()
    s.fill.fore_color.rgb = colour
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def diamond(slide, cx, cy, size, colour):
    s = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, int(cx - size / 2), int(cy - size / 2),
                               int(size), int(size))
    s.fill.solid()
    s.fill.fore_color.rgb = colour
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def txt(slide, x, y, w, h=Inches(0.32), lines=None, size=16, bold=False, colour=INK, font=SANS,
        align=PP_ALIGN.LEFT, space=2, line=1.18, caps=False):
    # callers may pass text where the height goes; accept either order
    if isinstance(h, (str, list)):
        lines, h = h, Inches(0.32)
    if lines is None:
        lines = ""
    box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    items = lines if isinstance(lines, list) else [lines]
    for i, s in enumerate(items):
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


def label(slide, x, y, w, s, colour=GREY, size=11):
    txt(slide, x, y, w, Inches(0.2), s, size=size, bold=True, colour=colour, caps=True)


def title(slide, s, colour=INK, size=36, y=Inches(0.62), w=None):
    txt(slide, M, y, w or CW, Inches(1.05), s, size=size, bold=True, colour=colour, line=1.03)


def footer(slide, s, page, dark=False):
    c = DARKGREY if dark else GREY
    txt(slide, M, Inches(7.02), Inches(11.0), Inches(0.22), s, size=11, colour=c)
    txt(slide, W - M - Inches(0.9), Inches(7.02), Inches(0.9), Inches(0.22), str(page),
        size=11, colour=c, align=PP_ALIGN.RIGHT, font=MONO)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def grid(slide, x, y, rows, cols, cell, gap, cells, alt_text=None):
    """Draw a small square grid; `cells` is a set of (r, c) -> colour."""
    for r in range(rows):
        for c in range(cols):
            col = cells.get((r, c))
            if col is None:
                continue
            rect(slide, x + c * (cell + gap), y + r * (cell + gap), cell, cell, col)


MASKS = {
    "full": lambda r, c: True,
    "causal": lambda r, c: c <= r,
    "band": lambda r, c: abs(r - c) <= 1,
    "grouped": lambda r, c: True,
    "cross": lambda r, c: c < 3,
    "slices": lambda r, c: (r + c) % 2 == 0,
}


def mask_pictogram(slide, x, y, kind, colour, cell=Inches(0.115), gap=Inches(0.028), n=5):
    test = MASKS[kind]
    for r in range(n):
        for c in range(n):
            if test(r, c):
                col = colour
                if kind == "grouped" and c >= 2:
                    col = TILE
                rect(slide, x + c * (cell + gap), y + r * (cell + gap), cell, cell, col)


prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# ==================================================================== 1 title
s = new(prs)
bg(s, DARK)
# two mirrored 8x8 matrices: the same problem, two ways of writing it
for gi, (gx, col) in enumerate([(Inches(8.45), CUDA_L), (Inches(11.05), TILE_L)]):
    for r in range(8):
        for c in range(8):
            shade = r * 8 + c
            rect(s, gx + c * Inches(0.31), Inches(2.1) + r * Inches(0.31),
                 Inches(0.26), Inches(0.26), col)
label(s, M, Inches(0.55), Inches(8.0), "CS-6023  \u00b7  GPU kernel programming  \u00b7  project OP2",
      colour=DARKGREY, size=12)
txt(s, M, Inches(1.5), Inches(7.7), Inches(1.9), "Attention,\nwritten two ways",
    size=56, bold=True, colour=DARKTX, line=1.0)
hline(s, M, Inches(4.05), Inches(5.4), colour=CUDA_L)
txt(s, M, Inches(4.25), Inches(6.6), Inches(0.9),
    "Memory, not arithmetic, is what limits attention.", size=22, colour=CUDA_L, line=1.2)
txt(s, M, Inches(4.95), Inches(6.4), Inches(0.8),
    "One fused kernel in CUDA C++ and in cuTile, measured against each other.",
    size=16, colour=DARKTX, line=1.25)
txt(s, M, Inches(5.95), Inches(9.0), Inches(0.9),
    ["surya harshith  \u00b7  cs23b010", "for pravin jagtap and sumanta datta  \u00b7  31 oct 2026"],
    size=12, colour=DARKGREY, font=MONO, line=1.5)
notes(s, "Seven days to the deadline. This deck is the whole picture: the operator, why it is a "
         "memory problem, which variant I build, what I build it with, how I will know it is right, "
         "and what could stop me.")

# ================================================================== 2 traffic
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the problem")
title(s, "Attention moves an N\u00d7N matrix, and that is the cost.", size=34)
label(s, M, Inches(1.62), Inches(6.0), "unfused: four passes over the matrix", colour=CUDA)
passes = ["write S", "read S", "write P", "read P"]
for i, name in enumerate(passes):
    y = Inches(1.92) + i * Inches(0.44)
    rect(s, M, y, Inches(3.3), Inches(0.34), CUDA)
    txt(s, M + Inches(3.42), y + Inches(0.04), Inches(1.6), name, size=12, colour=GREY, font=MONO)
txt(s, M, Inches(3.72), Inches(4.2), Inches(0.5), "4 \u00d7 268 MB  =  1.07 GB", size=18,
    bold=True, colour=CUDA, font=MONO)
label(s, M, Inches(4.35), Inches(6.0), "fused: the matrix never leaves the chip", colour=TILE)
rect(s, M, Inches(4.63), Inches(0.33), Inches(0.33), TILE)
txt(s, M + Inches(0.46), Inches(4.67), Inches(3.6), "17 MB  =  Q, K, V, O only", size=14,
    bold=True, colour=TILE, font=MONO)
txt(s, M, Inches(5.05), Inches(4.6), Inches(0.3),
    "drawn to scale, not to shape", size=11, colour=GREY)

# hero number
txt(s, Inches(4.55), Inches(4.55), Inches(3.2), Inches(1.4), "64\u00d7", size=96, bold=True,
    colour=TILE, line=0.9)
txt(s, Inches(4.55), Inches(5.95), Inches(3.4), Inches(1.0),
    ["less traffic,", "but at most ~3\u00d7 less time:", "compute becomes the limit."],
    size=12, colour=GREY, line=1.25)

# mini roofline: memory-bound diagonal, compute ceiling, ridge where they meet
rx, ry, rw, rh = Inches(7.9), Inches(2.15), Inches(4.7), Inches(3.2)
L, R, T, B = rx, rx + rw, ry, ry + rh
ridge_frac, ceil_frac = 0.42, 0.34
ridge_x = L + rw * ridge_frac
ceil_y = T + rh * ceil_frac
vline(s, L, T, rh, colour=INK, thick=Pt(1.0))
hline(s, L, B, rw, colour=INK, thick=Pt(1.0))
# memory-bound slope
diag = s.shapes.add_connector(1, int(L), int(B), int(ridge_x), int(ceil_y))
diag.line.color.rgb = CUDA
diag.line.width = Pt(2.0)
# compute ceiling
hline(s, ridge_x, ceil_y, rw * (1 - ridge_frac), colour=TILE, thick=Pt(2.0))
vline(s, ridge_x, ceil_y, rh * (1 - ceil_frac), colour=GREY, thick=Pt(0.75))
txt(s, L - Inches(0.05), T - Inches(0.3), Inches(2.4), "FLOP/s", size=10, colour=GREY, font=MONO)
txt(s, R - Inches(1.5), B + Inches(0.08), Inches(1.5), "bytes/s", size=10, colour=GREY,
    font=MONO, align=PP_ALIGN.RIGHT)
txt(s, ridge_x - Inches(0.9), ceil_y - Inches(0.28), Inches(1.3), "ridge 104", size=10,
    colour=INK, font=MONO, align=PP_ALIGN.RIGHT)
txt(s, ridge_x + Inches(0.16), ceil_y - Inches(0.26), Inches(1.9), "compute-bound", size=10,
    colour=TILE, font=MONO)
txt(s, L + Inches(0.12), B - Inches(0.42), Inches(1.8), "memory-bound", size=10, colour=CUDA,
    font=MONO)
# the two kernels sit on the memory-bound slope at their own arithmetic intensity
for frac, lab, col in [(0.11, "32", CUDA), (0.36, "2,048", TILE)]:
    px = L + rw * frac
    py_ = B - (px - L) / (ridge_x - L) * (B - ceil_y)
    rect(s, px - Inches(0.05), py_ - Inches(0.05), Inches(0.1), Inches(0.1), col)
    txt(s, px - Inches(0.85), py_ + Inches(0.04), Inches(1.0), lab, size=11, bold=True,
        colour=col, font=MONO, align=PP_ALIGN.RIGHT)
txt(s, L + rw * 0.16, B + Inches(0.3), Inches(3.2), "arithmetic intensity, log scale", size=10,
    colour=GREY)
footer(s, "computed: N=4096, 8 heads, D=64, fp16. Ridge assumes a 20 TFLOP/s tensor peak, which I "
          "measure first and then recompute.", 2)
notes(s, "Four passes over a 268 MB matrix is the 1.07 GB. The fused kernel never writes it. Do not "
         "read 64 times as a speedup: both time figures are roofline bounds and I will measure the "
         "real ones.")

# ============================================================== 3 buys/costs
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the trade")
title(s, "Fusion buys traffic and spends control.", size=34)
vline(s, W / 2, Inches(1.75), Inches(2.65), colour=HAIR)
label(s, M, Inches(1.75), Inches(5.0), "buys", colour=TILE)
txt(s, M, Inches(2.05), Inches(5.3), Inches(2.4),
    ["One kernel instead of three launches.",
     "No 268 MB tensor ever reaches device memory.",
     "A softmax that fits in registers, not in HBM."],
    size=17, colour=INK, line=1.3, space=10)
label(s, W / 2 + Inches(0.45), Inches(1.75), Inches(5.0), "costs", colour=CUDA)
txt(s, W / 2 + Inches(0.45), Inches(2.05), Inches(5.3), Inches(2.4),
    ["Shared memory and register pressure become yours to manage.",
     "The softmax is only ever visible one slice at a time.",
     "A tail that is not masked is a silent error, not a crash."],
    size=17, colour=INK, line=1.3, space=10)
hline(s, M, Inches(4.55), CW, colour=INK, thick=Pt(1.0))
label(s, M, Inches(4.72), Inches(6.0), "the goal", colour=INK)
txt(s, M, Inches(5.02), Inches(11.6), Inches(1.1),
    "Build the same fused operator twice, hand-written and tile-level, and measure which "
    "abstraction is worth the control it asks you to give up.", size=24, colour=INK, line=1.25)
txt(s, M, Inches(6.05), Inches(11.6), Inches(0.6),
    "Same mathematics, same sizes, same timing protocol, same reference. One novelty is not the "
    "goal; a fair comparison is.", size=14, colour=GREY, line=1.2)
footer(s, "The tile model takes the second cost off your hands. Whether that is worth it is the "
          "question the deck ends on.", 3)
notes(s, "This is the trade in one line. Fusion is what makes attention fast, and the price is "
         "control over memory. The two models divide that price differently, which is the whole "
         "project.")

# ================================================================ 4 variants
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the operator  \u00b7  six common variants, one gets built")
title(s, "One variant gets built. The other five decide why.", size=32)
txt(s, M, Inches(1.5), Inches(11.9), Inches(0.3),
    "O = softmax( Q K\u1d40 / \u221ad + mask ) \u00b7 V    self-attention, forward only, exact, "
    "causal as a flag", size=12, colour=GREY, font=MONO)

cards = [
    ("full", "Vanilla SDPA", "3 launches", "no help", False),
    ("causal", "Causal", "loop bound + if", "loop bound + where", False),
    ("slices", "Fused, online softmax", "rescale by hand", "one tile multiply", True),
    ("grouped", "GQA / MQA", "index maths", "stacked tile dim", False),
    ("band", "Sliding window", "two bounds", "two bounds", False),
    ("cross", "Cross-attention", "one row per block", "wastes 63/64", False),
]
for i, (kind, name, c_cuda, c_tile, hero) in enumerate(cards):
    col_i, row_i = i % 3, i // 3
    x = M + col_i * Inches(4.05)
    y = Inches(1.95) + row_i * Inches(2.45)
    if hero:
        mask_pictogram(s, x + Inches(0.05), y + Inches(0.1), kind, TILE)
        rect(s, x + Inches(1.05), y + Inches(0.06), Inches(0.86), Inches(0.26), TILE)
        txt(s, x + Inches(1.05), y + Inches(0.1), Inches(0.86), "build", size=10, bold=True,
            colour=PAPER, font=MONO, align=PP_ALIGN.CENTER)
        hline(s, x, y + Inches(0.86), Inches(3.6), colour=TILE, thick=Pt(1.5))
    else:
        mask_pictogram(s, x, y + Inches(0.16), kind, GREY)
    txt(s, x, y + Inches(0.92), Inches(3.6), name, size=17, bold=hero, colour=INK if hero else GREY)
    txt(s, x, y + Inches(1.22), Inches(3.6), "CUDA  " + c_cuda, size=12,
        colour=CUDA if hero else GREY, font=MONO)
    txt(s, x, y + Inches(1.46), Inches(3.6), "cuTile " + c_tile, size=12,
        colour=TILE if hero else GREY, font=MONO)
txt(s, M, Inches(6.55), Inches(11.9), Inches(0.4),
    "Rows 1 and 2 are the baseline and the flag. Rows 4 to 6 are read and reported, not built.",
    size=12, colour=GREY)
footer(s, "Why #3: it is the only variant where both models face the same hard problem, a softmax "
          "seen one slice at a time.", 4)
notes(s, "These are the six I would call common, and what each costs in each model. I build row "
         "three only. The CUDA and cuTile costs are my reading of the docs, not measurements.")

# ==================================================================== 5 code
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the kernel")
title(s, "The rescale: 32 registers, or one multiply.", size=34)

cuda_lines = [
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
]
tile_lines = [
    "bh, qi = ct.bid(1), ct.bid(0)",
    "q = ct.load(Q, (bh, qi, 0), (1, BQ, D))",
    "m = full(-1e30); l = zeros; acc = zeros",
    "for k in range(num_kv_tiles):",
    "    kt = ct.load(K, (bh, k, 0), (1, BK, D))",
    "    s = ct.mma(q, ct.transpose(kt)) * scale",
    "    s = ct.where(keep(qi, k), s, -1e30)",
    "    m_new = ct.maximum(m, ct.max(s, axis=2))",
    "    p = ct.exp2(s - ct.expand_dims(m_new, 2))",
    "    a = ct.exp2(m - m_new)    # rescale",
    "    l = l * a + ct.sum(p, axis=2)",
    "    acc = ct.mma(p.astype(f16), load_v(k),",
    "                 acc * ct.expand_dims(a, 2))",
    "    m = m_new",
    "ct.store(O, (bh, qi, 0), acc / ct.expand_dims(l, 2))",
]
LH = Inches(0.225)


def code_panel(slide, x, y, w, h, name, lines, colour, band_rows):
    vline(slide, x, y, h, colour=colour, thick=Pt(3.0))
    txt(slide, x + Inches(0.16), y - Inches(0.02), Inches(3.0), name, size=13, bold=True, colour=colour)
    text_top = y + Inches(0.34)
    for r0, r1 in band_rows:
        rect(slide, x + Inches(0.1), text_top + LH * r0 - Inches(0.02),
             w - Inches(0.2), LH * (r1 - r0 + 1) + Inches(0.02), RGBColor(0xEC, 0xE9, 0xE1))
    box = slide.shapes.add_textbox(int(x + Inches(0.14)), int(text_top),
                                   int(w - Inches(0.2)), int(h - Inches(0.4)))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.line_spacing = 1.0
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(13)
        r.font.name = MONO
        r.font.color.rgb = GREY if ln.strip().startswith(("#", "//")) else INK


code_panel(s, M, Inches(1.66), Inches(5.85), Inches(3.95), "cuda c++  \u00b7  sketch", cuda_lines,
           CUDA, [(11, 13)])
code_panel(s, Inches(6.85), Inches(1.66), Inches(5.85), Inches(3.95), "cutile  \u00b7  sketch",
           tile_lines, TILE, [(9, 12)])

txt(s, M, Inches(5.72), Inches(5.6), Inches(0.3), "32 registers, rescaled by hand", size=14,
    bold=True, colour=CUDA)
txt(s, Inches(6.85), Inches(5.72), Inches(5.6), Inches(0.3), "one tile multiply", size=14,
    bold=True, colour=TILE)
hline(s, M, Inches(6.12), CW, colour=HAIR)
txt(s, M, Inches(6.22), Inches(11.9), Inches(0.34),
    "Both scale before the max, mask the tail, and rescale before accumulating. That order is what "
    "keeps the denominator from underflowing.", size=12, colour=GREY)
footer(s, "Sketches, not compiled. cuTile API details are unverified until the kernel runs.", 5)
notes(s, "Left is what I write per thread, right is what I write per tile. The rescale is 32 "
         "registers by hand on the left and one multiply on the right. Whether that one line costs "
         "speed is what the profiler will tell me.")

# =================================================================== 6 stack
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the stack  \u00b7  installed and verified here")
title(s, "Two lanes to the same cubin.", size=34)

lane_y = [Inches(1.68), Inches(2.5)]
lanes = [("cuda c++", CUDA, ["attention.cu", "nvcc", "ptx", "ptxas", "cubin"]),
         ("cutile", TILE, ["attention.py", "cuda.tile", "tileir", "tileiras", "cubin"])]
for li, (name, col, steps) in enumerate(lanes):
    y = lane_y[li]
    label(s, M, y + Inches(0.1), Inches(1.25), name, colour=col, size=11)
    x = M + Inches(1.3)
    bw = Inches(1.55)
    for si, st in enumerate(steps):
        vline(s, x, y, Inches(0.44), colour=col, thick=Pt(2.0))
        txt(s, x + Inches(0.12), y + Inches(0.11), bw - Inches(0.15), st, size=12,
            colour=INK, font=MONO)
        if si < len(steps) - 1:
            txt(s, x + bw - Inches(0.04), y + Inches(0.08), Inches(0.26), "\u2192", size=13,
                colour=GREY)
        x += bw + Inches(0.24)
    txt(s, x + Inches(0.06), y + Inches(0.11), Inches(1.6), "GPU sm_86", size=12, bold=True,
        colour=INK, font=MONO)

label(s, M, Inches(3.35), Inches(6.0), "hardware, as measured not quoted")
specs = [("RTX 3050", "laptop GPU"), ("4 GB", "GDDR6, ~3.8 free"),
         ("~192 GB/s", "measured next"), ("sm_86", "Ampere, CC 8.6")]
for i, (big, small) in enumerate(specs):
    x = M + i * Inches(3.0)
    txt(s, x, Inches(3.6), Inches(2.8), big, size=26, bold=True, colour=INK, font=MONO)
    txt(s, x, Inches(3.98), Inches(2.8), small, size=11, colour=GREY)

label(s, M, Inches(4.62), Inches(4.6), "inert on this gpu, in both models")
txt(s, M + Inches(4.75), Inches(4.62), Inches(0.9), "cuda", size=10, colour=GREY, font=MONO)
txt(s, M + Inches(5.5), Inches(4.62), Inches(0.9), "cutile", size=10, colour=GREY, font=MONO)
inert = [("tma", "sm_90 and later"), ("thread-block clusters", "sm_90 and later"),
         ("fp8 / block-scaled mma", "Ada and later")]
for i, (feat, since) in enumerate(inert):
    y = Inches(4.9) + i * Inches(0.28)
    txt(s, M, y, Inches(2.3), feat, size=12, colour=GREY, font=MONO)
    txt(s, M + Inches(2.4), y, Inches(2.2), since, size=12, colour=GREY)
    txt(s, M + Inches(4.85), y, Inches(0.4), "\u2014", size=12, colour=GREY, font=MONO)
    txt(s, M + Inches(5.6), y, Inches(0.4), "\u2014", size=12, colour=GREY, font=MONO)

label(s, Inches(7.55), Inches(4.62), Inches(5.0), "installed, not intended")
versions = ["cuda-tile  1.7.0",
            "tileiras, nvvm, nvrtc, nvcc  13.4.92",
            "cupy-cuda13x  14.2.0   numpy  2.5.4"]
for i, v in enumerate(versions):
    txt(s, Inches(7.55), Inches(4.92) + i * Inches(0.28), Inches(5.1), v, size=12,
        colour=INK, font=MONO)

txt(s, M, Inches(5.85), Inches(11.9), Inches(0.8),
    ["Three of cuTile's knobs do nothing on Ampere, and CUDA C++ cannot reach those instruction "
     "families either. That is a result for the portability axis, not a limitation to apologise for.",
     "Still to install, needed only for the CUDA half and profiling: ncu, compute-sanitizer."],
    size=12, colour=GREY, line=1.28, space=4)
footer(s, "Verified with pip list, nvidia-smi and a smoke test on this machine. Missing: ncu and "
          "compute-sanitizer.", 6)
notes(s, "Two lanes, one destination. nvcc turns C++ into PTX and ptxas assembles it; cuTile goes "
         "through TileIR and tileiras. The versions are what is actually installed.")

# ============================================================= 7 measurement
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the measurement")
title(s, "One iteration, drawn, and the four figures it feeds.", size=34)

ty = Inches(2.05)
# warm-up block
rect(s, M, ty, Inches(3.5), Inches(0.44), RGBColor(0xE4, 0xE1, 0xD8))
txt(s, M + Inches(0.12), ty + Inches(0.11), Inches(3.3), "warm-up \u00d7 50", size=12,
    colour=INK, font=MONO)
# flush, deliberately OUTSIDE the bracket
fx = M + Inches(3.7)
rect(s, fx, ty, Inches(1.35), Inches(0.44), RGBColor(0xE4, 0xE1, 0xD8))
txt(s, fx + Inches(0.1), ty + Inches(0.11), Inches(1.2), "L2 flush", size=12, colour=INK, font=MONO)
# timed window
wx = fx + Inches(1.55)
rect(s, wx, ty, Inches(3.4), Inches(0.44), TILE)
txt(s, wx + Inches(0.1), ty + Inches(0.11), Inches(3.2), "event start \u2192 kernel \u2192 stop",
    size=12, colour=PAPER, font=MONO)
# bracket around the timed window only
by = ty + Inches(0.56)
vline(s, wx, by, Inches(0.2), colour=INK, thick=Pt(1.0))
vline(s, wx + Inches(3.4), by, Inches(0.2), colour=INK, thick=Pt(1.0))
hline(s, wx, by + Inches(0.18), Inches(3.4), colour=INK, thick=Pt(1.0))
txt(s, wx, by + Inches(0.24), Inches(3.4), "only this is timed", size=11, colour=INK,
    align=PP_ALIGN.CENTER)
txt(s, M, by + Inches(0.24), Inches(5.1),
    "the flush is outside the window on purpose", size=11, colour=GREY)
txt(s, M, Inches(3.1), Inches(11.9), Inches(0.6),
    "Time to a duration, not a count: at least one second of GPU work per configuration, so the "
    "number is steady state and not clock ramp.", size=15, colour=INK, line=1.25)
txt(s, M, Inches(3.72), Inches(11.9), Inches(0.6),
    "Report p50 and p90, and log clocks, power and throttle state on every row. Three timers have "
    "to agree, or it is launch overhead wearing a lab coat.", size=15, colour=INK, line=1.25)

label(s, M, Inches(4.55), Inches(6.0), "the four figures the report will contain")
figs = [("Fig 1", "latency vs N"), ("Fig 2", "DRAM bytes vs N"),
        ("Fig 3", "TFLOP/s vs peak"), ("Fig 4", "velocity, tunability")]
for i, (num, name) in enumerate(figs):
    x = M + i * Inches(3.0)
    vline(s, x, Inches(5.05), Inches(1.1), colour=INK, thick=Pt(1.0))
    hline(s, x, Inches(6.15), Inches(2.5), colour=INK, thick=Pt(1.0))
    txt(s, x, Inches(4.85), Inches(2.5), num, size=11, bold=True, colour=GREY, font=MONO)
    txt(s, x, Inches(6.24), Inches(2.5), name, size=12, colour=INK)
footer(s, "Empty axes on purpose: they are the axes the CSV will fill. Artefacts are results.csv, "
          "profiles/*.ncu-rep, figs/*.pdf, report.md.", 7)
notes(s, "The protocol exists to survive scrutiny. The flush sits outside the timed window, and if "
         "it drifted inside, fusion would look free. Every row carries clock and temperature, "
         "because this is a laptop that throttles.")

# ============================================================ 8 definition
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "definition of done  \u00b7  each of these can fail")
title(s, "Six checks decide whether this is finished.", size=34)

checks = [
    ("Correct on every size", "max|out \u2212 ref_fp64| \u2264 5e-3 \u00b7 max|ref|"),
    ("Numerically stable", "no NaN or Inf; inputs \u00d7100 stay finite"),
    ("Both models, one contract", "identical shapes, produced by the same harness"),
    ("Performance is measured", "p50/p90 at N=2048, DRAM bytes from ncu"),
    ("Profiling evidence exists", "one ncu report per kernel, section Tile for cuTile"),
    ("The argument is made", "five axes, one figure each, a recommendation"),
]
row_h = Inches(0.58)
gate_top = Inches(1.78)
for i, (name, cond) in enumerate(checks):
    y = gate_top + i * row_h
    rect(s, M, y + Inches(0.1), Inches(0.22), Inches(0.22), PAPER)
    vline(s, M, y + Inches(0.1), Inches(0.22), colour=INK, thick=Pt(1.0))
    hline(s, M, y + Inches(0.1), Inches(0.22), colour=INK, thick=Pt(1.0))
    txt(s, M + Inches(0.4), y + Inches(0.06), Inches(3.9), name, size=15, bold=True, colour=INK)
    txt(s, M + Inches(4.45), y + Inches(0.1), Inches(4.3), cond, size=13, colour=GREY)
    txt(s, W - M - Inches(0.95), y + Inches(0.08), Inches(0.95), "pending", size=11,
        colour=GREY, font=MONO, align=PP_ALIGN.RIGHT)
    if i < 5:
        hline(s, M, y + row_h - Inches(0.06), CW, colour=HAIR)

# hard gate bracket across checks 1-4, clear of the status column
bx = W - M - Inches(2.0)
vline(s, bx, gate_top + Inches(0.1), row_h * 4 - Inches(0.24), colour=CUDA, thick=Pt(2.0))
txt(s, bx - Inches(1.85), gate_top - Inches(0.06), Inches(1.75), "hard gate  \u25b8", size=10,
    bold=True, colour=CUDA, font=MONO, align=PP_ALIGN.RIGHT)

label(s, M, Inches(5.4), Inches(6.0), "sizes, with the ones that catch the tail bug marked")
chips = ["1", "65", "130", "1000", "256", "512", "1024", "2048", "4096"]
x = M
for c in chips:
    awkward = c in ("1", "65", "130", "1000")
    w = Inches(0.72)
    rect(s, x, Inches(5.68), w, Inches(0.34), CUDA if awkward else RGBColor(0xE4, 0xE1, 0xD8))
    txt(s, x, Inches(5.73), w, c, size=12, colour=(PAPER if awkward else INK), font=MONO,
        align=PP_ALIGN.CENTER)
    x += w + Inches(0.12)
txt(s, M, Inches(6.12), Inches(11.9), Inches(0.5),
    "A one-key tail error is about 1e-3, below tolerance: N=4095 passes with the bug, N=1000 does "
    "not. That is why the marked sizes exist.", size=12, colour=GREY)
footer(s, "The tolerance is my choice, the brief prescribes none, and it is stated so it can be "
          "argued with.", 8)
notes(s, "Every row can fail, which is the point. The marked sizes are the ones where a missing mask "
         "produces an error above tolerance. Row six is where a quarter of the grade sits, and it "
         "is writing, not code.")

# ================================================================ 9 the plan
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "the plan  \u00b7  seven days, five gates")
title(s, "What happens on which day, and what gets cut first.", size=34)

days = ["Sat 31", "Sun 1", "Mon 2", "Tue 3", "Wed 4", "Thu 5", "Fri 6", "Sat 7"]
gx0 = M + Inches(3.15)
col_w = (CW - Inches(3.15)) / 8
for i, d in enumerate(days):
    x = gx0 + i * col_w
    txt(s, x, Inches(1.72), col_w, d, size=11, bold=True,
        colour=INK if i else CUDA, font=MONO)
vline(s, gx0 + Inches(0.02), Inches(1.95), Inches(4.05), colour=CUDA, thick=Pt(1.5))
txt(s, gx0 + Inches(0.08), Inches(1.5), Inches(1.4), "today", size=10, colour=CUDA, font=MONO)

rows = [
    ("gate green", 0, 0, PASS, "environment"),
    ("bench.py", 1, 1, PASS, "contract"),
    ("cuTile kernel", 2, 2, TILE, "kernels"),
    ("CUDA kernel", 3, 3, CUDA, "kernels"),
    ("sweeps", 4, 4, PASS, "numbers"),
    ("ncu profiles", 5, 5, PASS, "evidence"),
    ("report", 6, 6, PASS, "argument"),
    ("submit", 7, 7, PASS, "submit"),
]
for i, (name, d0, d1, col, gate) in enumerate(rows):
    y = Inches(2.05) + i * Inches(0.5)
    txt(s, M, y + Inches(0.08), Inches(3.0), name, size=13, colour=INK)
    txt(s, M + Inches(2.05), y + Inches(0.12), Inches(1.0), gate, size=10, colour=GREY, font=MONO)
    rect(s, gx0 + d0 * col_w + Inches(0.05), y + Inches(0.1),
         col_w * (d1 - d0 + 1) - Inches(0.1), Inches(0.3), col)
    diamond(s, gx0 + (d1 + 1) * col_w - Inches(0.05), y + Inches(0.25), Inches(0.17), INK)

hline(s, M, Inches(6.15), CW, colour=HAIR)
label(s, M, Inches(6.26), Inches(6.0), "cut first, in this order", colour=CUDA)
txt(s, M, Inches(6.5), Inches(6.2), "GQA and the sweep figure  \u2192  Fig 4  \u2192  the debuggability "
    "table", size=12, colour=GREY)
label(s, Inches(7.0), Inches(6.26), Inches(4.0), "never cut", colour=PASS)
txt(s, Inches(7.0), Inches(6.5), Inches(5.6),
    "both kernels, the correctness numbers, the profiles", size=12, colour=GREY)
footer(s, "The brief's own milestones, 10/07 to 10/28, have passed. This is the recovery plan.", 9)
notes(s, "I am behind the brief's milestones and this is the recovery plan. Two kernels, measured "
         "properly, with a report that takes a position. The cut list is written now so that "
         "decision is not made at midnight on the sixth.")

# ============================================================ 10 references
s = new(prs)
bg(s, PAPER)
label(s, M, Inches(0.5), Inches(9.0), "references")
title(s, "Where this comes from.", size=34)

label(s, M, Inches(1.7), Inches(6.0), "the three that carry the most weight")
primary = [("cuTile Python documentation", "docs.nvidia.com/cuda/cutile-python"),
           ("NVIDIA/TileGym: matmul and attention kernels", "github.com/NVIDIA/TileGym"),
           ("Dao et al., FlashAttention, 2022", "arxiv.org/abs/2205.14135")]
for i, (name, url) in enumerate(primary):
    y = Inches(2.0) + i * Inches(0.72)
    txt(s, M, y, Inches(7.0), name, size=17, bold=True, colour=INK)
    txt(s, M, y + Inches(0.3), Inches(7.0), url, size=12, colour=TILE, font=MONO)

vline(s, Inches(8.05), Inches(1.7), Inches(3.6), colour=HAIR)
label(s, Inches(8.3), Inches(1.7), Inches(4.0), "supporting")
secondary = [("Milakov and Gimelshein, online softmax", "arxiv.org/abs/1801.05824"),
             ("Dao, FlashAttention-2", "arxiv.org/abs/2307.08691"),
             ("Vaswani et al., Attention Is All You Need", "arxiv.org/abs/1706.03762"),
             ("cutile-python samples", "github.com/NVIDIA/cutile-python"),
             ("Nsight Compute guide", "docs.nvidia.com/nsight-compute/ProfilingGuide")]
for i, (name, url) in enumerate(secondary):
    y = Inches(2.0) + i * Inches(0.6)
    txt(s, Inches(8.3), y, Inches(4.4), name, size=12, colour=INK)
    txt(s, Inches(8.3), y + Inches(0.24), Inches(4.4), url, size=10, colour=GREY, font=MONO)

hline(s, M, Inches(5.62), CW, colour=HAIR)
label(s, M, Inches(5.74), Inches(6.0), "learned from, not cited")
txt(s, M, Inches(6.0), Inches(11.9),
    "3Blue1Brown, Deep Learning ch.6   \u00b7   Alammar, The Illustrated Transformer   \u00b7   "
    "Georgia Tech, Transformer Explainer", size=12, colour=GREY)
footer(s, "Repository, plan and decision log: " + REPO + "   (the repo name is as created, typo "
          "included)", 10)
notes(s, "Three sources carry the technical weight: the cuTile docs, TileGym's attention kernels, "
         "and the FlashAttention paper. The rest is support and background.")

# ============================================================ 11 risks/asks
s = new(prs)
bg(s, DARK)
label(s, M, Inches(0.5), Inches(9.0), "risks and asks", colour=DARKGREY)
title(s, "What could stop this, and what I need from you.", colour=DARKTX, size=34)

# 2x2 likelihood / impact plot
px, py, pw, ph = M, Inches(1.85), Inches(4.6), Inches(3.2)
vline(s, px, py, ph, colour=DARKGREY, thick=Pt(1.0))
hline(s, px, py + ph, pw, colour=DARKGREY, thick=Pt(1.0))
txt(s, px, py - Inches(0.28), Inches(3.0), "impact \u2191", size=10, colour=DARKGREY, font=MONO)
txt(s, px + pw - Inches(1.4), py + ph + Inches(0.08), Inches(1.4), "likelihood \u2192", size=10,
    colour=DARKGREY, font=MONO, align=PP_ALIGN.RIGHT)
risks = [(0.28, 0.78, "laptop throttling"), (0.30, 0.34, "4 GB of VRAM"),
         (0.72, 0.72, "cuTile on sm_86"), (0.78, 0.30, "seven days, one person")]
for fx, fy, name in risks:
    cx, cy = px + pw * fx, py + ph * (1 - fy)
    diamond(s, cx, cy, Inches(0.17), CUDA_L)
    txt(s, cx + Inches(0.14), cy - Inches(0.11), Inches(1.8), name, size=11, colour=DARKTX, font=MONO)
txt(s, px, py + ph + Inches(0.34), Inches(4.7),
    "Each is handled: clocks logged per row, sizes fixed now, kernel before harness, cut list "
    "written.", size=12, colour=DARKGREY, line=1.25)

# the three questions, as the hero
qx = Inches(6.2)
label(s, qx, Inches(1.6), Inches(6.0), "three answers i need", colour=TILE_L)
qa = [("Attention only, no GEMM?", "the brief says choose either but lists artefacts for both"),
      ("Does 07 November hold?", "given the restart on 14 October"),
      ("Is N = 2048 the right target?", "it sets every benchmark size I fix tomorrow")]
for i, (q, why) in enumerate(qa):
    y = Inches(1.95) + i * Inches(1.22)
    txt(s, qx, y, Inches(6.4), q, size=21, bold=True, colour=DARKTX, line=1.1)
    txt(s, qx, y + Inches(0.42), Inches(6.4), "if silent: " + why, size=12, colour=DARKGREY,
        font=MONO)
rect(s, qx, Inches(5.65), Inches(6.4), Inches(0.34), RGBColor(0x1B, 0x22, 0x29))
txt(s, qx + Inches(0.12), Inches(5.72), Inches(6.2),
    "default from Tue 3 Nov: OP2 only, N = 2048, GQA stays stretch", size=12,
    colour=TILE_L, font=MONO)
txt(s, qx, Inches(6.2), Inches(6.4), "arrives 07 nov: two kernels, three commands, one report",
    size=12, colour=DARKGREY, font=MONO)
footer(s, REPO, 11, dark=True)
notes(s, "Three questions, and what I will assume if I hear nothing. The risks are plotted by "
         "likelihood and impact so the ones worth arguing about are visible. What arrives on the "
         "seventh is one line on purpose.")

# ---------------------------------------------------------------------- save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
