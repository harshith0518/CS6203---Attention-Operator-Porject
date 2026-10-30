"""Generate the 7-slide OP2 update deck.

Structure is SCQA, answer first: the complication (memory) leads, because it is
what makes the operator worth optimising. Seven slides, seven distinct layouts —
a repeated template is the clearest "generated" tell.

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py

Output: docs/op2-attention-slides.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---- palette: one dark ground, one accent, one warm tone for the problem
DEEP = RGBColor(0x0C, 0x33, 0x3B)
TEAL = RGBColor(0x02, 0x76, 0x86)
SEA = RGBColor(0x00, 0x9E, 0x8F)
MINT = RGBColor(0x5C, 0xD6, 0xB4)
INK = RGBColor(0x13, 0x22, 0x29)
MUTED = RGBColor(0x5E, 0x6E, 0x76)
PAPER = RGBColor(0xFF, 0xFF, 0xFF)
PANEL = RGBColor(0xEC, 0xF3, 0xF4)
WARM = RGBColor(0xC2, 0x5B, 0x22)
CODEBG = RGBColor(0x0E, 0x21, 0x27)
CODEFG = RGBColor(0xD6, 0xE9, 0xEB)

HEAD = "Cambria"
BODY = "Calibri"
MONO = "Courier New"

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.8)                      # left margin, consistent everywhere
CW = Inches(11.7)                    # content width
REPO = "github.com/harshith0518/CS6203---Attention-Operator-Porject"


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def tb(slide, x, y, w, h, lines, size=16, bold=False, color=INK, font=BODY,
       align=PP_ALIGN.LEFT, space=6, line=1.12):
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


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def tiles(slide, x, y, cell=0.2, gap=0.05, rows=3, cols=3,
          filled=((0, 0), (1, 1), (2, 2)), color=SEA):
    for r in range(rows):
        for c in range(cols):
            rect(slide, Inches(x + c * (cell + gap)), Inches(y + r * (cell + gap)),
                 Inches(cell), Inches(cell),
                 color if (r, c) in filled else RGBColor(0x1B, 0x4B, 0x54))


def label(slide, s, x=M, y=Inches(0.62), color=TEAL, size=11.5):
    tb(slide, x, y, Inches(10.4), Inches(0.28), s.upper(), size=size, bold=True, color=color)


def statement(slide, s, y=Inches(0.95), size=38, color=INK, w=CW, x=M):
    tb(slide, x, y, w, Inches(1.1), s, size=size, bold=True, color=color,
       font=HEAD, line=1.02)


def code(slide, x, y, w, h, lines, size=9.5, title=None, fg=CODEFG):
    panel = rect(slide, x, y, w, h, CODEBG, MSO_SHAPE.ROUNDED_RECTANGLE)
    panel.adjustments[0] = 0.03
    ty = y + Inches(0.4) if title else y + Inches(0.14)
    if title:
        tb(slide, x + Inches(0.24), y + Inches(0.13), w - Inches(0.4), Inches(0.26),
           title, size=11, bold=True, color=MINT)
    box = slide.shapes.add_textbox(x + Inches(0.24), ty, w - Inches(0.46),
                                   h - (ty - y) - Inches(0.08))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.line_spacing = 1.05
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.name = MONO
        r.font.color.rgb = MUTED if ln.strip().startswith(("#", "//")) else fg


def arrow(slide, x, y, w, h, color=SEA):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.fill.background()
    a.shadow.inherit = False


prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# ==================================================================== 1 title
s = new(prs)
rect(s, 0, 0, W, H, DEEP)
tiles(s, 9.95, 1.15, cell=0.4, gap=0.09, rows=5, cols=5,
      filled=((0, 0), (1, 1), (2, 2), (3, 3), (4, 4), (1, 3), (3, 1)))

tb(s, M, Inches(1.7), Inches(8.6), Inches(0.3),
   "CS-6023  \u00b7  GPU KERNEL PROGRAMMING  \u00b7  PROJECT OP2", size=12.5, bold=True, color=MINT)
tb(s, M, Inches(2.2), Inches(8.8), Inches(1.9),
   ["Attention,", "written two ways"], size=58, bold=True, color=PAPER, font=HEAD, line=0.98)
tb(s, M, Inches(4.35), Inches(8.2), Inches(1.0),
   "One operator, two programming models, and the measurement that decides between them.",
   size=18, color=RGBColor(0xC4, 0xDD, 0xDF), line=1.28)
tb(s, M, Inches(6.35), Inches(9.2), Inches(0.8),
   ["Surya Harshith (cs23b010)  \u00b7  for Pravin Jagtap and Sumanta Datta",
    "30 October 2026  \u00b7  " + REPO], size=11.5,
   color=RGBColor(0x8F, 0xB3, 0xB9), line=1.4)

# ============================================================ 2 the problem
s = new(prs)
label(s, "the complication")
statement(s, "Attention is not slow because it computes. It is slow because it moves.")

tb(s, M, Inches(2.35), Inches(4.6), Inches(0.3), "IF NOTHING IS FUSED", size=11, bold=True, color=WARM)
tb(s, M, Inches(2.65), Inches(4.8), Inches(0.9), "1.07 GB", size=54, bold=True, color=WARM, font=HEAD)
tb(s, M, Inches(3.55), Inches(4.6), Inches(0.8),
   ["S and P written out and read back", "N = 4096, 8 heads, fp16"], size=12.5, color=MUTED, line=1.25)

tb(s, Inches(6.15), Inches(2.35), Inches(4.6), Inches(0.3), "FUSED", size=11, bold=True, color=SEA)
tb(s, Inches(6.15), Inches(2.65), Inches(4.6), Inches(0.9), "17 MB", size=54, bold=True, color=TEAL, font=HEAD)
tb(s, Inches(6.15), Inches(3.55), Inches(4.6), Inches(0.8),
   ["the same result, one kernel", "S never reaches memory"], size=12.5, color=MUTED, line=1.25)

# to-scale bars: 1.07 GB vs 17 MB (a sliver)
rect(s, M, Inches(4.65), Inches(9.4), Inches(0.26), RGBColor(0xF0, 0xD3, 0xC1))
rect(s, M, Inches(5.05), Inches(0.15), Inches(0.26), TEAL)
tb(s, M, Inches(5.42), Inches(4.4), Inches(0.3), "bar width is proportional to bytes moved",
   size=10.5, color=MUTED)

tb(s, M, Inches(5.95), CW, Inches(1.0),
   ["32 FLOP per byte against a ridge point near 104 \u2014 so the naive version is memory-bound, "
    "not compute-bound.",
    "Fusion is therefore the design, not a tuning trick. It is also why we expect ~60\u00d7 less "
    "traffic but only ~3\u00d7 less time."], size=13.5, color=INK, line=1.35)

# ======================================================== 3 what we understood
s = new(prs)
label(s, "what we understood")
statement(s, "Same mathematics. Two very different jobs for the person writing it.")

rect(s, M, Inches(2.25), Inches(5.7), Inches(3.5), DEEP)
tb(s, Inches(1.15), Inches(2.5), Inches(5.0), Inches(0.4), "CUDA C++", size=22, bold=True,
   color=PAPER, font=HEAD)
tb(s, Inches(1.15), Inches(2.92), Inches(5.0), Inches(0.3), "thread level", size=12, color=MINT)
for i, (k, v) in enumerate([("You write", "code for one thread"),
                            ("You own", "shared memory, registers, sync"),
                            ("The compiler owns", "nothing \u2014 you already chose")]):
    y = 3.45 + i * 0.72
    tb(s, Inches(1.15), Inches(y), Inches(5.0), Inches(0.24), k.upper(), size=9.5,
       bold=True, color=SEA)
    tb(s, Inches(1.15), Inches(y + 0.24), Inches(5.0), Inches(0.34), v, size=13.5, color=PAPER)

rect(s, Inches(6.65), Inches(2.25), Inches(5.7), Inches(3.5), PANEL)
tb(s, Inches(7.0), Inches(2.5), Inches(5.0), Inches(0.4), "cuTile", size=22, bold=True,
   color=INK, font=HEAD)
tb(s, Inches(7.0), Inches(2.92), Inches(5.0), Inches(0.3), "tile level", size=12, color=TEAL)
for i, (k, v) in enumerate([("You write", "operations on whole tiles"),
                            ("You own", "tile shape and which loop to skip"),
                            ("The compiler owns", "layout, staging, thread mapping")]):
    y = 3.45 + i * 0.72
    tb(s, Inches(7.0), Inches(y), Inches(5.0), Inches(0.24), k.upper(), size=9.5,
       bold=True, color=TEAL)
    tb(s, Inches(7.0), Inches(y + 0.24), Inches(5.0), Inches(0.34), v, size=13.5, color=INK)

tb(s, M, Inches(6.05), CW, Inches(0.7),
   "That is the whole question the project answers: what does the compiler hide, and what does "
   "that hiding cost you?", size=15, color=TEAL, line=1.3)

# ============================================================ 4 the operator
s = new(prs)
label(s, "the operator")
statement(s, "One variant to build. Five to know.")
tb(s, M, Inches(1.78), CW, Inches(0.4),
   "O = softmax(Q K^T / \u221ad + mask) \u00b7 V      B=1, H=8, D=64, fp16 in, fp32 accumulate",
   size=14.5, color=MUTED)

names = [("1", "Vanilla", "baseline only"), ("2", "Causal", "comes free"),
         ("3", "Fused + online softmax", "this one"), ("4", "GQA", "stretch"),
         ("5", "Sliding window", "out"), ("6", "Cross-attention", "out")]
for i, (n, name, tag) in enumerate(names):
    x = M + i * Inches(1.97)
    chosen = i == 2
    box = rect(s, x, Inches(2.6), Inches(1.82), Inches(1.5),
               TEAL if chosen else PANEL, MSO_SHAPE.ROUNDED_RECTANGLE)
    box.adjustments[0] = 0.08
    tb(s, x + Inches(0.16), Inches(2.72), Inches(0.5), Inches(0.3), n, size=11, bold=True,
       color=MINT if chosen else MUTED)
    tb(s, x + Inches(0.16), Inches(3.02), Inches(1.55), Inches(0.7), name, size=12,
       bold=chosen, color=PAPER if chosen else INK, line=1.1)
    tb(s, x + Inches(0.16), Inches(3.68), Inches(1.55), Inches(0.3), tag, size=10.5,
       bold=chosen, color=MINT if chosen else MUTED)

tb(s, M, Inches(4.4), Inches(7.2), Inches(1.6),
   ["Why #3: it is the only variant where the two models face the same problem \u2014 a softmax "
    "you can only see one slice of.",
    "Causal then costs almost nothing (skip the tiles), and GQA stays a stretch goal rather than "
    "a second project."], size=13.5, color=INK, line=1.35)

rect(s, Inches(8.4), Inches(4.35), Inches(3.95), Inches(1.85), PANEL)
tb(s, Inches(8.65), Inches(4.5), Inches(3.5), Inches(0.28), "THE CONTRACT", size=9.5,
   bold=True, color=TEAL)
tb(s, Inches(8.65), Inches(4.78), Inches(3.5), Inches(1.2),
   ["multi-head self-attention", "forward only, exact", "causal as a flag",
    "max|out \u2212 ref| \u2264 5e-3"], size=11.5, color=INK, line=1.24)

# ================================================================= 5 the plan
s = new(prs)
label(s, "the plan")
statement(s, "The same loop, written twice. The difference is who owns the machine.")
code(s, M, Inches(2.15), Inches(6.0), Inches(3.5), [
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
    "}"], size=9, title="CUDA C++  \u00b7  you own the mapping")
code(s, Inches(7.1), Inches(2.15), Inches(5.45), Inches(3.5), [
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
], size=9, title="cuTile  \u00b7  the compiler owns the mapping")

tb(s, M, Inches(5.95), CW, Inches(0.9),
   ["Three files:  cuda/attn/flash.cu   \u00b7   cutile/attn/flash.py   \u00b7   bench/bench.py",
    "The scaling happens before the row maximum \u2014 after it, the denominator underflows to zero."],
   size=12.5, color=MUTED, line=1.35)

# ========================================================= 6 proof and marks
s = new(prs)
label(s, "how we will know it works")
statement(s, "A kernel is not \u201cworking\u201d until a test that could have failed says so.",
          size=34)

tb(s, M, Inches(2.15), Inches(6.3), Inches(0.3), "THE TEST SIZES, AND WHAT EACH ONE CATCHES",
   size=10.5, bold=True, color=TEAL)
rows = [("1, 65, 130, 1000", "a missing tail mask", "phantom keys are a big fraction"),
        ("1023, 2047, 4095", "almost nothing", "one key = 1e-3, under tolerance"),
        ("256 \u2026 4096", "the latency curve", "crossover predicted near 512"),
        ("16384", "the out-of-memory case", "4.29 GB per tensor; fused still runs")]
for i, (n, catches, why) in enumerate(rows):
    y = 2.5 + i * 0.84
    rect(s, M, Inches(y), Inches(6.3), Inches(0.76), PANEL if i % 2 == 0 else PAPER)
    tb(s, M + Inches(0.15), Inches(y + 0.1), Inches(2.0), Inches(0.3), n, size=12.5,
       bold=True, color=INK)
    tb(s, M + Inches(0.15), Inches(y + 0.4), Inches(2.6), Inches(0.3), catches, size=10.5,
       color=WARM if i == 0 else MUTED)
    tb(s, M + Inches(2.9), Inches(y + 0.22), Inches(3.3), Inches(0.4), why, size=10.5, color=MUTED)

tb(s, M, Inches(6.05), Inches(6.3), Inches(0.7),
   "Plus a \u00d7100 magnitude stress, NaN/Inf checks, and three timers that must agree.",
   size=12, color=MUTED, line=1.25)

tb(s, Inches(7.75), Inches(2.15), Inches(4.7), Inches(0.3), "WHERE THE MARKS ARE", size=10.5,
   bold=True, color=TEAL)
marks = [("Correctness", 25), ("Performance", 35), ("Analysis", 25), ("Code + profiling", 15)]
for i, (name, wt) in enumerate(marks):
    y = 2.52 + i * 0.84
    tb(s, Inches(7.75), Inches(y), Inches(2.6), Inches(0.28), name, size=12, color=INK)
    rect(s, Inches(7.75), Inches(y + 0.34), Inches(3.2), Inches(0.2), RGBColor(0xE2, 0xEC, 0xEE))
    rect(s, Inches(7.75), Inches(y + 0.34), Inches(3.2 * wt / 35.0), Inches(0.2), TEAL)
    tb(s, Inches(11.1), Inches(y + 0.28), Inches(0.9), Inches(0.3), f"{wt}%", size=11,
       bold=True, color=TEAL, align=PP_ALIGN.RIGHT)
tb(s, Inches(7.75), Inches(6.05), Inches(4.7), Inches(0.9),
   "40% of it is the argument, not the kernel. That is where the next three slides of effort go.",
   size=12, color=MUTED, line=1.3)

# ========================================================== 7 where we are, ask
s = new(prs)
rect(s, 0, 0, W, H, DEEP)
tiles(s, 11.5, 5.65, cell=0.26, gap=0.06, rows=4, cols=4,
      filled=((0, 0), (2, 2)))
label(s, "where we are", y=Inches(0.55), color=MINT)
statement(s, "Environment done. The kernel is next.", y=Inches(0.85), size=36, color=PAPER)

phases = [("DONE", MINT, ["Scope fixed to OP2 only", "Driver 617.42, preflight clean",
                          "Repo, plan, survey, this deck"]),
          ("NOW", RGBColor(0xF2, 0xB1, 0x7E), ["Gate: cuTile smoke test passes",
                                               "bench.py with fp64 + SDPA rows",
                                               "Fused kernel in cuTile"]),
          ("THEN", RGBColor(0x8F, 0xB3, 0xB9), ["CUDA kernel, sweep, ablations",
                                                "Nsight Compute, report",
                                                "Submit by 07 November"])]
for i, (name, col, items) in enumerate(phases):
    x = M + i * Inches(4.0)
    bar = rect(s, x, Inches(2.15), Inches(3.6), Inches(0.06), col)
    tb(s, x, Inches(1.85), Inches(3.6), Inches(0.26), name, size=11, bold=True, color=col)
    for j, it in enumerate(items):
        tb(s, x, Inches(2.42 + j * 0.42), Inches(3.6), Inches(0.38), it, size=12,
           color=RGBColor(0xD3, 0xE5, 0xE8), line=1.15)
    if i < 2:
        arrow(s, x + Inches(3.68), Inches(2.02), Inches(0.24), Inches(0.32),
              color=RGBColor(0x2C, 0x5C, 0x66))

tb(s, M, Inches(4.05), Inches(6.4), Inches(0.3), "WHAT I AM ASKING FOR", size=10.5,
   bold=True, color=MINT)
tb(s, M, Inches(4.35), Inches(6.6), Inches(1.5),
   ["1.  Confirm OP2 only, in writing.",
    "2.  Does 07 November hold?",
    "3.  GQA before the deadline, or stay stretch?"],
   size=13.5, color=PAPER, line=1.5)

tb(s, Inches(7.55), Inches(4.05), Inches(4.9), Inches(0.3), "RUNNING IN THE OPEN", size=10.5,
   bold=True, color=MINT)
tb(s, Inches(7.55), Inches(4.35), Inches(5.2), Inches(0.3), REPO, size=9.5, color=MINT, font=MONO)
tb(s, Inches(7.55), Inches(4.8), Inches(4.9), Inches(1.2),
   ["plan, survey and decision log are in it.",
    "Blocked, not on code: pip install into .venv,",
    "CUDA Toolkit, and ncu counter permission."],
   size=11.5, color=RGBColor(0xA8, 0xC6, 0xCB), line=1.35)

# ------------------------------------------------------------------ save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming - Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
