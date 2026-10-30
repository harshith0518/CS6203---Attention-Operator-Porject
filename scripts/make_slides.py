"""Generate the OP2 survey deck for the first mentor meeting.

One screen per idea: the six attention variants, each shown as "what it costs in
CUDA C++" vs "what it costs in cuTile", then the plan to 07 Nov.

    .\\.venv\\Scripts\\python.exe scripts\\make_slides.py

Output: docs/op2-attention-slides.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---------------------------------------------------------------- look and feel
INK = RGBColor(0x0F, 0x17, 0x2A)
MUTED = RGBColor(0x5B, 0x6B, 0x84)
ACCENT = RGBColor(0x0E, 0x74, 0x90)      # teal
WARN = RGBColor(0xB4, 0x4A, 0x0B)        # burnt orange
PANEL = RGBColor(0xF1, 0xF5, 0xF9)
RULE = RGBColor(0xCB, 0xD5, 0xE1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

UI = "Segoe UI"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)


def new(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def text(slide, x, y, w, h, s, size=16, bold=False, color=INK, font=UI,
         align=PP_ALIGN.LEFT, space_after=6, line=1.15):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    lines = s if isinstance(s, list) else [s]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
    return box


def header(slide, kicker, title, size=30):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                 Inches(0.14), H)
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()

    text(slide, Inches(0.75), Inches(0.5), Inches(11.6), Inches(0.3),
         kicker.upper(), size=11, bold=True, color=ACCENT)
    text(slide, Inches(0.75), Inches(0.82), Inches(11.9), Inches(0.9),
         title, size=size, bold=True, color=INK, line=1.05)
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.75), Inches(1.72),
                                Inches(11.85), Pt(1.1))
    ln.fill.solid()
    ln.fill.fore_color.rgb = RULE
    ln.line.fill.background()


def bullets(slide, items, x=Inches(0.75), y=Inches(1.95), w=Inches(11.85),
            size=16, gap=0.16):
    """items: list of (level, text, style) — style in {'', 'b', 'muted', 'accent'}"""
    box = slide.shapes.add_textbox(x, y, w, Inches(4.9))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    first = True
    for item in items:
        level, s = item[0], item[1]
        style = item[2] if len(item) > 2 else ""
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = 0
        p.space_after = Pt(gap * 72)
        p.line_spacing = 1.12
        p.space_before = Pt(0)
        pre = "—  " if level == 1 else ""
        r = p.add_run()
        r.text = pre + s
        r.font.size = Pt(size - (1 if level == 1 else 0))
        r.font.name = UI
        r.font.bold = style == "b"
        r.font.color.rgb = (
            MUTED if style == "muted" else
            ACCENT if style == "accent" else
            WARN if style == "warn" else
            INK
        )
        if level == 1:
            p.space_after = Pt(gap * 72 * 0.7)


def table(slide, head, rows, x=Inches(0.75), y=Inches(1.95),
          widths=None, size=11, head_size=11, row_h=0.42):
    nrow, ncol = len(rows) + 1, len(head)
    total = Inches(11.85)
    tbl_shape = slide.shapes.add_table(nrow, ncol, x, y, total, Inches(row_h * nrow))
    tbl = tbl_shape.table
    if widths:
        for i, frac in enumerate(widths):
            tbl.columns[i].width = int(total * frac)
    tbl.rows[0].height = Inches(0.38)
    for r in range(1, nrow):
        tbl.rows[r].height = Inches(row_h)

    for c, s in enumerate(head):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = INK
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = s
        r.font.size = Pt(head_size)
        r.font.bold = True
        r.font.color.rgb = WHITE
        r.font.name = UI

    for r, row in enumerate(rows, start=1):
        for c, s in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = PANEL if r % 2 else WHITE
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = str(s)
            run.font.size = Pt(size)
            run.font.name = UI
            bold = str(s).startswith("**") and str(s).endswith("**")
            run.font.bold = bold or c == 0
            run.font.color.rgb = INK
    return tbl


def code(slide, lines, x=Inches(0.75), y=Inches(2.0), w=Inches(11.85),
         h=Inches(4.6), size=10.5):
    panel = slide.shapes.add_shape(1, x, y, w, h)
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(0x0B, 0x12, 0x20)
    panel.line.fill.background()
    tb = slide.shapes.add_textbox(x + Inches(0.22), y + Inches(0.16),
                                 w - Inches(0.4), h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = False
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(1)
        p.line_spacing = 1.02
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.name = MONO
        r.font.color.rgb = RGBColor(0xD7, 0xE3, 0xF0)
        if ln.strip().startswith("#"):
            r.font.color.rgb = RGBColor(0x7C, 0x8C, 0xA4)


def caption(slide, s, y=Inches(6.62), color=MUTED, size=11.5):
    text(slide, Inches(0.75), y, Inches(11.85), Inches(0.5), s,
         size=size, color=color, line=1.1)


# ---------------------------------------------------------------- build
prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# ---- 1 title
s = new(prs)
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), W, Inches(0.42))
band.fill.solid()
band.fill.fore_color.rgb = ACCENT
band.line.fill.background()

text(s, Inches(0.95), Inches(2.05), Inches(11.4), Inches(0.4),
     "CS-6023 GPU PROGRAMMING  ·  PROJECT OP2", size=13, bold=True, color=ACCENT)
text(s, Inches(0.95), Inches(2.55), Inches(11.4), Inches(1.8),
     ["Attention, written two ways", "CUDA C++ and cuTile"], size=44, bold=True, line=1.05)
text(s, Inches(0.95), Inches(4.45), Inches(11.0), Inches(1.2),
     ["Six attention variants, one new kernel-engineering problem each —",
      "and what each problem costs in each programming model."],
     size=18, color=MUTED, line=1.2)
text(s, Inches(0.95), Inches(5.9), Inches(11.0), Inches(0.9),
     ["Surya Harshith (cs23b010)  ·  first meeting with Pravin Jagtap and Sumanta Datta",
      "30 October 2026"], size=12, color=MUTED, line=1.3)

# ---- 2 the argument
s = new(prs)
header(s, "the argument of this deck",
       "Six variants, one new problem each")
table(s,
      ["#", "Variant", "The one new problem", "CUDA C++", "cuTile"],
      [["1", "Vanilla SDPA", "why fusion is needed at all", "low", "low"],
       ["2", "Causal", "masking, and skipping whole tiles", "medium", "low"],
       ["3", "Fused + online softmax", "streaming softmax, live accumulator rescale", "high", "medium"],
       ["4", "GQA / MQA", "head broadcast; decode is memory-bound", "medium", "very low"],
       ["5", "Sliding window", "sparse tile traversal, ragged grids", "med-high", "low-med"],
       ["6", "Cross-attention", "asymmetric shapes — the one rung CUDA wins", "low-med", "low"]],
      widths=[0.05, 0.21, 0.40, 0.17, 0.13], size=12, row_h=0.52)
caption(s, "The claim: these are not six features. Each rung adds exactly one new "
           "kernel-engineering problem, and each problem costs something different in the two models.")

# ---- 3 what attention computes
s = new(prs)
header(s, "the operator", "What attention computes")
bullets(s, [
    (0, "S = QKᵀ / √D      →  P = softmax(S) row-wise  →  O = PV", "b"),
    (0, "Shapes per head: Q,K,V are [N × D]; S and P are [N × N]; O is [N × D]", ""),
    (0, "FLOPs = 4 · B · H · N² · D  (two matmuls, each 2·N²·D per head)", ""),
    (1, "N = sequence length, D = head dim, B = batch, H = heads", "muted"),
    (0, "Softmax is a row reduction — not a reduction over all N² values", "accent"),
    (0, "Attention starts after the projections: no weights, no RoPE in the kernel", "muted"),
], size=18)
caption(s, "Fixed contract: plain multi-head, self-attention, forward pass, exact, causal as a flag.")

# ---- 4 memory bound
s = new(prs)
header(s, "why fusion is not optional", "The naive version is memory-bound, not compute-bound")
bullets(s, [
    (0, "Naive = 3 kernels: compute S, write it to global memory, read it back for softmax, write P, read P for PV", ""),
    (0, "S and P are [N × N] — at N = 4096, H = 8, fp16 that is 268 MB per tensor", ""),
    (0, "≈ 1.07 GB of intermediate traffic, versus ~17 MB for the fused kernel", "b"),
    (0, "Arithmetic intensity ≈ 64 FLOP/byte vs this card's ridge point ≈ 36 FLOP/byte", ""),
    (1, "ridge point = peak FLOP/s ÷ peak bytes/s = 7 TFLOP/s ÷ 192 GB/s", "muted"),
    (0, "Fusion removes the N² tensors from memory entirely — that is the design, not a tweak", "accent"),
], size=17)
caption(s, "Traffic is O(N²·D). No compiler trick saves you here; only not writing the matrix does.")

# ---- 5 variant 1
s = new(prs)
header(s, "variant 1", "Vanilla scaled dot-product — the reference that justifies everything")
table(s,
      ["", "CUDA C++", "cuTile"],
      [["Shape", "3 launches: GEMM → row-softmax kernel → GEMM",
        "3 @ct.kernel functions"],
       ["Softmax", "one block per row, two passes (max, then sum), block reduction",
        "ct.max(s, axis=1) → ct.exp2 → ct.sum(s, axis=1)"],
       ["Cost", "two N×N tensors round-trip through HBM",
        "the same — cuTile does not stop you being naive"],
       ["Tells us", "why fusion is needed, measured not asserted",
        "the abstraction is not a performance guarantee"]],
      widths=[0.12, 0.44, 0.44], size=12.5, row_h=0.86)
caption(s, "Neither model is saved by its abstraction at this rung. That is the point.")

# ---- 6 variant 2
s = new(prs)
header(s, "variant 2", "Causal — the mask is trivial, skipping the work is not")
bullets(s, [
    (0, "A token may not attend to the future. Mask j > i before the softmax.", ""),
    (0, "CUDA C++: bound the k-loop at the diagonal, plus an if on the diagonal tile", ""),
    (1, "cost: blocks now have wildly unequal runtimes — load balance becomes your problem", "muted"),
    (0, "cuTile:  ct.where(cols[None,:] <= rows[:,None], s, -1e30)", "b"),
    (1, "broadcast gives the [BQ × BK] mask for free; the loop bound is just a dynamic range", "muted"),
    (0, "Numerics: -inf is correct but exp2(-inf − -inf) = NaN. Use fp32 -1e30.", "warn"),
], size=17)
caption(s, "FLOPs halve — but only if you actually skip the tiles. A rectangular grid over a "
           "triangle does all the work at ~50% efficiency.")

# ---- 7 variant 3 algorithm
s = new(prs)
header(s, "variant 3 — the algorithm", "Online softmax: a softmax you can compute in slices")
code(s, [
    "m = -1e30 (finite!),  l = 0,  O = 0        # running max, running denominator, accumulator",
    "for each KV tile k:",
    "    S_k   = Q_tile · K_kᵀ · scale        # scale = log2(e)/√D  -> base-2 exp",
    "    m_new = max(m, rowmax(S_k))",
    "    P_k   = exp2(S_k − m_new)",
    "    α     = exp2(m − m_new)              # rescale factor for everything already summed",
    "    l     = l·α + rowsum(P_k)",
    "    O     = O·α + P_k · V_k              # rescale BEFORE accumulate",
    "    m     = m_new",
    "O = O / l",
], size=12.5, h=Inches(2.55))
caption(s, "The problem: softmax is a global row reduction, but a tile holds only a slice of the row."
           "  So the reduction must be streaming and correct under rescaling.  (Milakov 2018 · Dao 2022)")

# ---- 8 the centrepiece
s = new(prs)
header(s, "variant 3 — two implementations", "The same algorithm, two costs")
text(s, Inches(0.75), Inches(1.9), Inches(4.9), Inches(0.3),
     "CUDA C++  —  what you write by hand", size=12, bold=True, color=WARN)
bullets(s, [
    (1, "K, V tiles in shared memory, double-buffered with cp.async", ""),
    (1, "O accumulator in registers: ~64 regs/thread", ""),
    (1, "the α rescale applied by hand to those registers", ""),
    (1, "matched mma.sync fragment layouts between QKᵀ and PV", ""),
    (1, "swizzled smem to kill bank conflicts on the K transpose", ""),
    (1, "every line is a place to be 3× off the roofline", "warn"),
], y=Inches(2.3), w=Inches(4.75), size=13, gap=0.13)

code(s, [
    "@ct.kernel",
    "def flash_fwd(Q, K, V, O, scale, causal, BQ, BK, D):",
    "    qi = ct.bid(0)",
    "    q   = Q.tiled_view((BQ,D)).load((qi,0))",
    "    m   = ct.full((BQ,), -1e30, ct.float32)",
    "    l   = ct.zeros((BQ,), ct.float32)",
    "    acc = ct.zeros((BQ,D), ct.float32)",
    "    for k in range(num_kv_tiles(causal, qi)):",
    "        s = ct.mma(q, ct.transpose(K.load(k)), zeros)",
    "        if causal: s = masked(s, qi, k)",
    "        m_new = ct.maximum(m, ct.max(s, axis=1))",
    "        p = ct.exp2(s*scale - expand(m_new, 1))",
    "        a = ct.exp2(m - m_new)",
    "        l = l*a + ct.sum(p, axis=1)",
    "        acc = ct.mma(p.as_f16(), V.load(k), acc*expand(a,1))",
    "        m = m_new",
    "    ct.store(O, (qi,), (acc/expand(l,1)).as_f16())",
], x=Inches(5.75), y=Inches(1.9), w=Inches(6.85), h=Inches(3.25), size=9.5)
caption(s, "In CUDA C++ you pay for fusion in shared-memory orchestration. In cuTile you pay for it "
           "in tile shapes. The abstraction moves the work — it does not remove it.")

# ---- 9 what fusion buys
s = new(prs)
header(s, "variant 3 — the payoff", "Fusion drops traffic 60×, but latency only ~3×")
bullets(s, [
    (0, "Traffic:  O(N²·D)  →  O(N·D)", "b"),
    (1, "at N = 4096: ~1.07 GB moved → ~17 MB", "muted"),
    (0, "Latency: ~5.6 ms → ~1.7 ms", "b"),
    (1, "because once the traffic wall is gone you hit the compute wall:", "muted"),
    (1, "4·B·H·N²·D = 34 GFLOP at ~20 TFLOP/s fp16 ≈ 1.7 ms", "muted"),
    (0, "Fusion converts a memory-bound kernel into a compute-bound one", "accent"),
    (0, "Prediction to test: the crossover is at N ≈ 512 — below it, launch overhead and low "
        "occupancy dominate and fusion buys little", ""),
], size=17)
caption(s, "Anyone reporting a 60× speedup is quoting the traffic number as if it were the time. "
           "The honest figure is 3× time for 60× traffic.")

# ---- 10 GQA
s = new(prs)
header(s, "variant 4", "GQA: one line of index math vs one broadcast")
text(s, Inches(0.75), Inches(1.95), Inches(11.6), Inches(0.4),
     "H query heads share H_kv key/value heads.  group  G = H / H_kv   (MQA is the G = 1 case)",
     size=15, color=MUTED)
table(s,
      ["", "CUDA C++", "cuTile"],
      [["Head map", "int kv_head = q_head / G;  plus hand-written strides",
        "stack the group into the leading tile dim"],
       ["Reuse", "you decide whether to fuse G query heads into one block so the KV tile "
        "loads into shared memory once — and you reason about L2 reuse",
        "ct.mma broadcasts the 2-D KV tile across a 3-D query tile "
        "(documented 3-D × 2-D broadcast)"],
       ["Legality", "any G", "G must be a power of two — 2, 4, 8 are, by construction"]],
      widths=[0.12, 0.44, 0.44], size=12, row_h=1.0, y=Inches(2.55))
caption(s, "This is the slide that decides the argument: the tile abstraction turns a scheduling "
           "decision into a type.  Whether that is a win is exactly what the project measures.")

# ---- 11 decode
s = new(prs)
header(s, "variant 4 — the other regime", "Decode: N = 1, and the roofline flips")
bullets(s, [
    (0, "Prefill (N large): two big matmuls, compute-bound, fusion wins", ""),
    (0, "Decode (N = 1): one long skinny GEMV against the KV cache", "b"),
    (1, "a handful of FLOP per byte ⇒ ~0% of peak FLOP/s, ~100% of bandwidth", "muted"),
    (0, "So the metric changes: report achieved GB/s, not TFLOP/s", "accent"),
    (0, "And it is the one place naive and fused converge again — there is no N² tensor to avoid", ""),
    (0, "Worth one microbenchmark: it is where 'achieved memory bandwidth' from the brief "
        "becomes a real number instead of a phrase", ""),
], size=17)

# ---- 12 window + cross
s = new(prs)
header(s, "variants 5 and 6", "Where the abstraction holds — and where it breaks")
table(s,
      ["", "Sliding window  j ∈ [i−W, i]", "Cross-attention  N ≠ M"],
      [["New problem", "the grid is a band: two bounds, two-sided mask, ragged edges",
        "two different sequence lengths ⇒ two tilings"],
       ["CUDA C++", "k_start = max(0,(q_begin−W)/BK), k_end = min(n,(q_end+BK−1)/BK) + an if",
        "natural: a 1-D block of 128 threads handles a row"],
       ["cuTile", "same two integers as a dynamic range() — barely longer than causal",
        "honest cost: a BQ=64 tile computing one query row wastes 63/64 of it"],
       ["Verdict", "tests the abstraction on irregular work; if the edge evaporates, that is a result",
        "the first rung where CUDA C++ wins"]],
      widths=[0.13, 0.44, 0.43], size=12, row_h=0.92)
caption(s, "A comparison with no case where the hand-written version wins is an advertisement, not an analysis.")

# ---- 13 knobs
s = new(prs)
header(s, "the tunability axis, measured on my hardware",
       "Two-thirds of cuTile's tuning surface is inert on my GPU")
table(s,
      ["Knob", "CUDA C++", "cuTile", "Active on sm_86?"],
      [["tile / block shape", "free (any 3-D block)", "BQ, BK, D — powers of two only", "yes"],
       ["threads / warps", "free", "num_worker_warps, occupancy", "yes"],
       ["smem layout, swizzle", "hand-written", "hidden", "—"],
       ["async pipelining", "cp.async, double buffer", "hidden (compiler)", "—"],
       ["prefetch hint", "manual", "latency = 1..10", "yes"],
       ["TMA", "no such instruction", "allow_tma", "inert (pre-Hopper)"],
       ["thread-block clusters", "no", "num_ctas", "inert (sm_90+)"],
       ["fp8 / block-scaled MMA", "no", "mma_scaled", "raises TileUnsupportedFeatureError"]],
      widths=[0.24, 0.22, 0.26, 0.28], size=11.5, row_h=0.46)
caption(s, "RTX 3050 Laptop = compute capability 8.6 (Ampere). No TMA, no clusters, no FP8. "
           "The abstraction's portability is delivered by exactly the mechanism that empties its tuning surface.")

# ---- 14 plan
s = new(prs)
header(s, "the plan", "OP2 only, three code files, eight days")
table(s,
      ["Day", "Date", "Work", "Done when"],
      [["0", "Fri 30 Oct", "environment gate + this deck", "smoke test passes; deck sent"],
       ["1", "Sat 31 Oct", "bench.py + fp64 and SDPA references", "one command produces the CSV"],
       ["2", "Sun 1 Nov", "fused kernel in cuTile", "matches fp64 in fp16 tolerance"],
       ["3", "Mon 2 Nov", "fused kernel in CUDA C++", "same; both rows in CSV"],
       ["4", "Tue 3 Nov", "tile / block sweep", "best config recorded for both"],
       ["5", "Wed 4 Nov", "Nsight Compute (incl. --section Tile)", "profiles/ populated"],
       ["6", "Thu 5 Nov", "report draft, five axes", "no placeholders"],
       ["7", "Fri 6 Nov", "report final, reproduction script", "stranger regenerates figures"],
       ["8", "Sat 7 Nov", "submit", "submitted"]],
      widths=[0.06, 0.16, 0.42, 0.36], size=11.5, row_h=0.44)
caption(s, "Scope is deliberately minimal: one fused kernel per model, causal as a flag. "
           "GQA and sliding window are stretch — attempted only after both kernels are correct and benchmarked.")

# ---- 15 questions
s = new(prs)
header(s, "what I need from this meeting", "Three questions")
bullets(s, [
    (0, "OP2 only — can you confirm it in writing? The brief's comparison section says "
        "\"choose either OP1 or OP2\", but its artifact list still says \"both operators\".", ""),
    (0, "Does 07 November hold, given we restarted from this survey on 14 October?", ""),
    (0, "Is a fused MHA kernel with causal as a flag the right variant — or would you rather "
        "I push into GQA before 07 Nov?", ""),
], y=Inches(2.1), size=18, gap=0.5)
text(s, Inches(0.75), Inches(5.1), Inches(11.6), Inches(1.2),
     "What I am deliberately not doing: GEMM, paged attention, MLA, the backward pass, "
     "and a hand-written naive baseline (PyTorch's MATH backend already is one, on my card, for free).",
     size=14, color=MUTED, line=1.25)

# ---------------------------------------------------------------- save
out = Path(__file__).resolve().parent.parent / "docs" / "op2-attention-slides.pptx"
prs.core_properties.title = "Attention in CUDA C++ and cuTile"
prs.core_properties.author = "Surya Harshith (cs23b010)"
prs.core_properties.subject = "CS-6023 GPU Programming — Project OP2"
prs.save(out)
print(f"wrote {out}  ({len(prs.slides)} slides, {out.stat().st_size/1024:.0f} KB)")
