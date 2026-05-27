"""
Build the enhanced presentation with:
- Cover page (IOM302, group members, Amazon orange)
- Table of Contents
- Conclusion slides per section
- Page numbers
- Acknowledgements + Q&A slides

Fix: do NOT call clear_placeholders (causes OOXML validity errors in WPS/PPT).
Fix: use direct XML <a:ln><a:noFill/></a:ln> instead of line.fill.background().
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── Colors ───────────────────────────────────────────────────────────────────
ORANGE   = RGBColor(0xFF, 0x99, 0x00)
NAVY     = RGBColor(0x16, 0x1E, 0x2D)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
LTGRAY   = RGBColor(0xF4, 0xF4, 0xF4)
DKGRAY   = RGBColor(0x33, 0x33, 0x33)
MDGRAY   = RGBColor(0x77, 0x77, 0x77)
LTORANGE = RGBColor(0xFF, 0xCC, 0x66)

# ── Helpers ───────────────────────────────────────────────────────────────────

def no_border(shape):
    """Write <a:ln><a:noFill/></a:ln> directly – compatible with WPS & PPT."""
    spPr = shape._element.spPr
    for old in spPr.findall(qn('a:ln')):
        spPr.remove(old)
    ln = etree.SubElement(spPr, qn('a:ln'))
    etree.SubElement(ln, qn('a:noFill'))


def rect(slide, L, T, W, H, fill):
    """Add a filled rectangle with no visible border."""
    s = slide.shapes.add_shape(1, L, T, W, H)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    no_border(s)
    return s


def txt(slide, text, L, T, W, H,
        sz=16, bold=False, italic=False,
        color=None, align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(L, T, W, H)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size   = Pt(sz)
    r.font.bold   = bold
    r.font.italic = italic
    r.font.name   = 'Calibri'
    if color:
        r.font.color.rgb = color
    return tb


def move_to(prs, from_idx, to_idx):
    lst  = prs.slides._sldIdLst
    elem = lst[from_idx]
    lst.remove(elem)
    lst.insert(to_idx, elem)


def pagenum(slide, W, H, num, light=False):
    color = WHITE if light else MDGRAY
    t = slide.shapes.add_textbox(W - Inches(1.1), H - Inches(0.45),
                                  Inches(0.9), Inches(0.38))
    p = t.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = str(num)
    r.font.size  = Pt(11)
    r.font.name  = 'Calibri'
    r.font.color.rgb = color


def new_blank(prs):
    """Add a new slide using the blank layout.
    We do NOT remove placeholders – that can break OOXML validity.
    Our full-slide background rect will visually cover any transparent placeholders."""
    return prs.slides.add_slide(prs.slide_layouts[1])   # layout 1 = blank (no name)


# ── Load ──────────────────────────────────────────────────────────────────────
prs = Presentation('/home/user/clawdbot-space/presentation.pptx')
W, H = prs.slide_width, prs.slide_height


# ════════════════════════════════════════════════════════════════════════════
# COVER SLIDE  (reuse existing empty slide 0)
# ════════════════════════════════════════════════════════════════════════════
cover = prs.slides[0]

rect(cover, 0, 0, W, H, NAVY)                           # full background
rect(cover, 0, 0, W, Inches(0.65), ORANGE)              # top bar
rect(cover, 0, H - Inches(0.65), W, Inches(0.65), ORANGE)  # bottom bar
rect(cover, Inches(0.55), Inches(0.65), Inches(0.08),   # left accent
     H - Inches(1.3), ORANGE)

txt(cover, 'IOM302  |  E-Commerce Business Strategy',
    Inches(0.8), Inches(0.8), W - Inches(1.6), Inches(0.5),
    sz=14, bold=True, color=LTORANGE)

txt(cover, 'E-Commerce Business', Inches(0.8), Inches(1.4),
    W - Inches(1.6), Inches(0.9), sz=42, bold=True, color=WHITE)
txt(cover, 'Strategies Analysis', Inches(0.8), Inches(2.2),
    W - Inches(1.6), Inches(0.9), sz=42, bold=True, color=ORANGE)

rect(cover, Inches(0.8), Inches(3.12), Inches(4.5), Inches(0.055), ORANGE)  # divider

txt(cover, 'Group Number:  [TBD]',
    Inches(0.8), Inches(3.25), Inches(5), Inches(0.42), sz=13, color=MDGRAY)

txt(cover, 'Group Members', Inches(0.8), Inches(3.75),
    Inches(3.5), Inches(0.38), sz=13, bold=True, color=ORANGE)

members = [
    ('Yiwen Hao',    '2253659'),
    ('Kaiwen Jiang', '2251793'),
    ('Xin He',       '2145653'),
    ('Xuan Wu',      '2254778'),
    ('Youtai Wang',  '2251996'),
    ('Zezheng Liu',  '2143890'),
]
yy = Inches(4.18)
for name, sid in members:
    txt(cover, f'{name:<22}  {sid}', Inches(0.9), yy,
        Inches(5.2), Inches(0.33), sz=12, color=WHITE)
    yy += Inches(0.33)

# Right panel
rect(cover, W - Inches(3.3), Inches(1.3), Inches(2.9), Inches(4.0),
     RGBColor(0x1F, 0x2B, 0x3E))
txt(cover, 'Case Studies', W - Inches(3.1), Inches(1.5),
    Inches(2.5), Inches(0.4), sz=12, bold=True, color=ORANGE)
rect(cover, W - Inches(3.1), Inches(1.95), Inches(2.5), Inches(0.04), ORANGE)
co_list = ['Alibaba', 'Amazon', 'Warby Parker', 'Tesla', 'Zara']
yy = Inches(2.1)
for co in co_list:
    txt(cover, f'▸  {co}', W - Inches(3.1), yy,
        Inches(2.6), Inches(0.38), sz=13, color=WHITE)
    yy += Inches(0.44)


# ════════════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS  (new slide → moved to index 1)
# ════════════════════════════════════════════════════════════════════════════
toc_entries = [
    ('01', 'Alibaba',               7),
    ('02', 'Amazon',               14),
    ('03', 'Warby Parker',         18),
    ('04', 'Tesla',                23),
    ('05', 'Zara',                 28),
    ('06', 'Acknowledgements & Q&A', 35),
]

toc = new_blank(prs)
rect(toc, 0, 0, W, H, WHITE)
rect(toc, 0, 0, W, Inches(1.15), NAVY)
rect(toc, 0, Inches(1.15), W, Inches(0.07), ORANGE)
rect(toc, 0, 0, Inches(0.08), H, ORANGE)

txt(toc, 'Table of Contents',
    Inches(0.4), Inches(0.22), W - Inches(1), Inches(0.75),
    sz=34, bold=True, color=WHITE)

row_h  = Inches(0.80)
start_y = Inches(1.38)
for i, (num, label, pg) in enumerate(toc_entries):
    yy = start_y + i * row_h
    if i % 2 == 0:
        rect(toc, Inches(0.4), yy - Inches(0.06),
             W - Inches(0.55), row_h - Inches(0.04), LTGRAY)
    rect(toc, Inches(0.5), yy + Inches(0.1), Inches(0.46), Inches(0.46), ORANGE)
    txt(toc, num, Inches(0.5), yy + Inches(0.08), Inches(0.46), Inches(0.46),
        sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(toc, label, Inches(1.15), yy + Inches(0.08),
        W - Inches(2.8), Inches(0.5), sz=19, color=DKGRAY)
    txt(toc, str(pg), W - Inches(1.5), yy + Inches(0.08),
        Inches(1.2), Inches(0.5), sz=17, bold=True, color=ORANGE,
        align=PP_ALIGN.RIGHT)
    if i < len(toc_entries) - 1:
        rect(toc, Inches(1.1), yy + row_h - Inches(0.1),
             W - Inches(1.6), Inches(0.012), RGBColor(0xE0, 0xE0, 0xE0))

rect(toc, 0, H - Inches(0.42), W, Inches(0.42), NAVY)
txt(toc, 'IOM302  |  E-Commerce Business Strategies Analysis',
    Inches(0.5), H - Inches(0.38), W - Inches(1), Inches(0.35),
    sz=10, color=LTORANGE)

move_to(prs, len(prs.slides) - 1, 1)

# ─── After TOC insertion, structure is:
# 0=Cover 1=TOC 2-5=imgs 6-10=Alibaba 11=divider 12-14=Amazon
# 15-18=WP 19-22=Tesla 23-27=Zara 28=References  (total 29 slides)


# ════════════════════════════════════════════════════════════════════════════
# CONCLUSION SLIDE FACTORY
# ════════════════════════════════════════════════════════════════════════════
def make_conclusion(prs, company, points):
    s = new_blank(prs)
    rect(s, 0, 0, W, H, WHITE)
    rect(s, 0, 0, Inches(0.38), H, ORANGE)
    rect(s, 0, 0, W, Inches(1.05), NAVY)
    rect(s, 0, Inches(1.05), W, Inches(0.07), ORANGE)
    txt(s, f'{company}  —  Section Conclusion',
        Inches(0.55), Inches(0.18), W - Inches(0.75), Inches(0.75),
        sz=26, bold=True, color=WHITE)
    txt(s, 'KEY TAKEAWAYS',
        Inches(0.55), Inches(1.25), Inches(3.5), Inches(0.38),
        sz=11, bold=True, color=MDGRAY)
    rect(s, Inches(0.55), Inches(1.6), Inches(6), Inches(0.045), ORANGE)
    yy = Inches(1.75)
    for i, point in enumerate(points, 1):
        rect(s, Inches(0.55), yy + Inches(0.05), Inches(0.37), Inches(0.37), ORANGE)
        txt(s, str(i), Inches(0.55), yy + Inches(0.03), Inches(0.37), Inches(0.37),
            sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(s, point, Inches(1.05), yy, W - Inches(1.5), Inches(0.65),
            sz=15, color=DKGRAY)
        yy += Inches(0.88)
    rect(s, 0, H - Inches(0.42), W, Inches(0.42), NAVY)
    txt(s, 'IOM302  |  E-Commerce Business Strategies Analysis',
        Inches(0.5), H - Inches(0.38), W - Inches(1), Inches(0.35),
        sz=10, color=LTORANGE)
    return s


conclusions_data = [
    ('Alibaba', [
        '"User First, AI-Driven" strategy leverages AI personalisation and 88VIP membership to deepen customer intimacy.',
        'Platform ecosystem model generates 71.7% of revenue from customer management services rather than direct sales.',
        'Margin pressure from AI investment and intensifying competition requires a balanced efficiency and growth strategy.',
    ]),
    ('Amazon', [
        'Customer intimacy, long-tail marketplace, and FBA fulfilment form a self-reinforcing e-commerce flywheel.',
        'Hybrid revenue model (products, seller fees, ads, Prime) diversifies income and reduces reliance on single streams.',
        'Governance of data privacy, seller relationships, cost efficiency, and sustainability are key ongoing challenges.',
    ]),
    ('Warby Parker', [
        'DTC omnichannel strategy successfully disrupted the traditional eyewear duopoly with design-forward, affordable products.',
        'Home try-on innovation and physical retail expansion create a differentiated, high-engagement customer journey.',
        'Scalability of cost structure and defensibility of brand differentiation determine long-term competitive position.',
    ]),
    ('Tesla', [
        'DTC e-commerce model with online customisation eliminates dealer costs and enables direct customer data collection.',
        'Vertical integration (battery, software, manufacturing) underpins both cost leadership and product differentiation.',
        'Reputational risks and intensifying EV competition require resilient brand management and strategy diversification.',
    ]),
    ('Zara', [
        'Agile supply chain enables a 2-week design-to-shelf turnaround, creating trend-responsiveness and scarcity urgency.',
        'Omnichannel integration aligns online discovery with in-store fulfilment for a seamless premium retail experience.',
        'Environmental pressure from fast fashion and data privacy regulations demand proactive sustainability governance.',
    ]),
]

for company, points in conclusions_data:
    make_conclusion(prs, company, points)

# 34 slides now (0-33). Conclusions at indices 29-33.
# Move each from tail (always idx=33) to target position, last-to-first:
move_to(prs, 33, 28)   # Zara.c   → after Zara.risk  (idx 27)
move_to(prs, 33, 23)   # Tesla.c  → after Tesla.sugg (idx 22)
move_to(prs, 33, 19)   # WP.c     → after WP.recs    (idx 18)
move_to(prs, 33, 15)   # Amz.c    → after Amz.risks  (idx 14)
move_to(prs, 33, 11)   # Ali.c    → after Ali.recs   (idx 10)


# ════════════════════════════════════════════════════════════════════════════
# ACKNOWLEDGEMENTS  (index 34)
# ════════════════════════════════════════════════════════════════════════════
ack = new_blank(prs)
rect(ack, 0, 0, W, H, NAVY)
rect(ack, 0, 0, W, Inches(0.65), ORANGE)
rect(ack, 0, H - Inches(0.65), W, Inches(0.65), ORANGE)

txt(ack, 'Acknowledgements',
    Inches(0.7), Inches(0.9), W - Inches(1.4), Inches(0.8),
    sz=36, bold=True, color=WHITE)
rect(ack, Inches(0.7), Inches(1.75), Inches(4), Inches(0.06), ORANGE)

ack_text = (
    'We would like to express our sincere gratitude to our course instructor '
    'and teaching assistants for their valuable guidance throughout IOM302.\n\n'
    'We also thank the authors and publishers of the academic works cited in this '
    'presentation for their foundational contributions to e-commerce strategy research.\n\n'
    'This group project was a collaborative effort, and we are grateful for '
    "each team member's dedication and hard work."
)
txt(ack, ack_text, Inches(0.7), Inches(2.0), W - Inches(1.4), Inches(3.2),
    sz=17, color=RGBColor(0xCC, 0xCC, 0xCC))

txt(ack, 'Group Members:',
    Inches(0.7), H - Inches(1.6), Inches(3), Inches(0.4),
    sz=12, bold=True, color=ORANGE)
members_str = '  ·  '.join([f'{n} ({s})' for n, s in members])
txt(ack, members_str,
    Inches(0.7), H - Inches(1.2), W - Inches(1.4), Inches(0.45),
    sz=11, color=RGBColor(0x99, 0x99, 0x99))


# ════════════════════════════════════════════════════════════════════════════
# Q&A  (index 35)
# ════════════════════════════════════════════════════════════════════════════
qa = new_blank(prs)
rect(qa, 0, 0, W, H, NAVY)
rect(qa, 0, 0, W, Inches(0.65), ORANGE)
rect(qa, 0, H - Inches(0.65), W, Inches(0.65), ORANGE)

txt(qa, 'Q & A',
    Inches(0), Inches(1.8), W, Inches(2.0),
    sz=80, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
txt(qa, 'Questions & Answers',
    Inches(0.5), Inches(3.85), W - Inches(1.0), Inches(0.55),
    sz=22, color=WHITE, align=PP_ALIGN.CENTER)
rect(qa, Inches(3.2), Inches(4.5), W - Inches(6.4), Inches(0.055), ORANGE)
txt(qa, 'Thank you for listening!',
    Inches(0.5), Inches(4.7), W - Inches(1.0), Inches(0.5),
    sz=17, color=RGBColor(0xAA, 0xAA, 0xAA), align=PP_ALIGN.CENTER)
txt(qa, 'IOM302  |  E-Commerce Business Strategies Analysis',
    Inches(0.5), Inches(5.4), W - Inches(1.0), Inches(0.45),
    sz=13, color=LTORANGE, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════════
# PAGE NUMBERS
# ════════════════════════════════════════════════════════════════════════════
dark_pages = {1, 35, 36}   # 1-indexed page numbers with dark (navy) bg
for idx, slide in enumerate(prs.slides):
    pg = idx + 1
    pagenum(slide, W, H, pg, light=(pg in dark_pages))


# ── Save ──────────────────────────────────────────────────────────────────────
out = '/home/user/clawdbot-space/presentation_final.pptx'
prs.save(out)
print(f'Saved → {out}')
print(f'Total slides: {len(prs.slides)}')
