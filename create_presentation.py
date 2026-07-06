from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import copy
import os

# ─── Color Palette ───────────────────────────────────────────────────────────
GREEN       = RGBColor(0x00, 0xC8, 0x53)   # vibrant green
RED         = RGBColor(0xFF, 0x17, 0x44)   # vibrant red
GOLD        = RGBColor(0xFF, 0xD7, 0x00)   # gold
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK_BG     = RGBColor(0x0A, 0x0E, 0x1A)   # deep navy-black
DARK_CARD   = RGBColor(0x10, 0x18, 0x2E)
ACCENT_BLUE = RGBColor(0x00, 0x8B, 0xFF)
LIGHT_GRAY  = RGBColor(0xCC, 0xD6, 0xE8)

OUT_PATH = r"C:\Users\mdsom\Desktop\Tahorim AI\Abid\BPL_Presentation.pptx"

# ─── Helpers ─────────────────────────────────────────────────────────────────
def new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs

def blank_slide(prs):
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)

def fill_bg(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, l, t, w, h, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             font_size=18, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, italic=False):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Segoe UI"
    return txb

def add_multiline(slide, lines, l, t, w, h,
                  font_size=16, bold=False, color=WHITE,
                  align=PP_ALIGN.LEFT, spacing_after=6):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for line_text, line_size, line_bold, line_color in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(spacing_after)
        run = p.add_run()
        run.text = line_text
        run.font.size = Pt(line_size if line_size else font_size)
        run.font.bold = line_bold if line_bold is not None else bold
        run.font.color.rgb = line_color if line_color else color
        run.font.name = "Segoe UI"
    return txb

def add_gradient_rect(slide, l, t, w, h, c1: RGBColor, c2: RGBColor, angle=0):
    """Add a shape; python-pptx doesn't support gradients natively via API easily,
    so we inject XML."""
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    # remove existing fill
    for old in spPr.findall(qn('a:solidFill')):
        spPr.remove(old)
    for old in spPr.findall(qn('a:gradFill')):
        spPr.remove(old)
    c1_hex = f"{c1[0]:02x}{c1[1]:02x}{c1[2]:02x}"
    c2_hex = f"{c2[0]:02x}{c2[1]:02x}{c2[2]:02x}"
    grad_xml = f"""<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0">
      <a:srgbClr val="{c1_hex}"/>
    </a:gs>
    <a:gs pos="100000">
      <a:srgbClr val="{c2_hex}"/>
    </a:gs>
  </a:gsLst>
  <a:lin ang="{angle * 60000}" scaled="0"/>
</a:gradFill>"""
    grad_el = parse_xml(grad_xml)
    spPr.insert(list(spPr).index(spPr.find(qn('a:ln'))) if spPr.find(qn('a:ln')) is not None else len(list(spPr)), grad_el)
    return shape

def add_overlay_rect(slide, l, t, w, h, color, opacity=0.85):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    for old in spPr.findall(qn('a:solidFill')):
        spPr.remove(old)
    for old in spPr.findall(qn('a:gradFill')):
        spPr.remove(old)
    clr_hex = f"{color[0]:02x}{color[1]:02x}{color[2]:02x}"
    alpha_val = int(opacity * 100000)
    xml = f"""<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:srgbClr val="{clr_hex}">
    <a:alpha val="{alpha_val}"/>
  </a:srgbClr>
</a:solidFill>"""
    el = parse_xml(xml)
    spPr.insert(list(spPr).index(spPr.find(qn('a:ln'))) if spPr.find(qn('a:ln')) is not None else len(list(spPr)), el)
    return shape

# ─── Animation helpers ────────────────────────────────────────────────────────
NSMAP = {
    'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

def ensure_timing(slide):
    spTree = slide.shapes._spTree
    sld_el = spTree.getparent()
    timing = sld_el.find(qn('p:timing'))
    if timing is None:
        timing = etree.SubElement(sld_el, qn('p:timing'))
    tnLst = timing.find(qn('p:tnLst'))
    if tnLst is None:
        tnLst = etree.SubElement(timing, qn('p:tnLst'))
    par = tnLst.find(qn('p:par'))
    if par is None:
        par = etree.SubElement(tnLst, qn('p:par'))
    cTn_outer = par.find(qn('p:cTn'))
    if cTn_outer is None:
        cTn_outer = etree.SubElement(par, qn('p:cTn'))
        cTn_outer.set('id', '1')
        cTn_outer.set('dur', 'indefinite')
        cTn_outer.set('restart', 'whenNotActive')
        cTn_outer.set('nodeType', 'tmRoot')
    childTnLst = cTn_outer.find(qn('p:childTnLst'))
    if childTnLst is None:
        childTnLst = etree.SubElement(cTn_outer, qn('p:childTnLst'))
    seq = childTnLst.find(qn('p:seq'))
    if seq is None:
        seq = etree.SubElement(childTnLst, qn('p:seq'))
        seq.set('concurrent', '1')
        seq.set('nextAc', 'seek')
    cTn_seq = seq.find(qn('p:cTn'))
    if cTn_seq is None:
        cTn_seq = etree.SubElement(seq, qn('p:cTn'))
        cTn_seq.set('id', '2')
        cTn_seq.set('dur', 'indefinite')
        cTn_seq.set('nodeType', 'mainSeq')
    childTnLst2 = cTn_seq.find(qn('p:childTnLst'))
    if childTnLst2 is None:
        childTnLst2 = etree.SubElement(cTn_seq, qn('p:childTnLst'))
    prevCondLst = seq.find(qn('p:prevCondLst'))
    if prevCondLst is None:
        prevCondLst = etree.SubElement(seq, qn('p:prevCondLst'))
        cond_prev = etree.SubElement(prevCondLst, qn('p:cond'))
        cond_prev.set('evt', 'onPrevClick')
        cond_prev.set('delay', '0')
        tgtEl_prev = etree.SubElement(cond_prev, qn('p:tgtEl'))
        etree.SubElement(tgtEl_prev, qn('p:sldTgt'))
    nextCondLst = seq.find(qn('p:nextCondLst'))
    if nextCondLst is None:
        nextCondLst = etree.SubElement(seq, qn('p:nextCondLst'))
        cond_next = etree.SubElement(nextCondLst, qn('p:cond'))
        cond_next.set('evt', 'onNextClick')
        cond_next.set('delay', '0')
        tgtEl_next = etree.SubElement(cond_next, qn('p:tgtEl'))
        etree.SubElement(tgtEl_next, qn('p:sldTgt'))
    return childTnLst2

_anim_id_counter = 3

def add_fly_in_animation(slide, shape, delay_ms=0, direction='fromBottom'):
    global _anim_id_counter
    childTnLst = ensure_timing(slide)
    shape_id = shape.shape_id
    dur = 600
    par_xml = f"""<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cTn id="{_anim_id_counter}" presetID="2" presetClass="entr" presetSubtype="8" fill="hold"
         grpId="0" nodeType="clickEffect">
    <p:stCondLst>
      <p:cond delay="{delay_ms}"/>
    </p:stCondLst>
    <p:childTnLst>
      <p:animEffect transition="in" filter="fly">
        <p:cBhvr>
          <p:cTn id="{_anim_id_counter+1}" dur="{dur}" fill="hold"/>
          <p:tgtEl>
            <p:spTgt spid="{shape_id}"/>
          </p:tgtEl>
        </p:cBhvr>
      </p:animEffect>
    </p:childTnLst>
  </p:cTn>
</p:par>"""
    _anim_id_counter += 2
    par_el = parse_xml(par_xml)
    childTnLst.append(par_el)

def add_fade_animation(slide, shape, delay_ms=0):
    global _anim_id_counter
    childTnLst = ensure_timing(slide)
    shape_id = shape.shape_id
    dur = 700
    par_xml = f"""<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cTn id="{_anim_id_counter}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"
         grpId="0" nodeType="clickEffect">
    <p:stCondLst>
      <p:cond delay="{delay_ms}"/>
    </p:stCondLst>
    <p:childTnLst>
      <p:animEffect transition="in" filter="fade">
        <p:cBhvr>
          <p:cTn id="{_anim_id_counter+1}" dur="{dur}" fill="hold"/>
          <p:tgtEl>
            <p:spTgt spid="{shape_id}"/>
          </p:tgtEl>
        </p:cBhvr>
      </p:animEffect>
    </p:childTnLst>
  </p:cTn>
</p:par>"""
    _anim_id_counter += 2
    par_el = parse_xml(par_xml)
    childTnLst.append(par_el)

def add_zoom_animation(slide, shape, delay_ms=0):
    global _anim_id_counter
    childTnLst = ensure_timing(slide)
    shape_id = shape.shape_id
    dur = 500
    par_xml = f"""<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cTn id="{_anim_id_counter}" presetID="22" presetClass="entr" presetSubtype="0" fill="hold"
         grpId="0" nodeType="clickEffect">
    <p:stCondLst>
      <p:cond delay="{delay_ms}"/>
    </p:stCondLst>
    <p:childTnLst>
      <p:animEffect transition="in" filter="zoom">
        <p:cBhvr>
          <p:cTn id="{_anim_id_counter+1}" dur="{dur}" fill="hold"/>
          <p:tgtEl>
            <p:spTgt spid="{shape_id}"/>
          </p:tgtEl>
        </p:cBhvr>
      </p:animEffect>
    </p:childTnLst>
  </p:cTn>
</p:par>"""
    _anim_id_counter += 2
    par_el = parse_xml(par_xml)
    childTnLst.append(par_el)

# ─── Slide Builders ──────────────────────────────────────────────────────────

def slide_01_title(prs):
    """Slide 1 – Title Slide"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)

    # Background image (stadium split screen)
    img_path = r"C:\Users\mdsom\Desktop\Tahorim AI\Abid\user_split_image.jpg"
    if os.path.exists(img_path):
        s.shapes.add_picture(img_path, Inches(0), Inches(0), Inches(13.33), Inches(7.5))

    # Translucent overlay
    overlay = add_overlay_rect(s, 0, 0, 13.33, 7.5, DARK_BG, 0.80)

    # Green accent top bar
    top = add_rect(s, 0, 0, 13.33, 0.12, fill_color=GREEN)
    add_fade_animation(s, top, 0)

    # Red accent line
    red_line = add_rect(s, 0, 0.12, 13.33, 0.05, fill_color=RED)
    add_fade_animation(s, red_line, 100)

    # Football emoji decoration
    emoji = add_text(s, "⚽ 🏆", 0.3, 0.3, 3, 0.8, font_size=28, bold=True, color=GOLD)
    add_fade_animation(s, emoji, 200)

    # Main Title
    t1 = add_text(s, "Why Bangladesh Premier League", 0.5, 1.1, 12.3, 1.2,
                  font_size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t1, 300)

    t2 = add_text(s, "Fails to Attract Urban Youth", 0.5, 2.2, 12.3, 1.0,
                  font_size=38, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t2, 500)

    # Subtitle divider line
    div = add_rect(s, 3.5, 3.3, 6.33, 0.05, fill_color=GOLD)
    add_fade_animation(s, div, 700)

    # Subtitle
    sub = add_text(s, "A Research-Based Presentation on Football Culture & Fan Behavior", 0.5, 3.5, 12.3, 0.6,
                   font_size=15, bold=False, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, italic=True)
    add_fade_animation(s, sub, 800)

    # Members card
    card = add_rect(s, 3.8, 4.3, 5.7, 2.5, fill_color=DARK_CARD, line_color=ACCENT_BLUE, line_width=1.5)
    add_fade_animation(s, card, 900)

    pres_by = add_text(s, "Presented by", 3.8, 4.3, 5.7, 0.45,
                       font_size=13, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_fade_animation(s, pres_by, 1000)

    members = [
        "▸  Member 1",
        "▸  Member 2",
        "▸  Member 3",
        "▸  Member 4",
        "▸  Member 5",
    ]
    for i, m in enumerate(members):
        mt = add_text(s, m, 4.1, 4.75 + i * 0.34, 5.1, 0.35,
                      font_size=13, bold=False, color=WHITE, align=PP_ALIGN.LEFT)
        add_fade_animation(s, mt, 1100 + i * 150)

    # Bottom bar
    bot = add_rect(s, 0, 7.3, 13.33, 0.2, fill_color=RGBColor(0x10,0x18,0x30))
    add_fade_animation(s, bot, 0)
    bot_txt = add_text(s, "© Copyright by Tahorim Somrat  |  BPL Football Presentation", 0, 7.28, 13.33, 0.25,
                       font_size=10, color=RGBColor(0x88,0x99,0xBB), align=PP_ALIGN.CENTER)
    add_fade_animation(s, bot_txt, 0)

def slide_02_reality(prs):
    """Slide 2 – The Reality"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)

    # Top accent
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=GREEN)

    # Section badge
    badge = add_rect(s, 0.3, 0.2, 2.8, 0.5, fill_color=ACCENT_BLUE)
    add_zoom_animation(s, badge, 0)
    badge_t = add_text(s, "THE REALITY", 0.3, 0.2, 2.8, 0.5,
                       font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_zoom_animation(s, badge_t, 0)

    # Slide number
    add_text(s, "02", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    # Title
    t = add_text(s, "Bangladesh Loves Football ❤️", 0.5, 0.9, 12.3, 0.75,
                 font_size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 200)

    but = add_text(s, "BUT…", 0.5, 1.65, 12.3, 0.6,
                   font_size=26, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, but, 400)

    no_watch = add_text(s, "Bangladesh Doesn't Watch Local Football ❌", 0.5, 2.25, 12.3, 0.65,
                        font_size=24, bold=True, color=RED, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, no_watch, 600)

    # Divider
    div = add_rect(s, 0.5, 3.05, 12.33, 0.04, fill_color=RGBColor(0x33,0x44,0x66))
    add_fade_animation(s, div, 700)

    # Two comparison cards
    # Left – European Football
    euro_card = add_rect(s, 0.5, 3.2, 5.8, 3.7, fill_color=RGBColor(0x08,0x2A,0x18),
                         line_color=GREEN, line_width=2)
    add_fly_in_animation(s, euro_card, 800)

    add_text(s, "🌍  European Football", 0.7, 3.3, 5.4, 0.55,
             font_size=16, bold=True, color=GREEN, align=PP_ALIGN.CENTER)

    euro_items = ["✔  Millions of Bangladeshi Fans",
                  "✔  Huge Social Media Buzz",
                  "✔  Packed Stadiums",
                  "✔  Global Star Players",
                  "✔  Premium Broadcasting"]
    for i, item in enumerate(euro_items):
        it = add_text(s, item, 0.8, 3.95 + i * 0.53, 5.2, 0.5,
                      font_size=14, color=WHITE)
        add_fade_animation(s, it, 900 + i * 150)

    # Right – BPL
    bpl_card = add_rect(s, 7.03, 3.2, 5.8, 3.7, fill_color=RGBColor(0x2A,0x08,0x08),
                        line_color=RED, line_width=2)
    add_fly_in_animation(s, bpl_card, 800)

    add_text(s, "🇧🇩  Bangladesh Premier League", 7.23, 3.3, 5.4, 0.55,
             font_size=16, bold=True, color=RED, align=PP_ALIGN.CENTER)

    bpl_items = ["❌  Low Stadium Attendance",
                 "❌  Limited Media Attention",
                 "❌  Less Excitement",
                 "❌  No Star Promotion",
                 "❌  Weak Broadcasting"]
    for i, item in enumerate(bpl_items):
        it = add_text(s, item, 7.23, 3.95 + i * 0.53, 5.4, 0.5,
                      font_size=14, color=WHITE)
        add_fade_animation(s, it, 900 + i * 150)

    # VS badge
    vs = add_rect(s, 6.0, 4.5, 1.33, 1.33, fill_color=GOLD)
    add_zoom_animation(s, vs, 850)
    vs_t = add_text(s, "VS", 6.0, 4.6, 1.33, 1.0,
                    font_size=22, bold=True, color=DARK_BG, align=PP_ALIGN.CENTER)
    add_zoom_animation(s, vs_t, 850)

def slide_03_section_header(prs, title="Main Reasons", subtitle="Why Youth Choose European Football Over BPL", color=ACCENT_BLUE):
    """Section divider slide"""
    s = blank_slide(prs)
    add_gradient_rect(s, 0, 0, 13.33, 7.5, RGBColor(0x06,0x0E,0x20), RGBColor(0x0C,0x1E,0x40), 135)

    # Decorative circles
    circ1 = add_rect(s, -1, -1, 4, 4, fill_color=RGBColor(0x00,0x60,0xFF))
    add_fade_animation(s, circ1, 0)

    circ2 = add_rect(s, 10, 4.5, 4, 4, fill_color=RGBColor(0x00,0xA0,0x40))
    add_fade_animation(s, circ2, 0)

    line = add_rect(s, 4.5, 3.4, 4.33, 0.06, fill_color=color)
    add_fade_animation(s, line, 300)

    t1 = add_text(s, title, 0.5, 2.5, 12.3, 1.2,
                  font_size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_zoom_animation(s, t1, 100)

    t2 = add_text(s, subtitle, 0.5, 3.7, 12.3, 0.7,
                  font_size=18, bold=False, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, italic=True)
    add_fade_animation(s, t2, 400)

def slide_04_why_euro(prs):
    """Slide 4 – Why Youth Prefer European Football"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=ACCENT_BLUE)
    add_text(s, "04", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "Why Youth Prefer European Football", 0.5, 0.2, 12.3, 0.75,
                 font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    sub = add_text(s, "Key factors driving youth engagement with European clubs", 0.5, 0.95, 12.3, 0.4,
                   font_size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, italic=True)
    add_fade_animation(s, sub, 200)

    items = [
        ("⚽", "Better Quality Football",    "High-level skills, tactics & competitive matches",    ACCENT_BLUE),
        ("⭐", "World-Famous Players",        "Mbappe, Haaland, Vinicius Jr – global icons",         GOLD),
        ("🎉", "Electric Stadium Atmosphere","Chants, fan culture, sold-out arenas worldwide",       GREEN),
        ("📺", "Superior Broadcasting",      "HD telecasts, multiple camera angles, analytics",      RGBColor(0xAA,0x00,0xFF)),
        ("🌐", "Global Brand Power",         "Massive merchandise, documentaries & digital reach",   RED),
    ]

    for i, (icon, heading, desc, clr) in enumerate(items):
        row = i // 1
        cx = 0.4
        cy = 1.55 + i * 1.08

        card = add_rect(s, cx, cy, 12.5, 0.95, fill_color=DARK_CARD, line_color=clr, line_width=1.5)
        add_fly_in_animation(s, card, 300 + i * 200)

        ic = add_text(s, icon, cx + 0.15, cy + 0.1, 0.8, 0.8, font_size=26, align=PP_ALIGN.CENTER)
        add_fade_animation(s, ic, 400 + i * 200)

        ht = add_text(s, heading, cx + 1.1, cy + 0.08, 4.5, 0.45,
                      font_size=16, bold=True, color=clr)
        add_fade_animation(s, ht, 500 + i * 200)

        dt = add_text(s, desc, cx + 1.1, cy + 0.50, 10.5, 0.4,
                      font_size=13, color=LIGHT_GRAY)
        add_fade_animation(s, dt, 600 + i * 200)

def slide_05_bpl_struggles(prs):
    """Slide 5 – Why BPL Struggles"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=RED)
    add_text(s, "05", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "Why Bangladesh Premier League Struggles", 0.5, 0.2, 12.3, 0.75,
                 font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    sub = add_text(s, "Core problems holding back BPL's growth & popularity", 0.5, 0.95, 12.3, 0.4,
                   font_size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, italic=True)
    add_fade_animation(s, sub, 200)

    items = [
        ("📣", "Poor Marketing",          "No aggressive campaigns; clubs unknown outside hardcore fans",   RED),
        ("📡", "Weak Broadcasting",       "Low-quality streams, no international coverage or distribution", RGBColor(0xFF,0x77,0x00)),
        ("🌟", "No Star Promotion",       "Local players not built into media personalities or heroes",     GOLD),
        ("🏟️", "Poor Match Experience",   "Uncomfortable stadiums, lack of fan zones & entertainment",     RGBColor(0xAA,0x00,0xFF)),
        ("📢", "Lack of Awareness",       "Youth simply don't know fixtures, standings, or club stories",   ACCENT_BLUE),
    ]

    for i, (icon, heading, desc, clr) in enumerate(items):
        cy = 1.55 + i * 1.08
        card = add_rect(s, 0.4, cy, 12.5, 0.95, fill_color=DARK_CARD, line_color=clr, line_width=1.5)
        add_fly_in_animation(s, card, 300 + i * 200)
        ic = add_text(s, icon, 0.55, cy + 0.1, 0.8, 0.8, font_size=26, align=PP_ALIGN.CENTER)
        add_fade_animation(s, ic, 400 + i * 200)
        ht = add_text(s, heading, 1.5, cy + 0.08, 4.0, 0.45, font_size=16, bold=True, color=clr)
        add_fade_animation(s, ht, 500 + i * 200)
        dt = add_text(s, desc, 1.5, cy + 0.50, 11.2, 0.4, font_size=13, color=LIGHT_GRAY)
        add_fade_animation(s, dt, 600 + i * 200)

def slide_06_section_digital(prs):
    slide_03_section_header(prs, "Digital Media & Fan Psychology",
                            "How Online Platforms Shape Football Preferences", GREEN)

def slide_07_social_media(prs):
    """Slide 7 – Social Media Influence"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=GREEN)
    add_text(s, "07", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "Social Media Influence on Football Preference", 0.5, 0.2, 12.3, 0.75,
                 font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    # Platform icons row
    platforms = [
        ("📷", "Instagram", RGBColor(0xE1,0x30,0x6C)),
        ("📘", "Facebook",  ACCENT_BLUE),
        ("🎵", "TikTok",    RGBColor(0x00,0xF2,0xEA)),
        ("▶️", "YouTube",   RED),
        ("🐦", "X (Twitter)", WHITE),
    ]
    for i, (icon, name, clr) in enumerate(platforms):
        cx = 0.7 + i * 2.4
        circ = add_rect(s, cx, 1.2, 1.8, 1.8, fill_color=DARK_CARD, line_color=clr, line_width=2)
        add_zoom_animation(s, circ, 200 + i * 150)
        ic_t = add_text(s, icon, cx, 1.3, 1.8, 0.9, font_size=30, align=PP_ALIGN.CENTER)
        add_zoom_animation(s, ic_t, 200 + i * 150)
        nm_t = add_text(s, name, cx, 2.1, 1.8, 0.45, font_size=13, bold=True, color=clr, align=PP_ALIGN.CENTER)
        add_fade_animation(s, nm_t, 350 + i * 150)

    # Comparison section
    div = add_rect(s, 0.5, 3.3, 12.33, 0.04, fill_color=RGBColor(0x33,0x44,0x66))
    add_fade_animation(s, div, 1000)

    # European clubs column
    euro_h = add_text(s, "🌍  European Clubs", 0.5, 3.5, 5.8, 0.5,
                      font_size=18, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, euro_h, 1100)

    euro_items = [
        "📈  Daily Content Uploads",
        "🎬  Reels, Shorts & Highlights",
        "💬  Millions of Comments & Shares",
        "🌐  Global Viral Campaigns",
        "⚡  Real-time Engagement",
    ]
    for i, item in enumerate(euro_items):
        it = add_text(s, item, 0.5, 4.1 + i * 0.55, 5.8, 0.5, font_size=13, color=WHITE)
        add_fade_animation(s, it, 1200 + i * 100)

    # BPL column
    bpl_h = add_text(s, "🇧🇩  Bangladesh Clubs", 7.0, 3.5, 5.8, 0.5,
                     font_size=18, bold=True, color=RED, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, bpl_h, 1100)

    bpl_items = [
        "📉  Limited Content Output",
        "😴  Outdated Posts & Graphics",
        "🔇  Almost No Online Presence",
        "🚫  No Viral Campaigns",
        "❄️  Very Low Fan Interaction",
    ]
    for i, item in enumerate(bpl_items):
        it = add_text(s, item, 7.0, 4.1 + i * 0.55, 5.8, 0.5, font_size=13, color=WHITE)
        add_fade_animation(s, it, 1200 + i * 100)

    vs2 = add_rect(s, 6.0, 4.2, 1.33, 1.33, fill_color=GOLD)
    add_zoom_animation(s, vs2, 1150)
    add_text(s, "VS", 6.0, 4.3, 1.33, 1.0,
             font_size=22, bold=True, color=DARK_BG, align=PP_ALIGN.CENTER)

def slide_08_fan_psychology(prs):
    """Slide 8 – Fan Psychology"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=GOLD)
    add_text(s, "08", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "Fan Psychology: What Drives People to Support a Club?", 0.5, 0.2, 12.3, 0.75,
                 font_size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    sub = add_text(s, "Understanding what fans truly seek in football", 0.5, 0.95, 12.3, 0.4,
                   font_size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, italic=True)
    add_fade_animation(s, sub, 200)

    factors = [
        ("⭐", "Heroes & Role Models",
         "Fans follow players they admire & aspire to be like.\nBangladesh needs local football heroes.",
         GOLD, 0.4, 1.5),
        ("❤️", "Emotional Connection",
         "Club identity, city pride & community create belonging.\nBPL clubs lack strong fan identities.",
         RED, 7.0, 1.5),
        ("🏆", "Winning Culture",
         "People love to associate with winners & champions.\nEuropean clubs have trophies & history.",
         GREEN, 0.4, 4.0),
        ("📱", "Entertainment Value",
         "Football must compete with Netflix, gaming & YouTube.\nBPL needs to be more thrilling & accessible.",
         ACCENT_BLUE, 7.0, 4.0),
    ]

    for icon, heading, desc, clr, cx, cy in factors:
        card = add_rect(s, cx, cy, 5.8, 2.7, fill_color=DARK_CARD, line_color=clr, line_width=2)
        add_fly_in_animation(s, card, 300)

        ic_t = add_text(s, icon, cx + 0.15, cy + 0.1, 1.0, 1.0, font_size=34)
        add_zoom_animation(s, ic_t, 400)

        ht = add_text(s, heading, cx + 1.3, cy + 0.15, 4.3, 0.5,
                      font_size=16, bold=True, color=clr)
        add_fade_animation(s, ht, 500)

        dt = add_text(s, desc, cx + 0.2, cy + 0.9, 5.4, 1.7,
                      font_size=13, color=LIGHT_GRAY)
        add_fade_animation(s, dt, 600)

def slide_09_section_financial(prs):
    slide_03_section_header(prs, "Financial & Structural Challenges",
                            "Deep-Rooted Problems in BPL Organization", RED)

def slide_10_financial(prs):
    """Slide 10 – Financial Challenges (Cycle)"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=RED)
    add_text(s, "10", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "Financial Challenges: A Vicious Cycle", 0.5, 0.2, 12.3, 0.75,
                 font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    cycle = [
        ("💰", "Low Revenue",       RED),
        ("📉", "Less Investment",   RGBColor(0xFF,0x77,0x00)),
        ("🏚️", "Poor Facilities",   GOLD),
        ("⬇️", "Lower Quality",     RGBColor(0xAA,0xCC,0x00)),
        ("👎", "Less Audience",     GREEN),
        ("🔁", "Again Low Revenue", ACCENT_BLUE),
    ]

    for i, (icon, label, clr) in enumerate(cycle):
        cx = 1.0 + (i % 3) * 3.8
        cy = 1.3 if i < 3 else 4.5
        node = add_rect(s, cx, cy, 3.2, 1.5, fill_color=DARK_CARD, line_color=clr, line_width=2)
        add_zoom_animation(s, node, 200 + i * 250)
        ic_t = add_text(s, icon, cx + 0.1, cy + 0.1, 0.9, 0.8, font_size=26)
        add_fade_animation(s, ic_t, 300 + i * 250)
        lb_t = add_text(s, label, cx + 1.0, cy + 0.4, 2.1, 0.7,
                        font_size=15, bold=True, color=clr)
        add_fade_animation(s, lb_t, 400 + i * 250)

        # Arrows between nodes
        if i < 5:
            if i < 2:
                arr = add_text(s, "→", cx + 3.2, cy + 0.5, 0.6, 0.5,
                               font_size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
            elif i == 2:
                arr = add_text(s, "↓", cx + 1.4, cy + 1.5, 0.5, 1.0,
                               font_size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
            elif i == 3:
                arr = add_text(s, "←", cx - 0.6, cy + 0.5, 0.6, 0.5,
                               font_size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
            elif i == 4:
                arr = add_text(s, "←", cx - 0.6, cy + 0.5, 0.6, 0.5,
                               font_size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
            add_fade_animation(s, arr, 500 + i * 250)

    cycle_label = add_text(s, "🔄  This cycle repeats — breaking it requires bold structural reform!", 
                           0.5, 6.6, 12.3, 0.6, font_size=14, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_fade_animation(s, cycle_label, 1800)

def slide_11_management(prs):
    """Slide 11 – Management Challenges"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=RGBColor(0xFF,0x77,0x00))
    add_text(s, "11", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "Management Challenges Holding BPL Back", 0.5, 0.2, 12.3, 0.75,
                 font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    sub = add_text(s, "Structural reforms are needed at every level of the league", 0.5, 0.95, 12.3, 0.4,
                   font_size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, italic=True)
    add_fade_animation(s, sub, 200)

    items = [
        ("📋", "League Management",  "Professional league governance & transparent operations",      RGBColor(0xFF,0x77,0x00)),
        ("📅", "Match Scheduling",   "Consistent, well-planned fixtures avoiding conflicts",         GOLD),
        ("🤝", "Sponsorship Deals",  "Corporate partnerships to inject capital into clubs",          GREEN),
        ("🧒", "Youth Academies",    "Long-term player development pipelines for future stars",      ACCENT_BLUE),
        ("🎨", "Club Branding",      "Professional logos, kits, identity & fan merchandise",         RGBColor(0xAA,0x00,0xFF)),
    ]

    for i, (icon, heading, desc, clr) in enumerate(items):
        cy = 1.55 + i * 1.08
        card = add_rect(s, 0.4, cy, 12.5, 0.95, fill_color=DARK_CARD, line_color=clr, line_width=1.5)
        add_fly_in_animation(s, card, 300 + i * 200)
        ic = add_text(s, icon, 0.55, cy + 0.1, 0.8, 0.8, font_size=26)
        add_fade_animation(s, ic, 400 + i * 200)
        ht = add_text(s, heading, 1.5, cy + 0.08, 3.5, 0.45, font_size=16, bold=True, color=clr)
        add_fade_animation(s, ht, 500 + i * 200)
        dt = add_text(s, desc, 5.3, cy + 0.25, 7.4, 0.5, font_size=14, color=LIGHT_GRAY)
        add_fade_animation(s, dt, 600 + i * 200)

def slide_12_section_solutions(prs):
    slide_03_section_header(prs, "Solutions & Conclusion",
                            "A Roadmap to Reviving Bangladesh Football", GREEN)

def slide_13_solutions(prs):
    """Slide 13 – How Can Bangladesh Improve?"""
    s = blank_slide(prs)
    fill_bg(s, DARK_BG)
    add_rect(s, 0, 0, 13.33, 0.1, fill_color=GREEN)
    add_text(s, "13", 12.5, 0.2, 0.8, 0.5, font_size=22, bold=True, color=RGBColor(0x33,0x44,0x66))

    t = add_text(s, "How Can Bangladesh Football Improve? 🚀", 0.5, 0.2, 12.3, 0.75,
                 font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t, 0)

    solutions = [
        ("📣", "Better Marketing",        "Run digital campaigns, influencer collaborations & PR drives",   GREEN),
        ("🏟️", "Better Stadium Experience","Fan zones, music, food & family-friendly match day events",      GOLD),
        ("📱", "Strong Social Media",     "Daily content, player features, BTS videos & live engagement",   ACCENT_BLUE),
        ("🏃", "Youth Development",       "Grassroots football programs, school leagues & academies",        RGBColor(0xAA,0x00,0xFF)),
        ("🎨", "Club Branding",           "Build club identity, merchandise, anthems & loyal fan bases",     RED),
    ]

    for i, (icon, heading, desc, clr) in enumerate(solutions):
        cy = 1.55 + i * 1.08
        card = add_rect(s, 0.4, cy, 12.5, 0.95, fill_color=DARK_CARD, line_color=clr, line_width=1.5)
        add_fly_in_animation(s, card, 200 + i * 200)

        chk = add_rect(s, 0.6, cy + 0.25, 0.45, 0.45, fill_color=clr)
        add_zoom_animation(s, chk, 300 + i * 200)
        chk_t = add_text(s, "✓", 0.6, cy + 0.22, 0.45, 0.45,
                         font_size=14, bold=True, color=DARK_BG, align=PP_ALIGN.CENTER)
        add_zoom_animation(s, chk_t, 300 + i * 200)

        ic = add_text(s, icon, 1.2, cy + 0.1, 0.8, 0.8, font_size=26)
        add_fade_animation(s, ic, 400 + i * 200)
        ht = add_text(s, heading, 2.1, cy + 0.08, 3.8, 0.45, font_size=16, bold=True, color=clr)
        add_fade_animation(s, ht, 500 + i * 200)
        dt = add_text(s, desc, 6.1, cy + 0.25, 6.7, 0.5, font_size=13, color=LIGHT_GRAY)
        add_fade_animation(s, dt, 600 + i * 200)

def slide_14_final_message(prs):
    """Slide 14 – Final Message / Conclusion"""
    s = blank_slide(prs)
    add_gradient_rect(s, 0, 0, 13.33, 7.5, RGBColor(0x06,0x10,0x20), RGBColor(0x0A,0x22,0x10), 135)

    # Top accent
    add_rect(s, 0, 0, 13.33, 0.12, fill_color=GREEN)
    add_rect(s, 0, 0.12, 13.33, 0.05, fill_color=GOLD)

    # Decorative shapes
    d1 = add_rect(s, 0, 5.5, 4, 4, fill_color=RGBColor(0x00,0x40,0x18))
    add_fade_animation(s, d1, 0)
    d2 = add_rect(s, 9.5, -1, 4, 4, fill_color=RGBColor(0x00,0x20,0x40))
    add_fade_animation(s, d2, 0)

    emoji = add_text(s, "⚽ 🌟 🇧🇩", 0.5, 0.4, 12.3, 0.8, font_size=32, align=PP_ALIGN.CENTER)
    add_zoom_animation(s, emoji, 100)

    t1 = add_text(s, '"If We Want Global Success,', 0.5, 1.5, 12.3, 1.0,
                  font_size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t1, 300)

    t2 = add_text(s, 'We Must First Build Strong Local Football."',
                  0.5, 2.4, 12.3, 1.0,
                  font_size=32, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_fly_in_animation(s, t2, 600)

    line = add_rect(s, 3.5, 3.6, 6.33, 0.06, fill_color=GOLD)
    add_fade_animation(s, line, 900)

    key_msg = add_text(s, "🔑  Key Takeaway", 0.5, 3.85, 12.3, 0.55,
                       font_size=18, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_fade_animation(s, key_msg, 1000)

    msgs = [
        "Bangladesh has the passion — it needs the structure, promotion & investment.",
        "BPL can grow if youth are engaged through smart digital strategies & quality football.",
        "A strong local league is the foundation of every great footballing nation.",
    ]
    for i, msg in enumerate(msgs):
        m = add_text(s, f"▸  {msg}", 1.0, 4.55 + i * 0.55, 11.3, 0.5,
                     font_size=13, color=LIGHT_GRAY)
        add_fade_animation(s, m, 1100 + i * 200)

    thank = add_text(s, "Thank You! 🙏", 0.5, 6.3, 12.3, 0.7,
                     font_size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_zoom_animation(s, thank, 1700)

    copy_t = add_text(s, "© Copyright by Tahorim Somrat  |  BPL Football Presentation", 0.5, 7.1, 12.3, 0.3,
                      font_size=10, color=RGBColor(0x55,0x77,0x55), align=PP_ALIGN.CENTER)
    add_fade_animation(s, copy_t, 0)

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    prs = new_prs()

    print("Building slides...")
    slide_01_title(prs)                  # 1
    slide_02_reality(prs)                # 2
    slide_03_section_header(prs,         # 3 – section break
        "Main Reasons",
        "Why Urban Youth Choose European Football Over BPL",
        ACCENT_BLUE)
    slide_04_why_euro(prs)               # 4
    slide_05_bpl_struggles(prs)          # 5
    slide_06_section_digital(prs)        # 6 – section break
    slide_07_social_media(prs)           # 7
    slide_08_fan_psychology(prs)         # 8
    slide_09_section_financial(prs)      # 9 – section break
    slide_10_financial(prs)              # 10
    slide_11_management(prs)             # 11
    slide_12_section_solutions(prs)      # 12 – section break
    slide_13_solutions(prs)              # 13
    slide_14_final_message(prs)          # 14

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    prs.save(OUT_PATH)
    print(f"[SUCCESS] Saved to: {OUT_PATH}")

if __name__ == "__main__":
    main()
