"""
BPL Presentation Generator — v3 (Images Prominent)
Why Bangladesh Premier League Fails to Attract Urban Youth
Batch 232 — Textile Department
(c) Copyright by Tahorim Somrat
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import os

# ── Color Palette ─────────────────────────────────────────────────────────────
GREEN       = RGBColor(0x00, 0xC8, 0x53)
DARK_GREEN  = RGBColor(0x00, 0x55, 0x22)
RED         = RGBColor(0xFF, 0x17, 0x44)
GOLD        = RGBColor(0xFF, 0xD7, 0x00)
ORANGE      = RGBColor(0xFF, 0x88, 0x00)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK_BG     = RGBColor(0x07, 0x0B, 0x16)
DARK_CARD   = RGBColor(0x0F, 0x16, 0x2A)
ACCENT_BLUE = RGBColor(0x00, 0x8B, 0xFF)
TEAL        = RGBColor(0x00, 0xD4, 0xAA)
LIGHT_GRAY  = RGBColor(0xCC, 0xD6, 0xE8)
MID_GRAY    = RGBColor(0x55, 0x66, 0x88)

OUT_PATH = r"C:\Users\mdsom\Desktop\Tahorim AI\Abid\BPL_Final.pptx"
DIR      = r"C:\Users\mdsom\Desktop\Tahorim AI\Abid"

IMG_BG      = os.path.join(DIR, "bpl_slide1_bg.png")
IMG_INTL    = os.path.join(DIR, "international_stars.png")
IMG_BD      = os.path.join(DIR, "bangladesh_players.png")
IMG_STADIUM = os.path.join(DIR, "bangladesh_stadium.png")
IMG_EURO    = os.path.join(DIR, "european_stars.png")
IMG_USER    = os.path.join(DIR, "user_split_image.jpg")

# ── Core helpers ──────────────────────────────────────────────────────────────
def new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs

def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def fill_bg(s, c):
    bg = s.background; bg.fill.solid(); bg.fill.fore_color.rgb = c

def rect(s, l, t, w, h, fill=None, border=None, bw=1.5):
    sh = s.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.line.fill.background()
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else:    sh.fill.background()
    if border: sh.line.color.rgb = border; sh.line.width = Pt(bw)
    else:       sh.line.fill.background()
    return sh

def txt(s, text, l, t, w, h, size=16, bold=False, color=WHITE,
        align=PP_ALIGN.LEFT, italic=False):
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p  = tf.paragraphs[0]; p.alignment = align
    r  = p.add_run()
    r.text = text; r.font.size = Pt(size); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color; r.font.name = "Segoe UI"
    return tb

def multiline(s, lines, l, t, w, h, align=PP_ALIGN.LEFT, spacing=8):
    """lines = list of (text, size, bold, color)"""
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for (tx, sz, bd, cl) in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False; p.alignment = align; p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = tx; r.font.size = Pt(sz or 14); r.font.bold = bool(bd)
        r.font.color.rgb = cl or WHITE; r.font.name = "Segoe UI"
    return tb

def grad(s, l, t, w, h, c1, c2, ang=90):
    sh = s.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.line.fill.background()
    sp = sh._element; spPr = sp.find(qn('p:spPr'))
    for x in spPr.findall(qn('a:solidFill')): spPr.remove(x)
    for x in spPr.findall(qn('a:gradFill')):  spPr.remove(x)
    h1 = f"{c1[0]:02x}{c1[1]:02x}{c1[2]:02x}"
    h2 = f"{c2[0]:02x}{c2[1]:02x}{c2[2]:02x}"
    xml = (f'<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
           f' rotWithShape="1"><a:gsLst>'
           f'<a:gs pos="0"><a:srgbClr val="{h1}"/></a:gs>'
           f'<a:gs pos="100000"><a:srgbClr val="{h2}"/></a:gs>'
           f'</a:gsLst><a:lin ang="{ang*60000}" scaled="0"/></a:gradFill>')
    spPr.insert(len(list(spPr)), parse_xml(xml))
    return sh

def overlay(s, l, t, w, h, color, opacity=0.75):
    sh = s.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.line.fill.background()
    sp = sh._element; spPr = sp.find(qn('p:spPr'))
    for x in spPr.findall(qn('a:solidFill')): spPr.remove(x)
    for x in spPr.findall(qn('a:gradFill')):  spPr.remove(x)
    ch = f"{color[0]:02x}{color[1]:02x}{color[2]:02x}"
    av = int(opacity * 100000)
    xml = (f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           f'<a:srgbClr val="{ch}"><a:alpha val="{av}"/></a:srgbClr></a:solidFill>')
    spPr.insert(len(list(spPr)), parse_xml(xml))
    return sh

def pic(s, path, l, t, w, h):
    """Safely add a picture; returns shape or None"""
    if not os.path.exists(path):
        return None
    try:
        return s.shapes.add_picture(path, Inches(l), Inches(t), Inches(w), Inches(h))
    except Exception as e:
        print(f"  [WARN] Could not add {path}: {e}")
        return None

# ── Layout helpers ────────────────────────────────────────────────────────────
def top_bar(s, c=GREEN):
    rect(s, 0, 0, 13.33, 0.08, fill=c)

def footer(s, note=""):
    rect(s, 0, 7.26, 13.33, 0.24, fill=RGBColor(0x06, 0x0C, 0x1E))
    label = note if note else "Why BPL Fails to Attract Urban Youth  |  Batch 232 — Textile Dept"
    txt(s, "(c) Copyright by Tahorim Somrat  |  " + label,
        0.2, 7.27, 12.9, 0.23, size=8, color=MID_GRAY, align=PP_ALIGN.CENTER)

def badge(s, label, l=0.3, t=0.13, w=3.2, c=ACCENT_BLUE):
    rect(s, l, t, w, 0.42, fill=c)
    txt(s, label, l, t, w, 0.42, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

def slide_num(s, n):
    txt(s, f"{n:02d}", 12.55, 0.13, 0.65, 0.42, size=20, bold=True,
        color=RGBColor(0x2A, 0x3A, 0x5A))

def heading(s, text, t=0.65, color=WHITE, size=28):
    return txt(s, text, 0.5, t, 12.33, 0.8, size=size, bold=True,
               color=color, align=PP_ALIGN.CENTER)

def hline(s, t=1.52, c=GREEN, l=0.5, w=12.33):
    rect(s, l, t, w, 0.05, fill=c)

# ── Animation (simple fade) ───────────────────────────────────────────────────
_id = [20]

def _nid():
    _id[0] += 1
    return _id[0]

def _get_seq(s):
    spTree = s.shapes._spTree
    sld    = spTree.getparent()
    timing = sld.find(qn('p:timing'))
    if timing is None: timing = etree.SubElement(sld, qn('p:timing'))
    tnLst = timing.find(qn('p:tnLst'))
    if tnLst is None: tnLst = etree.SubElement(timing, qn('p:tnLst'))
    par = tnLst.find(qn('p:par'))
    if par is None: par = etree.SubElement(tnLst, qn('p:par'))
    cTn = par.find(qn('p:cTn'))
    if cTn is None:
        cTn = etree.SubElement(par, qn('p:cTn'))
        for k,v in [('id','1'),('dur','indefinite'),('restart','whenNotActive'),('nodeType','tmRoot')]:
            cTn.set(k,v)
    cl = cTn.find(qn('p:childTnLst'))
    if cl is None: cl = etree.SubElement(cTn, qn('p:childTnLst'))
    seq = cl.find(qn('p:seq'))
    if seq is None:
        seq = etree.SubElement(cl, qn('p:seq'))
        seq.set('concurrent','1'); seq.set('nextAc','seek')
    c2 = seq.find(qn('p:cTn'))
    if c2 is None:
        c2 = etree.SubElement(seq, qn('p:cTn'))
        for k,v in [('id','2'),('dur','indefinite'),('nodeType','mainSeq')]:
            c2.set(k,v)
    cl2 = c2.find(qn('p:childTnLst'))
    if cl2 is None: cl2 = etree.SubElement(c2, qn('p:childTnLst'))
    if seq.find(qn('p:prevCondLst')) is None:
        pcl = etree.SubElement(seq, qn('p:prevCondLst'))
        c = etree.SubElement(pcl, qn('p:cond')); c.set('evt','onPrevClick'); c.set('delay','0')
        te = etree.SubElement(c, qn('p:tgtEl')); etree.SubElement(te, qn('p:sldTgt'))
    if seq.find(qn('p:nextCondLst')) is None:
        ncl = etree.SubElement(seq, qn('p:nextCondLst'))
        c = etree.SubElement(ncl, qn('p:cond')); c.set('evt','onNextClick'); c.set('delay','0')
        te = etree.SubElement(c, qn('p:tgtEl')); etree.SubElement(te, qn('p:sldTgt'))
    return cl2

def animate(s, shape, delay=0):
    try:
        el   = shape._element
        nvpr = el.find('.//' + qn('p:nvSpPr')) or el.find('.//' + qn('p:nvPicPr'))
        sid  = nvpr.find(qn('p:cNvPr')).get('id') if nvpr is not None else '1'
        seq  = _get_seq(s)
        p, c, e = str(_nid()), str(_nid()), str(_nid())
        xml = (f'<p:par xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
               f' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               f'<p:cTn id="{p}" fill="hold">'
               f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
               f'<p:childTnLst><p:par><p:cTn id="{c}" fill="hold">'
               f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
               f'<p:childTnLst>'
               f'<p:animEffect transition="in" filter="fade">'
               f'<p:cBhvr><p:cTn id="{e}" dur="500" fill="hold"/>'
               f'<p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl></p:cBhvr>'
               f'</p:animEffect>'
               f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
        seq.append(parse_xml(xml))
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════════════════════
#  SLIDES
# ══════════════════════════════════════════════════════════════════════════════

def s01_title(prs):
    """Slide 1 — Title"""
    s = blank(prs); fill_bg(s, DARK_BG)

    # Full-slide background image
    p = pic(s, IMG_BG, 0, 0, 13.33, 7.5)
    # Dark overlay so text is readable
    overlay(s, 0, 0, 13.33, 7.5, DARK_BG, opacity=0.68)

    # Top color bar
    grad(s, 0, 0, 13.33, 0.12, GREEN, ACCENT_BLUE, ang=0)

    # Title card
    overlay(s, 1.5, 0.8, 10.33, 2.8, DARK_CARD, opacity=0.90)
    rect(s,  1.5, 0.8, 10.33, 2.8, border=GREEN, bw=2)

    t1 = txt(s, "Why Bangladesh Premier League", 1.7, 1.0, 9.93, 0.85,
             size=34, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    animate(s, t1, 0)
    t2 = txt(s, "Fails to Attract Urban Youth", 1.7, 1.8, 9.93, 0.8,
             size=34, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    animate(s, t2, 250)
    rect(s, 3.6, 2.62, 6.13, 0.06, fill=GOLD)
    t3 = txt(s, "A Research Presentation — Bangladesh Football Fan Engagement & Youth Disconnect",
             1.7, 2.72, 9.93, 0.55, size=13, italic=True, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
    animate(s, t3, 450)

    # Batch badge
    grad(s, 4.3, 3.75, 4.73, 0.56, DARK_GREEN, RGBColor(0x00,0x44,0x18), ang=0)
    txt(s, "Batch 232   |   Textile Department", 4.3, 3.75, 4.73, 0.56,
        size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Members card
    overlay(s, 0.4, 4.5, 6.0, 2.6, DARK_CARD, opacity=0.92)
    rect(s, 0.4, 4.5, 6.0, 2.6, border=ACCENT_BLUE, bw=1.5)
    txt(s, "PRESENTED BY", 0.6, 4.58, 5.6, 0.4, size=11, bold=True,
        color=ACCENT_BLUE, align=PP_ALIGN.CENTER)
    rect(s, 0.6, 5.0, 5.6, 0.04, fill=RGBColor(0x15,0x25,0x48))

    members = [
        "2230800230  —  Mahmudul Hasan",
        "2230800231  —  Md. Rahat",
        "2230800232  —  Md. Rifat",
        "2230800233  —  Md. Tanvir",
        "2230800234  —  Md. Towhid",
    ]
    for i, m in enumerate(members):
        txt(s, "  " + m, 0.6, 5.08 + i*0.38, 5.6, 0.36, size=12, color=LIGHT_GRAY)

    # Topic overview card
    overlay(s, 7.0, 4.5, 6.0, 2.6, DARK_CARD, opacity=0.92)
    rect(s, 7.0, 4.5, 6.0, 2.6, border=GOLD, bw=1.5)
    txt(s, "PRESENTATION OVERVIEW", 7.2, 4.58, 5.6, 0.4, size=11, bold=True,
        color=GOLD, align=PP_ALIGN.CENTER)
    rect(s, 7.2, 5.0, 5.6, 0.04, fill=RGBColor(0x33,0x28,0x00))
    topics = ["The Reality of BPL vs European Football",
              "Statistical Analysis & Comparison",
              "Why Youth Prefer European Football",
              "Root Causes of BPL's Failure",
              "Actionable Recommendations & Vision"]
    for i, t_ in enumerate(topics):
        txt(s, "  >> " + t_, 7.2, 5.08 + i*0.38, 5.6, 0.36, size=11.5, color=LIGHT_GRAY)

    footer(s)


def s02_reality(prs):
    """Slide 2 — The Reality"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, RED); badge(s, "THE REALITY", c=RED); slide_num(s, 2)

    t1 = txt(s, "Bangladesh Loves Football", 0.5, 0.68, 12.33, 0.72,
             size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    animate(s, t1, 0)
    grad(s, 4.16, 1.46, 5.0, 0.62, RGBColor(0x33,0x10,0x00), RGBColor(0x1A,0x08,0x00), ang=0)
    t2 = txt(s, "B U T  . . .", 4.16, 1.46, 5.0, 0.62,
             size=22, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    animate(s, t2, 300)
    t3 = txt(s, "Bangladesh Doesn't Watch Local Football", 0.5, 2.18, 12.33, 0.68,
             size=27, bold=True, color=RED, align=PP_ALIGN.CENTER)
    animate(s, t3, 550)

    rect(s, 0.5, 3.0, 12.33, 0.05, fill=RGBColor(0x22, 0x33, 0x55))

    # European card
    grad(s, 0.4, 3.15, 5.9, 3.9, RGBColor(0x03,0x1A,0x0A), DARK_CARD, ang=135)
    rect(s, 0.4, 3.15, 5.9, 3.9, border=GREEN, bw=2.0)
    txt(s, "European Football", 0.6, 3.22, 5.5, 0.55,
        size=17, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    rect(s, 0.6, 3.78, 5.5, 0.04, fill=RGBColor(0x08,0x33,0x15))
    euro = ["Millions of Bangladeshi fans watch weekly",
            "Huge social media buzz every matchday",
            "Packed 80,000+ capacity stadiums",
            "International superstars & brand culture",
            "4K broadcast, OTT, highlights worldwide"]
    for i, e in enumerate(euro):
        txt(s, "  [+]  " + e, 0.55, 3.87 + i*0.56, 5.6, 0.52, size=13.5, color=LIGHT_GRAY)

    # BPL card
    grad(s, 7.0, 3.15, 5.9, 3.9, RGBColor(0x1A,0x03,0x03), DARK_CARD, ang=135)
    rect(s, 7.0, 3.15, 5.9, 3.9, border=RED, bw=2.0)
    txt(s, "Bangladesh Premier League", 7.2, 3.22, 5.5, 0.55,
        size=17, bold=True, color=RED, align=PP_ALIGN.CENTER)
    rect(s, 7.2, 3.78, 5.5, 0.04, fill=RGBColor(0x33,0x08,0x08))
    bpl = ["Very low stadium attendance per match",
           "Almost zero social media presence",
           "Limited and poor broadcast coverage",
           "No internationally known star players",
           "No youth marketing or fan programs"]
    for i, b in enumerate(bpl):
        txt(s, "  [x]  " + b, 7.15, 3.87 + i*0.56, 5.6, 0.52, size=13.5, color=LIGHT_GRAY)

    # VS badge
    grad(s, 6.16, 4.5, 1.0, 0.9, RED, GREEN, ang=90)
    txt(s, "VS", 6.16, 4.5, 1.0, 0.9, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    footer(s)


def s03_stats(prs):
    """Slide 3 — Key Statistics"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, GOLD); badge(s, "KEY STATISTICS", c=RGBColor(0xAA,0x77,0x00)); slide_num(s, 3)
    heading(s, "The Numbers Don't Lie", size=29, color=GOLD)
    hline(s, c=GOLD)
    txt(s, "Bangladesh vs. European Football — Viewership & Fan Engagement Data",
        0.5, 1.62, 12.33, 0.45, size=13.5, italic=True, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    cards = [
        # (l,  number,   label,                              color)
        (0.4,  "90%",   "Bangladeshis follow a\nEuropean club",          GOLD),
        (4.71, "3M+",   "Watch EPL live\nevery week",                    GREEN),
        (9.03, "<5%",   "Urban youth who watch\nBPL regularly",          RED),
        (0.4,  "8,000", "Average BPL stadium\nattendance per match",     ORANGE),
        (4.71, "60K",   "Stadium capacity —\nonly 13% ever filled",      RED),
        (9.03, "12%",   "Follow BPL on\nsocial media",                   TEAL),
    ]
    for i, (lx, num, lbl, col) in enumerate(cards):
        row = i // 3; ty = 2.18 + row * 2.3
        lpos = lx
        # Card bg
        grad(s, lpos, ty, 3.9, 2.0, DARK_CARD, RGBColor(0x0C,0x14,0x28), ang=90)
        rect(s, lpos, ty, 3.9, 2.0, border=col, bw=2)
        rect(s, lpos, ty, 3.9, 0.07, fill=col)
        txt(s, num, lpos, ty+0.1, 3.9, 1.0, size=38, bold=True, color=col, align=PP_ALIGN.CENTER)
        txt(s, lbl, lpos, ty+1.1, 3.9, 0.8, size=13, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    txt(s, "* Sources: BFF Reports, Statista 2024, Sports Research Bangladesh",
        0.5, 6.85, 12.33, 0.3, size=8.5, italic=True, color=MID_GRAY, align=PP_ALIGN.CENTER)
    footer(s)


def s04_why_euro(prs):
    """Slide 4 — Why Youth Prefer European Football"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, ACCENT_BLUE); badge(s, "WHY EUROPEAN FOOTBALL?", c=ACCENT_BLUE); slide_num(s, 4)
    heading(s, "5 Reasons Urban Youth Prefer European Football", size=25, color=ACCENT_BLUE)
    hline(s, c=ACCENT_BLUE)

    # RIGHT SIDE: International stars image — LARGE and visible
    p1 = pic(s, IMG_INTL, 7.9, 1.62, 5.1, 5.5)
    if p1:
        rect(s, 7.9, 1.62, 5.1, 5.5, border=ACCENT_BLUE, bw=2.0)

    reasons = [
        ("1", "Superstar Power",        "Ronaldo, Messi, Mbappe — global icons with 500M+ fans", GOLD),
        ("2", "World-Class Broadcast",  "4K, VR, 30+ cameras, official apps, DAZN, Prime Video",  ACCENT_BLUE),
        ("3", "Electric Atmosphere",    "80,000-seat packed stadiums, ultras, flares, chants",      GREEN),
        ("4", "Digital & Social Media", "Viral clips, TikTok, YouTube highlights available 24/7",  TEAL),
        ("5", "Global Brand Identity",  "Jerseys, FIFA game culture, lifestyle — a global movement", ORANGE),
    ]
    for i, (num, title, desc, col) in enumerate(reasons):
        ty = 1.7 + i * 1.05
        grad(s, 0.4, ty, 7.2, 0.9, DARK_CARD, RGBColor(0x0A,0x12,0x26), ang=0)
        rect(s, 0.4, ty, 0.5, 0.9, fill=col)
        txt(s, num, 0.4, ty, 0.5, 0.9, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(s, title, 1.02, ty+0.05, 3.3, 0.42, size=14.5, bold=True, color=col)
        txt(s, desc,  1.02, ty+0.48, 6.2, 0.4,  size=12,   color=LIGHT_GRAY)

    footer(s)


def s05_bd_players(prs):
    """Slide 5 — Bangladesh Players (IMAGE PROMINENT)"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, GREEN); badge(s, "BANGLADESHI PLAYERS", c=DARK_GREEN); slide_num(s, 5)
    heading(s, "Bangladesh Has Talented Players — Needs More Exposure!", size=23, color=GREEN)
    hline(s, c=GREEN)

    # ── LARGE player image on LEFT ──
    p1 = pic(s, IMG_BD, 0.3, 1.62, 5.5, 5.55)
    if p1:
        rect(s, 0.3, 1.62, 5.5, 5.55, border=GREEN, bw=2.5)
        txt(s, "Bangladesh National Football Team", 0.3, 7.02, 5.5, 0.35,
            size=9.5, color=MID_GRAY, align=PP_ALIGN.CENTER, italic=True)

    # ── Player info cards on RIGHT ──
    players = [
        ("Hamza Choudhury",  "British-Bangladeshi midfielder, Leicester City\nFirst Bangladeshi-origin Premier League player",  GOLD),
        ("Jamal Bhuyan",     "Bangladesh captain, played in Denmark & Spain\nMost experienced player in national team",         GREEN),
        ("Sheikh Morsalin",  "Top scorer of Bangladesh national team\nExplosive winger with clinical finishing",                 TEAL),
        ("Biplu Ahmed",      "BFF Player of the Year multiple times\nKey winger for Abahani Limited Dhaka",                     ACCENT_BLUE),
        ("Rakib Hossain",    "Rising star — next big prospect of BPL\nKnown for speed and creativity on the pitch",             ORANGE),
    ]
    for i, (name, desc, col) in enumerate(players):
        ty = 1.68 + i * 1.08
        grad(s, 6.1, ty, 6.9, 0.95, DARK_CARD, RGBColor(0x04,0x12,0x08), ang=0)
        rect(s, 6.1, ty, 6.9, 0.95, border=col, bw=1.3)
        rect(s, 6.1, ty, 0.08, 0.95, fill=col)
        txt(s, name, 6.3, ty+0.04, 6.5, 0.4, size=14, bold=True, color=col)
        txt(s, desc, 6.3, ty+0.48, 6.5, 0.45, size=11, color=LIGHT_GRAY)

    footer(s)


def s06_intl_stars(prs):
    """Slide 6 — International Stars (IMAGE PROMINENT)"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, GOLD); badge(s, "INTERNATIONAL STARS", c=RGBColor(0x99,0x66,0x00)); slide_num(s, 6)
    heading(s, "Why Global Stars Drive Massive Fan Engagement", size=25, color=GOLD)
    hline(s, c=GOLD)

    # ── LARGE international stars image on RIGHT ──
    p1 = pic(s, IMG_EURO, 7.5, 1.62, 5.5, 5.55)
    if p1:
        rect(s, 7.5, 1.62, 5.5, 5.55, border=GOLD, bw=2.5)
        txt(s, "World Football Superstars", 7.5, 7.02, 5.5, 0.35,
            size=9.5, color=MID_GRAY, align=PP_ALIGN.CENTER, italic=True)

    # ── Star info on LEFT ──
    stars = [
        ("Cristiano Ronaldo", "635M+ Instagram followers\nBrand value: $1.5 Billion USD",   GOLD),
        ("Lionel Messi",      "500M+ followers | Argentina World Cup 2022\nInter Miami — still breaking records", ORANGE),
        ("Kylian Mbappe",     "Real Madrid's golden boy\nFrance's captain & fastest striker in history", ACCENT_BLUE),
        ("Erling Haaland",    "52 goals in a single EPL season\nMan City's most watched striker globally", GREEN),
    ]
    for i, (name, desc, col) in enumerate(stars):
        ty = 1.72 + i * 1.35
        grad(s, 0.4, ty, 6.8, 1.2, DARK_CARD, RGBColor(0x0C,0x10,0x04), ang=0)
        rect(s, 0.4, ty, 6.8, 1.2, border=col, bw=1.5)
        txt(s, "  " + name, 0.55, ty+0.07, 6.4, 0.45, size=16, bold=True, color=col)
        txt(s, "  " + desc, 0.55, ty+0.55, 6.4, 0.6,  size=12.5, color=LIGHT_GRAY)

    txt(s, "BPL needs this kind of star culture to attract youth viewers!",
        0.4, 7.05, 12.33, 0.35, size=12, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    footer(s)


def s07_stadium(prs):
    """Slide 7 — Stadium Experience (IMAGE PROMINENT)"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, ORANGE); badge(s, "STADIUM EXPERIENCE", c=ORANGE); slide_num(s, 7)
    heading(s, "BPL Stadium Experience — Far Behind European Standards", size=22, color=ORANGE)
    hline(s, c=ORANGE)

    # ── LARGE stadium image on LEFT ──
    p1 = pic(s, IMG_STADIUM, 0.3, 1.62, 6.2, 5.55)
    if p1:
        rect(s, 0.3, 1.62, 6.2, 5.55, border=ORANGE, bw=2.5)
        txt(s, "Bangladesh Football Stadium", 0.3, 7.02, 6.2, 0.35,
            size=9.5, color=MID_GRAY, align=PP_ALIGN.CENTER, italic=True)

    # ── Issues on RIGHT ──
    issues = [
        ("Poor Infrastructure", "Old, broken seats, no modern amenities or facilities"),
        ("Security Concerns",   "Poor crowd control, scares families & female fans away"),
        ("No Entertainment",    "No food courts, fan zones, or pre-match events at all"),
        ("Transport Problems",  "Lack of easy transport links to stadiums across cities"),
        ("No Covered Seating",  "Fans exposed to rain and heat — no shelter at venues"),
        ("Poor Lighting",       "Outdated floodlights fail modern broadcast standards"),
    ]
    for i, (title, desc) in enumerate(issues):
        ty = 1.68 + i * 0.92
        grad(s, 6.8, ty, 6.2, 0.82, DARK_CARD, RGBColor(0x14,0x08,0x03), ang=0)
        rect(s, 6.8, ty, 6.2, 0.82, border=ORANGE, bw=1.2)
        txt(s, title, 7.0, ty+0.05, 5.8, 0.36, size=13.5, bold=True, color=ORANGE)
        txt(s, desc,  7.0, ty+0.43, 5.8, 0.35, size=11,   color=LIGHT_GRAY)

    footer(s)


def s08_bpl_struggles(prs):
    """Slide 8 — BPL Core Struggles"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, RED); badge(s, "BPL STRUGGLES", c=RED); slide_num(s, 8)
    heading(s, "Why Bangladesh Premier League Struggles", size=27, color=RED)
    hline(s, c=RED)

    issues = [
        ("Poor Marketing",          "No promotions, no social media campaigns, no hype events",      RED),
        ("Weak Broadcasting",       "Inconsistent TV schedule, no streaming platform or official app", ORANGE),
        ("No Star Players",         "No foreign marquee signings, local players lack global profile",  GOLD),
        ("Financial Instability",   "Clubs face salary delays, lack sponsors, no investor confidence", TEAL),
        ("Digital Absence",         "Zero short-form content, no YouTube highlights or TikTok reels", ACCENT_BLUE),
        ("No Fan Culture",          "No jersey culture, no songs, no supporter groups for urban youth", GREEN),
    ]
    for i, (title, desc, col) in enumerate(issues):
        ci = i % 2; ri = i // 2
        lx = 0.4 + ci * 6.5; ty = 1.72 + ri * 1.8
        grad(s, lx, ty, 6.1, 1.6, RGBColor(0x14,0x05,0x05), DARK_CARD, ang=135)
        rect(s, lx, ty, 6.1, 1.6, border=col, bw=1.5)
        rect(s, lx, ty, 0.6, 1.6, fill=col)
        txt(s, str(i+1), lx, ty, 0.6, 1.6, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(s, title, lx+0.72, ty+0.12, 5.1, 0.45, size=15, bold=True, color=col)
        txt(s, desc,  lx+0.72, ty+0.65, 5.1, 0.85, size=11.5, color=LIGHT_GRAY)

    footer(s)


def s09_marketing(prs):
    """Slide 9 — Marketing Failure"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, ORANGE); badge(s, "MARKETING FAILURE", c=ORANGE); slide_num(s, 9)
    heading(s, "BPL Marketing vs. European Football Marketing", size=26, color=ORANGE)
    hline(s, c=ORANGE)

    # BPL column
    grad(s, 0.4, 1.72, 6.1, 5.1, DARK_CARD, RGBColor(0x18,0x07,0x07), ang=90)
    rect(s, 0.4, 1.72, 6.1, 5.1, border=RED, bw=2)
    txt(s, "  BPL Marketing", 0.5, 1.78, 5.9, 0.52, size=16, bold=True, color=RED)
    rect(s, 0.55, 2.32, 5.8, 0.04, fill=RGBColor(0x33,0x0A,0x0A))
    bpl_pts = ["No social media strategy at all",
               "Zero youth ambassador programs",
               "No jersey or merchandise culture",
               "Matches not promoted on mainstream TV",
               "No fan zones or pre-match events",
               "Club websites outdated or non-existent"]
    for i, p in enumerate(bpl_pts):
        txt(s, "  [x]  " + p, 0.55, 2.45 + i*0.72, 5.7, 0.62, size=13, color=LIGHT_GRAY)

    # European column
    grad(s, 6.83, 1.72, 6.1, 5.1, DARK_CARD, RGBColor(0x03,0x18,0x0A), ang=90)
    rect(s, 6.83, 1.72, 6.1, 5.1, border=GREEN, bw=2)
    txt(s, "  European Football Marketing", 6.93, 1.78, 5.9, 0.52, size=16, bold=True, color=GREEN)
    rect(s, 6.98, 2.32, 5.8, 0.04, fill=RGBColor(0x0A,0x33,0x12))
    eu_pts = ["$100M+ marketing spend per top club",
              "Global brand ambassadors, fan events",
              "Official jerseys sold in 180+ countries",
              "Prime time dedicated broadcast channels",
              "Fan experience zones inside every stadium",
              "100M+ YouTube and TikTok followers"]
    for i, p in enumerate(eu_pts):
        txt(s, "  [+]  " + p, 6.98, 2.45 + i*0.72, 5.7, 0.62, size=13, color=LIGHT_GRAY)

    # VS divider
    grad(s, 6.43, 4.1, 0.5, 1.1, RED, GREEN, ang=90)
    txt(s, "VS", 6.43, 4.1, 0.5, 1.1, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    footer(s)


def s10_youth(prs):
    """Slide 10 — Youth Disconnect"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, ACCENT_BLUE); badge(s, "YOUTH DISCONNECT", c=ACCENT_BLUE); slide_num(s, 10)
    heading(s, "Why Urban Youth Are Disconnected from BPL", size=26, color=ACCENT_BLUE)
    hline(s, c=ACCENT_BLUE)

    # Quote box
    grad(s, 0.5, 1.72, 12.33, 1.1, RGBColor(0x07,0x14,0x2A), DARK_CARD, ang=0)
    rect(s, 0.5, 1.72, 0.1, 1.1, fill=ACCENT_BLUE)
    txt(s, '"I watch UCL highlights the moment they drop — '
         "BPL? I don't even know when the season starts or which team is winning.\"",
        0.75, 1.78, 11.8, 0.95, size=14, italic=True, color=LIGHT_GRAY)
    txt(s, "— Urban Bangladeshi Youth (Survey Response, 2024)",
        0.75, 2.68, 11.8, 0.28, size=10.5, italic=True, color=MID_GRAY)

    reasons = [
        ("No Social Media Content", "BPL clubs barely post; no reels, highlights, memes or live updates"),
        ("Gaming & FIFA Culture",   "Youth bond with EPL/LaLiga clubs through FIFA game career mode"),
        ("No Fan Identity",         "BPL clubs have no cool jersey, badge, or youth-relatable culture"),
        ("Content Unavailability",  "No YouTube channel, Shorts, TikTok, or Instagram presence for BPL"),
        ("No Role Model Players",   "Youth can't name a BPL star the way they name Messi or Ronaldo"),
    ]
    for i, (title, desc) in enumerate(reasons):
        ty = 3.05 + i * 0.85
        grad(s, 0.4, ty, 12.5, 0.74, DARK_CARD, RGBColor(0x05,0x0F,0x22), ang=0)
        rect(s, 0.4, ty, 12.5, 0.74, border=ACCENT_BLUE, bw=1)
        rect(s, 0.4, ty, 0.55, 0.74, fill=ACCENT_BLUE)
        txt(s, str(i+1), 0.4, ty, 0.55, 0.74, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(s, title, 1.1, ty+0.05, 3.5, 0.36, size=13.5, bold=True, color=ACCENT_BLUE)
        txt(s, desc,  1.1, ty+0.42, 11.3, 0.3, size=11.5, color=LIGHT_GRAY)

    footer(s)


def s11_broadcasting(prs):
    """Slide 11 — Broadcasting Problem"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, TEAL); badge(s, "BROADCASTING PROBLEM", c=RGBColor(0x00,0x88,0x77)); slide_num(s, 11)
    heading(s, "The Broadcasting Gap Between BPL and European Football", size=24, color=TEAL)
    hline(s, c=TEAL)

    points = [
        ("No Streaming App",       "BPL has no official OTT platform; Europeans have DAZN, Peacock, Prime Video, Disney+"),
        ("Poor Commentary Quality", "Local broadcasting has no professional commentary; no English/multilingual options"),
        ("Camera & Broadcast Tech", "Single-camera coverage vs. UEFA's 30+ cameras, 4K resolution, VR & drone coverage"),
        ("Irregular Match Timings", "BPL match times change last minute — fans can't plan; no fixed reliable schedule"),
        ("No International Feed",  "BPL matches not available outside Bangladesh; zero global audience being built"),
    ]
    for i, (title, desc) in enumerate(points):
        ty = 1.75 + i * 1.08
        grad(s, 0.4, ty, 12.5, 0.93, DARK_CARD, RGBColor(0x02,0x14,0x14), ang=0)
        rect(s, 0.4, ty, 12.5, 0.93, border=TEAL, bw=1.2)
        rect(s, 0.4, ty, 0.65, 0.93, fill=TEAL)
        txt(s, str(i+1), 0.4, ty, 0.65, 0.93, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(s, title, 1.22, ty+0.06, 3.5, 0.42, size=14.5, bold=True, color=TEAL)
        txt(s, desc,  1.22, ty+0.52, 11.3, 0.38, size=12, color=LIGHT_GRAY)

    footer(s)


def s12_recommendations(prs):
    """Slide 12 — Recommendations"""
    s = blank(prs); fill_bg(s, DARK_BG)
    top_bar(s, GREEN); badge(s, "RECOMMENDATIONS", c=DARK_GREEN); slide_num(s, 12)
    heading(s, "6 Key Recommendations to Fix BPL", size=28, color=GREEN)
    hline(s, c=GREEN)

    recs = [
        ("1", "Digital Revolution",      "Launch OTT app, YouTube, TikTok & Instagram for highlights",    TEAL),
        ("2", "Star Player Signings",    "Bring 2-3 marquee international players to create excitement",   GOLD),
        ("3", "Youth Engagement",        "Free student tickets, campus fan clubs, fantasy BPL leagues",    GREEN),
        ("4", "Upgrade Stadiums",        "Modern covered stands, fan zones, food courts, LED screens",     ORANGE),
        ("5", "Marketing Campaigns",     "Billboard, TV & social ads targeting 18-30 urban demographics",  ACCENT_BLUE),
        ("6", "Pro Broadcasting Deals",  "4K international broadcast, multi-language expert commentary",   RED),
    ]
    for i, (num, title, desc, col) in enumerate(recs):
        ci = i % 2; ri = i // 2
        lx = 0.4 + ci * 6.5; ty = 1.72 + ri * 1.8
        grad(s, lx, ty, 6.1, 1.6, DARK_CARD, RGBColor(0x04,0x10,0x04), ang=90)
        rect(s, lx, ty, 6.1, 1.6, border=col, bw=1.8)
        rect(s, lx, ty, 0.62, 1.6, fill=col)
        txt(s, num, lx, ty, 0.62, 1.6, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txt(s, title, lx+0.74, ty+0.12, 5.1, 0.48, size=15, bold=True, color=col)
        txt(s, desc,  lx+0.74, ty+0.68, 5.1, 0.82, size=11.5, color=LIGHT_GRAY)

    footer(s)


def s13_vision(prs):
    """Slide 13 — Vision"""
    s = blank(prs); fill_bg(s, DARK_BG)
    grad(s, 0, 0, 13.33, 7.5, RGBColor(0x02,0x14,0x07), DARK_BG, ang=90)
    top_bar(s, GREEN); badge(s, "THE VISION", c=DARK_GREEN); slide_num(s, 13)
    heading(s, "A Future Where BPL Becomes the Pride of Bangladesh", size=25, color=GREEN)
    hline(s, c=GOLD)

    overlay(s, 0.5, 1.72, 12.33, 5.1, DARK_CARD, opacity=0.72)
    rect(s,  0.5, 1.72, 12.33, 5.1, border=GREEN, bw=2)

    vision = [
        "  30,000+ average stadium attendance per BPL match by 2027",
        "  1 Million+ BPL social media followers across all platforms by 2026",
        "  International broadcast deal with Al Jazeera, Willow TV & DAZN",
        "  Bangladeshi players competing in AFC Asian Champions League",
        "  BPL jersey culture adopted by urban youth across Dhaka & Chittagong",
        "  Major sponsorships: Grameenphone, Bashundhara, bKash, Walton Group",
    ]
    for i, v in enumerate(vision):
        ty = 1.98 + i * 0.78
        rect(s, 0.7, ty+0.22, 0.18, 0.28, fill=GREEN)
        txt(s, v, 1.1, ty, 11.4, 0.68, size=15, color=LIGHT_GRAY)

    txt(s, '"Football is not just a sport — it is the heartbeat of a nation."',
        0.5, 6.97, 12.33, 0.36, size=13, italic=True, color=GOLD, align=PP_ALIGN.CENTER)

    footer(s)


def s14_conclusion(prs):
    """Slide 14 — Conclusion"""
    s = blank(prs); fill_bg(s, DARK_BG)

    # Background image faded
    p1 = pic(s, IMG_BG, 0, 0, 13.33, 7.5)
    overlay(s, 0, 0, 13.33, 7.5, DARK_BG, opacity=0.82)

    grad(s, 0, 0, 13.33, 0.12, RED, GREEN, ang=0)
    badge(s, "CONCLUSION", c=DARK_GREEN); slide_num(s, 14)
    heading(s, "Conclusion", size=34, color=WHITE)
    hline(s, c=GREEN, l=4.5, w=4.33)

    overlay(s, 1.0, 1.72, 11.33, 3.7, DARK_CARD, opacity=0.90)
    rect(s,   1.0, 1.72, 11.33, 3.7, border=GREEN, bw=2.0)

    txt(s, "Bangladesh Premier League's failure to attract urban youth",
        1.3, 1.88, 10.73, 0.65, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(s, "is not a football problem — it is a MANAGEMENT & MARKETING problem.",
        1.3, 2.48, 10.73, 0.65, size=18, bold=True, color=RED, align=PP_ALIGN.CENTER)

    rect(s, 1.3, 3.18, 10.73, 0.05, fill=RGBColor(0x22,0x33,0x22))

    txt(s, "With proper digital marketing, quality broadcasting,\n"
         "modern stadiums, and a star-player culture —\n"
         "BPL CAN become the pride of Bangladeshi youth.",
        1.8, 3.3, 9.73, 1.1, size=15.5, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

    grad(s, 1.5, 5.55, 10.33, 0.78, DARK_GREEN, RGBColor(0x00,0x44,0x16), ang=0)
    txt(s, "BPL — From Local League to National Pride",
        1.5, 5.55, 10.33, 0.78, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    txt(s, "Thank You for Your Attention!",
        0.5, 6.42, 12.33, 0.48, size=16, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

    footer(s)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    prs = new_prs()
    steps = [
        (s01_title,          " 1/14  Title Slide"),
        (s02_reality,        " 2/14  The Reality"),
        (s03_stats,          " 3/14  Key Statistics"),
        (s04_why_euro,       " 4/14  Why Euro Football + Image"),
        (s05_bd_players,     " 5/14  BD Players + IMAGE"),
        (s06_intl_stars,     " 6/14  International Stars + IMAGE"),
        (s07_stadium,        " 7/14  Stadium Experience + IMAGE"),
        (s08_bpl_struggles,  " 8/14  BPL Struggles"),
        (s09_marketing,      " 9/14  Marketing Failure"),
        (s10_youth,          "10/14  Youth Disconnect"),
        (s11_broadcasting,   "11/14  Broadcasting Problem"),
        (s12_recommendations,"12/14  Recommendations"),
        (s13_vision,         "13/14  Vision"),
        (s14_conclusion,     "14/14  Conclusion"),
    ]
    for fn, label in steps:
        print(f"[{label}] ...")
        fn(prs)

    prs.save(OUT_PATH)
    print(f"\n[DONE] Saved: {OUT_PATH}")
    print(f"       Slides: {len(prs.slides)}")

if __name__ == "__main__":
    main()
