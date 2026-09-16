"""
presentation.py - Builds richly designed PowerPoint decks for JARVIS.

This is a big step up from actions.py's plain `_write_pptx`, which just
dumps title + bullets onto PowerPoint's default white template. Here we
get a dark "JARVIS" theme (teal glow on near-black), several slide
layouts, native charts, web images, and slide transitions.

Input is JSON (see build_presentation) rather than the loose
"---"-separated text format, because the richer layouts need structure
that free text can't express unambiguously.

Two things worth knowing about how this is implemented:

1. TRANSITIONS: python-pptx has no API for slide transitions or entrance
   animations - it's a long-standing gap in the library. _set_transition()
   works around it by injecting the raw <mc:AlternateContent> transition
   XML directly into each slide's XML tree. This works in PowerPoint and
   is the standard workaround, but it IS a hack against an unsupported
   surface: if a future python-pptx or PowerPoint release changes how
   that XML is parsed, transitions may silently stop applying. The deck
   itself stays valid either way - you'd just lose the transitions, not
   the content.

   Entrance/exit animations on individual elements (text flying in, etc.)
   are deliberately NOT attempted. Their XML (<p:timing>) is far more
   intricate and interdependent than transitions, and malformed timing
   trees make PowerPoint declare the whole file corrupt rather than
   degrading gracefully. Transitions give most of the "animated" feel at
   a tiny fraction of the risk.

2. IMAGES: fetch_image() pulls from Pexels, which needs a free API key
   in config.json as "pexels_api_key". Without one, image slides fall
   back to a styled placeholder panel rather than failing. Pexels content
   is free to use commercially without attribution required, but do check
   their current license if these decks go somewhere public.
"""
import io
import json
import os
import re
import urllib.parse
import urllib.request

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- theme

# Dark JARVIS palette: teal glow on near-black, matching the app's HUD.
BG = RGBColor(0x0A, 0x0F, 0x14)        # near-black, slight blue cast
BG_PANEL = RGBColor(0x11, 0x1A, 0x22)  # raised card surface
GLOW = RGBColor(0x2C, 0xF5, 0xD4)      # primary teal accent
GLOW_DIM = RGBColor(0x1B, 0x8F, 0x7E)  # muted teal for secondary marks
TEXT = RGBColor(0xEA, 0xFF, 0xFA)      # near-white body text
TEXT_MUTED = RGBColor(0x8A, 0xA5, 0xA0)  # captions, footnotes

# Chart series colors - teal-forward, enough separation to read at a glance.
SERIES_COLORS = ["2CF5D4", "1B8F7E", "7BDCE8", "4A9FD4", "A88BE8", "E8C34A"]

# Safe fonts per the deck design guidance: these ship with Office and
# render at true width, so text that fits here fits on the user's machine.
FONT_HEAD = "Cambria"    # serif headers, for some personality
FONT_BODY = "Calibri"    # sans body

SLIDE_W = Inches(13.333)  # 16:9 widescreen
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)


def _solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def _textbox(slide, x, y, w, h, text, size, color=TEXT, bold=False,
             font=FONT_BODY, align=PP_ALIGN.LEFT, italic=False,
             anchor=MSO_ANCHOR.TOP, line_spacing=None):
    """Adds a text box with our theme defaults. Returns the text frame."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0

    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = font
    return tf


def _bullets(slide, x, y, w, h, items, size=16, color=TEXT):
    """Bulleted list with a teal glyph and comfortable paragraph spacing."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        # Teal bullet glyph as its own run so it can be colored separately
        # from the text - simpler and more predictable than buChar XML.
        marker = p.add_run()
        marker.text = "\u25B8  "
        marker.font.size = Pt(size)
        marker.font.color.rgb = GLOW
        marker.font.name = FONT_BODY

        run = p.add_run()
        run.text = str(item)
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.name = FONT_BODY
    return tf


def _bg(slide, color=BG):
    """Paints the whole slide background (no native 'slide fill' in the
    default template, so we lay a rectangle behind everything)."""
    from pptx.enum.shapes import MSO_SHAPE
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    _solid(rect, color)
    # Push it to the back of the z-order.
    spTree = slide.shapes._spTree
    spTree.remove(rect._element)
    spTree.insert(2, rect._element)
    return rect


def _panel(slide, x, y, w, h, color=BG_PANEL, radius=True):
    """A raised card surface - our repeated visual motif."""
    from pptx.enum.shapes import MSO_SHAPE
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    card = slide.shapes.add_shape(shape_type, x, y, w, h)
    _solid(card, color)
    return card


def _title(slide, text, subtitle=None):
    """Standard content-slide heading. No underline accent bar - those
    read as generic AI filler; whitespace does the separating instead."""
    _textbox(slide, MARGIN, Inches(0.5), SLIDE_W - 2 * MARGIN, Inches(0.9),
             text, 34, TEXT, bold=True, font=FONT_HEAD)
    if subtitle:
        _textbox(slide, MARGIN, Inches(1.28), SLIDE_W - 2 * MARGIN, Inches(0.4),
                 subtitle, 14, TEXT_MUTED, italic=True)


# ------------------------------------------------------- transitions

# Transition XML injected per-slide. python-pptx exposes no API for this
# (see module docstring); we build the element by hand and append it to
# the slide's <p:sld>. Wrapped in mc:AlternateContent with a p14 fallback
# so older PowerPoint versions ignore it gracefully instead of erroring.
_TRANSITION_XML = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">'
    '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
    'spd="slow" p14:dur="{dur}">{inner}</p:transition>'
    '</mc:Choice>'
    '<mc:Fallback>'
    '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow">'
    '{fallback}</p:transition>'
    '</mc:Fallback>'
    '</mc:AlternateContent>'
)

_TRANSITIONS = {
    "fade": ('<p:fade/>', '<p:fade/>'),
    "push": ('<p:push dir="u"/>', '<p:push dir="u"/>'),
    "wipe": ('<p:wipe dir="d"/>', '<p:wipe dir="d"/>'),
    "morph": ('<p14:morph option="byObject"/>', '<p:fade/>'),
    "split": ('<p:split orient="horz" dir="out"/>', '<p:split orient="horz" dir="out"/>'),
}


def _set_transition(slide, kind="fade", duration_ms=700):
    """Attaches a slide transition. Silently no-ops on unknown kinds."""
    if kind not in _TRANSITIONS:
        return
    from lxml import etree
    inner, fallback = _TRANSITIONS[kind]
    xml = _TRANSITION_XML.format(dur=duration_ms, inner=inner, fallback=fallback)
    try:
        frag = etree.fromstring(xml)
        slide._element.append(frag)
    except Exception as e:
        # A failed transition must never cost the user their deck.
        print(f"[presentation.py] couldn't apply '{kind}' transition: {e!r}")


# ------------------------------------------------------------- images

def fetch_image(query, api_key, orientation="landscape"):
    """
    Downloads a stock photo matching `query` from Pexels. Returns a
    BytesIO of image bytes, or None on any failure (no key, no network,
    no results) - callers fall back to a placeholder panel so a missing
    image never breaks deck generation.
    """
    if not api_key:
        return None
    url = ("https://api.pexels.com/v1/search?"
           + urllib.parse.urlencode({
               "query": query, "per_page": 1, "orientation": orientation,
           }))
    try:
        req = urllib.request.Request(url, headers={"Authorization": api_key})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        photos = data.get("photos") or []
        if not photos:
            print(f"[presentation.py] no Pexels results for {query!r}")
            return None
        src = photos[0]["src"].get("large") or photos[0]["src"].get("original")
        with urllib.request.urlopen(src, timeout=20) as resp:
            return io.BytesIO(resp.read())
    except Exception as e:
        print(f"[presentation.py] image fetch failed for {query!r}: {e!r}")
        return None


def _image_or_placeholder(slide, img, x, y, w, h, label="", radius=True):
    """Places a fetched image cropped-to-fill, or a themed placeholder.
    `radius` off for full-bleed placements, where rounded corners would
    leave odd gaps against the slide edge."""
    if img is not None:
        try:
            pic = slide.shapes.add_picture(img, x, y, width=w, height=h)
            return pic
        except Exception as e:
            print(f"[presentation.py] couldn't place image: {e!r}")
    card = _panel(slide, x, y, w, h, BG_PANEL, radius=radius)
    _textbox(slide, x, y + h / 2 - Inches(0.25), w, Inches(0.5),
             label or "figure", 13, TEXT_MUTED, align=PP_ALIGN.CENTER)
    return card


# ------------------------------------------------------------ layouts

def _blank(prs):
    """Layout 6 is the blank layout - we draw everything ourselves."""
    return prs.slides.add_slide(prs.slide_layouts[6])


def _slide_title(prs, spec, api_key):
    slide = _blank(prs)
    _bg(slide)

    img = fetch_image(spec["image"], api_key) if spec.get("image") else None
    if img is not None:
        # Half-bleed image on the right, content overlaid left.
        _image_or_placeholder(slide, img, SLIDE_W / 2, 0, SLIDE_W / 2, SLIDE_H,
                              spec.get("image", ""), radius=False)
        text_w = SLIDE_W / 2 - MARGIN - Inches(0.4)
    else:
        text_w = SLIDE_W - 2 * MARGIN

    _textbox(slide, MARGIN, Inches(2.4), text_w, Inches(1.8),
             spec.get("title", ""), 52, TEXT, bold=True, font=FONT_HEAD,
             line_spacing=1.0)
    if spec.get("subtitle"):
        _textbox(slide, MARGIN, Inches(4.35), text_w, Inches(0.9),
                 spec["subtitle"], 19, GLOW, font=FONT_BODY)
    if spec.get("footer"):
        _textbox(slide, MARGIN, SLIDE_H - Inches(1.0), text_w, Inches(0.4),
                 spec["footer"], 12, TEXT_MUTED)
    return slide


def _slide_bullets(prs, spec, api_key):
    slide = _blank(prs)
    _bg(slide)
    _title(slide, spec.get("title", ""), spec.get("subtitle"))

    items = spec.get("bullets", [])
    if spec.get("image"):
        img = fetch_image(spec["image"], api_key)
        col_w = (SLIDE_W - 2 * MARGIN - Inches(0.5)) / 2
        _bullets(slide, MARGIN, Inches(2.0), col_w, Inches(4.4), items)
        _image_or_placeholder(slide, img, MARGIN + col_w + Inches(0.5),
                              Inches(2.0), col_w, Inches(4.0), spec["image"])
    else:
        _bullets(slide, MARGIN, Inches(2.0), SLIDE_W - 2 * MARGIN, Inches(4.4), items)
    return slide


def _slide_stats(prs, spec, api_key):
    """Big-number callouts in a row of cards."""
    slide = _blank(prs)
    _bg(slide)
    _title(slide, spec.get("title", ""), spec.get("subtitle"))

    stats = spec.get("stats", [])[:4]
    if not stats:
        return slide
    gap = Inches(0.35)
    total_w = SLIDE_W - 2 * MARGIN
    card_w = (total_w - gap * (len(stats) - 1)) / len(stats)
    top = Inches(2.3)
    card_h = Inches(2.6)

    for i, st in enumerate(stats):
        x = MARGIN + i * (card_w + gap)
        _panel(slide, x, top, card_w, card_h)
        _textbox(slide, x, top + Inches(0.45), card_w, Inches(1.1),
                 str(st.get("value", "")), 46, GLOW, bold=True,
                 font=FONT_HEAD, align=PP_ALIGN.CENTER)
        _textbox(slide, x + Inches(0.2), top + Inches(1.6),
                 card_w - Inches(0.4), Inches(0.8),
                 str(st.get("label", "")), 13, TEXT_MUTED,
                 align=PP_ALIGN.CENTER)
    return slide


def _slide_compare(prs, spec, api_key):
    """Two-column comparison (before/after, pros/cons, option A/B)."""
    slide = _blank(prs)
    _bg(slide)
    _title(slide, spec.get("title", ""), spec.get("subtitle"))

    left = spec.get("left", {})
    right = spec.get("right", {})
    gap = Inches(0.5)
    col_w = (SLIDE_W - 2 * MARGIN - gap) / 2
    top = Inches(2.1)
    col_h = Inches(4.5)

    for i, col in enumerate((left, right)):
        x = MARGIN + i * (col_w + gap)
        _panel(slide, x, top, col_w, col_h)
        _textbox(slide, x + Inches(0.35), top + Inches(0.3),
                 col_w - Inches(0.7), Inches(0.5),
                 col.get("heading", ""), 21, GLOW, bold=True, font=FONT_HEAD)
        _bullets(slide, x + Inches(0.35), top + Inches(1.0),
                 col_w - Inches(0.7), col_h - Inches(1.3),
                 col.get("points", []), size=14)
    return slide


def _slide_steps(prs, spec, api_key):
    """Numbered process flow - rows of circle + heading + description."""
    slide = _blank(prs)
    _bg(slide)
    _title(slide, spec.get("title", ""), spec.get("subtitle"))

    from pptx.enum.shapes import MSO_SHAPE
    steps = spec.get("steps", [])[:5]
    top = Inches(2.0)
    row_h = Inches(0.95)

    for i, step in enumerate(steps):
        y = top + i * row_h
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, MARGIN, y,
                                      Inches(0.6), Inches(0.6))
        _solid(circ, GLOW_DIM)
        _textbox(slide, MARGIN, y + Inches(0.13), Inches(0.6), Inches(0.35),
                 str(i + 1), 16, TEXT, bold=True, align=PP_ALIGN.CENTER)

        tx = MARGIN + Inches(0.9)
        tw = SLIDE_W - tx - MARGIN
        _textbox(slide, tx, y, tw, Inches(0.35),
                 step.get("heading", ""), 17, TEXT, bold=True, font=FONT_HEAD)
        if step.get("detail"):
            _textbox(slide, tx, y + Inches(0.36), tw, Inches(0.5),
                     step["detail"], 13, TEXT_MUTED)
    return slide


def _slide_chart(prs, spec, api_key):
    """Native PowerPoint chart - stays editable/animatable in PowerPoint,
    unlike a rendered image."""
    slide = _blank(prs)
    _bg(slide)
    _title(slide, spec.get("title", ""), spec.get("subtitle"))

    chart_spec = spec.get("chart", {})
    kind = chart_spec.get("type", "bar").lower()
    cats = chart_spec.get("categories", [])
    series = chart_spec.get("series", [])
    if not cats or not series:
        _textbox(slide, MARGIN, Inches(3.0), SLIDE_W - 2 * MARGIN, Inches(0.6),
                 "(no chart data provided)", 14, TEXT_MUTED)
        return slide

    type_map = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
        "area": XL_CHART_TYPE.AREA,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
    }
    ct = type_map.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)

    data = CategoryChartData()
    data.categories = cats
    for s in series:
        data.add_series(s.get("name", "Series"), s.get("values", []))

    x, y = MARGIN, Inches(2.0)
    w, h = SLIDE_W - 2 * MARGIN, Inches(4.5)
    gframe = slide.shapes.add_chart(ct, x, y, w, h, data)
    chart = gframe.chart

    chart.has_title = False
    # Legend only earns its space with more than one series.
    chart.has_legend = len(series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(12)
        chart.legend.font.color.rgb = TEXT_MUTED

    try:
        for i, plot_series in enumerate(chart.series):
            color = RGBColor.from_string(SERIES_COLORS[i % len(SERIES_COLORS)])
            plot_series.format.fill.solid()
            plot_series.format.fill.fore_color.rgb = color
    except Exception as e:
        print(f"[presentation.py] couldn't color chart series: {e!r}")

    # Quiet the axis furniture so the data reads, not the frame.
    try:
        for axis in (chart.category_axis, chart.value_axis):
            axis.tick_labels.font.size = Pt(12)
            axis.tick_labels.font.color.rgb = TEXT_MUTED
            axis.format.line.color.rgb = GLOW_DIM
    except Exception:
        pass  # pie/doughnut charts have no category/value axes
    return slide


def _slide_image(prs, spec, api_key):
    """Full-bleed image with a caption band."""
    slide = _blank(prs)
    _bg(slide)
    img = fetch_image(spec["image"], api_key) if spec.get("image") else None
    _image_or_placeholder(slide, img, 0, 0, SLIDE_W, SLIDE_H - Inches(1.5),
                          spec.get("image", ""), radius=False)
    _textbox(slide, MARGIN, SLIDE_H - Inches(1.25), SLIDE_W - 2 * MARGIN,
             Inches(0.6), spec.get("title", ""), 28, TEXT, bold=True,
             font=FONT_HEAD)
    if spec.get("caption"):
        _textbox(slide, MARGIN, SLIDE_H - Inches(0.62),
                 SLIDE_W - 2 * MARGIN, Inches(0.4),
                 spec["caption"], 12, TEXT_MUTED)
    return slide


def _slide_quote(prs, spec, api_key):
    slide = _blank(prs)
    _bg(slide)
    _textbox(slide, Inches(1.4), Inches(2.2), SLIDE_W - Inches(2.8), Inches(2.6),
             '"' + spec.get("quote", "") + '"', 32, TEXT, italic=True,
             font=FONT_HEAD, align=PP_ALIGN.CENTER, line_spacing=1.2)
    if spec.get("attribution"):
        _textbox(slide, Inches(1.4), Inches(5.0), SLIDE_W - Inches(2.8), Inches(0.5),
                 "\u2014 " + spec["attribution"], 16, GLOW,
                 align=PP_ALIGN.CENTER)
    return slide


def _slide_section(prs, spec, api_key):
    """Divider slide - dark, big, minimal, resets the viewer's attention."""
    slide = _blank(prs)
    _bg(slide)
    if spec.get("eyebrow"):
        _textbox(slide, MARGIN, Inches(2.8), SLIDE_W - 2 * MARGIN, Inches(0.4),
                 spec["eyebrow"].upper(), 14, GLOW, bold=True)
    _textbox(slide, MARGIN, Inches(3.25), SLIDE_W - 2 * MARGIN, Inches(1.4),
             spec.get("title", ""), 44, TEXT, bold=True, font=FONT_HEAD)
    return slide


LAYOUTS = {
    "title": _slide_title,
    "bullets": _slide_bullets,
    "stats": _slide_stats,
    "compare": _slide_compare,
    "steps": _slide_steps,
    "chart": _slide_chart,
    "image": _slide_image,
    "quote": _slide_quote,
    "section": _slide_section,
}


# ------------------------------------------------------------- build

def build_presentation(path, spec, api_key=None):
    """
    Builds a themed .pptx at `path` from `spec` (a dict, usually parsed
    from JSON). Returns a dict describing what was built.

    spec = {
      "transition": "fade",          # optional, applied to every slide
      "slides": [ {...}, {...} ]
    }

    Each slide dict needs a "layout" key naming one of LAYOUTS, plus that
    layout's own fields (see the _slide_* functions above). Unknown
    layouts fall back to "bullets" so a bad layout name degrades to a
    readable slide rather than raising.
    """
    if isinstance(spec, str):
        spec = json.loads(spec)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    transition = spec.get("transition", "fade")
    slides = spec.get("slides", [])
    if not slides:
        raise ValueError("Presentation spec contains no slides.")

    built = []
    for i, s in enumerate(slides):
        layout_name = s.get("layout", "bullets")
        fn = LAYOUTS.get(layout_name)
        if fn is None:
            print(f"[presentation.py] unknown layout {layout_name!r} on slide "
                  f"{i + 1}, falling back to 'bullets'")
            fn = _slide_bullets
            layout_name = "bullets"
        slide = fn(prs, s, api_key)
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]
        _set_transition(slide, s.get("transition", transition))
        built.append(layout_name)

    prs.save(path)
    return {
        "ok": True,
        "path": path,
        "slide_count": len(built),
        "layouts_used": built,
        "images_enabled": bool(api_key),
    }
