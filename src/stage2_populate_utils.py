import io, os, copy
from PIL import Image
from lxml import etree
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# Exact template monochrome palette (from Slide 26 of Systems Engineering Business Plan)
COL_CHARCOAL = RGBColor(0x38, 0x38, 0x38)
COL_MUTED    = RGBColor(0x59, 0x59, 0x59)
COL_SILVER   = RGBColor(0xBA, 0xBA, 0xBA)
COL_PLATINUM = RGBColor(0xEA, 0xEA, 0xEA)
COL_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
COL_MIDGRAY  = RGBColor(0x88, 0x88, 0x88)
COL_SOFTBG   = RGBColor(0xF5, 0xF5, 0xF5)


def add_dark_badge(slide, x, y, w, h, text, font_size=7.6):
    """Crisp monochrome charcoal badge (#383838 fill + bold #FFFFFF text) for course/speaker tags."""
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.adjustments[0] = 0.22
    box.fill.solid()
    box.fill.fore_color.rgb = COL_CHARCOAL
    box.line.fill.background()
    _disable_shadow(box)

    tf = box.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = FONT_HEAD
    r.font.size = Pt(font_size)
    r.font.bold = True
    r.font.color.rgb = COL_WHITE
    return box

FONT_HEAD = "Schibsted Grotesk"
FONT_BODY = "Figtree"
FONT_CODE = "JetBrains Mono"


def _apply_rpr(rPr, size_pt=None, bold=None, color_hex=None, font_name=None):
    if size_pt is not None:
        rPr.set("sz", str(int(round(size_pt * 100))))
    if bold is not None:
        rPr.set("b", "1" if bold else "0")
    if color_hex is not None:
        for child in list(rPr):
            if child.tag in (f"{{{A_NS}}}solidFill", f"{{{A_NS}}}noFill", f"{{{A_NS}}}gradFill"):
                rPr.remove(child)
        sf = etree.Element(f"{{{A_NS}}}solidFill")
        srgb = etree.SubElement(sf, f"{{{A_NS}}}srgbClr")
        srgb.set("val", color_hex.lstrip("#").upper())
        rPr.insert(0, sf)
    if font_name is not None:
        for tag in ("latin", "cs", "ea"):
            el = rPr.find(f"{{{A_NS}}}{tag}")
            if el is None:
                el = etree.SubElement(rPr, f"{{{A_NS}}}{tag}")
            el.set("typeface", font_name)


def set_text(shape, paragraphs, size_pt=None, bold=None, color_hex=None, font_name=None, align=None, space_after_pt=None, line_spacing=None):
    """
    Replace text in a template shape while preserving original pPr & rPr XML styles.
    `paragraphs` can be:
      - str (split on '\n' into paragraphs)
      - list of str
      - list of list of (text, is_bold) or (text, is_bold, opt_color_hex) tuples
    """
    txBody = shape._element.find(f".//{{{P_NS}}}txBody")
    if txBody is None:
        txBody = shape._element.find(f".//{{{A_NS}}}txBody")
    if txBody is None:
        return

    existing_ps = txBody.findall(f"{{{A_NS}}}p")
    template_pPr = None
    template_rPr = None
    for p in existing_ps:
        if template_pPr is None:
            pPr = p.find(f"{{{A_NS}}}pPr")
            if pPr is not None:
                template_pPr = copy.deepcopy(pPr)
        for r in p.findall(f"{{{A_NS}}}r"):
            rPr = r.find(f"{{{A_NS}}}rPr")
            if rPr is not None:
                template_rPr = copy.deepcopy(rPr)
                break
        if template_rPr is not None:
            break

    if template_rPr is None:
        template_rPr = etree.Element(f"{{{A_NS}}}rPr")
        template_rPr.set("lang", "en-US")

    # Remove all old paragraphs
    for p in existing_ps:
        txBody.remove(p)

    if isinstance(paragraphs, str):
        para_list = paragraphs.split("\n")
    else:
        para_list = list(paragraphs)

    align_map = {"left": "l", "center": "ctr", "right": "r", "justify": "just"}

    for item in para_list:
        p = etree.SubElement(txBody, f"{{{A_NS}}}p")
        pPr = copy.deepcopy(template_pPr) if template_pPr is not None else etree.Element(f"{{{A_NS}}}pPr")
        if align is not None:
            pPr.set("algn", align_map.get(align, align))
        if space_after_pt is not None:
            for ch in list(pPr):
                if ch.tag == f"{{{A_NS}}}spcAft":
                    pPr.remove(ch)
            sa = etree.SubElement(pPr, f"{{{A_NS}}}spcAft")
            sp = etree.SubElement(sa, f"{{{A_NS}}}spcPts")
            sp.set("val", str(int(round(space_after_pt * 100))))
        if line_spacing is not None:
            for ch in list(pPr):
                if ch.tag == f"{{{A_NS}}}lnSpc":
                    pPr.remove(ch)
            ls = etree.SubElement(pPr, f"{{{A_NS}}}lnSpc")
            sp = etree.SubElement(ls, f"{{{A_NS}}}spcPct")
            sp.set("val", str(int(round(line_spacing * 100000))))
        p.append(pPr)

        if isinstance(item, str):
            runs = [(item, bold, color_hex)]
        else:
            runs = []
            for seg in item:
                if len(seg) == 2:
                    runs.append((seg[0], seg[1], color_hex))
                else:
                    runs.append((seg[0], seg[1], seg[2]))

        for txt, r_bold, r_col in runs:
            r = etree.SubElement(p, f"{{{A_NS}}}r")
            rPr = copy.deepcopy(template_rPr)
            _apply_rpr(rPr, size_pt=size_pt, bold=r_bold, color_hex=r_col, font_name=font_name)
            r.append(rPr)
            t = etree.SubElement(r, f"{{{A_NS}}}t")
            if txt.startswith(" ") or txt.endswith(" "):
                t.set(f"{{{XML_NS}}}space", "preserve")
            t.text = txt


def set_cell_text(cell, text, size_pt=9.5, bold=None, color_hex=None, font_name=FONT_BODY, align="left"):
    tc = cell._tc
    txBody = tc.find(f"{{{A_NS}}}txBody")
    if txBody is None:
        return
    existing_ps = txBody.findall(f"{{{A_NS}}}p")
    template_pPr = None
    template_rPr = None
    for p in existing_ps:
        if template_pPr is None:
            pPr = p.find(f"{{{A_NS}}}pPr")
            if pPr is not None:
                template_pPr = copy.deepcopy(pPr)
        for r in p.findall(f"{{{A_NS}}}r"):
            rPr = r.find(f"{{{A_NS}}}rPr")
            if rPr is not None:
                template_rPr = copy.deepcopy(rPr)
                break
        if template_rPr is not None:
            break
    if template_rPr is None:
        template_rPr = etree.Element(f"{{{A_NS}}}rPr")
        template_rPr.set("lang", "en-US")

    for p in existing_ps:
        txBody.remove(p)

    align_map = {"left": "l", "center": "ctr", "right": "r"}
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for ln in lines:
        p = etree.SubElement(txBody, f"{{{A_NS}}}p")
        pPr = copy.deepcopy(template_pPr) if template_pPr is not None else etree.Element(f"{{{A_NS}}}pPr")
        if align is not None:
            pPr.set("algn", align_map.get(align, align))
        p.append(pPr)
        r = etree.SubElement(p, f"{{{A_NS}}}r")
        rPr = copy.deepcopy(template_rPr)
        _apply_rpr(rPr, size_pt=size_pt, bold=bold, color_hex=color_hex, font_name=font_name)
        r.append(rPr)
        t = etree.SubElement(r, f"{{{A_NS}}}t")
        if ln.startswith(" ") or ln.endswith(" "):
            t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = ln


def remove_shape(shape):
    el = shape._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def replace_pic_with_cropped(slide, old_pic, img_path):
    """Replace a template picture with a center-cropped new image, preserving rounded-rect mask & z-order."""
    left, top, width, height = old_pic.left, old_pic.top, old_pic.width, old_pic.height
    target_ar = width / height

    im = Image.open(img_path).convert("RGB")
    iw, ih = im.size
    cur_ar = iw / ih
    if cur_ar > target_ar:
        new_w = int(round(ih * target_ar))
        x0 = (iw - new_w) // 2
        im = im.crop((x0, 0, x0 + new_w, ih))
    else:
        new_h = int(round(iw / target_ar))
        y0 = (ih - new_h) // 2
        im = im.crop((0, y0, iw, y0 + new_h))

    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=93)
    buf.seek(0)

    # Preserve exact prstGeom (e.g. roundRect with custom corner radius) from original pic
    old_spPr = old_pic._element.find(f"{{{P_NS}}}spPr")
    old_geom = copy.deepcopy(old_spPr.find(f"{{{A_NS}}}prstGeom")) if old_spPr is not None else None

    new_pic = slide.shapes.add_picture(buf, left, top, width=width, height=height)
    if old_geom is not None:
        new_spPr = new_pic._element.find(f"{{{P_NS}}}spPr")
        cur_geom = new_spPr.find(f"{{{A_NS}}}prstGeom")
        if cur_geom is not None:
            new_spPr.replace(cur_geom, old_geom)

    # Place new pic at exact z-order of old pic and remove old pic
    old_el = old_pic._element
    new_el = new_pic._element
    parent = old_el.getparent()
    idx = list(parent).index(old_el)
    parent.remove(new_el)
    parent.insert(idx, new_el)
    parent.remove(old_el)
    return new_pic


def _disable_shadow(shape):
    shape.shadow.inherit = False
    spPr = shape._element.find(f"{{{P_NS}}}spPr")
    if spPr is not None:
        old = spPr.find(f"{{{A_NS}}}effectLst")
        if old is not None:
            spPr.remove(old)
        etree.SubElement(spPr, f"{{{A_NS}}}effectLst")


def add_code_panel(slide, x, y, w, h, title, code_text, font_size=7.8, title_size=8.2):
    """Calm monochrome charcoal code box matching the template's #383838 primary tone."""
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    bg.adjustments[0] = 0.04
    bg.fill.solid()
    bg.fill.fore_color.rgb = COL_CHARCOAL
    bg.line.color.rgb = COL_MUTED
    bg.line.width = Pt(0.75)
    _disable_shadow(bg)

    # Subtle top header bar text
    hdr = slide.shapes.add_textbox(Inches(x + 0.14), Inches(y + 0.08), Inches(w - 0.28), Inches(0.24))
    tf_h = hdr.text_frame
    tf_h.word_wrap = True
    tf_h.margin_left = tf_h.margin_right = tf_h.margin_top = tf_h.margin_bottom = 0
    p_h = tf_h.paragraphs[0]
    r_h = p_h.add_run()
    r_h.text = title
    r_h.font.name = FONT_CODE
    r_h.font.size = Pt(title_size)
    r_h.font.bold = True
    r_h.font.color.rgb = COL_SILVER

    # Hairline divider
    div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.14), Inches(y + 0.34), Inches(w - 0.28), Inches(0.012))
    div.fill.solid()
    div.fill.fore_color.rgb = COL_MUTED
    div.line.fill.background()

    # Code body
    tb = slide.shapes.add_textbox(Inches(x + 0.14), Inches(y + 0.40), Inches(w - 0.28), Inches(h - 0.48))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = code_text.strip("\n").split("\n")
    for idx, ln in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = Pt(1.2)
        r = p.add_run()
        r.text = ln
        r.font.name = FONT_CODE
        r.font.size = Pt(font_size)
        r.font.color.rgb = COL_WHITE
    return bg


def add_pill_banner(slide, x, y, w, h, tag, text, font_size=8.8):
    """Template-native crisp white horizontal banner strip (#FFFFFF + #383838 tag + #595959 copy)."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = COL_WHITE
    box.line.color.rgb = COL_SILVER
    box.line.width = Pt(0.6)
    _disable_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.14)
    tf.margin_right = Inches(0.14)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if tag:
        r1 = p.add_run()
        r1.text = f"{tag}  |  "
        r1.font.name = FONT_HEAD
        r1.font.size = Pt(font_size)
        r1.font.bold = True
        r1.font.color.rgb = COL_CHARCOAL
    r2 = p.add_run()
    r2.text = text
    r2.font.name = FONT_BODY
    r2.font.size = Pt(font_size)
    r2.font.color.rgb = COL_MUTED
    return box


def add_monochrome_footprint_chart(slide, x, y, w, h):
    """Native PowerPoint grouped bar chart styled in the template's exact monochrome charcoal/silver tones."""
    chart_data = CategoryChartData()
    chart_data.categories = ["Kernel\n(Flash)", "Kernel\n(RAM)", "Full IoT*\n(Flash)", "Full IoT*\n(RAM)"]
    chart_data.add_series("FreeRTOS + Vendor HAL", (9, 2, 195, 44))
    chart_data.add_series("Zephyr RTOS", (14, 4, 220, 48))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), chart_data
    )
    chart = chart_shape.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Real Flash & RAM Footprint (KB)"
    for r in chart.chart_title.text_frame.paragraphs[0].runs:
        r.font.name = FONT_HEAD
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = COL_CHARCOAL

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.name = FONT_BODY
    chart.legend.font.size = Pt(8.5)
    chart.legend.font.color.rgb = COL_MUTED

    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.gap_width = 110
    plot.overlap = -15

    # Style series in exact template monochrome palette
    colors = [COL_MIDGRAY, COL_CHARCOAL]
    for s_idx, ser in enumerate(plot.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = colors[s_idx]
        ser.format.line.fill.background()

    dl = plot.data_labels
    dl.show_value = True
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.name = FONT_BODY
    dl.font.size = Pt(8)
    dl.font.bold = True
    dl.font.color.rgb = COL_CHARCOAL

    val_axis = chart.value_axis
    val_axis.maximum_scale = 260
    val_axis.minimum_scale = 0
    val_axis.major_unit = 50
    val_axis.tick_labels.font.name = FONT_BODY
    val_axis.tick_labels.font.size = Pt(8)
    val_axis.tick_labels.font.color.rgb = COL_MUTED
    val_axis.major_gridlines.format.line.color.rgb = COL_SILVER
    val_axis.major_gridlines.format.line.width = Pt(0.5)
    val_axis.format.line.color.rgb = COL_SILVER

    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.name = FONT_BODY
    cat_axis.tick_labels.font.size = Pt(8.5)
    cat_axis.tick_labels.font.color.rgb = COL_CHARCOAL
    cat_axis.format.line.color.rgb = COL_SILVER
    cat_axis.has_major_gridlines = False
    return chart_shape


def add_chain_box(slide, x, y, w, h, header, body, dark=False):
    """Clean minimalist monochrome architectural block matching template styling."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = COL_CHARCOAL if dark else COL_WHITE
    box.line.color.rgb = COL_CHARCOAL if dark else COL_SILVER
    box.line.width = Pt(0.75)
    _disable_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(2)
    r1 = p1.add_run()
    r1.text = header
    r1.font.name = FONT_HEAD
    r1.font.size = Pt(8.8)
    r1.font.bold = True
    r1.font.color.rgb = COL_WHITE if dark else COL_CHARCOAL

    if body:
        for ln in body.split("\n"):
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_after = Pt(1)
            r2 = p2.add_run()
            r2.text = ln
            r2.font.name = FONT_BODY
            r2.font.size = Pt(7.6)
            r2.font.color.rgb = COL_PLATINUM if dark else COL_MUTED
    return box


def add_verdict_card(slide, x, y, w, h, title, bullets, dark=False):
    """Spacious editorial left-aligned verdict card matching the template's monochrome minimalism."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = COL_CHARCOAL if dark else COL_WHITE
    box.line.color.rgb = COL_CHARCOAL if dark else COL_SILVER
    box.line.width = Pt(0.75)
    _disable_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.24)
    tf.margin_right = Inches(0.24)
    tf.margin_top = Inches(0.22)
    tf.margin_bottom = Inches(0.18)

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    p0.space_after = Pt(10)
    r0 = p0.add_run()
    r0.text = title
    r0.font.name = FONT_HEAD
    r0.font.size = Pt(13.5)
    r0.font.bold = True
    r0.font.color.rgb = COL_WHITE if dark else COL_CHARCOAL

    for pre, desc in bullets:
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(7)
        r1 = p.add_run()
        r1.text = f"• {pre}: "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(9.6)
        r1.font.bold = True
        r1.font.color.rgb = COL_WHITE if dark else COL_CHARCOAL
        r2 = p.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(9.6)
        r2.font.color.rgb = COL_PLATINUM if dark else COL_MUTED
    return box


def add_right_arrow(slide, x, y, w=0.18, h=0.14):
    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    arr.fill.solid()
    arr.fill.fore_color.rgb = COL_MUTED
    arr.line.fill.background()
    _disable_shadow(arr)
    return arr


def add_down_arrow(slide, x, y, w=0.14, h=0.18):
    arr = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    arr.fill.solid()
    arr.fill.fore_color.rgb = COL_MUTED
    arr.line.fill.background()
    _disable_shadow(arr)
    return arr


def add_up_arrow(slide, x, y, w=0.14, h=0.18):
    arr = slide.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    arr.fill.solid()
    arr.fill.fore_color.rgb = COL_MUTED
    arr.line.fill.background()
    _disable_shadow(arr)
    return arr


def add_left_right_arrow(slide, x, y, w=0.28, h=0.14):
    arr = slide.shapes.add_shape(MSO_SHAPE.LEFT_RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    arr.fill.solid()
    arr.fill.fore_color.rgb = COL_CHARCOAL
    arr.line.fill.background()
    _disable_shadow(arr)
    return arr


def add_lead_subtitle(slide, x, y, w, h, text, size_pt=10.0, bold=False, color=None, align="left"):
    """Clean borderless lead-in sentence or label that builds narrative flow below slide title."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    al = PP_ALIGN.CENTER if align == "center" else (PP_ALIGN.RIGHT if align == "right" else PP_ALIGN.LEFT)
    for idx_l, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if idx_l == 0 else tf.add_paragraph()
        p.alignment = al
        r = p.add_run()
        r.text = ln
        r.font.name = FONT_BODY
        r.font.size = Pt(size_pt)
        r.font.bold = bold
        r.font.color.rgb = color if color is not None else COL_MUTED
    return tb


def add_card(slide, x, y, w, h, title, body_lines, dark=False, align="left", title_pt=11.0, body_pt=9.2, soft=False):
    """Structured monochrome card matching template palette (#FFFFFF, #F5F5F5, or #383838)."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    if dark:
        box.fill.fore_color.rgb = COL_CHARCOAL
        box.line.color.rgb = COL_CHARCOAL
    elif soft:
        box.fill.fore_color.rgb = COL_SOFTBG
        box.line.color.rgb = COL_SILVER
    else:
        box.fill.fore_color.rgb = COL_WHITE
        box.line.color.rgb = COL_SILVER
    box.line.width = Pt(0.75)
    _disable_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP if align == "left" else MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.14)
    tf.margin_right = Inches(0.14)
    tf.margin_top = Inches(0.11)
    tf.margin_bottom = Inches(0.10)

    al = PP_ALIGN.CENTER if align == "center" else PP_ALIGN.LEFT
    p0 = tf.paragraphs[0]
    p0.alignment = al
    p0.space_after = Pt(4 if body_lines else 0)
    r0 = p0.add_run()
    r0.text = title
    r0.font.name = FONT_HEAD
    r0.font.size = Pt(title_pt)
    r0.font.bold = True
    r0.font.color.rgb = COL_WHITE if dark else COL_CHARCOAL

    if isinstance(body_lines, str):
        body_lines = body_lines.split("\n")
    for idx_b, line in enumerate(body_lines):
        p = tf.add_paragraph()
        p.alignment = al
        p.space_after = Pt(3 if idx_b < len(body_lines) - 1 else 0)
        if isinstance(line, tuple):
            pre, rest = line
            r1 = p.add_run()
            r1.text = pre
            r1.font.name = FONT_BODY
            r1.font.size = Pt(body_pt)
            r1.font.bold = True
            r1.font.color.rgb = COL_WHITE if dark else COL_CHARCOAL
            r2 = p.add_run()
            r2.text = rest
            r2.font.name = FONT_BODY
            r2.font.size = Pt(body_pt)
            r2.font.color.rgb = COL_PLATINUM if dark else COL_MUTED
        else:
            r = p.add_run()
            r.text = line
            r.font.name = FONT_BODY
            r.font.size = Pt(body_pt)
            r.font.color.rgb = COL_PLATINUM if dark else COL_MUTED
    return box


def add_lane_block(slide, x, y, w, h, label, dark=False, muted=False, dashed=False, font_pt=8.2):
    """Crisp timeline segment block for execution lanes, state graphs & heap maps."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    if dark:
        box.fill.fore_color.rgb = COL_CHARCOAL
        box.line.color.rgb = COL_CHARCOAL
    elif muted:
        box.fill.fore_color.rgb = COL_PLATINUM
        box.line.color.rgb = COL_MIDGRAY
    else:
        box.fill.fore_color.rgb = COL_WHITE
        box.line.color.rgb = COL_SILVER
    box.line.width = Pt(0.75)
    _disable_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    for idx_l, ln in enumerate(label.split("\n")):
        p = tf.paragraphs[0] if idx_l == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ln
        r.font.name = FONT_BODY
        r.font.size = Pt(font_pt)
        r.font.bold = True if (idx_l == 0 or dark) else False
        r.font.color.rgb = COL_WHITE if dark else COL_CHARCOAL
    return box


def clear_fill_and_line(shape):
    """Remove solid background fill & border from a template placeholder (e.g. Slide 1 subtitle slab)."""
    spPr = shape._element.find(f"{{{P_NS}}}spPr")
    if spPr is None:
        spPr = etree.SubElement(shape._element, f"{{{P_NS}}}spPr")
    for ch in list(spPr):
        if ch.tag in (f"{{{A_NS}}}solidFill", f"{{{A_NS}}}gradFill", f"{{{A_NS}}}noFill", f"{{{A_NS}}}ln", f"{{{A_NS}}}effectLst"):
            spPr.remove(ch)
    etree.SubElement(spPr, f"{{{A_NS}}}noFill")
    ln = etree.SubElement(spPr, f"{{{A_NS}}}ln")
    etree.SubElement(ln, f"{{{A_NS}}}noFill")
    etree.SubElement(spPr, f"{{{A_NS}}}effectLst")


def set_notes(slide, notes_text):
    ns = slide.notes_slide
    ns.notes_text_frame.text = notes_text
