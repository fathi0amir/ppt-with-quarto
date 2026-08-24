"""
Build a customized PowerPoint template from Quarto's default reference doc
and patch rendered decks (fonts, colors, geometry, math images, text autofit).

Usage (uv):
    uv run style_template.py                      # build template.pptx
    uv run style_template.py --config custom.yml  # build with an alternate config
    uv run style_template.py --patch-deck main.pptx
    uv run style_template.py --render main.qmd    # quarto render + patch in one step
    uv run style_template.py --no-regen ...       # skip regenerating the template

All styling knobs live in style_config.yml (see that file for the full list).
"""

import argparse
import re
import struct
import subprocess
import sys
from pathlib import Path

import yaml
from lxml import etree
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Pt

# ---------------------------------------------------------------------------
# Constants & defaults
# ---------------------------------------------------------------------------

EMU_PER_INCH = 914400
DEFAULT_DPI = 96.0
QUARTO_DEFAULT_SLIDE = (9144000, 5143500)  # 16:9 reference doc

# Table style GUIDs (must be defined in ppt/tableStyles.xml)
TABLE_STYLE_DEFAULT_GUID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"  # Medium Style 2 - Accent 1
TABLE_STYLE_SIMPLE_GUID = "{8F8B1A6E-4C7E-4A9A-9E3B-1234567890AB}"  # Simple Grid (defined below)

# Simple table style: white cells, thin dark borders, no banding/emphasis.
SIMPLE_TABLE_STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{TABLE_STYLE_SIMPLE_GUID}">
  <a:tblStyle styleId="{TABLE_STYLE_SIMPLE_GUID}" styleName="Simple Grid">
    <a:wholeTbl>
      <a:tcTxStyle>
        <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>
        <a:schemeClr val="tx1"/>
      </a:tcTxStyle>
      <a:tcStyle>
        <a:tcBdr>
          <a:left><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:left>
          <a:right><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:right>
          <a:top><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:top>
          <a:bottom><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:bottom>
          <a:insideH><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:insideH>
          <a:insideV><a:ln w="6350"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:insideV>
        </a:tcBdr>
      </a:tcStyle>
    </a:wholeTbl>
  </a:tblStyle>
</a:tblStyleLst>'''

RATIOS = {
    "16:9": (9144000, 5143500),
    "16:10": (9144000, 6250000),
    "4:3": (9144000, 6858000),
}

# OOXML child ordering (used to insert elements at schema-valid positions)
PPR_ORDER = [
    "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts",
    "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "tabLst", "defRPr",
]
RPR_ORDER = [
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
]
FONT_SCHEME_ORDER = ["latin", "ea", "cs", "font"]

TITLE_PH = {"title", "ctrTitle"}
BODY_PH = {"body", "subTitle"}
FOOTER_PH = {"dt", "ftr", "sldNum"}

DEFAULTS = {
    "slide": {"ratio": "4:3", "width": None, "height": None},
    "font": {
        "name": "Candara",
        "title_name": None,
        "mono": "Consolas",
        "title_size": 28,
        "subtitle_size": 20,
        "body_size": 18,
        "code_size": 12,
        "level_sizes": None,
    },
    "colors": {
        "title_text": "1F3864",
        "title_fill": "ADD8E6",
        "body_text": "262626",
        "background": "",
        "accent1": "", "accent2": "", "accent3": "", "accent4": "",
        "accent5": "", "accent6": "",
        "link": "",
    },
    "layout": {
        "title_align": "ctr",
        "subtitle_align": "ctr",
        "body_align": "l",
        "title_bold": False,
        "body_bold": False,
        "vertical_anchor": "ctr",
        "autofit": False,
        "insets": {"left": 0.10, "right": 0.10, "top": 0.05, "bottom": 0.05},
        "line_spacing": 1.0,
        "space_after": 4.0,
        "body_indent": 0.37,
        "footer_bottom_margin": 1.5,
        "content_gap": 1.0,
        "caption_ratio": 0.28,
        "expand_body_to_footer": True,
        "geometry_overrides": {},
    },
    "patch": {
        "autofit": True,
        "apply_insets": True,
        "normalize_indents": True,
        "code_font": True,
        "fix_math_images": True,
        "clamp_math_to_slide": True,
        "math_image_prefix": "math-png-cache/",
    },
    "table": {"style": "simple"},  # "simple" | "default"
}


def deep_merge(base, override):
    """Recursively merge override dict into base dict (override wins)."""
    out = dict(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = deep_merge(out[key], value)
        else:
            out[key] = value
    return out


def load_config(path):
    cfg = deep_merge(DEFAULTS, {})
    path = Path(path)
    if path.exists():
        with open(path, encoding="utf-8") as fh:
            user = yaml.safe_load(fh) or {}
        cfg = deep_merge(cfg, user)
    else:
        print(f"Note: config {path} not found, using built-in defaults.")
    return cfg


def norm_hex(value):
    """Normalize a hex color (accepts with/without '#', None, '')."""
    if not value:
        return ""
    return str(value).lstrip("#").upper()


def emu_from_inches(value):
    return int(round(float(value) * EMU_PER_INCH))


def pct_to_emu(value, total):
    """Interpret a value as a percentage (0-100) of `total` EMUs."""
    value = float(value)
    if 0 <= value <= 100:
        return int(round(value / 100.0 * total))
    return int(round(value))


def replace_zip_entry(path, entry_name, content):
    """Rewrite a .pptx, swapping one entry's bytes (used for tableStyles.xml)."""
    import shutil
    import zipfile
    path = str(path)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = content if item.filename == entry_name else zin.read(item.filename)
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------


def _insert_ordered(parent, new_elem, order):
    """Insert new_elem before the first child that sorts after it."""
    new_rank = order.index(etree.QName(new_elem).localname) if etree.QName(new_elem).localname in order else len(order)
    pos = len(parent)
    for i, child in enumerate(parent):
        tag = etree.QName(child).localname
        if tag in order and order.index(tag) > new_rank:
            pos = i
            break
    parent.insert(pos, new_elem)


def iter_placeholder_sp(element):
    """Yield (sp, ph) for every placeholder shape under an element."""
    for sp in element.iter(qn("p:sp")):
        ph = sp.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}ph")
        if ph is not None:
            yield sp, ph


def ensure_xfrm(sp):
    """Return (off, ext) creating spPr/xfrm/off/ext if missing.
    Only use when a placeholder is SUPPOSED to carry explicit geometry
    (e.g. user geometry overrides); the layout functions below must not
    invent geometry, because pandoc copies the layout's title geometry
    into slides and an invented 0x0 box makes titles invisible."""
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        spPr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = spPr.find(qn("a:xfrm"))
    if xfrm is None:
        xfrm = etree.SubElement(spPr, qn("a:xfrm"))
    off = xfrm.find(qn("a:off"))
    if off is None:
        off = etree.SubElement(xfrm, qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if ext is None:
        ext = etree.SubElement(xfrm, qn("a:ext"))
    return off, ext


def existing_xfrm(sp):
    """Return (off, ext) if the placeholder already carries geometry, else None.
    Layouts that inherit geometry from the master have empty <p:spPr/> and
    must be left untouched."""
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        return None
    xfrm = spPr.find(qn("a:xfrm"))
    if xfrm is None:
        return None
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    return off, ext


def ensure_lst_style(sp):
    txBody = sp.find(qn("p:txBody"))
    if txBody is None:
        return None
    lstStyle = txBody.find(qn("a:lstStyle"))
    if lstStyle is None:
        lstStyle = etree.SubElement(txBody, qn("a:lstStyle"))
    return lstStyle


def ensure_lvl_pPr(lstStyle, lvl):
    tag = qn(f"a:lvl{lvl}pPr")
    pPr = lstStyle.find(tag)
    if pPr is None:
        pPr = etree.SubElement(lstStyle, tag)
    return pPr


def ensure_def_rpr(lvl_pPr):
    defRPr = lvl_pPr.find(qn("a:defRPr"))
    if defRPr is None:
        defRPr = etree.SubElement(lvl_pPr, qn("a:defRPr"))
    return defRPr


def set_def_rpr_fill(defRPr, hex_color):
    solid = defRPr.find(qn("a:solidFill"))
    if solid is None:
        solid = etree.Element(qn("a:solidFill"))
        _insert_ordered(defRPr, solid, RPR_ORDER)
    for child in list(solid):
        solid.remove(child)
    etree.SubElement(solid, qn("a:srgbClr")).set("val", hex_color)


def set_ph_fill(sp, hex_color):
    """Set the solid background fill of a placeholder shape."""
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        spPr = etree.SubElement(sp, qn("p:spPr"))
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        solid = etree.Element(qn("a:solidFill"))
        _insert_ordered(spPr, solid, RPR_ORDER)
    for child in list(solid):
        solid.remove(child)
    etree.SubElement(solid, qn("a:srgbClr")).set("val", hex_color)


def set_vertical_anchor(sp, anchor):
    txBody = sp.find(qn("p:txBody"))
    if txBody is None:
        return
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txBody, qn("a:bodyPr"))
    bodyPr.set("anchor", anchor)


def set_insets(sp, insets):
    txBody = sp.find(qn("p:txBody"))
    if txBody is None:
        return
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txBody, qn("a:bodyPr"))
    bodyPr.set("lIns", str(emu_from_inches(insets["left"])))
    bodyPr.set("rIns", str(emu_from_inches(insets["right"])))
    bodyPr.set("tIns", str(emu_from_inches(insets["top"])))
    bodyPr.set("bIns", str(emu_from_inches(insets["bottom"])))


def set_autofit(sp):
    txBody = sp.find(qn("p:txBody"))
    if txBody is None:
        return
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txBody, qn("a:bodyPr"))
    bodyPr.set("wrap", "square")
    for tag in (qn("a:noAutofit"), qn("a:spAutoFit"), qn("a:normAutofit")):
        node = bodyPr.find(tag)
        if node is not None:
            bodyPr.remove(node)
    norm = etree.SubElement(bodyPr, qn("a:normAutofit"))
    norm.set("fontScale", "100000")


# ---------------------------------------------------------------------------
# Text defaults (sizes, bold, color, alignment, spacing)
# ---------------------------------------------------------------------------


def style_placeholder_defaults(sp, ph_type, cfg, kind):
    """Apply font size/bold/color/alignment/paragraph spacing to a placeholder's
    lstStyle defaults for all 9 levels (these flow through to rendered slides,
    because pandoc writes empty <a:rPr/> on slide runs)."""
    lstStyle = ensure_lst_style(sp)
    if lstStyle is None:
        return

    font = cfg["font"]
    layout = cfg["layout"]
    colors = cfg["colors"]

    if kind == "title":
        sizes = [font["title_size"]] * 9
        bold = layout["title_bold"]
        color = norm_hex(colors["title_text"])
        align = layout["title_align"]
    elif kind == "subtitle":
        sizes = [font["subtitle_size"]] * 9
        bold = layout["body_bold"]
        color = norm_hex(colors["body_text"])
        align = layout["subtitle_align"]
    else:
        level_sizes = font.get("level_sizes")
        if level_sizes:
            sizes = [int(s) for s in level_sizes] + [font["body_size"]] * 9
            sizes = sizes[:9]
        else:
            sizes = [font["body_size"]] * 9
        bold = layout["body_bold"]
        color = norm_hex(colors["body_text"])
        align = layout["body_align"]

    for lvl in range(1, 10):
        pPr = ensure_lvl_pPr(lstStyle, lvl)
        if lvl == 1:
            pPr.set("algn", align)
        defRPr = ensure_def_rpr(pPr)
        defRPr.set("sz", str(int(sizes[lvl - 1] * 100)))
        defRPr.set("b", "1" if bold else "0")
        if color:
            set_def_rpr_fill(defRPr, color)

        if kind == "body":
            # line spacing
            lnSpc = pPr.find(qn("a:lnSpc"))
            if lnSpc is not None:
                pPr.remove(lnSpc)
            lnSpc = etree.Element(qn("a:lnSpc"))
            _insert_ordered(pPr, lnSpc, PPR_ORDER)
            etree.SubElement(lnSpc, qn("a:spcPct")).set(
                "val", str(int(float(layout["line_spacing"]) * 100000))
            )
            # space after
            spcAft = pPr.find(qn("a:spcAft"))
            if spcAft is not None:
                pPr.remove(spcAft)
            spcAft = etree.Element(qn("a:spcAft"))
            _insert_ordered(pPr, spcAft, PPR_ORDER)
            etree.SubElement(spcAft, qn("a:spcPts")).set(
                "val", str(int(float(layout["space_after"]) * 100))
            )
            # bullet indent
            if layout.get("body_indent") is not None:
                indent_emu = emu_from_inches(layout["body_indent"])
                pPr.set("marL", str(indent_emu))
                pPr.set("indent", str(-indent_emu))


def apply_text_defaults(prs, cfg):
    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            for sp, ph in iter_placeholder_sp(element):
                ph_type = ph.get("type", "body")
                if ph_type in TITLE_PH:
                    style_placeholder_defaults(sp, ph_type, cfg, "title")
                elif ph_type == "subTitle":
                    style_placeholder_defaults(sp, ph_type, cfg, "subtitle")
                elif ph_type in BODY_PH:
                    style_placeholder_defaults(sp, ph_type, cfg, "body")
                set_vertical_anchor(sp, cfg["layout"]["vertical_anchor"])
                set_insets(sp, cfg["layout"]["insets"])
                if cfg["layout"]["autofit"]:
                    set_autofit(sp)


# ---------------------------------------------------------------------------
# Fonts (theme + masters/layouts)
# ---------------------------------------------------------------------------


def apply_theme(prs, cfg):
    """Patch theme font scheme (major/minor) and color scheme (accents/link)."""
    font_name = cfg["font"]["name"]
    title_name = cfg["font"]["title_name"] or font_name
    colors = cfg["colors"]

    for master in prs.slide_masters:
        part = _theme_part(master)
        if part is None:
            print("  warning: could not locate theme part; skipping theme fonts/colors.")
            continue
        theme = etree.fromstring(part.blob)

        # major/minor font schemes
        for tag, typeface in (("majorFont", title_name), ("minorFont", font_name)):
            scheme = theme.find(f".//{qn('a:' + tag)}")
            if scheme is None:
                continue
            for child_tag in ("latin", "ea", "cs"):
                node = scheme.find(qn("a:" + child_tag))
                if node is None:
                    node = etree.Element(qn("a:" + child_tag))
                    _insert_ordered(scheme, node, FONT_SCHEME_ORDER)
                node.set("typeface", typeface)

        # color scheme accents + hyperlink color
        scheme = theme.find(".//" + qn("a:clrScheme"))
        if scheme is None:
            continue
        for tag in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                    "hlinkClr", "folHlinkClr"):
            hex_color = norm_hex(colors.get(tag, ""))
            if not hex_color:
                continue
            node = scheme.find(qn("a:" + tag))
            if node is None:
                continue
            for child in list(node):
                node.remove(child)
            etree.SubElement(node, qn("a:srgbClr")).set("val", hex_color)

        # persist the modified theme back into the package
        part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)


def _theme_part(master):
    """Return the slide master's theme part (as an OPC Part)."""
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            return rel.target_part
    return None


def set_slide_size(prs, cfg):
    """Apply the configured slide dimensions (and a matching sldSz type)."""
    slide_size = resolve_slide_size(cfg)
    prs.slide_width, prs.slide_height = slide_size
    sldSz = prs.part._element.find(qn("p:sldSz"))
    if sldSz is not None:
        ratio = cfg["slide"].get("ratio")
        sldSz.set("type", {"16:9": "screen16x9", "16:10": "screen16x10", "4:3": "screen4x3"}.get(ratio, "custom"))
    return slide_size


def apply_fonts_to_masters(prs, cfg):
    """Set the typeface on every placeholder in masters + layouts."""
    font_name = cfg["font"]["name"]
    title_name = cfg["font"]["title_name"] or font_name

    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            for sp, ph in iter_placeholder_sp(element):
                typeface = title_name if ph.get("type", "body") in TITLE_PH else font_name
                for tag in (qn("a:latin"), qn("a:cs"), qn("a:ea")):
                    for node in sp.iter(tag):
                        node.set("typeface", typeface)


# ---------------------------------------------------------------------------
# Colors (text, fills, slide background)
# ---------------------------------------------------------------------------


def apply_placeholder_fills(prs, cfg):
    """Optional background fill on title placeholders (e.g. a highlight bar)."""
    title_hex = norm_hex(cfg["colors"].get("title_fill", ""))
    if not title_hex:
        return
    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            for sp, ph in iter_placeholder_sp(element):
                if ph.get("type", "body") in TITLE_PH:
                    set_ph_fill(sp, title_hex)


def apply_slide_background(prs, cfg):
    """Set a solid slide background on the master and every layout."""
    hex_color = norm_hex(cfg["colors"].get("background", ""))
    if not hex_color:
        return
    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            cSld = element.find(qn("p:cSld"))
            if cSld is None:
                continue
            for bg in cSld.findall(qn("p:bg")):
                cSld.remove(bg)
            bg = etree.Element(qn("p:bg"))
            bgPr = etree.SubElement(bg, qn("p:bgPr"))
            fill = etree.SubElement(bgPr, qn("a:solidFill"))
            etree.SubElement(fill, qn("a:srgbClr")).set("val", hex_color)
            cSld.insert(0, bg)


# ---------------------------------------------------------------------------
# Geometry: aspect ratio, overrides, footer anchoring
# ---------------------------------------------------------------------------


def resolve_slide_size(cfg):
    ratio = cfg["slide"].get("ratio")
    if ratio and ratio in RATIOS:
        return RATIOS[ratio]
    width_in = cfg["slide"].get("width")
    height_in = cfg["slide"].get("height")
    if width_in and height_in:
        return emu_from_inches(width_in), emu_from_inches(height_in)
    return RATIOS["4:3"]


def scale_placeholders_for_size(prs, target, source=QUARTO_DEFAULT_SLIDE):
    """Stretch placeholder geometry from the source aspect ratio to the target.
    The Quarto default reference doc is 16:9; when the target ratio grows taller
    (4:3, 16:10) placeholders scale along that axis so content fills the slide."""
    sx = target[0] / source[0]
    sy = target[1] / source[1]
    if abs(sx - 1.0) < 1e-9 and abs(sy - 1.0) < 1e-9:
        return
    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            for sp, _ph in iter_placeholder_sp(element):
                geo = existing_xfrm(sp)
                if geo is None:
                    continue  # layout inherits geometry from the master; don't invent any
                off, ext = geo
                off.set("x", str(int(round(int(off.get("x", 0)) * sx))))
                off.set("y", str(int(round(int(off.get("y", 0)) * sy))))
                ext.set("cx", str(int(round(int(ext.get("cx", 0)) * sx))))
                ext.set("cy", str(int(round(int(ext.get("cy", 0)) * sy))))


def apply_geometry_overrides(prs, cfg, slide_size):
    """Apply explicit per-placeholder geometry (percent of slide) if configured."""
    overrides = cfg["layout"].get("geometry_overrides") or {}
    if not overrides:
        return
    sw, sh = slide_size

    def _apply(sp, ph, geom):
        off, ext = ensure_xfrm(sp)
        x = pct_to_emu(geom.get("x", 0), sw)
        y = pct_to_emu(geom.get("y", 0), sh)
        w = pct_to_emu(geom.get("w", 100), sw)
        h = pct_to_emu(geom.get("h", 100), sh)
        if ph.get("type") == "ctrTitle":
            x = (sw - w) // 2
        off.set("x", str(x))
        off.set("y", str(y))
        ext.set("cx", str(w))
        ext.set("cy", str(h))

    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            for sp, ph in iter_placeholder_sp(element):
                ph_type = ph.get("type", "body")
                if ph_type in TITLE_PH and "title" in overrides:
                    _apply(sp, ph, overrides["title"])
                elif ph_type in BODY_PH and "body" in overrides:
                    _apply(sp, ph, overrides["body"])


def anchor_footer_placeholders_to_bottom(prs, slide_height, bottom_margin_pct=1.5):
    """Move date/footer/slide-number placeholders to the visual bottom."""
    margin_emu = int(slide_height * (bottom_margin_pct / 100.0))
    for master in prs.slide_masters:
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            for sp, ph in iter_placeholder_sp(element):
                if ph.get("type") not in FOOTER_PH:
                    continue
                geo = existing_xfrm(sp)
                if geo is None:
                    continue  # inherits from master; leave untouched
                off, ext = geo
                current_h = int(ext.get("cy", 0))
                new_y = max(0, slide_height - current_h - margin_emu)
                off.set("y", str(new_y))


def expand_body_placeholders_to_footer(prs, slide_height, content_gap_pct=1.0):
    """Extend body placeholders downward to just above the footer row."""
    gap_emu = int(slide_height * (content_gap_pct / 100.0))

    def _footer_top(element):
        top = slide_height
        for sp, ph in iter_placeholder_sp(element):
            if ph.get("type") not in FOOTER_PH:
                continue
            geo = existing_xfrm(sp)
            if geo is None:
                continue
            off, ext = geo
            top = min(top, int(off.get("y", 0)))
        return top

    for master in prs.slide_masters:
        master_top = _footer_top(master.element)
        for element in (master.element, *[layout.element for layout in master.slide_layouts]):
            footer_top = _footer_top(element)
            target_bottom = max(0, min(footer_top, master_top) - gap_emu)
            for sp, ph in iter_placeholder_sp(element):
                if ph.get("type") not in BODY_PH:
                    continue
                geo = existing_xfrm(sp)
                if geo is None:
                    continue  # inherits master geometry (already expanded)
                off, ext = geo
                y = int(off.get("y", 0))
                new_h = target_bottom - y
                if new_h > 0:
                    ext.set("cy", str(new_h))


def rearrange_content_with_caption_layouts(prs, slide_height, vertical_gap_pct=1.0, caption_ratio=0.28):
    """Rearrange Content with Caption layouts: title full width, content on top,
    caption below (instead of the default side-by-side arrangement)."""
    pml_ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    gap_emu = int(slide_height * (vertical_gap_pct / 100.0))

    def _master_geom(master, ph_type, ph_idx=None):
        for sp, ph in iter_placeholder_sp(master.element):
            if ph.get("type", "body") != ph_type:
                continue
            if ph_idx is not None and ph.get("idx") != str(ph_idx):
                continue
            off, ext = ensure_xfrm(sp)
            return (int(off.get("x", 0)), int(off.get("y", 0)),
                    int(ext.get("cx", 0)), int(ext.get("cy", 0)))
        return None

    def _normalize_title(sp):
        lstStyle = ensure_lst_style(sp)
        if lstStyle is None:
            return
        pPr = ensure_lvl_pPr(lstStyle, 1)
        pPr.set("algn", "ctr")
        set_def_rpr_fill(ensure_def_rpr(pPr), "1F3864")

    for master in prs.slide_masters:
        title_geom = _master_geom(master, "title")
        body_geom = _master_geom(master, "body", ph_idx=1)
        if body_geom is None:
            continue
        body_x, body_y, body_w, body_h = body_geom
        caption_h = int(body_h * caption_ratio)
        content_h = body_h - caption_h - gap_emu
        if content_h <= 0:
            continue

        for layout in master.slide_layouts:
            if "content with caption" not in layout.name.lower():
                continue
            title_sp = content_sp = caption_sp = None
            body_candidates = []
            for sp, ph in iter_placeholder_sp(layout.element):
                ph_type = ph.get("type", "body")
                if ph_type in TITLE_PH:
                    title_sp = sp
                elif ph_type == "body":
                    body_candidates.append((ph.get("idx"), sp))
            for ph_idx, sp in body_candidates:
                if ph_idx == "1":
                    content_sp = sp
                elif ph_idx == "2":
                    caption_sp = sp
            if content_sp is None and body_candidates:
                content_sp = body_candidates[0][1]
            if caption_sp is None and len(body_candidates) > 1:
                caption_sp = body_candidates[1][1]

            if title_sp is not None and title_geom is not None:
                off, ext = ensure_xfrm(title_sp)
                off.set("x", str(title_geom[0]))
                off.set("y", str(title_geom[1]))
                ext.set("cx", str(title_geom[2]))
                ext.set("cy", str(title_geom[3]))
                _normalize_title(title_sp)

            if caption_sp is not None:
                off, ext = ensure_xfrm(caption_sp)
                off.set("x", str(body_x))
                off.set("y", str(body_y))
                ext.set("cx", str(body_w))
                ext.set("cy", str(caption_h))

            if content_sp is not None:
                off, ext = ensure_xfrm(content_sp)
                off.set("x", str(body_x))
                off.set("y", str(body_y + caption_h + gap_emu))
                ext.set("cx", str(body_w))
                ext.set("cy", str(max(0, content_h)))


# ---------------------------------------------------------------------------
# Deck patching (post-render fixes)
# ---------------------------------------------------------------------------


def _shape_description(shape):
    cNvPr = shape.element.find(".//" + qn("p:cNvPr"))
    if cNvPr is None:
        return ""
    return cNvPr.get("descr", "")


def _image_size_to_emu(image_path):
    suffix = image_path.suffix.lower()
    if suffix == ".png":
        return _png_size_to_emu(image_path)
    if suffix == ".svg":
        return _svg_size_to_emu(image_path)
    raise ValueError(f"Unsupported math image format: {image_path}")


def _png_size_to_emu(image_path):
    with image_path.open("rb") as fh:
        if fh.read(8) != b"\x89PNG\r\n\x1a\n":
            raise ValueError(f"Not a PNG file: {image_path}")
        width_px = height_px = None
        dpi_x = dpi_y = DEFAULT_DPI
        while True:
            length_bytes = fh.read(4)
            if len(length_bytes) != 4:
                break
            chunk_length = struct.unpack(">I", length_bytes)[0]
            chunk_type = fh.read(4)
            chunk_data = fh.read(chunk_length)
            fh.read(4)  # CRC
            if chunk_type == b"IHDR":
                width_px, height_px = struct.unpack(">II", chunk_data[:8])
            elif chunk_type == b"pHYs" and len(chunk_data) >= 9:
                ppu_x, ppu_y, unit = struct.unpack(">IIB", chunk_data[:9])
                if unit == 1:
                    dpi_x = ppu_x * 0.0254
                    dpi_y = ppu_y * 0.0254
            elif chunk_type == b"IEND":
                break
    if width_px is None or height_px is None:
        raise ValueError(f"Could not read PNG size: {image_path}")
    return int(round(width_px / dpi_x * EMU_PER_INCH)), int(round(height_px / dpi_y * EMU_PER_INCH))


def _svg_size_to_emu(image_path):
    root = etree.parse(str(image_path)).getroot()
    width, height = root.get("width"), root.get("height")
    view_box = root.get("viewBox")
    if width and height:
        return _svg_length_to_emu(width), _svg_length_to_emu(height)
    if view_box:
        parts = [float(v) for v in view_box.replace(",", " ").split()]
        return _svg_length_to_emu(str(parts[2])), _svg_length_to_emu(str(parts[3]))
    raise ValueError(f"Could not determine SVG size: {image_path}")


def _svg_length_to_emu(length):
    match = re.fullmatch(r"\s*([0-9]*\.?[0-9]+)\s*([a-zA-Z%]*)", length)
    if not match:
        raise ValueError(f"Unsupported SVG length: {length}")
    value = float(match.group(1))
    unit = match.group(2) or "px"
    if unit == "pt":
        inches = value / 72.0
    elif unit == "pc":
        inches = value / 6.0
    elif unit == "in":
        inches = value
    elif unit == "cm":
        inches = value / 2.54
    elif unit == "mm":
        inches = value / 25.4
    elif unit == "px":
        inches = value / DEFAULT_DPI
    else:
        raise ValueError(f"Unsupported SVG unit: {unit}")
    return int(round(inches * EMU_PER_INCH))


def patch_deck(cfg, deck_path):
    """Post-process a rendered PPTX: text autofit, insets, code font,
    paragraph indents, and true-size math images."""
    deck_path = str(deck_path)
    deck = Presentation(deck_path)
    deck_dir = Path(deck_path).resolve().parent
    patch = cfg["patch"]
    insets = cfg["layout"]["insets"]
    indent_emu = emu_from_inches(cfg["layout"]["body_indent"]) if patch["normalize_indents"] else None
    code_font = cfg["font"]["mono"]
    code_size = cfg["font"]["code_size"]

    for slide in deck.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                if patch["autofit"]:
                    text_frame.word_wrap = True
                    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                if patch["apply_insets"]:
                    bodyPr = shape.text_frame._txBody.find(qn("a:bodyPr"))
                    if bodyPr is not None:
                        bodyPr.set("lIns", str(emu_from_inches(insets["left"])))
                        bodyPr.set("rIns", str(emu_from_inches(insets["right"])))
                        bodyPr.set("tIns", str(emu_from_inches(insets["top"])))
                        bodyPr.set("bIns", str(emu_from_inches(insets["bottom"])))

                if patch["normalize_indents"]:
                    for paragraph in text_frame.paragraphs:
                        pPr = paragraph._p.get_or_add_pPr()
                        lvl = int(pPr.get("lvl") or 0)
                        pPr.set("marL", str(indent_emu * (lvl + 1)))
                        pPr.set("indent", str(-indent_emu))

                if patch["code_font"]:
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.font.name == "Courier":
                                run.font.name = code_font
                                if code_size:
                                    run.font.size = Pt(code_size)

            if shape.has_table and cfg["table"]["style"] != "default":
                _patch_table_style(shape.table._tbl, cfg["table"]["style"])

            descr = _shape_description(shape)
            if not patch["fix_math_images"] or not descr.startswith(patch["math_image_prefix"]):
                continue

            image_path = deck_dir / Path(descr)
            if not image_path.exists():
                print(f"  Skipped missing math image: {image_path}")
                continue

            width_emu, height_emu = _image_size_to_emu(image_path)
            if patch["clamp_math_to_slide"] and width_emu > 0 and height_emu > 0:
                scale = min(1.0, shape.width / width_emu, shape.height / height_emu)
                if scale < 1.0:
                    width_emu = int(width_emu * scale)
                    height_emu = int(height_emu * scale)
            shape.left += (shape.width - width_emu) // 2
            shape.top += (shape.height - height_emu) // 2
            shape.width = width_emu
            shape.height = height_emu

    deck.save(deck_path)
    print(f"Patched {deck_path} (autofit={patch['autofit']}, "
          f"math images={patch['fix_math_images']}, code font={patch['code_font']}, "
          f"tables={cfg['table']['style']}).")


def _patch_table_style(tbl, style):
    """Point a table at a simple style and disable banding/header emphasis.
    tblPr is the first child of <a:tbl>; it carries the firstRow/bandRow flags
    and the <a:tableStyleId> child."""
    tblPr = tbl.find(qn("a:tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn("a:tblPr"))
    for attr in ("firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol"):
        tblPr.set(attr, "0")
    if style == "simple":
        style_id = tblPr.find(qn("a:tableStyleId"))
        if style_id is None:
            style_id = etree.SubElement(tblPr, qn("a:tableStyleId"))
        style_id.text = TABLE_STYLE_SIMPLE_GUID


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------


def build_template(cfg, no_regen=False, input_path="template.pptx", output_path="template.pptx"):
    if not no_regen:
        subprocess.run(
            ["quarto", "pandoc", "-o", input_path, "--print-default-data-file", "reference.pptx"],
            check=True,
        )
        print(f"Regenerated {input_path} from Quarto default.")

    prs = Presentation(input_path)
    slide_size = set_slide_size(prs, cfg)

    scale_placeholders_for_size(prs, slide_size)
    apply_geometry_overrides(prs, cfg, slide_size)
    apply_theme(prs, cfg)
    apply_fonts_to_masters(prs, cfg)
    apply_text_defaults(prs, cfg)
    apply_placeholder_fills(prs, cfg)
    apply_slide_background(prs, cfg)
    anchor_footer_placeholders_to_bottom(prs, slide_size[1], cfg["layout"]["footer_bottom_margin"])
    if cfg["layout"]["expand_body_to_footer"]:
        expand_body_placeholders_to_footer(prs, slide_size[1], cfg["layout"]["content_gap"])
    rearrange_content_with_caption_layouts(
        prs, slide_size[1],
        vertical_gap_pct=cfg["layout"]["content_gap"],
        caption_ratio=cfg["layout"]["caption_ratio"],
    )

    prs.save(output_path)

    # Install a simple table style into the template's tableStyles.xml so
    # patched decks can reference it (the Quarto default only defines the
    # blue "Medium Style 2" GUID).
    if cfg["table"]["style"] != "default":
        replace_zip_entry(output_path, "ppt/tableStyles.xml", SIMPLE_TABLE_STYLES_XML.encode("utf-8"))

    print(f"Saved styled template to {output_path} "
          f"(slide {slide_size[0] / EMU_PER_INCH:.1f}\"x{slide_size[1] / EMU_PER_INCH:.2f}\", "
          f"font '{cfg['font']['name']}', tables '{cfg['table']['style']}').")


def render_and_patch(cfg, qmd_path, no_regen=False):
    qmd_path = Path(qmd_path)
    deck_path = qmd_path.with_suffix(".pptx")
    if not no_regen:
        subprocess.run(["quarto", "render", str(qmd_path), "--to", "pptx"], check=True)
    if not deck_path.exists():
        sys.exit(f"Expected rendered deck {deck_path} not found.")
    patch_deck(cfg, deck_path)


def main():
    parser = argparse.ArgumentParser(description="Build styled PPTX template and/or patch a rendered deck.")
    parser.add_argument("--config", default="style_config.yml", help="YAML config file (default: style_config.yml)")
    parser.add_argument("--patch-deck", dest="patch_deck", metavar="DECK.pptx",
                        help="Patch an existing rendered deck and exit.")
    parser.add_argument("--render", dest="render_qmd", metavar="DECK.qmd",
                        help="Run 'quarto render' on the qmd, then patch the resulting deck.")
    parser.add_argument("--no-regen", action="store_true",
                        help="Skip regenerating template.pptx from the Quarto default.")
    args = parser.parse_args()

    cfg = load_config(args.config)

    if args.render_qmd:
        render_and_patch(cfg, args.render_qmd, no_regen=args.no_regen)
    elif args.patch_deck:
        patch_deck(cfg, args.patch_deck)
    else:
        build_template(cfg, no_regen=args.no_regen)


if __name__ == "__main__":
    main()
