#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = ["lxml", "cssselect", "tinycss2"]
# ///
"""
mmd_svg_fix.py - make a mermaid-cli (mmdc) SVG open correctly in Inkscape.

WHAT IT FIXES
-------------
mmdc SVGs render fine in a browser but break in Inkscape for four reasons.
This script applies one fix for each, in this order:

  1. CSS INLINING        (black boxes / missing colors)
     Mermaid puts all colors in one <style> block with class selectors.
     Inkscape's CSS parser chokes on parts of it (@keyframes, !important,
     custom properties, rgba() ...) and falls back to SVG's default fill:
     black. We resolve every rule against the document with a real selector
     engine, write the result into each element's style="..." attribute and
     delete the <style> block.

  2. foreignObject -> <text>   (missing text)
     Labels that are still HTML inside <foreignObject> are invisible in
     Inkscape. They are converted to plain, editable SVG <text>. Only runs
     if any foreignObject is left (with htmlLabels: false there normally
     are none). Automatic line wrapping of HTML labels cannot be
     reproduced; use <br/> in the Mermaid source for forced line breaks.

  3. NODE / EDGE LABEL CENTERING   (labels not vertically centered)
     Mermaid positions label rows with em offsets measured in Chromium;
     Inkscape lays them out differently. Rows are re-placed at explicit
     pixel positions, centered on the node (or edge-label) center.

  4. SUBGRAPH TITLE PLACEMENT  (tiny / off-center subgraph titles)
     Titles are re-placed relative to the subgraph rectangle: aligned
     left / center / right and pinned near the top edge.

USAGE
-----
  uv run mmd_svg_fix.py diagram.svg                  # -> diagram.fixed.svg
  uv run mmd_svg_fix.py diagram.svg -o final.svg
  uv run mmd_svg_fix.py diagram.svg --cluster-align left --cluster-font-size 14

  (uv installs lxml, cssselect and tinycss2 automatically. Without uv:
   pip install lxml cssselect tinycss2 && python mmd_svg_fix.py ...)

TUNING
------
Every tunable is both a command-line option (see --help) and a default in
the DEFAULTS block below, so you can either pass flags or edit the file.
Quick guide:

  Text looks too high / too low in the box   -> --baseline (0.30-0.40) or --label-dy
  Rows too tight / too loose                 -> --line-height (Mermaid uses 1.1)
  Text slightly left / right                 -> --label-dx
  Text size wrong                            -> --font-size (default: read from SVG)
  Subgraph title too small / large           -> --cluster-font-size
  Subgraph title too close to / far from top -> --cluster-pad-top
  Subgraph title position                    -> --cluster-align left|center|right

Steps can be switched off individually with --skip-inline-css,
--skip-foreign-objects, --skip-node-labels and --skip-cluster-labels, which
helps to find out which step causes a given artifact.

NOTES
-----
* Install the font that the diagram uses (e.g. Noto Sans) or Inkscape will
  substitute one with different widths and text will not fit the boxes.
* The script is idempotent enough to re-run on its own output, but run it
  on the original mmdc export when you change tuning values.
* Transparent backgrounds show the Inkscape canvas color. Use
  `mmdc -b white` if you want a white background baked in.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, fields
from pathlib import Path
from typing import Optional

import cssselect
import tinycss2
from lxml import etree

# =============================================================================
# DEFAULTS - edit these or override them with command-line flags
# =============================================================================


@dataclass
class Config:
    # ---- node and edge labels (step 3) --------------------------------------
    font_size: Optional[float] = None   # px. None = read from the SVG (fallback 16)
    line_height: float = 1.1            # row spacing in em (Mermaid's own value)
    baseline: float = 0.35              # em. Shifts the text block down (+) / up (-)
    label_dx: float = 0.0               # px nudge applied to every node/edge label
    label_dy: float = 0.0               # px nudge applied to every node/edge label

    # ---- subgraph titles (step 4) -------------------------------------------
    cluster_font_size: Optional[float] = None  # px. None = same as font_size
    cluster_align: str = "center"       # "left" | "center" | "right"
    cluster_pad_top: float = 6.0        # px between the subgraph's top edge and title
    cluster_pad_x: float = 8.0          # px from the side edge (left / right align)

    # ---- foreignObject conversion (step 2) ----------------------------------
    fo_font_family: Optional[str] = None  # None = inherit from the SVG
    fo_fill: Optional[str] = None         # None = read from the SVG (fallback #333333)

    # ---- which steps run ----------------------------------------------------
    inline_css: bool = True
    foreign_objects: bool = True
    node_labels: bool = True
    cluster_labels: bool = True


# CSS properties that are meaningless in a static drawing; dropped when inlining
SKIP_PROPERTIES = {"animation", "cursor", "pointer-events"}
# CSS values Inkscape does not understand; the declaration is dropped
SKIP_VALUES = {"revert", "revert-layer"}

SVG_NS_URI = "http://www.w3.org/2000/svg"
SVG_NS_ATTR = f' xmlns="{SVG_NS_URI}"'
TRANSLATE_RE = re.compile(r"translate\(\s*([-+\d.eE]+)(?:[\s,]+([-+\d.eE]+))?\s*\)")
HTML_BLOCK_TAGS = {"p", "div", "li"}

# =============================================================================
# Small helpers
# =============================================================================


def load_svg(path: Path) -> etree._Element:
    """Parse the SVG. The default namespace is stripped so that plain tag
    names and CSS selectors match; dump_svg() puts it back."""
    raw = path.read_text(encoding="utf-8")
    raw = re.sub(r"^\s*<\?xml[^>]*\?>", "", raw)
    raw = re.sub(r'\sxmlns="' + re.escape(SVG_NS_URI) + '"', "", raw, count=1)
    return etree.fromstring(raw.encode("utf-8"))


def dump_svg(root: etree._Element, path: Path) -> None:
    xml = etree.tostring(root, encoding="unicode")
    xml = xml.replace("<svg", "<svg" + SVG_NS_ATTR, 1)
    path.write_text('<?xml version="1.0" encoding="UTF-8"?>\n' + xml, encoding="utf-8")


def parent_map(root: etree._Element) -> dict:
    return {child: parent for parent in root.iter() for child in parent}


def get_style(el: etree._Element) -> dict:
    props = {}
    for part in (el.get("style") or "").split(";"):
        if ":" in part:
            key, _, value = part.partition(":")
            props[key.strip().lower()] = value.strip()
    return props


def set_style(el: etree._Element, **props) -> None:
    """Merge properties into the style attribute (underscores become dashes)."""
    style = get_style(el)
    for key, value in props.items():
        style[key.replace("_", "-")] = str(value)
    el.set("style", "; ".join(f"{k}: {v}" for k, v in style.items()))


def translate_of(el: etree._Element) -> tuple[float, float]:
    m = TRANSLATE_RE.search(el.get("transform") or "")
    if not m:
        return 0.0, 0.0
    return float(m.group(1)), float(m.group(2) or 0.0)


def localname(el: etree._Element) -> Optional[str]:
    return etree.QName(el).localname if isinstance(el.tag, str) else None


def has_class(el: etree._Element, name: str) -> bool:
    return name in (el.get("class") or "").split()


def detect_font_size(root: etree._Element) -> float:
    m = re.search(r"font-size:\s*([\d.]+)px", root.get("style") or "")
    if not m:  # CSS not inlined (yet): look into the <style> block
        css = "".join("".join(s.itertext()) for s in root.iter("style"))
        m = re.search(r"font-size:\s*([\d.]+)px", css)
    return float(m.group(1)) if m else 16.0


def detect_fill(root: etree._Element) -> str:
    fill = get_style(root).get("fill")
    if not fill:
        css = "".join("".join(s.itertext()) for s in root.iter("style"))
        m = re.search(r"#[\w-]+\{[^}]*?[;{]fill:\s*(#[0-9A-Fa-f]{3,8})", css)
        fill = m.group(1) if m else None
    return fill or "#333333"


def place_text(text: etree._Element, rows: list, cx: float, first_y: float,
               line_px: float, size: float, anchor: str) -> None:
    """Give a <text> explicit coordinates: one pixel x/y per row, no em offsets."""
    text.attrib.pop("x", None)
    text.attrib.pop("y", None)
    text.set("text-anchor", anchor)
    text.set("font-size", f"{size:g}px")
    # the inlined style attribute would otherwise win over the attributes above
    set_style(text, text_anchor=anchor, font_size=f"{size:g}px")
    if rows:
        for i, row in enumerate(rows):
            row.set("x", f"{cx:.2f}")
            row.set("y", f"{first_y + i * line_px:.2f}")
            row.attrib.pop("dy", None)
    else:
        text.set("x", f"{cx:.2f}")
        text.set("y", f"{first_y:.2f}")


def text_rows(text: etree._Element) -> list:
    """The row <tspan>s of a Mermaid label (direct children of <text>)."""
    return [c for c in text if localname(c) == "tspan"]


# =============================================================================
# Step 1 - inline the CSS
# =============================================================================


def inline_css(root: etree._Element) -> int:
    css = "\n".join("".join(s.itertext()) for s in root.iter("style"))
    if not css.strip():
        return 0

    translator = cssselect.GenericTranslator()
    matches = []  # (important, specificity, rule order, element, property, value)

    for order, rule in enumerate(
        tinycss2.parse_stylesheet(css, skip_comments=True, skip_whitespace=True)
    ):
        if rule.type != "qualified-rule":  # skips @keyframes, @media ...
            continue
        decls = []
        for d in tinycss2.parse_declaration_list(
            rule.content, skip_comments=True, skip_whitespace=True
        ):
            if d.type != "declaration" or d.lower_name.startswith("--"):
                continue
            value = tinycss2.serialize(d.value).strip()
            if d.lower_name in SKIP_PROPERTIES or value in SKIP_VALUES:
                continue
            decls.append((d.lower_name, value, d.important))
        if not decls:
            continue
        try:
            selectors = cssselect.parse(tinycss2.serialize(rule.prelude))
        except Exception:
            continue
        for sel in selectors:
            try:
                elements = root.xpath(translator.selector_to_xpath(sel))
            except Exception:
                continue
            for el in elements:
                if isinstance(el, etree._Element):
                    for name, value, important in decls:
                        matches.append(
                            (important, sel.specificity(), order, el, name, value)
                        )

    # cascade: normal before !important, low specificity before high, then source order
    matches.sort(key=lambda m: (m[0], m[1], m[2]))
    styles: dict = {}
    for _imp, _spec, _order, el, name, value in matches:
        styles.setdefault(el, {})[name] = value

    for el, props in styles.items():
        props.update(get_style(el))  # an existing inline style beats the stylesheet
        el.set("style", "; ".join(f"{k}: {v}" for k, v in props.items()))

    for style_el in list(root.iter("style")):
        style_el.getparent().remove(style_el)
    return len(styles)


# =============================================================================
# Step 2 - foreignObject -> <text>
# =============================================================================


def html_lines(fo: etree._Element) -> list:
    """Flatten the HTML inside a foreignObject into a list of text lines."""
    lines = [""]

    def walk(el):
        if el.text:
            lines[-1] += el.text
        for child in el:
            name = localname(child)
            if name is None:  # comment / processing instruction
                if child.tail:
                    lines[-1] += child.tail
                continue
            if name == "br":
                lines.append("")
            elif name in HTML_BLOCK_TAGS and lines[-1].strip():
                lines.append("")
            walk(child)
            if name in HTML_BLOCK_TAGS:
                lines.append("")
            if child.tail:
                lines[-1] += child.tail

    walk(fo)
    return [" ".join(line.split()) for line in lines if line.strip()]


def convert_foreign_objects(root: etree._Element, cfg: Config, size: float) -> int:
    parents = parent_map(root)
    fill = cfg.fo_fill or detect_fill(root)
    count = 0
    for fo in list(root.iter("foreignObject")):
        parent = parents[fo]
        index = list(parent).index(fo)
        lines = html_lines(fo)
        w = float(fo.get("width") or 0)
        h = float(fo.get("height") or 0)
        x = float(fo.get("x") or 0)
        y = float(fo.get("y") or 0)
        parent.remove(fo)
        if not lines:
            continue

        text = etree.Element("text", {"fill": fill})
        if cfg.fo_font_family:
            text.set("font-family", cfg.fo_font_family)
        for line in lines:
            row = etree.SubElement(text, "tspan")
            row.text = line
        # provisional position (box center); steps 3 and 4 refine it
        n = len(lines)
        lh = cfg.line_height * size
        first = y + h / 2 - (n - 1) * lh / 2 + cfg.baseline * size
        place_text(text, text_rows(text), x + w / 2, first, lh, size, "middle")
        parent.insert(index, text)
        count += 1
    return count


# =============================================================================
# Step 3 - center node and edge labels
# =============================================================================


def center_node_labels(root: etree._Element, cfg: Config, size: float) -> int:
    parents = parent_map(root)
    line_px = cfg.line_height * size
    count = 0
    for text in list(root.iter("text")):
        # climb to the node / edgeLabel group, summing the translates below it:
        # that group's origin is the label's center
        sx = sy = 0.0
        el = parents.get(text)
        found = False
        while el is not None:
            if has_class(el, "node") or has_class(el, "edgeLabel"):
                found = True
                break
            tx, ty = translate_of(el)
            sx += tx
            sy += ty
            el = parents.get(el)
        if not found:  # subgraph titles etc. are handled in step 4
            continue

        rows = text_rows(text)
        n = max(len(rows), 1)
        cx = -sx + cfg.label_dx
        first = -sy - (n - 1) * line_px / 2 + cfg.baseline * size + cfg.label_dy
        place_text(text, rows, cx, first, line_px, size, "middle")
        count += 1
    return count


# =============================================================================
# Step 4 - subgraph (cluster) titles
# =============================================================================


def place_cluster_labels(root: etree._Element, cfg: Config, size: float) -> int:
    parents = parent_map(root)
    line_px = cfg.line_height * size
    anchor = {"left": "start", "center": "middle", "right": "end"}[cfg.cluster_align]
    count = 0

    for cluster in root.iter("g"):
        if not has_class(cluster, "cluster"):
            continue
        rect = next((c for c in cluster if localname(c) == "rect"), None)
        if rect is None:
            continue
        rx = float(rect.get("x", 0))
        ry = float(rect.get("y", 0))
        rw = float(rect.get("width", 0))

        if cfg.cluster_align == "left":
            box_x = rx + cfg.cluster_pad_x
        elif cfg.cluster_align == "right":
            box_x = rx + rw - cfg.cluster_pad_x
        else:
            box_x = rx + rw / 2

        for text in cluster.iter("text"):
            # translates between the text and the cluster group convert the
            # rectangle's coordinates into the text's local coordinates
            sx = sy = 0.0
            el = parents.get(text)
            while el is not None and el is not cluster:
                tx, ty = translate_of(el)
                sx += tx
                sy += ty
                el = parents.get(el)

            rows = text_rows(text)
            first = ry + cfg.cluster_pad_top + size / 2 + cfg.baseline * size - sy
            place_text(text, rows, box_x - sx, first, line_px, size, anchor)
            count += 1
    return count


# =============================================================================
# Command line
# =============================================================================


def build_parser() -> argparse.ArgumentParser:
    d = Config()
    p = argparse.ArgumentParser(
        description="Fix a mermaid-cli SVG export so it opens correctly in Inkscape.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("input", type=Path, help="SVG exported by mmdc")
    p.add_argument("-o", "--output", type=Path,
                   help="output file (default: <input>.fixed.svg)")

    g = p.add_argument_group("node and edge labels")
    g.add_argument("--font-size", type=float, default=d.font_size, metavar="PX",
                   help="label font size; default: read from the SVG")
    g.add_argument("--line-height", type=float, default=d.line_height, metavar="EM",
                   help="row spacing")
    g.add_argument("--baseline", type=float, default=d.baseline, metavar="EM",
                   help="vertical shift of text blocks; raise to move text down")
    g.add_argument("--label-dx", type=float, default=d.label_dx, metavar="PX",
                   help="extra horizontal nudge for all node/edge labels")
    g.add_argument("--label-dy", type=float, default=d.label_dy, metavar="PX",
                   help="extra vertical nudge for all node/edge labels")

    g = p.add_argument_group("subgraph titles")
    g.add_argument("--cluster-font-size", type=float, default=d.cluster_font_size,
                   metavar="PX", help="title font size; default: same as --font-size")
    g.add_argument("--cluster-align", choices=["left", "center", "right"],
                   default=d.cluster_align, help="horizontal title position")
    g.add_argument("--cluster-pad-top", type=float, default=d.cluster_pad_top,
                   metavar="PX", help="gap between the top edge and the title")
    g.add_argument("--cluster-pad-x", type=float, default=d.cluster_pad_x,
                   metavar="PX", help="gap to the side edge (left/right align)")

    g = p.add_argument_group("foreignObject conversion")
    g.add_argument("--fo-font-family", default=d.fo_font_family,
                   help="font for converted text; default: inherit")
    g.add_argument("--fo-fill", default=d.fo_fill,
                   help="text color for converted text; default: read from the SVG")

    g = p.add_argument_group("steps (for troubleshooting)")
    g.add_argument("--skip-inline-css", action="store_true")
    g.add_argument("--skip-foreign-objects", action="store_true")
    g.add_argument("--skip-node-labels", action="store_true")
    g.add_argument("--skip-cluster-labels", action="store_true")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    cfg = Config(**{f.name: getattr(args, f.name) for f in fields(Config)
                    if hasattr(args, f.name)})
    cfg.inline_css = not args.skip_inline_css
    cfg.foreign_objects = not args.skip_foreign_objects
    cfg.node_labels = not args.skip_node_labels
    cfg.cluster_labels = not args.skip_cluster_labels

    src: Path = args.input
    dst: Path = args.output or src.with_name(src.stem + ".fixed.svg")
    if not src.is_file():
        print(f"error: {src} not found", file=sys.stderr)
        return 1

    root = load_svg(src)

    # sizes are read before the CSS is inlined and removed
    size = cfg.font_size or detect_font_size(root)
    cluster_size = cfg.cluster_font_size or size

    if cfg.inline_css:
        n = inline_css(root)
        print(f"[1] CSS inlined onto {n} elements, <style> removed")
    if cfg.foreign_objects:
        n = convert_foreign_objects(root, cfg, size)
        print(f"[2] foreignObject converted to <text>: {n}")
    if cfg.node_labels:
        n = center_node_labels(root, cfg, size)
        print(f"[3] node/edge labels centered: {n} (font {size:g}px)")
    if cfg.cluster_labels:
        n = place_cluster_labels(root, cfg, cluster_size)
        print(f"[4] subgraph titles placed: {n} (font {cluster_size:g}px, "
              f"{cfg.cluster_align})")

    dump_svg(root, dst)
    print(f"written: {dst}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
