# Quarto → PowerPoint pipeline

## Project Introduction

This project was developed to bridge the gap between Quarto's programmatic slide generation and the high-fidelity styling expectations of professional PowerPoint presentations. While Quarto is an exceptional tool for technical documentation, its native PPTX writer is hindered by limited styling exposure and suboptimal LaTeX math rendering. To overcome these barriers, I implemented a sophisticated pipeline that replaces the need for manual MS PowerPoint intervention.

By utilizing a custom Python-based styling engine, the project programmatically regenerates the `template.pptx` from Quarto's defaults. This ensures that master slides are perfectly aligned and styled—eliminating the reliance on external office suites like LibreOffice or Google Docs, which often produce incompatible or "strange" master slide structures that conflict with Pandoc's expectations. The workflow further integrates a specialized Lua filter to convert complex math into high-quality SVG/PNG images and employs a post-render "patching" phase. This critical final step corrects layout anomalies that only emerge after the render process, such as forcing text word-wrap, implementing "shrink-to-fit" logic for overflowing content, and precisely clamping math images to their containers to ensure they never bleed off the slide.

## Why this project was created (Problem vs. Solution)

**1. Limitations in Template Styling**


- **The Problem:** Pandoc’s PPTX writer only exposes a few basic options (like theme and aspect ratio), making it nearly impossible to achieve a truly "branded" or professional look without manual editing.

- **The Solution:** A Python script that uses `python-pptx` to programmatically build a customized `template.pptx` from the ground up, allowing for precise control over fonts, colors, and placeholder geometry.


**2. Poor Math Rendering in PPTX**


- **The Problem:** Native LaTeX math rendering in PowerPoint via Pandoc is often poorly formatted or fails to render complex equations correctly.

- **The Solution:** A custom Lua filter that intercepts math blocks, rendering inline math as clean Unicode and display math as high-resolution SVG/PNG images via XeLaTeX/pdfLaTeX.

**3. The "Master Slide" Compatibility Issue**

- **The Problem:** Creating a reference template in non-Microsoft software (LibreOffice, Google Docs) often results in a corrupted Master Slide structure that Pandoc cannot interpret correctly, leading to broken layouts.


- **The Solution:** By programmatically generating the template from the official Quarto default reference doc, the project ensures 100% compatibility and removes the dependency on having MS PowerPoint installed just to create a template.


**4. Post-Render Layout Failures (Overflow)**


- **The Problem:** Quarto/Pandoc often ignores text boundaries, resulting in text that overflows the slide or images that are scaled incorrectly.


- **The Solution:** A "Patching" phase that runs after the render to:


  - Force **Word-Wrap** and **Shrink-to-Fit** on text boxes.


  - **Clamp and Resize** math images to their true physical size so they stay within their designated boxes.


**5. Rigid Default Formatting**

- **The Problem:** Pandoc hardcodes certain elements, such as the `Courier` font for code blocks, regardless of the template settings.

- **The Solution:** The patching script scans the final XML of the deck to swap hardcoded fonts with a user-configured monospace font and applies a "Simple Grid" style to tables to remove distracting banding.

## Project Overview
A personal workflow for building styled PowerPoint decks from Quarto, working around two pandoc/Quarto limitations:

1.  **Template styling** — pandoc's PPTX writer only exposes a few options (`theme`, `aspectratio`), so a `template.pptx` reference doc is built with `python-pptx` from Quarto's default and heavily customized.
2.  **Math rendering** — pandoc's PPTX output renders LaTeX math poorly, so a Lua filter converts inline math to Unicode text and display math to SVG/PNG images.

## Pipeline

```         
uv run style_template.py              ──► builds template.pptx (from Quarto default + style_config.yml)
quarto render main.qmd                ──► renders deck using template.pptx + math-to-png.lua
uv run style_template.py --render main.qmd   ──► does BOTH steps, then patches the deck
```

The final patch step fixes things that only exist in the rendered deck:

- forces text word-wrap + shrink-to-fit on overflow (the main overflow fix),
- resizes display-math images to their true physical size and clamps them to their box so they never overflow the slide,
- swaps pandoc's hardcoded `Courier` code font for the configured monospace font and size,
- restyles tables (simple grid, no banding) via a style defined in the template.

## Commands

``` bash
uv sync                                   # first time: install python-pptx + PyYAML

uv run style_template.py                  # (re)build template.pptx from Quarto default
uv run style_template.py --config custom.yml
uv run style_template.py --patch-deck main.pptx          # patch an existing deck
uv run style_template.py --render main.qmd               # quarto render + patch
uv run style_template.py --no-regen --patch-deck main.pptx  # skip template rebuild
```

## Configuration

### `style_config.yml` (template + deck patching)

Everything the Python script does is driven from this file — fonts, sizes, colors, aspect ratio, placeholder geometry, spacing, footer handling, and the deck-patch behavior. Every key is optional; see the comments in the file for the full list. Highlights:

| Section | Example knobs |
|-----------------------------|-------------------------------------------|
| `slide` | `ratio: "16:9" | "16:10" | "4:3"` or custom `width`/`height` in inches |
| `font` | `name`, `title_name`, `mono`, `title_size`, `subtitle_size`, `body_size`, `code_size`, optional per-level `level_sizes` |
| `colors` | `title_text`, `title_fill` (highlight bar), `body_text`, `background`, theme `accent1..6`, `link` |
| `layout` | title/subtitle/body alignment, bold, vertical anchor, text insets, line spacing, space-after, bullet indent, footer margins, `geometry_overrides` |
| `table` | `style: "simple"` (white cells, thin grid, no banding) or `"default"` (pandoc's blue table) |
| `patch` | autofit, insets, indent normalization, code font, math-image resize/clamp, table style |

Because the template's placeholder text defaults (`lstStyle`) are what pandoc's empty `<a:rPr/>` runs inherit, size/color/bold changes made in the template flow through to every slide automatically.

### `math-config` (in the qmd front matter or `_quarto.yml`)

The Lua filter reads its settings from document metadata:

``` yaml
math-config:
  format: svg          # "svg" (xelatex + dvisvgm) | "png" (pdflatex + ghostscript)
  font-size: 20        # display math typeset size (pt)
  png-dpi: 300
  use-cache: true      # reuse rendered images; disable to force regeneration
  inline-style: italic # "italic" (default) | "normal"
  svg-fonts: false     # true = embed fonts (selectable text); false = paths
```

Rendered images are cached in `math-png-cache/`, so rebuilds are fast.

> **Known limitation:** pandoc's PPTX writer starts a new slide after the first image in a slide, so a display equation must be the last block of a slide — keep to one display equation per slide.

## How the pieces work

- **`style_template.py`** regenerates `template.pptx` from `quarto pandoc --print-default-data-file reference.pptx`, then restyles it. Geometry is scaled from the reference doc's 16:9 layout to your target aspect ratio (that's the part that was previously broken), and optional `geometry_overrides` can reposition title/body placeholders by percent. Footer placeholders are anchored to the bottom, body placeholders are expanded to the footer row, and the "Content with Caption" layout is rearranged to content-on-top/caption-below.
- **`math-to-png.lua`** converts inline math via pandoc's plain-text writer (real Unicode + sub/superscript runs, italic by default) and display math via LaTeX → SVG (or PNG). It never touches non-PPTX formats.
- **`main.qmd`** is a working example: it wires in `template.pptx` and the filter, and demonstrates inline math, display math, and `\begin{aligned}`.

## Requirements

- Quarto ≥ 1.4 (pandoc ≥ 3), `uv`
- SVG math: `xelatex` + `dvisvgm` · PNG math: `pdflatex` + Ghostscript
- `python-pptx` + `PyYAML` (installed by `uv sync`)