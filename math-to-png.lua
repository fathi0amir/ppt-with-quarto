--[[
math-to-png.lua

Fixes LaTeX math rendering in Quarto PowerPoint (pptx) output:
  - Inline math ($...$) is converted into Unicode text via Pandoc's native
    plain-text writer (optionally italicized to look like math).
  - Display math ($$...$$) is rendered to SVG (xelatex + dvisvgm) or PNG
    (pdflatex + Ghostscript) and embedded as an image.

Configuration via document metadata (qmd front matter or _quarto.yml):

    math-config:
      format: svg            # "svg" (default) or "png"
      font-size: 20          # pt, display math typeset size
      png-dpi: 300           # raster resolution for PNG output
      use-cache: true        # reuse rendered images across renders
      inline-style: italic   # "italic" (default) or "normal"
      svg-fonts: false       # true = embed font references (text selectable,
                             #         needs fonts installed); false = paths

Set `format: png` if dvisvgm/xelatex are not installed; requires
pdflatex + Ghostscript instead.

Usage in .qmd or _quarto.yml:
    format:
      pptx:
        filters:
          - math-to-png.lua
]]

local outdir = "math-png-cache"
local cache = {}
local dir_ready = false
local is_windows = package.config:sub(1, 1) == "\\"

local CONFIG = {
  format = "svg",
  font_size = 20,
  png_dpi = 300,
  use_cache = true,
  inline_style = "italic",
  svg_fonts = false,
}

--------------------------------------------------------------------------
-- configuration from document metadata
--------------------------------------------------------------------------

local function load_config()
  if not (quarto and quarto.doc and quarto.doc.meta) then return end
  local m = quarto.doc.meta["math-config"]
  if type(m) ~= "table" then return end

  local function get(k1, k2)
    local v = m[k1]
    if v == nil then v = m[k2] end
    return v
  end
  local function as_num(v)
    if type(v) == "number" then return v end
    if type(v) == "string" then return tonumber(v) end
    return nil
  end
  local function as_bool(v, default)
    if type(v) == "boolean" then return v end
    if type(v) == "string" then
      if v == "true" or v == "yes" then return true end
      if v == "false" or v == "no" then return false end
    end
    return default
  end

  local fmt = get("format", "format")
  if fmt == "png" or fmt == "svg" then CONFIG.format = fmt end
  local fs = as_num(get("font-size", "font_size"))
  if fs then CONFIG.font_size = fs end
  local dpi = as_num(get("png-dpi", "png_dpi"))
  if dpi then CONFIG.png_dpi = dpi end
  CONFIG.use_cache = as_bool(get("use-cache", "use_cache"), CONFIG.use_cache)
  local style = get("inline-style", "inline_style")
  if style == "italic" or style == "normal" then CONFIG.inline_style = style end
  CONFIG.svg_fonts = as_bool(get("svg-fonts", "svg_fonts"), CONFIG.svg_fonts)
end

--------------------------------------------------------------------------
-- filesystem & shell helpers
--------------------------------------------------------------------------

local function sh(cmd)
  local ok = os.execute(cmd)
  return ok == true or ok == 0
end

local function ensure_dir(dir)
  if dir_ready then return end
  if is_windows then
    os.execute('if not exist "' .. dir .. '" mkdir "' .. dir .. '"')
  else
    os.execute('mkdir -p "' .. dir .. '"')
  end
  dir_ready = true
end

local function hash(s)
  local h = 5381
  for i = 1, #s do
    h = (h * 33 + string.byte(s, i)) % 4294967296
  end
  return string.format("%x", h)
end

-- Reads image dimensions directly from the PNG binary header.
local function read_png_size_pt(pngfile, dpi)
  local f = io.open(pngfile, "rb")
  if not f then return nil, nil end
  local data = f:read(24)
  f:close()
  if not data or #data < 24 or data:sub(1, 4) ~= "\137PNG" then return nil, nil end

  local w1, w2, w3, w4 = data:byte(17, 20)
  local h1, h2, h3, h4 = data:byte(21, 24)
  local width_px = w1 * 16777216 + w2 * 65536 + w3 * 256 + w4
  local height_px = h1 * 16777216 + h2 * 65536 + h3 * 256 + h4

  return string.format("%.2f", (width_px / dpi) * 72),
         string.format("%.2f", (height_px / dpi) * 72)
end

local function read_svg_size_pt(svgfile)
  local f = io.open(svgfile, "r")
  if not f then return nil, nil end
  local data = f:read("*a")
  f:close()
  if not data then return nil, nil end

  local width = data:match("width=['\"]([0-9.]+)pt['\"]")
  local height = data:match("height=['\"]([0-9.]+)pt['\"]")
  if not width or not height then return nil, nil end
  return width, height
end

local function read_image_size_pt(imagefile)
  if CONFIG.format == "png" then
    return read_png_size_pt(imagefile, CONFIG.png_dpi)
  end
  return read_svg_size_pt(imagefile)
end

local function cleanup_temp_files(base)
  os.remove(base .. ".tex")
  os.remove(base .. ".pdf")
  os.remove(base .. ".xdv")
  os.remove(base .. ".aux")
  os.remove(base .. ".log")
  os.remove(base .. "-stdout.log")
  os.remove(base .. "-gs.log")
  os.remove(base .. "-dvisvgm.log")
end

--------------------------------------------------------------------------
-- display math -> SVG/PNG (with persistent disk caching)
--------------------------------------------------------------------------

local function find_gs()
  ensure_dir(outdir)
  for _, candidate in ipairs({ "gswin64c", "gswin32c", "gs" }) do
    if sh(candidate .. ' -v > "' .. outdir .. '/gs-check.log" 2>&1') then
      return candidate
    end
  end
  return nil
end

local function find_dvisvgm()
  ensure_dir(outdir)
  if sh('dvisvgm --version > "' .. outdir .. '/dvisvgm-check.log" 2>&1') then
    return "dvisvgm"
  end
  return nil
end

local function tex_to_image(tex)
  local key = "d-" .. hash(tex)
  if CONFIG.use_cache and cache[key] then return cache[key] end

  ensure_dir(outdir)
  local base = outdir .. "/eq-" .. key
  local ext = CONFIG.format == "svg" and ".svg" or ".png"
  local imagepath = base .. ext

  -- DISK CACHE: reuse the rendered image without re-running LaTeX.
  if CONFIG.use_cache then
    local w_cached, h_cached = read_image_size_pt(imagepath)
    if w_cached and h_cached then
      cache[key] = { path = imagepath, width = w_cached, height = h_cached }
      return cache[key]
    end
  end

  local texfile = base .. ".tex"
  local f = io.open(texfile, "w")
  if not f then
    io.stderr:write("math-to-png.lua: could not create TeX file: " .. texfile .. "\n")
    return nil
  end

  f:write("\\documentclass[preview,border=1pt,varwidth]{standalone}\n")
  f:write("\\usepackage{amsmath,amssymb,amsfonts}\n")
  f:write("\\AtBeginDocument{\\fontsize{" .. CONFIG.font_size .. "}{"
          .. (CONFIG.font_size * 1.2) .. "}\\selectfont}\n")
  f:write("\\begin{document}\n")
  -- Environments like \begin{aligned} must live INSIDE math mode, so only
  -- skip the \[ \] wrapper for genuinely standalone display environments.
  local standalone = { equation = true, ["equation*"] = true, align = true,
                       ["align*"] = true, alignat = true, ["alignat*"] = true,
                       gather = true, ["gather*"] = true, multline = true,
                       ["multline*"] = true, displaymath = true }
  local env = tex:match("^%s*\\begin{([^}]+)}")
  if env and standalone[env] then
    f:write(tex .. "\n")
  else
    f:write("\\[" .. tex .. "\\]\n")
  end
  f:write("\\end{document}\n")
  f:close()

  if CONFIG.format == "png" then
    local pdflatex_ok = sh(string.format(
      'pdflatex -interaction=nonstopmode -halt-on-error -output-directory="%s" "%s" > "%s-stdout.log" 2>&1',
      outdir, texfile, base))
    if not pdflatex_ok then
      io.stderr:write("math-to-png.lua: pdflatex failed for '" .. tex .. "', see " .. base .. ".log\n")
      return nil
    end
    local gs = find_gs()
    if not gs then
      io.stderr:write("math-to-png.lua: Ghostscript not found on PATH (required for PNG output).\n")
      return nil
    end
    local png_ok = sh(string.format(
      '%s -q -dBATCH -dNOPAUSE -dSAFER -sDEVICE=pngalpha -r%d -dGraphicsAlphaBits=4 -dTextAlphaBits=4 -sOutputFile="%s" "%s.pdf" > "%s-gs.log" 2>&1',
      gs, CONFIG.png_dpi, imagepath, base, base))
    if not png_ok then
      io.stderr:write("math-to-png.lua: Ghostscript failed for '" .. tex .. "', see " .. base .. "-gs.log\n")
      return nil
    end
  else
    local xelatex_ok = sh(string.format(
      'xelatex -no-pdf -interaction=nonstopmode -halt-on-error -output-directory="%s" "%s" > "%s-stdout.log" 2>&1',
      outdir, texfile, base))
    if not xelatex_ok then
      io.stderr:write("math-to-png.lua: xelatex failed for '" .. tex .. "', see " .. base .. ".log\n")
      return nil
    end
    local dvisvgm = find_dvisvgm()
    if not dvisvgm then
      io.stderr:write("math-to-png.lua: dvisvgm not found on PATH (required for SVG output).\n")
      return nil
    end
    local no_fonts = CONFIG.svg_fonts and "" or "--no-fonts "
    local svg_ok = sh(string.format(
      '%s %s--exact-bbox --bbox=min -o "%s" "%s.xdv" > "%s-dvisvgm.log" 2>&1',
      dvisvgm, no_fonts, imagepath, base, base))
    if not svg_ok then
      io.stderr:write("math-to-png.lua: dvisvgm failed for '" .. tex .. "', see " .. base .. "-dvisvgm.log\n")
      return nil
    end
  end

  local w_pt, h_pt = read_image_size_pt(imagepath)
  cleanup_temp_files(base)

  cache[key] = { path = imagepath, width = w_pt, height = h_pt }
  return cache[key]
end

--------------------------------------------------------------------------
-- inline math -> Unicode text with real sub/superscript runs
--------------------------------------------------------------------------

local function text_to_inlines(s)
  local inlines = pandoc.Inlines({})
  local first = true
  for word in s:gmatch("%S+") do
    if not first then inlines:insert(pandoc.Space()) end
    inlines:insert(pandoc.Str(word))
    first = false
  end
  return inlines
end

local function find_closing_paren(str, start_pos)
  local depth = 1
  for i = start_pos + 1, #str do
    local ch = str:sub(i, i)
    if ch == "(" then
      depth = depth + 1
    elseif ch == ")" then
      depth = depth - 1
      if depth == 0 then
        return i
      end
    end
  end
  return nil
end

-- Converts the plain-text writer output (which uses "_" and "^" for
-- sub/superscripts) into real Subscript/Superscript inlines.
local function parse_script_string(str)
  local inlines = pandoc.Inlines({})
  local pos = 1
  local len = #str

  while pos <= len do
    local sub_pos = str:find("_", pos, true)
    local sup_pos = str:find("^", pos, true)

    local next_pos, script_type = nil, nil
    if sub_pos and sup_pos then
      if sub_pos < sup_pos then
        next_pos, script_type = sub_pos, "sub"
      else
        next_pos, script_type = sup_pos, "sup"
      end
    elseif sub_pos then
      next_pos, script_type = sub_pos, "sub"
    elseif sup_pos then
      next_pos, script_type = sup_pos, "sup"
    end

    if not next_pos then
      local remaining = str:sub(pos)
      if #remaining > 0 then
        inlines:extend(text_to_inlines(remaining))
      end
      break
    end

    if next_pos > pos then
      inlines:extend(text_to_inlines(str:sub(pos, next_pos - 1)))
    end

    pos = next_pos + 1
    if pos <= len then
      local content = ""
      if str:sub(pos, pos) == "(" then
        local close_pos = find_closing_paren(str, pos)
        if close_pos then
          content = str:sub(pos + 1, close_pos - 1)
          pos = close_pos + 1
        else
          content = str:sub(pos + 1)
          pos = len + 1
        end
      else
        content = str:sub(pos, pos)
        pos = pos + 1
      end

      local inner = parse_script_string(content)
      if script_type == "sub" then
        inlines:insert(pandoc.Subscript(pandoc.Inlines(inner)))
      else
        inlines:insert(pandoc.Superscript(pandoc.Inlines(inner)))
      end
    end
  end

  return inlines
end

local function tex_to_unicode_native(el)
  local doc = pandoc.Pandoc({ pandoc.Para({ el }) })
  local plain_text = pandoc.write(doc, "plain")
  plain_text = plain_text:gsub("%s+$", "")
  local inlines = parse_script_string(plain_text)
  if CONFIG.inline_style == "italic" then
    return pandoc.Emph(inlines)
  end
  return inlines
end

--------------------------------------------------------------------------
-- filter entry point
--------------------------------------------------------------------------

local function is_pptx()
  if quarto and quarto.doc and quarto.doc.is_format then
    return quarto.doc.is_format("pptx")
  end
  -- running under plain pandoc (pandoc -L)
  return FORMAT == "pptx"
end

function Math(el)
  if not is_pptx() then
    return nil -- leave html/pdf/docx untouched
  end

  if el.mathtype == "InlineMath" then
    return tex_to_unicode_native(el)
  end

  local img = tex_to_image(el.text)
  if not img then
    return nil
  end

  local attrs = {}
  if img.width then attrs.width = img.width .. "pt" end
  if img.height then attrs.height = img.height .. "pt" end
  return pandoc.Image({}, img.path, "", pandoc.Attr("", {}, attrs))
end

load_config()
