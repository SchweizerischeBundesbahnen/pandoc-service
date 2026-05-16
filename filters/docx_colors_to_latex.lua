-- docx_colors_to_latex.lua
--
-- Companion to app/DocxColorPreProcess.py. The preprocessor rewrites every
-- run with direct color formatting in a DOCX as a reference to a synthetic
-- character style named "PandocColor__FG_RRGGBB__BG_RRGGBB__HL_<word-name>"
-- (each segment optional). With `-f docx+styles`, pandoc surfaces those
-- references as Span nodes carrying `custom-style="PandocColor__..."`.
--
-- This filter intercepts those Spans and wraps their content in raw LaTeX
-- so the foreground color, background shading, and Word-style highlight
-- survive into the PDF/LaTeX output:
--
--   FG (color)      -> \textcolor[HTML]{RRGGBB}{...}
--   BG (shading)    -> \hl{...} via soul + \sethlcolor   (only for plain text)
--   HL (highlight)  -> \hl{...} via soul + \sethlcolor   (only for plain text)
--
-- Why \hl and not \colorbox: \colorbox produces a non-breakable hbox, so any
-- colored run longer than a line overflows the right margin (paragraphs end
-- up outside the page), and wrapping an image in \colorbox adds padding that
-- pushes wide images past \textwidth. soul's \hl is line-breakable.
--
-- Why \textcolor is OUTSIDE \hl: soul's \hl reconstructs the content
-- character-by-character and crashes on most macros it has not been told
-- about (\textcolor, \href, \textbf, ...). pdflatex emits "Argument of
-- \textcolor has an extra }" when the macro is inside \hl. Putting the
-- color directive outside means \hl only sees plain text characters,
-- which is the only input shape it reliably typesets.
--
-- Why we still drop \hl for non-trivial inline content: even with the
-- color macro hoisted out, the Span may itself contain pandoc-emitted
-- macros (\textbf for Strong, \href for Link, \includegraphics for Image,
-- math delimiters, raw LaTeX, etc.). Any of those inside \hl produce
-- corrupt output: the link disappears, math is mis-typeset, the image
-- gets dropped or clipped. We therefore only apply the highlight wrapper
-- when the span's content is pure text (Str/Space/SoftBreak/LineBreak).
-- For anything richer we silently drop the highlight and rely on the
-- foreground color alone — which \textcolor preserves losslessly,
-- including around images.

-- Word's 16 named highlight colors (ECMA-376 17.18.40) -> RGB hex.
local HIGHLIGHT_HEX = {
  yellow      = "FFFF00",
  green       = "00FF00",
  cyan        = "00FFFF",
  magenta     = "FF00FF",
  blue        = "0000FF",
  red         = "FF0000",
  darkBlue    = "000080",
  darkCyan    = "008080",
  darkGreen   = "008000",
  darkMagenta = "800080",
  darkRed     = "800000",
  darkYellow  = "808000",
  darkGray    = "808080",
  lightGray   = "C0C0C0",
  black       = "000000",
  white       = "FFFFFF",
}

local STYLE_PREFIX = "PandocColor"

-- The only AST inline node types soul's \hl can wrap safely. Pandoc's
-- LaTeX writer emits macros for everything else (\textbf for Strong,
-- \emph for Emph, \href for Link, \includegraphics for Image, math
-- delimiters, raw LaTeX, ...), and soul cannot parse those.
local SOUL_SAFE_INLINES = {
  Str = true,
  Space = true,
  SoftBreak = true,
  LineBreak = true,
}

local function content_is_soul_safe(content)
  for _, inline in ipairs(content) do
    if not SOUL_SAFE_INLINES[inline.t] then
      return false
    end
  end
  return true
end

-- Parse the synthetic style name into a {fg, bg, hl} table, or return nil
-- if it does not match our naming scheme. Segments are separated by "__".
-- Each segment is "<KEY>_<value>" with KEY in {FG, BG, HL} and value with
-- no internal underscore (6-char hex or a Word highlight identifier).
local function parse_style(name)
  if name:sub(1, #STYLE_PREFIX) ~= STYLE_PREFIX then
    return nil
  end
  local props = {}
  for key, value in name:gmatch("__(%a+)_([^_]+)") do
    if key == "FG" then
      props.fg = value
    elseif key == "BG" then
      props.bg = value
    elseif key == "HL" then
      props.hl = value
    end
  end
  return props
end

local filter = {}

-- Inject \usepackage{soul} into header-includes so the document preamble
-- loads soul before \hl appears in the body. Only meaningful when the
-- target writer is LaTeX/PDF; for any other writer this is a harmless
-- no-op (the metadata is simply ignored by non-LaTeX writers).
function filter.Meta(meta)
  if not FORMAT:match("latex") then
    return nil
  end
  local soul_block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", "\\usepackage{soul}") })
  local existing = meta["header-includes"]
  if existing == nil then
    meta["header-includes"] = pandoc.MetaList({ soul_block })
  elseif existing.t == "MetaList" then
    table.insert(existing, soul_block)
    meta["header-includes"] = existing
  else
    -- Single MetaInlines/MetaBlocks value — promote to MetaList and append.
    meta["header-includes"] = pandoc.MetaList({ existing, soul_block })
  end
  return meta
end

function filter.Span(el)
  local style = el.attributes and el.attributes["custom-style"]
  if not style then return nil end
  local props = parse_style(style)
  if not props then return nil end

  -- Resolve the background color. <w:shd> (BG) takes precedence over
  -- <w:highlight> (HL) when both are present — w:shd carries an explicit
  -- RGB value, w:highlight only a palette name.
  local bg_color = props.bg
  if not bg_color and props.hl then
    bg_color = HIGHLIGHT_HEX[props.hl]
  end

  -- Highlight wrapper only when soul can handle the content shape.
  local apply_bg = bg_color and content_is_soul_safe(el.content)

  if not props.fg and not apply_bg then
    return nil
  end

  local open, close = "", ""

  -- Outer: foreground color. \textcolor is a directive (no box, no
  -- padding, no line-break inhibition); applied first so it wraps the
  -- entire span including any highlight wrapper inside.
  if props.fg then
    open = open .. "\\textcolor[HTML]{" .. props.fg .. "}{"
    close = "}" .. close
  end

  -- Inner: highlight via soul \hl. The braces around \definecolor and
  -- \sethlcolor make the color binding local so other highlighted spans
  -- can redefine "pdc_hl" without interference. \hl sees only plain
  -- text characters (the content-safety check above guarantees that).
  if apply_bg then
    open = open .. "{\\definecolor{pdc_hl}{HTML}{" .. bg_color .. "}\\sethlcolor{pdc_hl}\\hl{"
    close = "}}" .. close
  end

  local result = { pandoc.RawInline("latex", open) }
  for _, inline in ipairs(el.content) do
    result[#result + 1] = inline
  end
  result[#result + 1] = pandoc.RawInline("latex", close)
  return result
end

return filter
