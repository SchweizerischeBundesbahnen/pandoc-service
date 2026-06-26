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
-- Bold/italic highlighted text: the same hoist-it-outside trick extends to
-- the inline-formatting macros. When the whole span content is uniformly
-- wrapped in Emph/Strong/Superscript/Subscript (the common case — a fully
-- italic or bold highlighted phrase), we peel those wrappers and emit them
-- as macros AROUND \hl (\emph{...\hl{plain}...}), so \hl still sees only
-- plain text. hoist_for_hl() does this recursively, so bold-italic
-- (Strong[Emph[...]]) hoists both. We deliberately do NOT hoist
-- Underline/Strikeout: soul's own \ul/\st cannot nest inside \hl, and the
-- core \underline is not line-breakable.
--
-- Why we still drop \hl for genuinely mixed content: if the highlight range
-- only partially overlaps a formatting run (e.g. [Str "a", Emph "b"]) or
-- contains a \href/\includegraphics/math, there is no single macro to hoist
-- and soul cannot wrap it. hoist_for_hl() returns nil in that case and we
-- fall back to dropping the highlight and keeping the foreground color
-- alone — which \textcolor preserves losslessly, including around images.

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

-- Inline-formatting nodes whose LaTeX macro uniformly wraps the node's whole
-- content and is safe to place OUTSIDE soul's \hl. These are core LaTeX macros
-- (not soul commands), so nesting them around \hl is fine. Underline/Strikeout
-- are intentionally absent — see the header comment.
local HOISTABLE = {
  Emph        = { "\\emph{", "}" },
  Strong      = { "\\textbf{", "}" },
  Superscript = { "\\textsuperscript{", "}" },
  Subscript   = { "\\textsubscript{", "}" },
}

-- Peel uniform formatting wrappers off `content` so soul's \hl can wrap the
-- plain text inside. Returns (open, close, inner_inlines) where open/close are
-- the hoisted macros to place around \hl, and inner_inlines is the soul-safe
-- remainder to render inside it. Returns nil when the content is mixed/partial
-- or contains something \hl can't handle (caller then drops the highlight).
local function hoist_for_hl(content)
  local open, close = "", ""
  while true do
    if content_is_soul_safe(content) then
      return open, close, content
    end
    -- Only a single wrapper covering the entire content can be hoisted; a list
    -- with siblings means the formatting is partial and soul can't wrap it.
    if #content == 1 and HOISTABLE[content[1].t] then
      local macro = HOISTABLE[content[1].t]
      open = open .. macro[1]
      close = macro[2] .. close
      content = content[1].content
    else
      return nil
    end
  end
end

-- Parse the synthetic style name into a {fg, bg, hl, sz} table, or return nil
-- if it does not match our naming scheme. Segments are separated by "__".
-- Each segment is "<KEY>_<value>" with KEY in {FG, BG, HL, SZ} and value with
-- no internal underscore (6-char hex, a Word highlight identifier, or — for
-- SZ — the font size in half-points).
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
    elseif key == "SZ" then
      props.sz = value
    end
  end
  return props
end

-- Build a "{\fontsize{pt}{baselineskip}\selectfont " opener for a half-point
-- size, or nil when the value isn't a clean positive integer. pt = half-points
-- / 2; the baseline skip follows LaTeX's usual 1.2x leading. string.format with
-- %g keeps the trust boundary tight — only digits and a dot reach the output.
local function font_size_open(half_points)
  local hp = tonumber(half_points)
  if not hp or hp <= 0 or hp ~= math.floor(hp) then
    return nil
  end
  local pt = hp / 2
  return "{\\fontsize{" .. string.format("%g", pt) .. "}{" .. string.format("%g", pt * 1.2) .. "}\\selectfont "
end

local filter = {}

-- Inject the LaTeX preamble fixups this pipeline needs into header-includes.
-- Only meaningful when the target writer is LaTeX/PDF; for any other writer
-- this is a harmless no-op (the metadata is ignored by non-LaTeX writers).
--
--   * \usepackage{soul} — defines \hl (the line-breakable highlight we emit
--     for background/highlight runs below) and pandoc's \ul/\st for
--     underline/strikeout.
--   * pandoc renders Underline/Strikeout as soul's \ul/\st, but soul
--     reconstructs its argument character-by-character and aborts with
--     "Reconstruction failed" inside the LR boxes that
--     \textsuperscript/\textsubscript build — which happens as soon as
--     underlined/struck text sits inside a <sup>/<sub>. ulem's \uline/\sout are
--     box-safe, so we route \ul/\st to them — but ONLY inside super/subscripts,
--     by redefining \textsuperscript/\textsubscript to swap the soul commands
--     locally. Globally, \ul/\st stay soul, so ordinary underlined/struck text
--     renders and line-breaks exactly as before (no document-wide change).
--     \hl has no box-safe drop-in, so inside a super/subscript we drop the
--     highlight (keep the text) — rare, and only the highlight is affected.
local PREAMBLE = table.concat({
  "\\usepackage{soul}",
  "\\usepackage[normalem]{ulem}",
  "\\providecommand{\\pdcDropHl}[1]{#1}",
  "\\let\\pdcOldSuperscript\\textsuperscript",
  "\\let\\pdcOldSubscript\\textsubscript",
  "\\renewcommand{\\textsuperscript}[1]{\\pdcOldSuperscript{\\let\\ul\\uline\\let\\st\\sout\\let\\hl\\pdcDropHl#1}}",
  "\\renewcommand{\\textsubscript}[1]{\\pdcOldSubscript{\\let\\ul\\uline\\let\\st\\sout\\let\\hl\\pdcDropHl#1}}",
}, "\n")

function filter.Meta(meta)
  if not FORMAT:match("latex") then
    return nil
  end
  local soul_block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", PREAMBLE) })
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

  -- Highlight wrapper only when soul can handle the content shape. We peel
  -- any uniform Emph/Strong/... wrappers out so \hl sees plain text; hl_inner
  -- is what goes inside \hl (the original content minus the hoisted wrappers).
  local hl_macros_open, hl_macros_close, hl_inner
  if bg_color then
    hl_macros_open, hl_macros_close, hl_inner = hoist_for_hl(el.content)
  end
  local apply_bg = hl_inner ~= nil

  local size_open = props.sz and font_size_open(props.sz) or nil

  if not props.fg and not apply_bg and not size_open then
    return nil
  end

  local open, close = "", ""

  -- Outermost: font size. \fontsize ... \selectfont changes the size for the
  -- rest of the group; it wraps everything else so colour/highlight inherit it.
  if size_open then
    open = open .. size_open
    close = "}" .. close
  end

  -- Outer: foreground color. \textcolor is a directive (no box, no
  -- padding, no line-break inhibition); applied first so it wraps the
  -- entire span including any formatting/highlight wrappers inside.
  if props.fg then
    open = open .. "\\textcolor[HTML]{" .. props.fg .. "}{"
    close = "}" .. close
  end

  -- Middle + inner: inline-formatting macros hoisted out of \hl (\emph,
  -- \textbf, ...) followed by the highlight itself. The braces around
  -- \definecolor and \sethlcolor make the color binding local so other
  -- highlighted spans can redefine "pdc_hl" without interference. \hl sees
  -- only plain text characters (hoist_for_hl guarantees that). All emitted
  -- together so the macro strings are only referenced when apply_bg is true.
  if apply_bg then
    open = open .. hl_macros_open .. "{\\definecolor{pdc_hl}{HTML}{" .. bg_color .. "}\\sethlcolor{pdc_hl}\\hl{"
    close = "}}" .. hl_macros_close .. close
  end

  -- When highlighting we render the peeled-down inner inlines (the hoisted
  -- Emph/Strong are now macros in `open`/`close`, so re-emitting them here
  -- would double them up). Otherwise render the original content untouched.
  local body = apply_bg and hl_inner or el.content

  local result = { pandoc.RawInline("latex", open) }
  for _, inline in ipairs(body) do
    result[#result + 1] = inline
  end
  result[#result + 1] = pandoc.RawInline("latex", close)
  return result
end

return filter
