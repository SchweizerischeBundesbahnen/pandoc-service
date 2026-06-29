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
--   BG (shading)    -> full-height \colorbox[HTML]{RRGGBB}{\strut ...}
--   HL (highlight)  -> full-height \colorbox[HTML]{RRGGBB}{\strut ...}
--   SZ (font size)  -> {\fontsize{pt}{baseline}\selectfont ...}
--
-- Why \colorbox and not soul's \hl: soul's \hl band hugs the glyph height, so a
-- plain-text highlight sits shorter and lower than the box we are forced to use
-- when a highlight ALSO carries underline/strikeout (soul and ulem cannot
-- compose, so decorated highlight must be a box). Mixing the two left a visible
-- step in the background wherever an underlined word met a non-underlined one.
-- Boxing every highlight with a leading \strut gives one uniform band height
-- across the whole run (the highlighter-pen look) and a \colorbox tolerates any
-- inline content, so bold/italic/underlined highlights need no special casing.
--
-- The deliberate trade-off: a \colorbox cannot break across a line, so a
-- highlighted run no longer hyphenates, and a run longer than MAX_BOXABLE_LEN
-- (or one containing an image/link/math/raw LaTeX — see BOX_UNSAFE) drops its
-- background rather than overflow the margin. The foreground color is still
-- applied in that case (\textcolor is line-breakable and harmless around
-- images), so only the shading is lost on those rare long/complex runs.
--
-- \textcolor and \fontsize wrap OUTSIDE the box so they apply to the whole run
-- regardless of whether it ends up boxed.

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

-- Every highlight (background shading or Word highlight) is rendered as a
-- full-height \colorbox rather than soul's \hl. soul's \hl band hugs the glyph
-- height, so a plain-text highlight (e.g. "de") sits visibly shorter and lower
-- than the box used for a highlight that also carries underline/strikeout (soul
-- and ulem can't compose, so decorated highlight MUST be a box). Mixing the two
-- left a step in the background band wherever an underlined word abutted a
-- non-underlined one. Boxing everything with a leading \strut gives one uniform
-- band height across the whole run — the highlighter-pen look — with no steps.
-- The trade-off (chosen deliberately): a \colorbox can't break across lines, so
-- a highlighted run no longer hyphenates and a run longer than MAX_BOXABLE_LEN
-- drops its background rather than overflow the margin.

-- Inline types that must NOT be wrapped in a \colorbox: a box is unbreakable
-- and adds no value here — an Image gets padded/pushed past \textwidth, and
-- Link/Code/Math/raw LaTeX can carry their own boxes or fragile macros. When a
-- highlight's content contains any of these we drop the background (keeping the
-- foreground colour, which is harmless around images); otherwise it is plain
-- text wrapped in decorations/inline-formatting (Underline, Strikeout, Emph,
-- Strong, sub/superscript, nested Spans), which boxes safely.
local BOX_UNSAFE = {
  Image = true,
  Link = true,
  Code = true,
  Math = true,
  RawInline = true,
  Note = true,
  Cite = true,
  -- A \colorbox cannot break, so a hard line break inside it (\\ ) would either
  -- be swallowed or overflow; such content must not be boxed.
  LineBreak = true,
}

-- A \colorbox is a single unbreakable hbox: content wider than the text block
-- overflows the margin (an "Overfull \hbox ... too wide" that runs metres off
-- the page). Most highlighted runs are a word or a short phrase, so we only box
-- runs whose text is at most this many characters; a longer run drops its
-- background rather than overflow. ~60 chars is well under one \textwidth line
-- at body size.
local MAX_BOXABLE_LEN = 60

-- Total length of the plain text in `content` (Str chars + one per Space /
-- SoftBreak), recursing into wrappers. Used only to keep \colorbox runs short.
local function content_text_length(content)
  local len = 0
  for _, inline in ipairs(content) do
    if inline.t == "Str" then
      len = len + #inline.text
    elseif inline.t == "Space" or inline.t == "SoftBreak" then
      len = len + 1
    elseif inline.content then
      len = len + content_text_length(inline.content)
    end
  end
  return len
end

local function content_is_boxable(content)
  for _, inline in ipairs(content) do
    if BOX_UNSAFE[inline.t] then
      return false
    end
    if inline.content and not content_is_boxable(inline.content) then
      return false
    end
  end
  return true
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
  -- Pin the underline depth/thickness for BOTH underline mechanisms so every
  -- underline sits at the same place regardless of font size or which mechanism
  -- drew it. Plain underlined text uses soul's \ul (it hyphenates); underlined
  -- text carrying other formatting uses ulem's \uline (it composes with macros).
  -- Both default to a depth that scales with the current font, so an underline
  -- under a larger run dropped lower than its normal-size neighbour — a visible
  -- step where the formatting changed. Fixing ulem's \ULdepth and soul's \setul
  -- to the same absolute 1.6pt/0.4pt keeps the rule level across size changes and
  -- across the soul/ulem boundary. 1.6pt is the body-size auto value, so ordinary
  -- underlines are unchanged.
  "\\setlength{\\ULdepth}{1.6pt}",
  "\\renewcommand{\\ULthickness}{0.4pt}",
  "\\setul{1.6pt}{0.4pt}",
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
  -- A white background is the page colour: invisible, and Polarion stamps
  -- background-color:#FFFFFF on virtually every run. Boxing it would wrap most
  -- of the document in non-breakable \colorboxes (lines overflow the margin),
  -- so treat white as no highlight at all.
  if bg_color and bg_color:upper() == "FFFFFF" then
    bg_color = nil
  end

  -- Highlight is a full-height \colorbox (uniform band, no soul \hl). We can
  -- only box content a box won't break or distort, and only while it is short
  -- enough to fit a line; otherwise the background is dropped (the foreground
  -- colour is still applied, which is harmless and line-breakable).
  local box_bg = bg_color ~= nil and content_is_boxable(el.content) and content_text_length(el.content) <= MAX_BOXABLE_LEN

  local size_open = props.sz and font_size_open(props.sz) or nil

  if not props.fg and not box_bg and not size_open then
    return nil
  end

  local open, close = "", ""

  -- Outermost: font size. \fontsize ... \selectfont changes the size for the
  -- rest of the group; it wraps everything else so colour/highlight inherit it.
  if size_open then
    open = open .. size_open
    close = "}" .. close
  end

  -- Outer: foreground color. \textcolor is a directive (no box, no padding, no
  -- line-break inhibition); applied first so it wraps the highlight box inside.
  if props.fg then
    open = open .. "\\textcolor[HTML]{" .. props.fg .. "}{"
    close = "}" .. close
  end

  -- Inner: the highlight box. \fboxsep=0pt so it hugs the text horizontally and
  -- a leading \strut forces full line height, so every highlighted run — plain,
  -- bold, italic, underlined, struck — shares one uniform band height and the
  -- background never steps. A \colorbox tolerates any inline content (unlike
  -- soul's \hl), so no wrapper hoisting is needed.
  if box_bg then
    -- \strut{} (not "\strut ") forces full line height: the {} terminates the
    -- control word so a following space — e.g. when this box bridges the gap
    -- between two highlighted words and its whole content IS a space — is
    -- typeset rather than swallowed as the macro's argument terminator.
    open = open .. "{\\setlength{\\fboxsep}{0pt}\\colorbox[HTML]{" .. bg_color .. "}{\\strut{}"
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
