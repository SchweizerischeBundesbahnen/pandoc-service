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
--   SZ (font size)  -> {\fontsize{bp}{baseline}\selectfont ...}
--   FN (font)       -> {\fontspec{Family}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}] ...}, the family as the hex of its UTF-8
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
-- highlighted run no longer hyphenates, and a word longer than MAX_BOXABLE_LEN
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
-- a highlighted run no longer hyphenates and a word longer than MAX_BOXABLE_LEN
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
-- Valid 6-char hex colour, the same guard docx_tables_to_latex.lua puts in
-- front of \cellcolor. A style name is document controlled: the preprocessor
-- writes these names, but a DOCX may carry a character style of its own called
-- "PandocColor__FG_AAAAAA}{\input{/etc/passwd}", and the value reaches a raw
-- LaTeX node which this filter emits after the strip filter has run. Anything
-- which is not a colour is dropped rather than written out. Verified against
-- this image.
local function valid_hex(s)
  return s and #s == 6 and s:match("^%x%x%x%x%x%x$") ~= nil
end

-- The family of an FN segment: the hex of its UTF-8, as app/docx_color_pre_process.py
-- writes it for a family the container has. Letters, digits, spaces, dots, hyphens and
-- underscores only reach LaTeX.
local function font_family(hex)
  if not hex:match("^%x+$") or #hex % 2 ~= 0 then
    return nil
  end
  local family = hex:gsub("%x%x", function(byte)
    return string.char(tonumber(byte, 16))
  end)
  return family:match("^[%w][%w %.%-_]*$") and family or nil
end

local function parse_style(name)
  if name:sub(1, #STYLE_PREFIX) ~= STYLE_PREFIX then
    return nil
  end
  local props = {}
  for key, value in name:gmatch("__(%a+)_([^_]+)") do
    if key == "FG" and valid_hex(value) then
      props.fg = value
    elseif key == "BG" and valid_hex(value) then
      props.bg = value
    elseif key == "HL" then
      props.hl = value
    elseif key == "SZ" then
      props.sz = value
    elseif key == "FN" then
      props.font = font_family(value)
      props.font_hex = value:upper()
    end
  end
  return props
end

-- Build a "{\fontsize{pt}{baselineskip}\selectfont " opener for a half-point
-- size, or nil when the value isn't a clean positive integer. pt = half-points
-- / 2; the baseline skip follows LaTeX's usual 1.2x leading. string.format with
-- %g keeps the trust boundary tight — only digits and a dot reach the output.
-- The size of the body text, which docx_document_styles_to_latex sets: see filter.Meta
local body_size = nil

local function font_size_open(half_points)
  local hp = tonumber(half_points)
  if not hp or hp <= 0 or hp ~= math.floor(hp) then
    return nil
  end
  local pt = hp / 2
  -- bp: Word counts in PostScript points, a pt of TeX is 0.4% smaller
  local open = "{\\fontsize{" .. string.format("%g", pt) .. "bp}{" .. string.format("%g", pt * 1.2) .. "bp}\\selectfont "
  if body_size and pt > body_size then
    -- Word sets a line as high as the ascent and as deep as the descent of the largest font in it
    open = open .. "\\pdcRunStrut{}"
  end
  return open
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
  -- ulem draws the line of a word before the word, which a highlight box then covers: Word draws it over
  "\\makeatletter\\def\\UL@putbox{\\ifx\\UL@start\\@empty\\else\\vrule\\@width\\z@\\LA@penalty\\@M"
    .. "{\\UL@skip\\wd\\UL@box\\box\\UL@box\\kern-\\UL@skip\\UL@leaders}\\fi}\\makeatother",
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
  -- Underline and strike-through in one pass: ulem breaks a line only at the spaces of its outermost
  -- command, so a \\sout inside a \\uline was one word that no line could break
  "\\newcommand{\\pdcUlineSout}{\\bgroup\\markoverwith{\\hbox{\\rlap{\\rule[-\\ULdepth]{.25pt}{\\ULthickness}}\\rule[.5ex]{.25pt}{\\ULthickness}}}\\ULon}",
  "\\providecommand{\\pdcDropHl}[1]{#1}",
  "\\let\\pdcOldSuperscript\\textsuperscript",
  "\\let\\pdcOldSubscript\\textsubscript",
  "\\renewcommand{\\textsuperscript}[1]{\\pdcOldSuperscript{\\let\\ul\\uline\\let\\st\\sout\\let\\hl\\pdcDropHl#1}}",
  "\\renewcommand{\\textsubscript}[1]{\\pdcOldSubscript{\\let\\ul\\uline\\let\\st\\sout\\let\\hl\\pdcDropHl#1}}",
}, "\n")

-- A strut of the ascent and the descent of the current font: without the metrics of its glyphs, XeTeX
-- sets a word as high as the ascent of its font and as deep as its descent
-- The line and the strut of a paragraph of runs in another font than the body's, see run_line
local PAR_LINE = "\\newcommand{\\pdcParLine}[3]{\\baselineskip=#1\\relax\\setbox\\strutbox\\hbox{\\vrule height #2 depth #3 width 0pt}}"

-- A superscript and a subscript as Word sets them in the font around them, see script_ratios of
-- app/docx_document_styles.py: \pdcScriptFont takes the size of a script, the raise and the drop of the font,
-- its ascent, at least an em, and its descent, as shares of its size. A script is that size, rounded to a half
-- point, its ascent meets the ascent of the text and its descent the descent, never further off the baseline than
-- the font asks for. A font without the metrics, Latin Modern among them, keeps the scripts of LaTeX.
local SCRIPTS = table.concat({
  "\\makeatletter",
  "\\newcommand{\\pdcScriptFont}[5]{\\def\\pdc@script{{#1}{#2}{#3}{#4}{#5}}}",
  "\\newcommand{\\pdcNoScriptFont}{\\let\\pdc@script\\@undefined}",
  "\\def\\pdc@scriptsize#1{\\dimexpr\\numexpr\\dimexpr#1\\dimexpr\\f@size pt\\relax\\relax/32768\\relax sp*32768\\relax}",
  -- #1 the size, #2 the most the font allows, #3 the share of the rest of the size, #4 \raise or \lower.
  -- \raise starts no paragraph, a script first in one does
  "\\def\\pdc@shift#1#2#3#4#5{\\leavevmode\\begingroup\\dimen@=\\pdc@scriptsize{#1}\\dimen\\tw@=#2\\dimexpr\\f@size pt\\relax"
    .. "\\dimen4=#3\\dimexpr\\dimexpr\\f@size pt\\relax-\\dimen@\\relax\\ifdim\\dimen4<\\dimen\\tw@\\dimen\\tw@=\\dimen4 \\fi"
    .. "#4\\dimen\\tw@\\hbox{\\fontsize{\\dimen@}{\\z@}\\selectfont #5}\\endgroup}",
  "\\def\\pdc@super#1#2#3#4#5#6{\\pdc@shift{#1}{#2}{#4}\\raise{#6}}",
  "\\def\\pdc@sub#1#2#3#4#5#6{\\pdc@shift{#1}{#3}{#5}\\lower{#6}}",
  "\\let\\pdc@oldsuper\\@textsuperscript",
  "\\let\\pdc@oldsub\\@textsubscript",
  "\\def\\@textsuperscript#1{\\ifdefined\\pdc@script\\expandafter\\pdc@super\\pdc@script{#1}\\else\\pdc@oldsuper{#1}\\fi}",
  "\\def\\@textsubscript#1{\\ifdefined\\pdc@script\\expandafter\\pdc@sub\\pdc@script{#1}\\else\\pdc@oldsub{#1}\\fi}",
  "\\makeatother",
}, "\n")

local RUN_STRUT = "\\newcommand{\\pdcRunStrut}{\\begingroup\\XeTeXuseglyphmetrics=0 \\setbox0\\hbox{x}"
  .. "\\rule[-\\dp0]{0pt}{\\dimexpr\\ht0+\\dp0\\relax}\\endgroup}"

-- The ascent and the single line of the font of each run, as shares of its size, by the hex of the name of the
-- font, and the line of the body: see app/docx_document_styles.py
local run_fonts = {}
local body_line = nil
-- The script ratios of the font of the body, as the arguments of \pdcScriptFont
local body_script = nil

local SCRIPT_FIELDS = "^:([%d.]+):([%d.]+):([%d.]+):([%d.]+):([%d.]+)$"

-- The script ratios of five numbers apart by colons as the arguments of \pdcScriptFont; nil for anything else.
local function script_arguments(field)
  local ratios = { field:match(SCRIPT_FIELDS) }
  if #ratios ~= 5 then
    return nil
  end
  return "{" .. table.concat(ratios, "}{") .. "}"
end

local function number(meta, key)
  return tonumber(pandoc.utils.stringify(meta[key] or ""))
end

local function read_body_size(meta)
  body_size = number(meta, "docx-body-size")
  for entry in pandoc.utils.stringify(meta["docx-run-fonts"] or ""):gmatch("[^,]+") do
    local hex, ascent, line, script = entry:match("^(%x+):([%d.]+):([%d.]+)(.*)$")
    if hex then
      run_fonts[hex:upper()] = { ascent = tonumber(ascent), line = tonumber(line), script = script_arguments(script) }
    end
  end
  body_script = script_arguments(":" .. pandoc.utils.stringify(meta["docx-body-script"] or ""))
  local ascent, descent, multiple = number(meta, "docx-body-ascent"), number(meta, "docx-body-descent"), number(meta, "docx-body-line-multiple")
  if body_size and ascent and descent and multiple then
    body_line = { ascent = ascent, descent = descent, multiple = multiple }
  end
  return nil
end

-- What a paragraph of runs may hold besides them: the spaces and breaks between them, the LaTeX of other
-- filters, and the formatting and the plain spans around them.
local RUN_GAPS = { Space = true, SoftBreak = true, LineBreak = true, RawInline = true }
local RUN_CONTAINERS = {
  Span = true, Emph = true, Strong = true, Underline = true, Strikeout = true, SmallCaps = true,
  Superscript = true, Subscript = true, Link = true, Quoted = true,
}

-- How far a line of the font and the size of a run reaches above and below the baseline, as Word sets it; nil
-- for a font without metrics.
local function run_metrics(props)
  if body_line == nil or body_size == nil then
    return nil
  end
  local size = props.sz and tonumber(props.sz) / 2 or body_size
  local font = props.font_hex and run_fonts[props.font_hex]
  if props.font_hex and font == nil then
    return nil
  end
  local ascent = font and font.ascent * size or body_line.ascent * size / body_size
  local line = font and font.line * size or (body_line.ascent + body_line.descent) * size / body_size
  return ascent, line - ascent
end

-- The font and the size of a run, as one key.
local function run_key(inline)
  local name = inline.t == "Span" and inline.attributes["custom-style"]
  local props = name and parse_style(name)
  if not props or not (props.font_hex or props.sz) then
    return nil
  end
  return (props.font_hex or "") .. ":" .. (props.sz or ""), props
end

-- Word sets a line as high as the font of its text. A paragraph of runs alone, all in one font and size
-- other than the body's, gets the line and the strut of that font, times the line spacing of its style:
-- "{\pdcParLine{line}{ascent}{descent}", which "\par}" closes. docx_document_styles_to_latex puts the
-- struts of the paragraph inside.
local function run_line(para)
  if body_line == nil then
    return nil
  end
  local key, props, mixed = nil, nil, false
  local function visit(inlines)
    for _, inline in ipairs(inlines) do
      local own, own_props = run_key(inline)
      if own then
        mixed = mixed or (key ~= nil and own ~= key)
        key, props = own, own_props
      elseif RUN_CONTAINERS[inline.t] then
        visit(inline.content)
      elseif not RUN_GAPS[inline.t] then
        -- Text in the font of the body, a picture, an equation
        mixed = true
      end
    end
  end
  visit(para.content)
  if key == nil or mixed then
    return nil
  end
  if props.sz == nil then
    -- Runs without a size of their own take the size of the paragraph, which only TeX knows: a caption, for one
    local font = run_fonts[props.font_hex]
    local body_ratio = body_line.ascent / body_size
    if font == nil or (math.abs(font.ascent - body_ratio) < 0.0001 and math.abs(font.line - (body_line.ascent + body_line.descent) / body_size) < 0.0001) then
      return nil
    end
    local size = "\\dimexpr\\csname f@size\\endcsname pt\\relax"
    return string.format("{\\pdcParLine{%.4f%s}{%.4f%s}{%.4f%s}", font.line * body_line.multiple, size, font.ascent, size, font.line - font.ascent, size)
  end
  local ascent, descent = run_metrics(props)
  if ascent == nil or (math.abs(ascent - body_line.ascent) < 0.01 and math.abs(descent - body_line.descent) < 0.01) then
    return nil
  end
  return string.format("{\\pdcParLine{%.4fbp}{%.4fbp}{%.4fbp}", (ascent + descent) * body_line.multiple, ascent, descent)
end

local function mark_run_line(para)
  if not FORMAT:match("latex") then
    return nil
  end
  local open = run_line(para)
  if open == nil then
    return nil
  end
  para.content:insert(1, pandoc.RawInline("latex", open))
  para.content:insert(pandoc.RawInline("latex", "\\par}"))
  return para
end

function filter.Meta(meta)
  if not FORMAT:match("latex") then
    return nil
  end
  local scripts = SCRIPTS .. (body_script and ("\n\\pdcScriptFont" .. body_script) or "")
  local soul_block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", PREAMBLE .. "\n" .. scripts .. "\n" .. RUN_STRUT .. "\n" .. PAR_LINE) })
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

-- The words of inlines and the spaces between them, as { word, "space", word, ... }: an element of
-- several words, as an emphasis, becomes a copy of it around each of them.
local SPLITTABLE = { Span = true, Emph = true, Strong = true, Underline = true, Strikeout = true, SmallCaps = true }

local function split_words(inlines)
  local words = { {} }
  for _, inline in ipairs(inlines) do
    if inline.t == "Space" or inline.t == "SoftBreak" then
      words[#words + 1] = "space"
      words[#words + 1] = {}
    elseif SPLITTABLE[inline.t] then
      for index, part in ipairs(split_words(inline.content)) do
        if part == "space" then
          words[#words + 1] = "space"
          words[#words + 1] = {}
        else
          local copy = inline:clone()
          copy.content = part
          if index > 1 or #part > 0 then
            table.insert(words[#words], copy)
          end
        end
      end
    else
      table.insert(words[#words], inline)
    end
  end
  return words
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
  -- A box per word, see below: each word has to fit a line, not the whole run
  local box_bg = bg_color ~= nil and content_is_boxable(el.content)
  if box_bg then
    for _, word in ipairs(split_words(el.content)) do
      if word ~= "space" and content_text_length(word) > MAX_BOXABLE_LEN then
        box_bg = false
      end
    end
  end

  local size_open = props.sz and font_size_open(props.sz) or nil

  if not props.fg and not box_bg and not size_open and not props.font then
    return nil
  end

  local open, close = "", ""

  -- Outermost of all: the family of the run, unkerned as Word sets it.
  if props.font then
    local font = props.font_hex and run_fonts[props.font_hex]
    local script = font and font.script and ("\\pdcScriptFont" .. font.script) or "\\pdcNoScriptFont{}"
    open = open .. "{\\fontspec{" .. props.font .. "}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}]" .. script
    close = "}" .. close
  end

  -- Outermost: font size. \fontsize ... \selectfont changes the size for the
  -- rest of the group; it wraps everything else so colour/highlight inherit it.
  if size_open then
    open = open .. size_open
    close = "}" .. close
  end
  -- A highlight box and its leaders take the strut: of the font of the run, where \\selectfont would set
  -- one of LaTeX, 0.7 and 0.3 of the line, deeper than the font of Word. A run without a size of its own
  -- takes the size around it, which only TeX knows: a heading, for one.
  local strut = nil
  if props.sz then
    local ascent, descent = run_metrics(props)
    if ascent then
      strut = string.format("height %.4fbp depth %.4fbp", ascent, descent)
    end
  elseif props.font_hex and run_fonts[props.font_hex] then
    local font = run_fonts[props.font_hex]
    strut = string.format("height %.4f\\dimexpr\\csname f@size\\endcsname pt\\relax depth %.4f\\dimexpr\\csname f@size\\endcsname pt\\relax",
      font.ascent, font.line - font.ascent)
  end
  if strut then
    open = open .. "\\setbox\\strutbox\\hbox{\\vrule " .. strut .. " width 0pt}"
  end

  -- Outer: foreground color. \textcolor is a directive (no box, no padding, no
  -- line-break inhibition); applied first so it wraps the highlight box inside.
  if props.fg then
    -- {}: \\color ignores the spaces after it, and a run may start with a space of its own
    local space = el.content[1] and el.content[1].t == "Space"
    open = open .. "\\textcolor[HTML]{" .. props.fg .. "}{" .. (space and "{}" or "")
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

  if box_bg then
    -- Word breaks a highlighted run at its spaces like any other: a box per word, and between them a
    -- space in the colour of the background, where a line may end
    local box_open = "{\\setlength{\\fboxsep}{0pt}\\colorbox[HTML]{" .. bg_color .. "}{\\strut{}"
    local space = "{\\color[HTML]{" .. bg_color .. "}\\leaders\\vrule height\\ht\\strutbox depth\\dp\\strutbox"
      .. "\\hskip\\fontdimen2\\font plus\\fontdimen3\\font minus\\fontdimen4\\font}"
    local words = split_words(el.content)
    if #words > 1 then
      local outer_open = open:sub(1, #open - #box_open)
      local outer_close = close:sub(3)
      local result = { pandoc.RawInline("latex", outer_open) }
      for _, word in ipairs(words) do
        if word == "space" then
          result[#result + 1] = pandoc.RawInline("latex", space)
        elseif #word > 0 then
          result[#result + 1] = pandoc.RawInline("latex", box_open)
          for _, inline in ipairs(word) do
            result[#result + 1] = inline
          end
          result[#result + 1] = pandoc.RawInline("latex", "}}")
        end
      end
      result[#result + 1] = pandoc.RawInline("latex", outer_close)
      return result
    end
  end

  local result = { pandoc.RawInline("latex", open) }
  for _, inline in ipairs(el.content) do
    result[#result + 1] = inline
  end
  result[#result + 1] = pandoc.RawInline("latex", close)
  return result
end

return { { Meta = read_body_size }, { Para = mark_run_line }, filter }
