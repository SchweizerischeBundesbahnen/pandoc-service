-- docx_comments_to_latex.lua
--
-- Print the comments of a docx in balloons right of the text, as Word does.
--
-- Word prints a document with comments on a smaller page, with a grey markup
-- area right of the text. Each comment gets a balloon there, "Commented [n]:"
-- and its text, as high as the line it comments or below the balloon before
-- it. A dashed line leads from the start of the comment to its balloon, and
-- brackets mark the start and the end. The reader turns a comment into a span
-- comment-start holding its text and an empty span comment-end, which the
-- LaTeX writer printed inline, in the middle of the text.
--
-- app/docx_page_geometry.py scales the page as Word does when the document
-- has comments, and passes the markup area as metadata. This filter replaces
-- the spans with anchors, which record their place on the page in the .aux
-- file; tectonic runs TeX again when it changes. The balloons are drawn when a
-- page is shipped out, at the places the run before recorded, so a comment in
-- a heading, a table or a footnote gets one as well.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx -> pdf/latex).

local MAG = "docx-comment-mag"
local AREA = "docx-comment-area"
local BALLOON = "docx-comment-balloon"
local ASCENT = "docx-body-ascent"
local DESCENT = "docx-body-descent"
local BASELINE = "docx-body-baseline"

-- The balloon as Word for Mac draws it, in points of the page before Word shrinks it. Word draws its text at
-- 9.12pt, but sets the glyphs as far apart as at 9pt: the place of the words counts
local TEXT_SIZE = 9
local TEXT_LINE = 10.94
local PADDING = 4.3
local FIRST_BASELINE = 10.6
local LAST_DEPTH = 4.7
local RADIUS = 2.7
local GAP = 1.2
-- How far below the top of its balloon the line from the comment ends
local LINE_END = 6.0
local BLUE = "0 0.471 0.831"
local FILL = "0.835 0.929 1"
local GREY = "0.949"
local LINE_WIDTH = 0.24
local DASH = 0.72
local BORDER = 0.48
local TICK = 0.27
-- When the font of the body is unknown: the share of a line above its baseline
local ASCENT_SHARE = 0.75
local DEFAULT_LINE = 14.4

local function numbers(value)
  local result = {}
  for item in pandoc.utils.stringify(value or ""):gmatch("[^%s,]+") do
    result[#result + 1] = tonumber(item)
  end
  return result
end

local function is_comment(inline, class)
  return inline.t == "Span" and inline.classes:includes(class)
end

-- The text of each comment by its number, in the order of the document.
local texts = {}

local function anchor(command)
  -- A heading carries its text into the bookmarks of the PDF, which take no command
  return pandoc.RawInline("latex", "\\texorpdfstring{" .. command .. "}{}")
end

local function replace_comments(inlines)
  local result = pandoc.Inlines({})
  for _, inline in ipairs(inlines) do
    if is_comment(inline, "comment-start") then
      local text = pandoc.write(pandoc.Pandoc({ pandoc.Plain(inline.content) }), "latex"):gsub("%s+$", "")
      texts[#texts + 1] = text
      result:insert(anchor(string.format("\\pdcCommentStart{%d}", #texts)))
    elseif is_comment(inline, "comment-end") then
      result:insert(anchor("\\pdcCommentEnd"))
    else
      result:insert(inline)
    end
  end
  return result
end

-- A paragraph of comments only is an empty line in Word.
local function keep_line(block)
  local comments = false
  for _, inline in ipairs(block.content) do
    if is_comment(inline, "comment-start") or is_comment(inline, "comment-end") then
      comments = true
    elseif inline.t ~= "Space" and inline.t ~= "SoftBreak" then
      return nil
    end
  end
  if not comments then
    return nil
  end
  block.content:insert(pandoc.RawInline("latex", "\\strut"))
  return block
end

local function bp(value)
  return string.format("%.4g", value)
end

local function preamble(mag, area, balloon, ascent, descent)
  local x, width, bend = balloon[1], balloon[2], balloon[3]
  local lines = {
    string.format("\\mag=%d", mag),
    -- XeTeX shrinks the page size by \mag once and the driver once more
    string.format("\\AtBeginDocument{\\pdfpagewidth=\\dimexpr\\paperwidth*1000/%d\\relax\\pdfpageheight=\\dimexpr\\paperheight*1000/%d\\relax}", mag, mag),
    "\\makeatletter",
    "\\newif\\ifpdc@anchors\\pdc@anchorstrue",
    -- A table of contents repeats the headings, and with them their anchors
    "\\AtBeginDocument{\\let\\pdc@starttoc\\@starttoc\\def\\@starttoc#1{\\begingroup\\pdc@anchorsfalse\\pdc@starttoc{#1}\\endgroup}}",
    "\\def\\pdc@bp#1{\\strip@pt\\dimexpr(#1)*7200/7227\\relax}",
    string.format("\\def\\pdc@bracket#1{\\special{pdf:content q %s RG %s w #1 %s m 0 %s l 0 -%s l #1 -%s l S Q}}", BLUE, bp(LINE_WIDTH), bp(ascent), bp(ascent), bp(descent), bp(descent)),
    string.format("\\protected\\def\\pdcCommentStart#1{\\ifpdc@anchors\\leavevmode\\pdc@bracket{%s}\\pdfsavepos\\write\\@auxout{\\string\\pdc@anchor{#1}{\\the\\ReadonlyShipoutCounter}{\\the\\pdflastxpos}{\\the\\pdflastypos}}\\fi}", bp(TICK)),
    string.format("\\protected\\def\\pdcCommentEnd{\\ifpdc@anchors\\leavevmode\\pdc@bracket{-%s}\\fi}", bp(TICK)),
    "\\def\\pdc@anchor#1#2#3#4{\\expandafter\\xdef\\csname pdc@at@#1\\endcsname{{#2}{#3}{#4}}}",
    string.format("\\def\\pdc@count{%d}", #texts),
    "\\newcount\\pdc@n",
    "\\newdimen\\pdc@next\\newdimen\\pdc@top\\newdimen\\pdc@ax\\newdimen\\pdc@ay\\newdimen\\pdc@h\\newdimen\\pdc@limit",
    "\\newbox\\pdc@text\\newbox\\pdc@box",
    "\\AtBeginDocument{\\IfFontExistsTF{Times New Roman}{\\newfontfamily\\pdc@labelfont{Times New Roman}[Scale=1]}{\\IfFontExistsTF{Liberation Serif}{\\newfontfamily\\pdc@labelfont{Liberation Serif}[Scale=1]}{\\let\\pdc@labelfont\\rmfamily}}}",
    -- The grey markup area, on every page
    string.format("\\AddToHook{shipout/background}{\\put(0,0){\\special{pdf:content q %s g %s %s %s %s re f Q}}}", GREY, bp(area[1]), bp(-area[2] - area[4]), bp(area[3]), bp(area[4])),
    -- The balloon of a comment: its frame, then its text
    table.concat({
      "\\def\\pdc@balloon#1{",
      string.format("\\setbox\\pdc@text\\vbox{\\hsize=%sbp\\normalfont\\ifdefined\\color\\color{black}\\fi\\fontsize{%sbp}{%sbp}\\selectfont", bp(width - 2 * PADDING), bp(TEXT_SIZE), bp(TEXT_LINE)),
      "\\parindent\\z@\\parskip\\z@\\rightskip\\z@ plus 1fil\\hyphenpenalty\\@M\\exhyphenpenalty\\@M",
      string.format("\\leavevmode\\rule{0pt}{%sbp}{\\pdc@labelfont\\bfseries Commented [#1]: }\\csname pdc@text@#1\\endcsname\\rule[-%sbp]{0pt}{0pt}\\par}", bp(FIRST_BASELINE), bp(LAST_DEPTH)),
      "\\pdc@h=\\dimexpr\\ht\\pdc@text+\\dp\\pdc@text\\relax",
      "\\edef\\pdc@frame{",
      string.format("%s rg %s RG %s w", FILL, BLUE, bp(BORDER)),
      string.format(" %s 0 m \\pdc@bp{%sbp-%sbp} 0 l", bp(RADIUS), bp(width), bp(RADIUS)),
      string.format(" %s 0 %s %s %s %s c", bp(width - RADIUS * 0.4477), bp(width), bp(RADIUS * 0.4477), bp(width), bp(RADIUS)),
      string.format(" %s \\pdc@bp{\\pdc@h-%sbp} l", bp(width), bp(RADIUS)),
      string.format(" %s \\pdc@bp{\\pdc@h-%sbp} %s \\pdc@bp{\\pdc@h} %s \\pdc@bp{\\pdc@h} c", bp(width), bp(RADIUS * 0.4477), bp(width - RADIUS * 0.4477), bp(width - RADIUS)),
      string.format(" %s \\pdc@bp{\\pdc@h} l", bp(RADIUS)),
      string.format(" %s \\pdc@bp{\\pdc@h} 0 \\pdc@bp{\\pdc@h-%sbp} 0 \\pdc@bp{\\pdc@h-%sbp} c", bp(RADIUS * 0.4477), bp(RADIUS * 0.4477), bp(RADIUS)),
      string.format(" 0 %s l 0 %s %s 0 %s 0 c h B}", bp(RADIUS), bp(RADIUS * 0.4477), bp(RADIUS * 0.4477), bp(RADIUS)),
      string.format("\\setbox\\pdc@box\\hbox{\\special{pdf:content q \\pdc@frame\\space Q}\\kern%sbp\\raise\\dp\\pdc@text\\box\\pdc@text}}", bp(PADDING)),
    }),
    -- Where the balloon of a comment anchored on this page goes: level with its anchor, below the one before
    table.concat({
      "\\def\\pdc@measure#1#2#3#4{\\ifnum#2=\\ReadonlyShipoutCounter",
      -- XeTeX takes the page from an inch in, the driver from a true inch
      "\\pdc@ax=\\dimexpr#3sp+1truein-1in\\relax\\pdc@ay=\\dimexpr\\pdfpageheight-#4sp+1truein-1in\\relax",
      string.format("\\pdc@top=\\dimexpr\\pdc@ay-%sbp\\relax\\ifdim\\pdc@top<\\pdc@next\\pdc@top=\\pdc@next\\fi", bp(ascent)),
      "\\pdc@balloon{#1}\\pdc@keep{#1}",
      string.format("\\global\\pdc@next=\\dimexpr\\pdc@top+\\pdc@h+%sbp\\relax\\fi}", bp(GAP)),
    }),
    "\\def\\pdc@keep#1{\\expandafter\\xdef\\csname pdc@pos@#1\\endcsname{{\\the\\pdc@ax}{\\the\\pdc@ay}{\\the\\pdc@top}{\\the\\pdc@h}}}",
    "\\def\\pdc@get#1#2#3#4{\\pdc@ax=#1\\pdc@ay=#2\\pdc@top=#3\\pdc@h=#4}",
    "\\def\\pdc@read#1{\\expandafter\\expandafter\\expandafter\\pdc@get\\csname pdc@pos@#1\\endcsname}",
    -- A balloon which reaches past the markup area moves up, and the ones above it with it, as Word moves them
    table.concat({
      "\\def\\pdc@lift#1{\\@ifundefined{pdc@pos@#1}{}{\\pdc@read{#1}",
      "\\ifdim\\dimexpr\\pdc@top+\\pdc@h\\relax>\\pdc@limit\\pdc@top=\\dimexpr\\pdc@limit-\\pdc@h\\relax\\fi",
      string.format("\\pdc@keep{#1}\\global\\pdc@limit=\\dimexpr\\pdc@top-%sbp\\relax}}", bp(GAP)),
    }),
    -- The balloon of a comment and the line from its anchor
    table.concat({
      "\\def\\pdc@draw#1{\\@ifundefined{pdc@pos@#1}{}{\\pdc@read{#1}\\pdc@balloon{#1}",
      string.format("\\put(0,0){\\special{pdf:content q %s RG %s w [%s %s] 0 d", BLUE, bp(LINE_WIDTH), bp(DASH), bp(DASH)),
      string.format(" \\pdc@bp{\\pdc@ax} \\pdc@bp{-\\pdc@ay-%sbp} m %s \\pdc@bp{-\\pdc@ay-%sbp} l", bp(descent), bp(bend), bp(descent)),
      string.format(" %s \\pdc@bp{-\\pdc@top-%sbp} l S Q}}", bp(x), bp(LINE_END)),
      string.format("\\put(\\strip@pt\\dimexpr%sbp\\relax,\\strip@pt\\dimexpr-\\pdc@top-\\pdc@h\\relax){\\box\\pdc@box}", bp(x)),
      "\\global\\expandafter\\let\\csname pdc@pos@#1\\endcsname\\relax}}",
    }),
    table.concat({
      "\\AddToHook{shipout/foreground}{\\global\\pdc@next=-\\maxdimen\\global\\pdc@n=\\z@",
      "\\@whilenum\\pdc@n<\\pdc@count\\do{\\global\\advance\\pdc@n\\@ne",
      "\\@ifundefined{pdc@at@\\the\\pdc@n}{}{\\edef\\pdc@args{{\\the\\pdc@n}\\csname pdc@at@\\the\\pdc@n\\endcsname}\\expandafter\\pdc@measure\\pdc@args}}",
      string.format("\\global\\pdc@limit=%sbp\\global\\pdc@n=\\pdc@count", bp(area[2] + area[4] - GAP)),
      "\\@whilenum\\pdc@n>\\z@\\do{\\pdc@lift{\\the\\pdc@n}\\global\\advance\\pdc@n\\m@ne}",
      "\\global\\pdc@n=\\z@\\@whilenum\\pdc@n<\\pdc@count\\do{\\global\\advance\\pdc@n\\@ne\\pdc@draw{\\the\\pdc@n}}}",
    }),
  }
  for number, text in ipairs(texts) do
    lines[#lines + 1] = string.format("\\long\\@namedef{pdc@text@%d}{%s}", number, text)
  end
  lines[#lines + 1] = "\\makeatother"
  return table.concat(lines, "\n")
end

local function add_header_includes(meta, text)
  local block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", text) })
  local existing = meta["header-includes"]
  if existing == nil then
    meta["header-includes"] = pandoc.MetaList({ block })
  elseif existing.t == "MetaList" then
    table.insert(existing, block)
    meta["header-includes"] = existing
  else
    meta["header-includes"] = pandoc.MetaList({ existing, block })
  end
end

function Pandoc(doc)
  local meta = doc.meta
  local mag = tonumber(pandoc.utils.stringify(meta[MAG] or ""))
  local area = numbers(meta[AREA])
  local balloon = numbers(meta[BALLOON])
  local ascent = tonumber(pandoc.utils.stringify(meta[ASCENT] or ""))
  local descent = tonumber(pandoc.utils.stringify(meta[DESCENT] or ""))
  local line = tonumber(pandoc.utils.stringify(meta[BASELINE] or "")) or DEFAULT_LINE
  -- The ascent and descent of the body stay for filters/docx/docx_document_styles_to_latex.lua
  for _, key in ipairs({ MAG, AREA, BALLOON }) do
    meta[key] = nil
  end
  if mag == nil or #area ~= 4 or #balloon ~= 3 or not FORMAT:match("latex") then
    return doc
  end
  ascent = ascent or line * ASCENT_SHARE
  descent = descent or line - ascent
  doc = doc:walk({ Para = keep_line, Plain = keep_line })
  doc = doc:walk({ Inlines = replace_comments })
  add_header_includes(meta, preamble(mag, area, balloon, ascent, descent))
  doc.meta = meta
  return doc
end
