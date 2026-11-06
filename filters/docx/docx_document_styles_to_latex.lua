-- docx_document_styles_to_latex.lua
--
-- Lay a docx out in LaTeX/PDF in the styles Word lays it out in.
--
-- LaTeX justifies and hyphenates its paragraphs, pandoc's template puts 6pt
-- between them, and \maketitle and \section set the title and headings in
-- the sizes of LaTeX. Word follows the styles of the document: for pandoc's
-- own template the body is ragged right, unhyphenated and 9pt apart, and the
-- title and headings take the font, size, colour and spacing of their styles.
--
-- app/docx_document_styles.py reads the styles and passes them as metadata:
--   docx-body-paragraph-skip  the space between two paragraphs, in points
--   docx-body-justified       true when the style justifies the text
--   docx-body-hyphenated      true when the document turns hyphenation on
--   docx-body-size            the size of the body text, in PostScript points
--   docx-body-baseline        the distance of its lines, from its font
--   docx-table-cell-space     the space before and after the paragraph of a
--                             table cell, of the Compact style
--   docx-table-plain-cell-space  the same for a cell paragraph without a
--                             style, of the Normal style
--   docx-table-cell-margin    the left and right margin of a table cell
--   docx-page-numbers         false when no section has a header or footer,
--                             so Word prints no page number
--   docx-style-<StyleId>      the Title or Heading1 to Heading9 style, as
--                             name:value fields separated by commas: font,
--                             size, line, bold, italic, color, before,
--                             after, align
-- This filter turns that metadata into LaTeX in header-includes, after the
-- preamble the other docx filters add there, and removes the metadata.
--
-- The font and the size of the body are not set here: the service passes
-- them as the mainfont and fontsize variables of the template.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx -> pdf/latex).

local SKIP = "docx-body-paragraph-skip"
local JUSTIFIED = "docx-body-justified"
local HYPHENATED = "docx-body-hyphenated"
local SIZE = "docx-body-size"
local BASELINE = "docx-body-baseline"
local DESCENT = "docx-body-descent"
local CELL_SPACE = "docx-table-cell-space"
local LIST_ITEM_SPACE = "docx-list-item-space"
local PARAGRAPH_SPACING = "docx-paragraph-spacing"
local BULLET_ASCENTS = "docx-list-bullet-ascents"
local LIST_INDENTS = "docx-list-indents"
local BULLET_LABELS = { "labelitemi", "labelitemii", "labelitemiii", "labelitemiv" }
local PLAIN_CELL_SPACE = "docx-table-plain-cell-space"
local CELL_MARGIN = "docx-table-cell-margin"
local PAGE_NUMBERS = "docx-page-numbers"
local STYLE_PREFIX = "docx-style-"

-- The LaTeX command of each heading style, and a letter for the macro of its font.
-- LaTeX has five levels of headings and Word nine: titlesec adds the four below \subparagraph.
local HEADINGS = {
  { style = "Heading1", command = "section", letter = "A" },
  { style = "Heading2", command = "subsection", letter = "B" },
  { style = "Heading3", command = "subsubsection", letter = "C" },
  { style = "Heading4", command = "paragraph", letter = "D" },
  { style = "Heading5", command = "subparagraph", letter = "E" },
  { style = "Heading6", command = "pdcHeadingF", letter = "F", deep = true },
  { style = "Heading7", command = "pdcHeadingG", letter = "G", deep = true },
  { style = "Heading8", command = "pdcHeadingH", letter = "H", deep = true },
  { style = "Heading9", command = "pdcHeadingI", letter = "I", deep = true },
}
local DEEPEST = 9
-- titlesec's filling of a line for each alignment, and the plain one for the title.
local TITLESEC_ALIGN = { left = "\\filright", center = "\\filcenter", right = "\\filleft" }
local PLAIN_ALIGN = { left = "\\raggedright", center = "\\centering", right = "\\raggedleft" }
-- The distance of the lines of a style without one: LaTeX's 1.2 times the size.
local LINE_SPREAD = 1.2

-- A metadata value given with -M arrives as a boolean or as a string.
local function is_true(value)
  if type(value) == "boolean" then
    return value
  end
  return value ~= nil and pandoc.utils.stringify(value) == "true"
end

-- A length in points, or nil when it is absent or not a non-negative number.
local function points(value)
  if value == nil then
    return nil
  end
  local number = tonumber(pandoc.utils.stringify(value))
  if number == nil or number < 0 then
    return nil
  end
  return number
end

-- A table as Word sets it: its paragraph skip before it, no space of LaTeX after it and
-- around its rules, and rows a line and the space around the paragraph of the cell
-- tall, shared above and below the text by stretching the strut of the row. The
-- packages of tables come with a document which has one, so each length is set
-- only where it is defined.
local function table_preamble(meta, lines)
  local margin = points(meta[CELL_MARGIN])
  if margin ~= nil then
    lines[#lines + 1] = string.format("\\setlength{\\tabcolsep}{%gpt}", margin)
  end
  local space = points(meta[CELL_SPACE])
  if space == nil then
    return
  end
  lines[#lines + 1] = "\\ifdefined\\LTpre\\setlength{\\LTpre}{\\parskip}\\setlength{\\LTpost}{0pt}\\fi"
  lines[#lines + 1] = "\\ifdefined\\aboverulesep\\setlength{\\aboverulesep}{0pt}\\setlength{\\belowrulesep}{0pt}\\fi"
  local baseline = points(meta[BASELINE])
  if baseline ~= nil and baseline > 0 then
    -- A cell paragraph without a style is in Normal, at the top of its cell with the
    -- space of Normal below it: filters/docx/docx_tables_to_latex.lua adds it to the depth
    -- of each cell, and switches a table of paragraphs of the Compact style to its rows
    local plain = points(meta[PLAIN_CELL_SPACE]) or space
    lines[#lines + 1] = string.format("\\newcommand{\\pdcPlainCellSpace}{%gpt}", plain)
    local descent = points(meta[DESCENT])
    if descent ~= nil and descent > 0 and descent < baseline then
      -- Word sets half the space of a row of Compact above its line and half below it: the stretch
      -- makes the depth of a row, and the extra height of array.sty takes off what it adds above
      local stretch = (descent + space / 2) / descent
      local ascent = baseline - descent
      lines[#lines + 1] = string.format("\\newcommand{\\pdcCompactCellStretch}{%.4f}", stretch)
      lines[#lines + 1] = string.format("\\newcommand{\\pdcCompactCellExtra}{%.4gbp}", (ascent + space / 2) / stretch - ascent)
      lines[#lines + 1] = "\\makeatletter\\newcommand{\\pdcCompactCellDepth}{\\rule[-\\dp\\@arstrutbox]{0pt}{0pt}}\\makeatother"
    else
      -- Without the metrics of the font, the space of a row shared out by the stretch of the line
      lines[#lines + 1] = string.format("\\newcommand{\\pdcCompactCellStretch}{%.4f}", (baseline + space) / baseline)
      lines[#lines + 1] = "\\newcommand{\\pdcCompactCellExtra}{0pt}"
    end
  end
end

-- The pages without a number when the document has none.
local function page_preamble(meta, lines)
  -- Word sets a picture taller than what a page holds at the top of a page, below the lines before it
  -- there, and lets it run past the bottom margin: TeX gets a picture that much shorter, the rest below
  lines[#lines + 1] = "\\newsavebox{\\pdcPictureBox}\\newlength{\\pdcPictureRoom}"
  lines[#lines + 1] = "\\newcommand{\\pdcPagePicture}[1]{\\sbox{\\pdcPictureBox}{#1}\\setlength{\\pdcPictureRoom}{\\dimexpr\\textheight-2\\baselineskip\\relax}"
    .. "\\ifdim\\ht\\pdcPictureBox>\\pdcPictureRoom\\raisebox{\\dimexpr\\pdcPictureRoom-\\ht\\pdcPictureBox\\relax}[\\pdcPictureRoom][0pt]{\\usebox{\\pdcPictureBox}}"
    .. "\\else\\usebox{\\pdcPictureBox}\\fi}"
  lines[#lines + 1] = "\\AtBeginDocument{\\ifdefined\\includegraphics\\NewCommandCopy{\\pdcIncludegraphics}{\\includegraphics}"
    .. "\\RenewDocumentCommand{\\includegraphics}{som}{\\pdcPagePicture{\\IfBooleanTF{#1}{\\IfValueTF{#2}{\\pdcIncludegraphics*[#2]{#3}}{\\pdcIncludegraphics*{#3}}}"
    .. "{\\IfValueTF{#2}{\\pdcIncludegraphics[#2]{#3}}{\\pdcIncludegraphics{#3}}}}}\\fi}"
  if meta[PAGE_NUMBERS] ~= nil and not is_true(meta[PAGE_NUMBERS]) then
    lines[#lines + 1] = "\\makeatletter\\let\\ps@plain\\ps@empty\\makeatother"
    lines[#lines + 1] = "\\pagestyle{empty}"
  end
end

local ROMAN = { "i", "ii", "iii", "iv" }

-- Word sets the rows of a matrix 1.175 times the size of the font apart, and an aligned row its own
-- height and 1.35pt below the one above: TeX the line of the body and 3pt
local MATRIX_ROW = 1.175
-- and rows which would overlap so that times the size of the font apart
local ROW_GAP = 0.25
-- Word sets the line of an equation 0.135 times the size of the font above its top
local EQUATION_TOP = 0.135
local ALIGNED_GAP = 1.35

-- The array of array.sty with the rows of Word: no strut, and a pitch whatever the rows hold
local EQUATION_ARRAY = table.concat({
  "\\makeatletter",
  "\\newcommand{\\pdcEquationArray}[2][c]{\\setbox\\@arstrutbox\\hbox{}%",
  "\\begingroup\\@mkpream{#2}\\xdef\\@preamble{\\noexpand\\ialign\\@halignto\\bgroup\\@arstrut\\@preamble\\tabskip\\z@\\cr}\\endgroup",
  "\\if #1t\\vtop\\else\\if#1b\\vbox\\else\\vcenter\\fi\\fi\\bgroup\\let\\@sharp##\\let\\protect\\relax",
  "\\lineskip\\pdcEquationGap\\baselineskip\\pdcEquationRow\\lineskiplimit\\z@\\mathsurround\\z@",
  "\\let\\\\\\@arraycr\\let\\tabularnewline\\\\\\let\\par\\@empty\\@preamble}",
  "\\newsavebox{\\pdcEquationBox}",
  "\\newcommand{\\pdcInnerFrac}[2]{\\genfrac{}{}{}{1}{\\displaystyle #1}{\\displaystyle #2}}",
  "\\newcommand{\\pdcEquationFrac}[2]{\\genfrac{}{}{}{0}{\\let\\frac\\pdcInnerFrac #1}{\\let\\frac\\pdcInnerFrac #2}}",
  "\\newcommand{\\pdcEquationArrays}{\\def\\@array[##1]##2{\\pdcEquationArray[##1]{##2}}}",
  "\\makeatother",
}, "\n")

-- Word indents the text of each level of a list past the level above it, and sets the label at the
-- start of the hanging indent, aligned left: LaTeX right-aligns it against the text.
local function list_indents(meta, lines)
  local levels = {}
  for left, hanging in pandoc.utils.stringify(meta[LIST_INDENTS] or ""):gmatch("([%d.]+):([%d.]+)") do
    levels[#levels + 1] = { left = tonumber(left), hanging = tonumber(hanging) }
  end
  if #levels ~= #ROMAN then
    return
  end
  lines[#lines + 1] = "\\makeatletter"
  for index, level in ipairs(levels) do
    lines[#lines + 1] = string.format("\\def\\pdcListLeft%s{%gbp}\\def\\pdcListHang%s{%gbp}", ROMAN[index], level.left, ROMAN[index], level.hanging)
  end
  -- Only the list of an itemize or enumerate: a quote keeps its margin, a description its labels
  lines[#lines + 1] = "\\let\\pdcNextList\\relax\\let\\pdcList\\list\\def\\list#1#2{\\pdcList{#1}{#2\\pdcNextList\\global\\let\\pdcNextList\\relax}}"
  lines[#lines + 1] = "\\def\\pdcWordLabel{\\ifcsname pdcListHang\\romannumeral\\@listdepth\\endcsname\\labelsep=0pt"
    .. "\\leftmargin=\\csname pdcListLeft\\romannumeral\\@listdepth\\endcsname\\relax"
    .. "\\labelwidth=\\csname pdcListHang\\romannumeral\\@listdepth\\endcsname\\relax\\def\\makelabel##1{##1\\hfil}\\fi}"
  lines[#lines + 1] = "\\AddToHook{env/itemize/begin}{\\global\\let\\pdcNextList\\pdcWordLabel}\\AddToHook{env/enumerate/begin}{\\global\\let\\pdcNextList\\pdcWordLabel}"
  lines[#lines + 1] = "\\makeatother"
end

local function body_preamble(meta, lines)
  -- Word counts in PostScript points, \fontsize in bp does the same; pt of TeX are 0.4% smaller
  local size, baseline = points(meta[SIZE]), points(meta[BASELINE])
  if size ~= nil and baseline ~= nil then
    -- A line of Word reaches as far down as the descent of its font, and the next
    -- line or picture starts right there: a strut that deep, and no line skip
    local descent = points(meta[DESCENT])
    local strut = descent and descent < baseline and string.format("\\setbox\\strutbox\\hbox{\\vrule height %gbp depth %gbp width 0pt}", baseline - descent, descent) or ""
    -- An equation of Word is a paragraph like any other: no space of a display of its own around it
    lines[#lines + 1] = string.format(
      "\\renewcommand{\\normalsize}{\\fontsize{%gbp}{%gbp}\\selectfont%s"
        .. "\\abovedisplayskip=0pt\\belowdisplayskip=0pt\\abovedisplayshortskip=0pt\\belowdisplayshortskip=0pt}",
      size,
      baseline,
      strut
    )
    lines[#lines + 1] = "\\normalsize"
    if strut ~= "" then
      lines[#lines + 1] = "\\lineskip=0pt\\lineskiplimit=0pt"
      -- tectonic keeps the glyph boxes of a font across its runs by the number of the font, and a
      -- table of contents loads its fonts first from the second run on: a word of it then gets the box
      -- of a glyph of another font. The fonts of the body loaded first keep their numbers in every run
      lines[#lines + 1] = "\\AtBeginDocument{\\setbox0\\hbox{\\normalfont x\\bfseries x\\itshape x\\normalfont\\itshape x}}"
      -- Word sets the first line of a page as far down as the ascent of its font
      lines[#lines + 1] = string.format("\\setlength{\\topskip}{%gbp}", baseline - descent)
      -- A table starts below the descent of the line before it, which TeX takes from its letters
      lines[#lines + 1] = "\\AddToHook{env/longtable/before}{\\par\\ifdim\\prevdepth>-1000pt\\ifdim\\prevdepth<\\dp\\strutbox\\vspace{\\dimexpr\\dp\\strutbox-\\prevdepth\\relax}\\fi\\fi}"
      -- longtable leaves the depth of its last row below the line, not in \\prevdepth: the line after it
      -- would stand a full line of the body below that, where Word sets it below the descent of the row
      lines[#lines + 1] = "\\AddToHook{env/longtable/after}{\\par\\prevdepth=\\dp\\strutbox}"
      -- A table goes on at the top of the next page right at the margin, below its header row if it has one:
      -- longtable sets an empty header there, which TeX would move down by the ascent of a line
      lines[#lines + 1] = "\\AddToHook{env/longtable/begin}{\\topskip=0pt}"
    end
  end
  if points(meta[DESCENT]) ~= nil then
    -- The space before a heading, set right before it: less by as much as the line before reaches less deep
    lines[#lines + 1] = "\\newskip\\pdcBefore\\newcommand{\\pdcSetBefore}[2]{\\par\\pdcBefore=#1\\relax\\ifdim\\prevdepth<#2\\relax\\ifdim\\prevdepth>-1000pt\\advance\\pdcBefore by \\dimexpr\\prevdepth-#2\\relax\\fi\\fi}"
  end
  local skip = points(meta[SKIP])
  -- Word puts the same space after a full stop as between words
  lines[#lines + 1] = "\\frenchspacing"
  -- Word breaks a line after a hyphen of the text as freely as after a space
  lines[#lines + 1] = "\\exhyphenpenalty=0"
  -- Word keeps the first and the last line of a paragraph with another line of it
  -- LaTeX sets \\clubpenalty back to \\@clubpenalty after each heading
  lines[#lines + 1] = "\\makeatletter\\AtBeginDocument{\\clubpenalty=10000\\widowpenalty=10000\\@clubpenalty=10000}\\makeatother"
  if skip ~= nil then
    lines[#lines + 1] = string.format("\\setlength{\\parskip}{%gpt}", skip)
  end
  local item_space = points(meta[LIST_ITEM_SPACE])
  if item_space ~= nil then
    -- pandoc's template sets the items of a tight list without space between them, Word with that of Compact
    lines[#lines + 1] = string.format("\\providecommand{\\tightlist}{}\\renewcommand{\\tightlist}{\\setlength{\\itemsep}{%gpt}\\setlength{\\parskip}{0pt}}", item_space)
  end
  -- Word sets the line of a bullet as high as the font of the bullet reaches, Symbol higher than the text
  local level = 0
  for value in pandoc.utils.stringify(meta[BULLET_ASCENTS] or ""):gmatch("[^,]+") do
    level = level + 1
    local ascent = tonumber(value)
    if BULLET_LABELS[level] and ascent and ascent > 0 then
      local label = BULLET_LABELS[level]
      -- As deep as a line of the body: Word sets the next line below its descent, whatever the letters
      local strut = "\\let\\pdc%s\\%s\\renewcommand{\\%s}{\\rule[-\\dp\\strutbox]{0pt}{\\dimexpr%gbp+\\dp\\strutbox\\relax}\\pdc%s}"
      lines[#lines + 1] = string.format(strut, label, label, label, ascent, label)
    end
  end
  list_indents(meta, lines)
  local size = points(meta[SIZE])
  if size then
    lines[#lines + 1] = string.format("\\newcommand{\\pdcEquationRow}{%gbp}\\newcommand{\\pdcEquationGap}{%gbp}", MATRIX_ROW * size, ROW_GAP * size)
    lines[#lines + 1] = EQUATION_ARRAY
  end
  if meta[JUSTIFIED] ~= nil and not is_true(meta[JUSTIFIED]) then
    -- \\raggedright makes \\\\ end the paragraph, so the lines of a paragraph of line breaks were
    -- paragraphs of their own, which a page broke apart: a line break of Word stays in its paragraph
    lines[#lines + 1] = "\\makeatletter\\AtBeginDocument{\\raggedright\\let\\\\\\@normalcr}\\makeatother"
  end
  if meta[HYPHENATED] ~= nil and not is_true(meta[HYPHENATED]) then
    -- Only the hyphenation of words: Word breaks a line after a hyphen the text holds.
    lines[#lines + 1] = "\\hyphenpenalty=10000"
  end
end

-- LaTeX's counter commands, which the number of a heading may use.
local NUMBER_FORMATS = { arabic = true, alph = true, Alph = true, roman = true, Roman = true }

-- The number of a heading in LaTeX, from its text: digits and punctuation, and
-- each level as {level:format}. nil for any other text.
local function heading_number(text)
  if text == nil or not text:gsub("{%d:%a+}", ""):match("^[%d%.%-%(%) ]*$") then
    return nil
  end
  local valid = true
  local latex = text:gsub("{(%d):(%a+)}", function(level, format)
    local heading = HEADINGS[tonumber(level)]
    if heading == nil or not NUMBER_FORMATS[format] then
      valid = false
      return ""
    end
    return string.format("\\%s{%s}", format, heading.command)
  end)
  return valid and latex or nil
end

-- The fields of a style, each checked: a family of letters, digits, spaces,
-- dots, hyphens and underscores only, numbers for sizes and spaces, six hex
-- digits for the colour. Anything else is left out.
-- The space before and after is the style's own; next to body text, Word puts
-- the larger of it and the paragraph skip.
local function style_fields(value, skip)
  if value == nil then
    return nil
  end
  local fields = {}
  for name, field in pandoc.utils.stringify(value):gmatch("([%a]+):([^,]*)") do
    fields[name] = field
  end
  local style = {}
  if fields.font and fields.font:match("^[%w][%w %.%-_]*$") then
    style.font = fields.font
  end
  style.size = points(fields.size)
  style.line = points(fields.line)
  style.ascent = points(fields.ascent)
  style.bold = fields.bold == "true"
  style.italic = fields.italic == "true"
  if fields.color and fields.color:match("^%x%x%x%x%x%x$") then
    style.color = fields.color
  end
  style.own_before = points(fields.before) or 0
  style.own_after = points(fields.after) or 0
  style.before = math.max(style.own_before, skip or 0)
  style.after = math.max(style.own_after, skip or 0)
  style.align = PLAIN_ALIGN[fields.align] and fields.align or "left"
  style.number = heading_number(fields.number)
  style.number_hang = points(fields.numberhang) or 0
  return style
end

-- Word sets the first line of a page as far down as the ascent of its font, TeX
-- as far as its letters reach: a strut as high as the ascent, for the title at
-- the top of the first page.
local function ascent_strut(style)
  return style.ascent and string.format("\\rule{0pt}{%gbp}", style.ascent) or ""
end

local body_descent = nil
local body_size, body_baseline = nil, nil
local body_ascent = nil
local body_skip = 0

-- Word sets a line below the descent of the line before it, after the space
-- between them, as far down as its own ascent. Each paragraph and heading has
-- a strut of Word's line, so TeX does the same, unless the line of the next
-- font is deeper than the line before: then it keeps a line of that font
-- between the baselines. The space before the next line makes up the difference.
local function heading_rise(style, descent_before)
  if style.line == nil or style.ascent == nil or descent_before == nil then
    return 0
  end
  return descent_before + style.ascent - style.line
end

-- The space before a heading: titlesec takes it from \pdcBefore, which the
-- block before each heading sets from the depth of the line before it.
local function space_before(style)
  if style.line == nil or style.ascent == nil or body_descent == nil then
    return string.format("%gpt", style.before)
  end
  return "\\pdcBefore"
end

local function set_before(style)
  if style == nil or style.line == nil or style.ascent == nil or body_descent == nil then
    return nil
  end
  return pandoc.RawBlock("latex", string.format("\\pdcSetBefore{%gbp}{%.4gbp}", style.before, style.line - style.ascent))
end

-- After the title Word puts the larger of its space after and the next one's
-- before, TeX the space after the title and the paragraph skip. A table of
-- contents has no skip.
local function below_title(style, toc)
  local after = toc and style.own_after or style.after
  if style.line == nil or style.ascent == nil or body_descent == nil then
    return after
  end
  -- The struts of the title and the next line keep their lines Word's distance apart: only the skip is TeX's
  return toc and after or after - body_skip
end

-- The LaTeX which switches to the font, size, weight, shape and colour of a style. With colour_apart the
-- colour comes back on its own, the second result.
local function font_switch(style, family_macro, lines, colour_apart)
  local parts = { "\\normalfont" }
  if style.font then
    lines[#lines + 1] = string.format("\\newfontfamily{\\%s}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}]{%s}", family_macro, style.font)
    parts[#parts + 1] = "\\" .. family_macro
  end
  if style.size then
    parts[#parts + 1] = string.format("\\fontsize{%gbp}{%gbp}\\selectfont", style.size, style.line or style.size * LINE_SPREAD)
    if style.line and style.ascent and style.ascent < style.line then
      -- titlesec sets a strut in a heading: as high and as deep as Word's line of the font
      parts[#parts + 1] = string.format("\\setbox\\strutbox\\hbox{\\vrule height %gbp depth %.4gbp width 0pt}", style.ascent, style.line - style.ascent)
    end
  end
  if style.bold then
    parts[#parts + 1] = "\\bfseries"
  end
  if style.italic then
    parts[#parts + 1] = "\\itshape"
  end
  local colour = style.color and string.format("\\color[HTML]{%s}", style.color) or ""
  if not colour_apart then
    parts[#parts + 1] = colour
  end
  return table.concat(parts), colour
end

-- Whether the document has a heading below level 5, which pandoc writes as a plain paragraph.
local deep_used = false

-- The headings of level 6 to 9: unnumbered, in the table of contents, bold in the body size without a style.
local function deep_levels(lines)
  lines[#lines + 1] = "\\makeatletter"
  for level = 6, DEEPEST do
    local command, parent = HEADINGS[level].command, HEADINGS[level - 1].command
    lines[#lines + 1] = string.format("\\titleclass{\\%s}{straight}[\\%s]", command, parent)
    lines[#lines + 1] = string.format("\\newcounter{%s}\\renewcommand{\\the%s}{}", command, command)
    lines[#lines + 1] = string.format("\\@namedef{toclevel@%s}{%d}", command, level)
    lines[#lines + 1] = string.format("\\titleformat{\\%s}[block]{\\normalfont\\bfseries}{}{0pt}{}", command)
    lines[#lines + 1] = string.format("\\titlespacing*{\\%s}{0pt}{0pt}{0pt}", command)
  end
  lines[#lines + 1] = "\\makeatother"
end

local function headings_preamble(meta, lines)
  local any = false
  local function packages()
    if not any then
      lines[#lines + 1] = "\\usepackage{titlesec}"
      lines[#lines + 1] = "\\usepackage{xcolor}"
      any = true
    end
  end
  if deep_used then
    packages()
    deep_levels(lines)
  end
  local numbered = nil
  for level, heading in ipairs(HEADINGS) do
    local style = style_fields(meta[STYLE_PREFIX .. heading.style], points(meta[SKIP]))
    if style and (deep_used or not heading.deep) then
      packages()
      -- The colour goes inside the line of the heading: in the format, before the line, it leaves a colour
      -- special on the page between the heading and what follows, where TeX may then break the page
      local switch, colour = font_switch(style, "docxHeadingFont" .. heading.letter, lines, true)
      if style.number then
        -- The number of the numbering of the style, in its hanging indent, as Word sets it
        numbered = level
        lines[#lines + 1] = string.format("\\renewcommand{\\the%s}{%s}", heading.command, style.number)
        local label = string.format("\\makebox[%gpt][l]{\\the%s}", style.number_hang, heading.command)
        if style.number_hang == 0 then
          label = string.format("\\the%s\\quad", heading.command)
        end
        lines[#lines + 1] = string.format("\\titleformat{\\%s}[hang]{%s%s}{%s%s}{0pt}{%s}", heading.command, switch, TITLESEC_ALIGN[style.align], colour, label, colour)
      else
        lines[#lines + 1] = string.format("\\titleformat{\\%s}[block]{%s%s}{}{0pt}{%s}", heading.command, switch, TITLESEC_ALIGN[style.align], colour)
      end
      lines[#lines + 1] = string.format("\\titlespacing*{\\%s}{0pt}{%s}{%.4gpt}", heading.command, space_before(style), style.after)
    end
  end
  if numbered then
    -- The template numbers no heading; a level of a numbered style and those above it count
    lines[#lines + 1] = string.format("\\AtBeginDocument{\\setcounter{secnumdepth}{%d}}", numbered)
  end
end

-- The Title style, read in the first pass: the reader makes the first Title
-- paragraph the title of the document, and keeps any later one as a Div.
local title_style = nil

-- \maketitle sets the title alone, in the Title style, at the top of the page as Word does.
-- \pdcTitleStyle switches to it for a later Title paragraph.
-- Whether the document starts with a table of contents: see app/docx_toc_pre_process.py.
local toc_first = false

local function title_preamble(lines)
  if title_style == nil then
    return
  end
  lines[#lines + 1] = "\\usepackage{xcolor}"
  local switch = font_switch(title_style, "docxTitleFont", lines)
  lines[#lines + 1] = string.format("\\newcommand{\\pdcTitleStyle}{%s%s}", PLAIN_ALIGN[title_style.align], switch)
  lines[#lines + 1] = "\\makeatletter"
  -- A table of contents right after the title: its entries have no space before them, so only the title's counts
  local after = below_title(title_style, toc_first)
  -- The strut of the title's font holds its first line as far down as Word, and its last as deep
  local text = (title_style.line and title_style.ascent and body_descent) and "\\strut\\@title\\strut" or ascent_strut(title_style) .. "\\@title"
  lines[#lines + 1] = string.format("\\renewcommand{\\maketitle}{\\par\\begingroup\\pdcTitleStyle%s\\par\\endgroup\\vspace{%gpt}}", text, after)
  lines[#lines + 1] = "\\makeatother"
end

-- The Caption style, which a Polarion caption keeps as PandocCaption: see app/docx_table_pre_process.py.
local caption_style = nil
local CAPTION_DIV = "PandocCaption"

-- A caption in the font of the Caption style; the paragraph skip is the space Word puts around it.
local function caption_preamble(lines)
  if caption_style == nil then
    return
  end
  lines[#lines + 1] = "\\usepackage{xcolor}"
  -- The switch first: it adds the font family to the lines itself
  local switch = font_switch(caption_style, "docxCaptionFont", lines)
  lines[#lines + 1] = string.format("\\newcommand{\\pdcCaptionStyle}{%s}", switch)
  -- The space between a table and its caption, below or above it, see filters/docx/docx_tables_to_latex.lua
  lines[#lines + 1] = string.format("\\newcommand{\\pdcCaptionBefore}{%gpt}", caption_style.own_before)
  lines[#lines + 1] = string.format("\\newcommand{\\pdcCaptionAfter}{%gpt}", caption_style.own_after)
end

local function caption_div(div)
  local blocks = { pandoc.RawBlock("latex", "\\par\\begingroup\\pdcCaptionStyle") }
  for _, block in ipairs(div.content) do
    blocks[#blocks + 1] = block
  end
  blocks[#blocks + 1] = pandoc.RawBlock("latex", "\\par\\endgroup")
  return blocks
end

-- A heading of level 6 to 9 as the command of its level; a deeper one as level 9, as in Word.
local function deep_heading(header)
  if not FORMAT:match("latex") then
    return nil
  end
  if body_descent ~= nil and #header.content > 0 then
    -- The last line of a heading reaches as deep as Word's line of its font, as a paragraph's
    header.content:insert(pandoc.RawInline("latex", "\\texorpdfstring{\\strut}{}"))
  end
  if header.level <= 5 then
    return header
  end
  deep_used = true
  local inlines = { pandoc.RawInline("latex", "\\" .. HEADINGS[math.min(header.level, DEEPEST)].command .. "{") }
  for _, inline in ipairs(header.content) do
    inlines[#inlines + 1] = inline
  end
  inlines[#inlines + 1] = pandoc.RawInline("latex", "}")
  if header.identifier:match("^[%w_%-:.]+$") then
    inlines[#inlines + 1] = pandoc.RawInline("latex", "\\label{" .. header.identifier .. "}")
  end
  -- A heading sets its own space: the space between paragraphs passes it by
  return pandoc.Div({ pandoc.Plain(inlines) }, { ["pdc-style"] = "Heading" })
end

-- An equation of Word is a centred paragraph. TeX opens an empty line above
-- a display which starts a paragraph, and its matrices take the row height
-- of the tables: the equation is set as a centred paragraph in display style.
local function display_equation(math)
  local open = "{\\centering\\renewcommand{\\arraystretch}{1}\\setlength{\\jot}{" .. ALIGNED_GAP .. "pt}"
  local close = "\\par}"
  if body_size then
    -- Word sets a fraction inside a fraction of an equation as large as the outer one, closer to its line.
    -- The rows of an array that far apart whatever they hold, and the array as high as its first row and
    -- as deep as its last: see EQUATION_ARRAY. Word sets the line of an equation that much above it.
    open = open .. "\\let\\frac\\pdcEquationFrac\\pdcEquationArrays\\sbox{\\pdcEquationBox}{"
    close = string.format("}\\rule{0pt}{\\dimexpr\\ht\\pdcEquationBox+%gbp\\relax}\\usebox{\\pdcEquationBox}\\par}", EQUATION_TOP * body_size)
  end
  return pandoc.Plain({
    pandoc.RawInline("latex", open),
    -- Word sets no limit where a sum has an empty one: TeX keeps room above or below it
    pandoc.Math("InlineMath", "\\displaystyle " .. math.text:gsub("[_^]{}", "")),
    pandoc.RawInline("latex", close),
  })
end

-- Word lets the full stop which ends a paragraph stand in the right margin
-- when the line is full: a full stop of no width, which nothing follows.
-- What else a line of only a picture may hold: the struts of the filters and the space around them
local function strut_only(inline)
  return inline.t == "Space" or inline.t == "SoftBreak" or (inline.t == "RawInline" and (inline.text:match("^\\strut") or inline.text:match("^\\rule{0pt}")))
end

-- Word sets a line of only a picture on the bottom of the picture, without the descent of the font,
-- and the next line right below it: the picture goes down by the descent, which the next line then
-- keeps between the baselines.
local function lower_lone_pictures(content)
  local start = 1
  for index = 1, #content + 1 do
    local inline = content[index]
    if inline == nil or inline.t == "LineBreak" then
      local picture, others = nil, false
      for position = start, index - 1 do
        local item = content[position]
        if item.t == "Image" and picture == nil then
          picture = position
        elseif not strut_only(item) then
          others = true
        end
      end
      if picture and not others then
        content[picture] = pandoc.Span({
          pandoc.RawInline("latex", string.format("\\raisebox{-%gbp}{", body_descent)),
          content[picture],
          pandoc.RawInline("latex", "}"),
        })
      end
      start = index + 1
    end
  end
end

local last_word

-- Word lets the period which ends a line hang past the right margin: only the word before it has to
-- fit. A period before a space takes no width of its own, and the space after it as much more: at the
-- end of a line the space goes, and the period stands past the margin.
local HANGING_PERIOD = "\\rlap{.}\\hskip\\dimexpr\\fontcharwd\\font`.+\\fontdimen2\\font\\relax "

local function hang_periods(inlines)
  local changed = false
  for index = #inlines - 1, 1, -1 do
    local str, after = inlines[index], inlines[index + 1]
    if str.t == "Str" and #str.text > 1 and str.text:sub(-1) == "." and (after.t == "Space" or after.t == "SoftBreak") then
      inlines[index] = pandoc.Str(str.text:sub(1, -2))
      inlines[index + 1] = pandoc.RawInline("latex", HANGING_PERIOD)
      changed = true
    end
  end
  return changed and inlines or nil
end

-- The list and the position of the last word of inlines, past the closing braces of raw groups and
-- into the last element which holds inlines
last_word = function(inlines)
  for index = #inlines, 1, -1 do
    local inline = inlines[index]
    if inline.t == "Str" then
      return inlines, index
    elseif inline.t == "RawInline" and inline.text:match("^}+$") then
      -- the end of a group of the run: the word is before it
    elseif inline.content and inline.t ~= "Note" then
      return last_word(inline.content)
    else
      return nil
    end
  end
  return nil
end

local function set_paragraph(para)
  if not FORMAT:match("latex") then
    return nil
  end
  local only = #para.content == 1 and para.content[1]
  if only and only.t == "Math" and only.mathtype == "DisplayMath" then
    return display_equation(only)
  end
  local changed = false
  local list, index = last_word(para.content)
  local last = list and list[index]
  if last and #last.text > 1 and last.text:sub(-1) == "." then
    list[index] = pandoc.Str(last.text:sub(1, -2))
    list:insert(index + 1, pandoc.RawInline("latex", "\\rlap{.}"))
    changed = true
  end
  local visible = false
  for _, inline in ipairs(para.content) do
    visible = visible or inline.t ~= "RawInline"
  end
  -- A paragraph of LaTeX alone, as a page break, sets no line of its own
  if body_descent ~= nil and visible then
    lower_lone_pictures(para.content)
    -- Every line of Word reaches the ascent and the descent of its font, whatever its letters:
    -- the first and the last line of a paragraph get the strut of its font
    para.content:insert(1, pandoc.RawInline("latex", "\\strut{}"))
    para.content:insert(pandoc.RawInline("latex", "\\strut{}"))
    changed = true
  end
  return changed and para or nil
end

-- A Title paragraph after the first one, in the Title style with its space around it.
local function title_div(div)
  if caption_style ~= nil and div.attributes["custom-style"] == CAPTION_DIV then
    -- Its space comes from the space between paragraphs, next to a table from docx_tables_to_latex
    return pandoc.Div(caption_div(div), { ["pdc-style"] = "Caption" })
  end
  if title_style == nil or div.attributes["custom-style"] ~= "Title" then
    return nil
  end
  -- The space between paragraphs puts the title's space around it. TeX keeps a line of the title's
  -- font between the baselines, Word the descent of the line above and the ascent of the title.
  local open = "\\par\\begingroup\\pdcTitleStyle"
  if title_style.line and title_style.ascent and body_descent then
    local descent = title_style.line - title_style.ascent
    open = string.format("\\par\\ifdim\\prevdepth<%.4gbp\\ifdim\\prevdepth>-1000pt\\vskip\\dimexpr\\prevdepth-%.4gbp\\relax\\fi\\fi\\begingroup\\pdcTitleStyle", descent, descent)
  end
  local blocks = { pandoc.RawBlock("latex", open) }
  for _, block in ipairs(div.content) do
    blocks[#blocks + 1] = block
  end
  blocks[#blocks + 1] = pandoc.RawBlock("latex", "\\par\\endgroup")
  return pandoc.Div(blocks, { ["pdc-style"] = "Title" })
end

local function add_header_includes(meta, lines)
  local block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", table.concat(lines, "\n")) })
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

-- The styles of the headings by level, read with the title.
local heading_styles = {}
local function read_first_block(doc)
  local first = doc.blocks[1]
  local style = first ~= nil and first.t == "Div" and first.attributes["custom-style"] or ""
  toc_first = style:sub(1, #"PandocToc__CONTENTS_") == "PandocToc__CONTENTS_"
  return nil
end

-- The space before and after each paragraph style of the body, by id and by name.
local paragraph_spacing = { ids = {}, names = {} }

-- The spacing of a block which is a paragraph of the body, or nil: a paragraph
-- without a style is Normal; the reader keeps another style as a Div, and
-- docx_paragraphs_to_latex the style of an aligned or indented one.
-- A table has no space before or after it: Word puts the space of the paragraph next to it
local TABLE_SPACING = { before = 0, after = 0 }
-- An entry of a list of figures or tables, in Word's TOC style: docx_toc_to_latex sets it without
-- the paragraph skip and with its 5pt after it
local ENTRY_SPACING = { before = 0, after = 5, no_skip = true, trailing = 5 }
-- Raw LaTeX between blocks, as the groups around a table, which keeps the paragraph before
local TRANSPARENT = {}

-- A table docx_tables_to_latex writes as LaTeX: a raw block, or a Plain of raw LaTeX and its pictures
local function raw_table(block)
  local first = block.t == "Plain" and block.content[1] or block
  return first ~= nil and (first.t == "RawBlock" or first.t == "RawInline") and first.text:find("\\begin{longtable}", 1, true) ~= nil
end

local function block_spacing(block)
  if block.t == "Table" or raw_table(block) then
    return TABLE_SPACING
  end
  if block.t == "Para" or block.t == "Plain" then
    return paragraph_spacing.ids.Normal
  end
  if block.t == "RawBlock" then
    return TRANSPARENT
  end
  if block.t == "BulletList" or block.t == "OrderedList" then
    -- The items of a tight list are Compact paragraphs, of a loose one Normal
    local first = block.content[1] and block.content[1][1]
    local tight = first ~= nil and first.t == "Plain"
    return tight and paragraph_spacing.ids.Compact or paragraph_spacing.ids.Normal
  end
  if block.t == "Div" and block.attributes["custom-style"] == "PandocToc__ENTRY" then
    return ENTRY_SPACING
  end
  if block.t == "Div" and block.attributes["pdc-style"] == "Caption" and caption_style ~= nil then
    return { before = caption_style.own_before, after = caption_style.own_after, caption = true }
  end
  if block.t == "Div" and block.attributes["pdc-style"] == "Title" and title_style ~= nil then
    return { before = title_style.own_before, after = title_style.own_after }
  end
  if block.t == "Div" then
    local id = block.attributes["pdc-style"]
    if id then
      return paragraph_spacing.ids[id]
    end
    return paragraph_spacing.names[block.attributes["custom-style"] or ""]
  end
  return nil
end

-- Word puts the larger of the space after a paragraph and before the next
-- between them, TeX the paragraph skip: the difference goes between them.
local function gap(previous, spacing)
  if previous == nil or spacing == nil then
    return nil
  end
  if (previous.caption and spacing == TABLE_SPACING) or (previous == TABLE_SPACING and spacing.caption) then
    -- docx_tables_to_latex spaces a caption next to its table
    return nil
  end
  local delta = math.max(previous.after, spacing.before) - (spacing.no_skip and 0 or body_skip) - (previous.trailing or 0)
  return math.abs(delta) > 0.01 and pandoc.RawBlock("latex", string.format("\\vspace{%gbp}", delta)) or nil
end

-- Word keeps the space before the first paragraph of a document at the top of
-- its page, where TeX drops it: a document without a title starts that far down.
local function paragraph_gaps(blocks, starts_page)
  local result, previous = {}, nil
  for _, block in ipairs(blocks) do
    local spacing = block_spacing(block)
    if spacing == TRANSPARENT then
      result[#result + 1] = block
      goto next_block
    end
    if starts_page and spacing ~= nil and spacing.before > 0 and body_ascent ~= nil then
      -- The first line as high as the space and the ascent: \\topskip holds no more than the ascent
      local first = block.t == "Div" and block.content[1] or block
      if first ~= nil and (first.t == "Para" or first.t == "Plain") then
        first.content:insert(1, pandoc.RawInline("latex", string.format("\\rule{0pt}{%gbp}", spacing.before + body_ascent)))
      end
    end
    starts_page = false
    result[#result + 1] = gap(previous, spacing)
    if spacing and block.t == "Div" and block.attributes["pdc-style"] == nil then
      -- The paragraphs of one style the reader groups into one Div: each one its gap
      local content = {}
      for index, inner in ipairs(block.content) do
        if index > 1 and (inner.t == "Para" or inner.t == "Plain") then
          content[#content + 1] = gap(spacing, spacing)
        end
        content[#content + 1] = inner
      end
      block.content = content
    end
    result[#result + 1] = block
    previous = spacing
    ::next_block::
  end
  return result
end

local function read_title(meta)
  local skip = points(meta[SKIP])
  title_style = style_fields(meta[STYLE_PREFIX .. "Title"], skip)
  body_descent = points(meta[DESCENT])
  body_ascent = points(meta["docx-body-ascent"])
  body_size, body_baseline = points(meta[SIZE]), points(meta[BASELINE])
  body_skip = skip or 0
  for entry in pandoc.utils.stringify(meta[PARAGRAPH_SPACING] or ""):gmatch("[^;]+") do
    local id, name, before, after = entry:match("^(%w+):([%w ]+):([%d.]+):([%d.]+)$")
    if id then
      local spacing = { before = tonumber(before), after = tonumber(after) }
      paragraph_spacing.ids[id], paragraph_spacing.names[name] = spacing, spacing
    end
  end
  caption_style = style_fields(meta[STYLE_PREFIX .. "Caption"], skip)
  for level, heading in ipairs(HEADINGS) do
    heading_styles[level] = style_fields(meta[STYLE_PREFIX .. heading.style], skip)
  end
  return nil
end

local function heading_level(block)
  return block ~= nil and block.t == "Header" and math.min(block.level, DEEPEST) or nil
end

-- Between two headings Word puts the larger of their own spaces, without the
-- paragraph skip the spacing of the preamble holds for body text. LaTeX sets
-- no space before a heading right after another, so a heading followed by
-- another gets that space after it, and its own spacing back after it.
local function heading_spacing(blocks)
  if not FORMAT:match("latex") then
    return nil
  end
  local result = {}
  for index, block in ipairs(blocks) do
    local level = heading_level(block)
    local style = level and heading_styles[level]
    local next = style and heading_styles[heading_level(blocks[index + 1])]
    result[#result + 1] = set_before(style)
    if next then
      local spacing = "\\titlespacing*{\\%s}{0pt}{%s}{%.4gpt}"
      local command = HEADINGS[level].command
      local before = space_before(style)
      -- The struts of both reach Word's line: TeX keeps the descent of the first and the ascent of the
      -- next apart unless the next line of its own font is deeper than that
      local descent = style.line and style.ascent and style.line - style.ascent or body_descent
      local after = math.max(style.own_after, next.own_before) + math.min(0, heading_rise(next, descent))
      result[#result + 1] = pandoc.RawBlock("latex", string.format(spacing, command, before, after))
      result[#result + 1] = block
      result[#result + 1] = pandoc.RawBlock("latex", string.format(spacing, command, before, style.after))
    elseif style then
      -- Word puts the larger of the space after the heading and before the next block, a table none
      local ahead = index + 1
      while blocks[ahead] and block_spacing(blocks[ahead]) == TRANSPARENT do
        ahead = ahead + 1
      end
      local following = blocks[ahead] and block_spacing(blocks[ahead])
      local after = following and math.max(style.own_after, following.before)
      if after and math.abs(after - style.after) > 0.01 then
        local spacing = "\\titlespacing*{\\%s}{0pt}{%s}{%.4gpt}"
        local command, before = HEADINGS[level].command, space_before(style)
        result[#result + 1] = pandoc.RawBlock("latex", string.format(spacing, command, before, after))
        result[#result + 1] = block
        result[#result + 1] = pandoc.RawBlock("latex", string.format(spacing, command, before, style.after))
      else
        result[#result + 1] = block
      end
    else
      result[#result + 1] = block
    end
  end
  return result
end

local function write_preamble(meta)
  if not FORMAT:match("latex") then
    return nil
  end
  local lines = {}
  body_preamble(meta, lines)
  page_preamble(meta, lines)
  table_preamble(meta, lines)
  headings_preamble(meta, lines)
  title_preamble(lines)
  caption_preamble(lines)
  for key, _ in pairs(meta) do
    if key == SKIP or key == JUSTIFIED or key == HYPHENATED or key == SIZE or key == PARAGRAPH_SPACING or key == LIST_ITEM_SPACE or key == BULLET_ASCENTS or key == LIST_INDENTS or key == BASELINE or key == DESCENT or key == "docx-body-ascent" or key == CELL_SPACE or key == PLAIN_CELL_SPACE or key == CELL_MARGIN or key == PAGE_NUMBERS or key:sub(1, #STYLE_PREFIX) == STYLE_PREFIX then
      meta[key] = nil
    end
  end
  if #lines > 0 then
    add_header_includes(meta, lines)
  end
  return meta
end

-- The Title, Caption and heading styles are read before the blocks are set, the preamble written after.
return {
  { Meta = read_title, Pandoc = read_first_block },
  { Blocks = heading_spacing },
  { Div = title_div, Header = deep_heading, Para = set_paragraph, Meta = write_preamble },
  { Para = function(para) return para:walk({ Inlines = hang_periods }) end, Plain = function(plain) return plain:walk({ Inlines = hang_periods }) end },
  {
    Pandoc = function(doc)
      if FORMAT:match("latex") then
        doc.blocks = paragraph_gaps(doc.blocks, doc.meta.title == nil)
      end
      return doc
    end,
  },
}
