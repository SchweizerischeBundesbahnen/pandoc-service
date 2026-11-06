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
--   docx-link-color           the colour of the Hyperlink style
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
local CELL_SPACE = "docx-table-cell-space"
local PLAIN_CELL_SPACE = "docx-table-plain-cell-space"
local CELL_MARGIN = "docx-table-cell-margin"
local LINK_COLOR = "docx-link-color"
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
    -- space of Normal below it: filters/docx_tables_to_latex.lua adds it to the depth
    -- of each cell, and switches a table of paragraphs of the Compact style to its rows
    local plain = points(meta[PLAIN_CELL_SPACE]) or space
    lines[#lines + 1] = string.format("\\newcommand{\\pdcPlainCellSpace}{%gpt}", plain)
    lines[#lines + 1] = string.format("\\newcommand{\\pdcCompactCellStretch}{%.4f}", (baseline + space) / baseline)
  end
end

-- Links in the colour of the Hyperlink style, the pages without a number when the document has none.
local function page_preamble(meta, lines)
  local color = meta[LINK_COLOR] and pandoc.utils.stringify(meta[LINK_COLOR]) or ""
  if color:match("^%x%x%x%x%x%x$") then
    lines[#lines + 1] = "\\usepackage{xcolor}"
    -- The service passes colorlinks and the colours of the template as pdcLink, set up after this preamble
    lines[#lines + 1] = string.format("\\definecolor{pdcLink}{HTML}{%s}", color)
  end
  if meta[PAGE_NUMBERS] ~= nil and not is_true(meta[PAGE_NUMBERS]) then
    lines[#lines + 1] = "\\makeatletter\\let\\ps@plain\\ps@empty\\makeatother"
    lines[#lines + 1] = "\\pagestyle{empty}"
  end
end

-- A new page of another size and margins: the paper, the PDF page and the text area of LaTeX.
local PAGE_GEOMETRY = table.concat({
  "\\makeatletter",
  "\\newcommand{\\pdcPageGeometry}[6]{\\clearpage",
  "\\setlength{\\paperwidth}{#1bp}\\setlength{\\paperheight}{#2bp}%",
  "\\setlength{\\pdfpagewidth}{\\paperwidth}\\setlength{\\pdfpageheight}{\\paperheight}%",
  "\\setlength{\\textwidth}{\\dimexpr\\paperwidth-#4bp-#6bp\\relax}%",
  "\\setlength{\\textheight}{\\dimexpr\\paperheight-#3bp-#5bp\\relax}%",
  "\\setlength{\\oddsidemargin}{\\dimexpr#6bp-1in\\relax}\\setlength{\\evensidemargin}{\\oddsidemargin}%",
  "\\setlength{\\topmargin}{\\dimexpr#3bp-1in-\\headheight-\\headsep\\relax}%",
  "\\setlength{\\linewidth}{\\textwidth}\\setlength{\\columnwidth}{\\textwidth}\\hsize\\textwidth",
  "\\global\\@colht\\textheight\\global\\@colroom\\textheight\\global\\vsize\\textheight}",
  "\\makeatother",
}, "\n")

local function body_preamble(meta, lines)
  lines[#lines + 1] = PAGE_GEOMETRY
  -- Word counts in PostScript points, \fontsize in bp does the same; pt of TeX are 0.4% smaller
  local size, baseline = points(meta[SIZE]), points(meta[BASELINE])
  if size ~= nil and baseline ~= nil then
    -- An equation of Word is a paragraph like any other: no space of a display of its own around it
    lines[#lines + 1] = string.format(
      "\\renewcommand{\\normalsize}{\\fontsize{%gbp}{%gbp}\\selectfont"
        .. "\\abovedisplayskip=0pt\\belowdisplayskip=0pt\\abovedisplayshortskip=0pt\\belowdisplayshortskip=0pt}",
      size,
      baseline
    )
    lines[#lines + 1] = "\\normalsize"
  end
  local skip = points(meta[SKIP])
  -- Word puts the same space after a full stop as between words
  lines[#lines + 1] = "\\frenchspacing"
  -- Word breaks a line after a hyphen of the text as freely as after a space
  lines[#lines + 1] = "\\exhyphenpenalty=0"
  if skip ~= nil then
    lines[#lines + 1] = string.format("\\setlength{\\parskip}{%gpt}", skip)
  end
  if meta[JUSTIFIED] ~= nil and not is_true(meta[JUSTIFIED]) then
    lines[#lines + 1] = "\\AtBeginDocument{\\raggedright}"
  end
  if meta[HYPHENATED] ~= nil and not is_true(meta[HYPHENATED]) then
    -- Only the hyphenation of words: Word breaks a line after a hyphen the text holds.
    lines[#lines + 1] = "\\hyphenpenalty=10000"
  end
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
  return style
end

-- Word sets the first line of a page as far down as the ascent of its font, TeX
-- as far as its letters reach: a strut as high as the ascent, for the title at
-- the top of the first page.
local function ascent_strut(style)
  return style.ascent and string.format("\\rule{0pt}{%gbp}", style.ascent) or ""
end

-- The LaTeX which switches to the font, size, weight, shape and colour of a style.
local function font_switch(style, family_macro, lines)
  local parts = { "\\normalfont" }
  if style.font then
    lines[#lines + 1] = string.format("\\newfontfamily{\\%s}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}]{%s}", family_macro, style.font)
    parts[#parts + 1] = "\\" .. family_macro
  end
  if style.size then
    parts[#parts + 1] = string.format("\\fontsize{%gbp}{%gbp}\\selectfont", style.size, style.line or style.size * LINE_SPREAD)
  end
  if style.bold then
    parts[#parts + 1] = "\\bfseries"
  end
  if style.italic then
    parts[#parts + 1] = "\\itshape"
  end
  if style.color then
    parts[#parts + 1] = string.format("\\color[HTML]{%s}", style.color)
  end
  return table.concat(parts)
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
  for _, heading in ipairs(HEADINGS) do
    local style = style_fields(meta[STYLE_PREFIX .. heading.style], points(meta[SKIP]))
    if style and (deep_used or not heading.deep) then
      packages()
      local switch = font_switch(style, "docxHeadingFont" .. heading.letter, lines)
      lines[#lines + 1] = string.format("\\titleformat{\\%s}[block]{%s%s}{}{0pt}{}", heading.command, switch, TITLESEC_ALIGN[style.align])
      lines[#lines + 1] = string.format("\\titlespacing*{\\%s}{0pt}{%gpt}{%gpt}", heading.command, style.before, style.after)
    end
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
  local after = toc_first and title_style.own_after or title_style.after
  lines[#lines + 1] = string.format("\\renewcommand{\\maketitle}{\\par\\begingroup\\pdcTitleStyle%s\\@title\\par\\endgroup\\vspace{%gpt}}", ascent_strut(title_style), after)
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
end

local function caption_div(div)
  local blocks = { pandoc.RawBlock("latex", "\\par\\begingroup\\pdcCaptionStyle") }
  for _, block in ipairs(div.content) do
    blocks[#blocks + 1] = block
  end
  blocks[#blocks + 1] = pandoc.RawBlock("latex", "\\par\\endgroup")
  return blocks
end

-- A section on a new page of its own size and margins: six numbers in points.
local function section_div(div)
  local numbers = {}
  for value in pandoc.utils.stringify(div):gmatch("%S+") do
    local number = tonumber(value)
    if number == nil or number < 0 then
      return pandoc.RawBlock("latex", "\\clearpage")
    end
    numbers[#numbers + 1] = string.format("%g", number)
  end
  if #numbers ~= 6 then
    return pandoc.RawBlock("latex", "\\clearpage")
  end
  return pandoc.RawBlock("latex", "\\pdcPageGeometry{" .. table.concat(numbers, "}{") .. "}")
end

-- A heading of level 6 to 9 as the command of its level; a deeper one as level 9, as in Word.
local function deep_heading(header)
  if header.level <= 5 or not FORMAT:match("latex") then
    return nil
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
  return pandoc.Plain(inlines)
end

-- An equation of Word is a centred paragraph. TeX opens an empty line above
-- a display which starts a paragraph, and its matrices take the row height
-- of the tables: the equation is set as a centred paragraph in display style.
local function display_equation(math)
  return pandoc.Plain({
    pandoc.RawInline("latex", "{\\centering\\renewcommand{\\arraystretch}{1}"),
    pandoc.Math("InlineMath", "\\displaystyle " .. math.text),
    pandoc.RawInline("latex", "\\par}"),
  })
end

-- Word lets the full stop which ends a paragraph stand in the right margin
-- when the line is full: a full stop of no width, which nothing follows.
local function set_paragraph(para)
  if not FORMAT:match("latex") then
    return nil
  end
  local only = #para.content == 1 and para.content[1]
  if only and only.t == "Math" and only.mathtype == "DisplayMath" then
    return display_equation(only)
  end
  local last = para.content[#para.content]
  if last ~= nil and last.t == "Str" and #last.text > 1 and last.text:sub(-1) == "." then
    para.content[#para.content] = pandoc.Str(last.text:sub(1, -2))
    para.content:insert(pandoc.RawInline("latex", "\\rlap{.}"))
    return para
  end
  return nil
end

-- A Title paragraph after the first one, in the Title style with its space around it.
local function title_div(div)
  -- A page break and a section of the docx: see app/docx_page_break_pre_process.py
  if div.attributes["custom-style"] == "PandocPageBreak" then
    return pandoc.RawBlock("latex", "\\clearpage")
  end
  if div.attributes["custom-style"] == "PandocSection" then
    return section_div(div)
  end
  if caption_style ~= nil and div.attributes["custom-style"] == CAPTION_DIV then
    return caption_div(div)
  end
  if title_style == nil or div.attributes["custom-style"] ~= "Title" then
    return nil
  end
  local blocks = { pandoc.RawBlock("latex", string.format("\\par\\vspace{%gpt}\\begingroup\\pdcTitleStyle", title_style.before)) }
  for _, block in ipairs(div.content) do
    blocks[#blocks + 1] = block
  end
  blocks[#blocks + 1] = pandoc.RawBlock("latex", string.format("\\par\\endgroup\\vspace{%gpt}", title_style.after))
  return blocks
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

local function read_title(meta)
  local skip = points(meta[SKIP])
  title_style = style_fields(meta[STYLE_PREFIX .. "Title"], skip)
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
    if next then
      local spacing = "\\titlespacing*{\\%s}{0pt}{%gpt}{%gpt}"
      local command = HEADINGS[level].command
      local after = math.max(style.own_after, next.own_before)
      result[#result + 1] = pandoc.RawBlock("latex", string.format(spacing, command, style.before, after))
      result[#result + 1] = block
      result[#result + 1] = pandoc.RawBlock("latex", string.format(spacing, command, style.before, style.after))
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
    if key == SKIP or key == JUSTIFIED or key == HYPHENATED or key == SIZE or key == BASELINE or key == LINK_COLOR or key == CELL_SPACE or key == PLAIN_CELL_SPACE or key == CELL_MARGIN or key == PAGE_NUMBERS or key:sub(1, #STYLE_PREFIX) == STYLE_PREFIX then
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
}
