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
--   docx-link-color           the colour of the Hyperlink style
--   docx-page-numbers         false when no section has a header or footer,
--                             so Word prints no page number
--   docx-style-<StyleId>      the Title or Heading1 to Heading5 style, as
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
local LINK_COLOR = "docx-link-color"
local PAGE_NUMBERS = "docx-page-numbers"
local STYLE_PREFIX = "docx-style-"

-- The LaTeX command of each heading style, and a letter for the macro of its font.
local HEADINGS = {
  { style = "Heading1", command = "section", letter = "A" },
  { style = "Heading2", command = "subsection", letter = "B" },
  { style = "Heading3", command = "subsubsection", letter = "C" },
  { style = "Heading4", command = "paragraph", letter = "D" },
  { style = "Heading5", command = "subparagraph", letter = "E" },
}
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
  local space = points(meta[CELL_SPACE])
  if space == nil then
    return
  end
  lines[#lines + 1] = "\\ifdefined\\LTpre\\setlength{\\LTpre}{\\parskip}\\setlength{\\LTpost}{0pt}\\fi"
  lines[#lines + 1] = "\\ifdefined\\aboverulesep\\setlength{\\aboverulesep}{0pt}\\setlength{\\belowrulesep}{0pt}\\fi"
  local baseline = points(meta[BASELINE])
  if baseline ~= nil and baseline > 0 then
    lines[#lines + 1] = string.format("\\renewcommand{\\arraystretch}{%.4f}", (baseline + space) / baseline)
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

local function body_preamble(meta, lines)
  -- Word counts in PostScript points, \fontsize in bp does the same; pt of TeX are 0.4% smaller
  local size, baseline = points(meta[SIZE]), points(meta[BASELINE])
  if size ~= nil and baseline ~= nil then
    lines[#lines + 1] = string.format("\\renewcommand{\\normalsize}{\\fontsize{%gbp}{%gbp}\\selectfont}", size, baseline)
    lines[#lines + 1] = "\\normalsize"
  end
  local skip = points(meta[SKIP])
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
local function style_fields(value)
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
  style.bold = fields.bold == "true"
  style.italic = fields.italic == "true"
  if fields.color and fields.color:match("^%x%x%x%x%x%x$") then
    style.color = fields.color
  end
  style.before = points(fields.before) or 0
  style.after = points(fields.after) or 0
  style.align = PLAIN_ALIGN[fields.align] and fields.align or "left"
  return style
end

-- The LaTeX which switches to the font, size, weight, shape and colour of a style.
local function font_switch(style, family_macro, lines)
  local parts = { "\\normalfont" }
  if style.font then
    lines[#lines + 1] = string.format("\\newfontfamily{\\%s}[Kerning=Off]{%s}", family_macro, style.font)
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

local function headings_preamble(meta, lines)
  local any = false
  for _, heading in ipairs(HEADINGS) do
    local style = style_fields(meta[STYLE_PREFIX .. heading.style])
    if style then
      if not any then
        lines[#lines + 1] = "\\usepackage{titlesec}"
        lines[#lines + 1] = "\\usepackage{xcolor}"
        any = true
      end
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
local function title_preamble(lines)
  if title_style == nil then
    return
  end
  lines[#lines + 1] = "\\usepackage{xcolor}"
  local switch = font_switch(title_style, "docxTitleFont", lines)
  lines[#lines + 1] = string.format("\\newcommand{\\pdcTitleStyle}{%s%s}", PLAIN_ALIGN[title_style.align], switch)
  lines[#lines + 1] = "\\makeatletter"
  lines[#lines + 1] = string.format("\\renewcommand{\\maketitle}{\\par\\begingroup\\pdcTitleStyle\\@title\\par\\endgroup\\vspace{%gpt}}", title_style.after)
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

-- A Title paragraph after the first one, in the Title style with its space around it.
local function title_div(div)
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

local function read_title(meta)
  title_style = style_fields(meta[STYLE_PREFIX .. "Title"])
  caption_style = style_fields(meta[STYLE_PREFIX .. "Caption"])
  return nil
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
    if key == SKIP or key == JUSTIFIED or key == HYPHENATED or key == SIZE or key == BASELINE or key == LINK_COLOR or key == CELL_SPACE or key == PAGE_NUMBERS or key:sub(1, #STYLE_PREFIX) == STYLE_PREFIX then
      meta[key] = nil
    end
  end
  if #lines > 0 then
    add_header_includes(meta, lines)
  end
  return meta
end

-- The Title and Caption styles are read before the Divs are set, the preamble written after.
return {
  { Meta = read_title },
  { Div = title_div, Meta = write_preamble },
}
