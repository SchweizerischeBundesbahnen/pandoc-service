-- docx_page_breaks_to_latex.lua
--
-- Keep the page and section breaks of a docx in LaTeX/PDF.
--
-- The docx reader drops both. app/docx_page_break_pre_process.py writes
-- them into the text as sentinels before pandoc reads the document:
--
--   Str "\u{E020}\u{E021}"                 a page break
--   Str "\u{E020}w,h,t,r,b,l,start\u{E021}"  a section break: the page and
--                                          margins it opens in twips, and
--                                          nextPage, oddPage or evenPage
--
-- This filter replaces a sentinel with \clearpage where it stands. A
-- section break then sets the page size and the margins of the next page,
-- so a landscape section comes out landscape. That setup goes after the
-- body block holding the break: docx_paragraphs_to_latex wraps a formatted
-- paragraph in a TeX group, which would undo it. The sentinel is document
-- text, so its numbers are checked and only they reach the raw LaTeX.
--
-- An oddPage or evenPage section adds an empty page when the next page
-- has the wrong number, as Word does.
--
-- A heading argument is copied to the table of contents and the page
-- headers, so a break never stays in it. A break before the text of the
-- heading goes before it, any other break after it.
--
-- LaTeX ignores \clearpage on an empty page, where Word starts one more
-- page. A break with nothing before it on its page gets \null first.
--
-- Only paragraphs of the body are changed, also inside a Div (the reader
-- wraps a styled paragraph in one). A sentinel anywhere else, a list item
-- or the title in the metadata, is removed: no page breaks there.
--
-- Must run before docx_empty_paragraphs_to_latex: a paragraph holding only
-- a break is no blank line. Only runs for the LaTeX writer (gated in
-- PandocController on docx -> pdf/latex).

local OPEN = "\u{E020}"
local CLOSE = "\u{E021}"
local SENTINEL_PATTERN = OPEN .. "([%w,]*)" .. CLOSE

-- The marker of a paragraph which starts a new page before it, see app/docx_page_break_pre_process.py
local BREAK_BEFORE = "before"

-- More than Word's largest page, and small enough for a TeX dimension.
local MAX_TWIPS = 99999
local TWIPS_PER_INCH = 1440

-- What a section start adds after \clearpage. This is a trust boundary: only
-- the literal values of this table reach the output.
local PAGE_STARTS = {
  nextPage = "",
  oddPage = "\\ifodd\\value{page}\\else\\null\\clearpage\\fi",
  evenPage = "\\ifodd\\value{page}\\null\\clearpage\\fi",
}

local BLANK_INLINES = {
  Space = true,
  SoftBreak = true,
  LineBreak = true,
}

-- The inlines which show on the page without holding a Str.
local VISIBLE_INLINES = {
  Code = true,
  Image = true,
  Math = true,
  Note = true,
}

-- Whether the current page has nothing on it yet.
local on_empty_page = true

-- The page setup of the last section break, until its body block ends.
local pending_page_setup = ""

-- Whether the pages keep the size of the first section: docx_comments_to_latex magnifies a document
-- with comments and puts the balloons where its first page has room for them.
local keep_first_page = false

-- Whether add_heading has passed visible content of the heading.
local heading_has_text = false

local function inches(twips)
  return string.format("%.4fin", twips / TWIPS_PER_INCH)
end

-- The LaTeX which sets the size and margins of a page, or "" when the numbers
-- describe no page LaTeX can set.
local function page_geometry(numbers)
  local w, h, top, right, bottom, left = numbers:match("^(%d+),(%d+),(%d+),(%d+),(%d+),(%d+)$")
  if not w or keep_first_page then
    return ""
  end
  w, h, top, right, bottom, left = tonumber(w), tonumber(h), tonumber(top), tonumber(right), tonumber(bottom), tonumber(left)
  if w > MAX_TWIPS or h > MAX_TWIPS or w <= left + right or h <= top + bottom then
    return ""
  end
  -- \newgeometry leaves the paper alone and measures the layout, so it gets both sizes.
  local width, height = inches(w), inches(h)
  return string.format(
    "\\pdfpagewidth=%s\\pdfpageheight=%s\\paperwidth=%s\\paperheight=%s\\newgeometry{layoutwidth=%s,layoutheight=%s,top=%s,right=%s,bottom=%s,left=%s}",
    width, height, width, height, width, height, inches(top), inches(right), inches(bottom), inches(left)
  )
end

-- The LaTeX which sets up the page a section break opens, or "" for a sentinel
-- which is no section break.
local function page_setup(marker)
  local numbers, start = marker:match("^([%d,]*),(%a+)$")
  local page_start = PAGE_STARTS[start]
  if not page_start then
    return ""
  end
  return page_start .. page_geometry(numbers)
end

local function break_tex(marker)
  -- A paragraph which starts a new page before it, on a page with nothing on it yet, needs none
  if marker == BREAK_BEFORE and on_empty_page then
    return ""
  end
  local before = on_empty_page and "\\null" or ""
  on_empty_page = true
  local setup = page_setup(marker)
  if setup ~= "" then
    pending_page_setup = setup
  end
  -- The braces end the command: text can follow right after the break.
  return before .. "\\clearpage{}"
end

local function note_text()
  on_empty_page = false
  heading_has_text = true
end

local function note_visible_inline(inline)
  if VISIBLE_INLINES[inline.t] then
    note_text()
  end
end

-- The Str with each sentinel replaced by what `place` adds to the result for
-- its marker. pandoc merges the sentinel run with the text next to it, so it
-- can stand anywhere in a Str.
local function split_str(str, place)
  local text = str.text
  if not text:find(OPEN, 1, true) then
    note_text()
    return nil
  end
  local result = {}
  local position = 1
  while true do
    local first, last, marker = text:find(SENTINEL_PATTERN, position)
    local before = text:sub(position, (first or 0) - 1)
    if before ~= "" then
      note_text()
      result[#result + 1] = pandoc.Str(before)
    end
    if not first then
      return result
    end
    place(result, marker)
    position = last + 1
  end
end

local function expand(str)
  return split_str(str, function(result, marker)
    result[#result + 1] = pandoc.RawInline("latex", break_tex(marker))
  end)
end

local function holds_only_breaks(inlines)
  for _, il in ipairs(inlines) do
    if il.t == "Span" then
      if il.identifier ~= "" or not holds_only_breaks(il.content) then
        return false
      end
    elseif il.t ~= "RawInline" and not BLANK_INLINES[il.t] then
      return false
    end
  end
  return true
end

-- A line break before a page break leaves LaTeX no line to end.
local function without_blanks(inlines)
  local result = {}
  for _, il in ipairs(inlines) do
    if not BLANK_INLINES[il.t] then
      result[#result + 1] = il
    end
  end
  return result
end

local function with_breaks(block)
  local had_break = pandoc.utils.stringify(block):find(OPEN, 1, true) ~= nil
  if not had_break then
    on_empty_page = false
    return block
  end
  block = block:walk({ Str = expand, Inline = note_visible_inline })
  if holds_only_breaks(block.content) then
    block.content = without_blanks(block.content)
  end
  return block
end

local function add_breaks(markers, out)
  for _, marker in ipairs(markers) do
    out[#out + 1] = pandoc.RawBlock("latex", break_tex(marker))
  end
end

-- Add the heading to `out`, with its breaks before and after it. The text is
-- read first and the breaks written after, so on_empty_page follows the order
-- of the output: the text after a break moved behind the heading stands before it.
local function add_heading(header, out)
  local leading, trailing = {}, {}
  local was_on_empty_page = on_empty_page
  heading_has_text = false
  header = header:walk({
    Str = function(str)
      return split_str(str, function(_, marker)
        local markers = heading_has_text and trailing or leading
        markers[#markers + 1] = marker
      end)
    end,
    Inline = note_visible_inline,
  })
  on_empty_page = was_on_empty_page
  add_breaks(leading, out)
  local had_breaks = #leading + #trailing > 0
  if not had_breaks or not holds_only_breaks(header.content) then
    out[#out + 1] = header
    on_empty_page = false
  end
  add_breaks(trailing, out)
end

local function add_block(block, out)
  if block.t == "Header" then
    add_heading(block, out)
  elseif block.t == "Para" or block.t == "Plain" then
    out[#out + 1] = with_breaks(block)
  elseif block.t == "Div" then
    local content = {}
    for _, inner in ipairs(block.content) do
      add_block(inner, content)
    end
    block.content = content
    out[#out + 1] = block
  else
    on_empty_page = false
    out[#out + 1] = block
  end
end

local function replace_breaks(blocks)
  local result = {}
  for _, block in ipairs(blocks) do
    add_block(block, result)
    if pending_page_setup ~= "" then
      result[#result + 1] = pandoc.RawBlock("latex", pending_page_setup)
      pending_page_setup = ""
    end
  end
  return result
end

local function strip_sentinels(str)
  local cleaned, count = str.text:gsub(SENTINEL_PATTERN, "")
  if count > 0 then
    return pandoc.Str(cleaned)
  end
end

function Pandoc(doc)
  keep_first_page = doc.meta["docx-comment-mag"] ~= nil
  doc.blocks = replace_breaks(doc.blocks)
  return doc:walk({ Str = strip_sentinels })
end
