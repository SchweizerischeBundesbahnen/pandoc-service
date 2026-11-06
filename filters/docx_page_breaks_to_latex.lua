-- docx_page_breaks_to_latex.lua
--
-- Keep the page and section breaks of a docx in LaTeX/PDF.
--
-- The docx reader drops both. app/docx_page_break_pre_process.py writes
-- them into the text as sentinels before pandoc reads the document:
--
--   Str "\u{E020}\u{E021}"                 a page break
--   Str "\u{E020}w,h,t,r,b,l\u{E021}"      a section break, in twips:
--                                          the page and margins it opens
--
-- This filter replaces a sentinel with \clearpage where it stands. A
-- section break then sets the page size and the margins of the next page,
-- so a landscape section comes out landscape. That setup goes after the
-- body block holding the break: docx_paragraphs_to_latex wraps a formatted
-- paragraph in a TeX group, which would undo it. The sentinel is document
-- text, so its numbers are checked and only they reach the raw LaTeX.
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
local SENTINEL_PATTERN = OPEN .. "([%d,]*)" .. CLOSE

-- More than Word's largest page, and small enough for a TeX dimension.
local MAX_TWIPS = 99999
local TWIPS_PER_INCH = 1440

local BLANK_INLINES = {
  Space = true,
  SoftBreak = true,
  LineBreak = true,
}

-- Whether the current page has nothing on it yet.
local on_empty_page = true

-- The page setup of the last section break, until its body block ends.
local pending_page_setup = ""

local function inches(twips)
  return string.format("%.4fin", twips / TWIPS_PER_INCH)
end

-- The LaTeX which sets up the page a section break opens, or "" when the
-- numbers describe no page LaTeX can set.
local function page_setup(numbers)
  local w, h, top, right, bottom, left = numbers:match("^(%d+),(%d+),(%d+),(%d+),(%d+),(%d+)$")
  if not w then
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

local function break_tex(numbers)
  local before = on_empty_page and "\\null" or ""
  on_empty_page = true
  local setup = page_setup(numbers)
  if setup ~= "" then
    pending_page_setup = setup
  end
  -- The braces end the command: text can follow right after the break.
  return pandoc.RawInline("latex", before .. "\\clearpage{}")
end

-- The Str with each sentinel replaced by its break. pandoc merges the
-- sentinel run with the text next to it, so it can stand anywhere in a Str.
local function expand(str)
  local text = str.text
  if not text:find(OPEN, 1, true) then
    on_empty_page = false
    return nil
  end
  local result = {}
  local position = 1
  while true do
    local first, last, numbers = text:find(SENTINEL_PATTERN, position)
    local before = text:sub(position, (first or 0) - 1)
    if before ~= "" then
      on_empty_page = false
      result[#result + 1] = pandoc.Str(before)
    end
    if not first then
      return result
    end
    result[#result + 1] = break_tex(numbers)
    position = last + 1
  end
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
  block = block:walk({ Str = expand })
  if holds_only_breaks(block.content) then
    block.content = without_blanks(block.content)
  end
  return block
end

local function replace_breaks_in(block)
  if block.t == "Para" or block.t == "Plain" or block.t == "Header" then
    return with_breaks(block)
  end
  if block.t == "Div" then
    local content = {}
    for _, inner in ipairs(block.content) do
      content[#content + 1] = replace_breaks_in(inner)
    end
    block.content = content
    return block
  end
  on_empty_page = false
  return block
end

local function replace_breaks(blocks)
  local result = {}
  for _, block in ipairs(blocks) do
    result[#result + 1] = replace_breaks_in(block)
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
  doc.blocks = replace_breaks(doc.blocks)
  return doc:walk({ Str = strip_sentinels })
end
