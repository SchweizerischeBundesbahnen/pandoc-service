-- docx_empty_paragraphs_to_latex.lua
--
-- Keep the empty paragraphs of a docx as empty lines in LaTeX/PDF.
--
-- Word prints an empty paragraph as a blank line. The docx reader drops
-- empty paragraphs unless the source format carries +empty_paragraphs,
-- and even then the LaTeX writer prints an empty Para as nothing. So a
-- document with a blank line and one without rendered to the same PDF.
--
-- This filter turns each empty paragraph into a paragraph holding a
-- \strut: one line of the normal height, with the normal paragraph skip.
-- A manual line break adds one more such line, as it does in Word. A
-- bookmark in the paragraph is kept, so a link to it still resolves.
--
-- Only paragraphs in the document body are changed, also inside a Div
-- (the reader wraps a styled paragraph in one). Table cells and list
-- items are left as they are: docx_tables_to_latex and
-- docx_lists_to_latex lay those out, and an empty cell must not grow.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx -> pdf/latex, together with +empty_paragraphs on the reader).

local BLANK_INLINES = {
  Space = true,
  SoftBreak = true,
  LineBreak = true,
}

-- True when the inlines print nothing: whitespace, breaks, or an empty
-- span (the reader turns a bookmark into an empty anchor span).
local function prints_nothing(inlines)
  for _, il in ipairs(inlines) do
    if il.t == "Span" then
      if not prints_nothing(il.content) then
        return false
      end
    elseif not BLANK_INLINES[il.t] then
      return false
    end
  end
  return true
end

-- Count the manual line breaks and collect the spans (the bookmark
-- anchors) of inlines which print nothing.
local function breaks_and_anchors(inlines, anchors)
  local breaks = 0
  for _, il in ipairs(inlines) do
    if il.t == "LineBreak" then
      breaks = breaks + 1
    elseif il.t == "Span" then
      if il.identifier ~= "" then
        anchors[#anchors + 1] = pandoc.Span({}, il.attr)
      end
      breaks = breaks + breaks_and_anchors(il.content, anchors)
    end
  end
  return breaks
end

-- The blank line an empty paragraph prints: its anchors, a \strut, and a
-- further \strut line for each manual line break.
local function blank_line(inlines)
  local anchors = {}
  local breaks = breaks_and_anchors(inlines, anchors)
  local result = anchors
  result[#result + 1] = pandoc.RawInline("latex", "\\strut")
  for _ = 1, breaks do
    result[#result + 1] = pandoc.LineBreak()
    result[#result + 1] = pandoc.RawInline("latex", "\\strut")
  end
  return pandoc.Para(result)
end

local function keep_empty_lines(blocks)
  local result = {}
  for _, block in ipairs(blocks) do
    if (block.t == "Para" or block.t == "Plain") and prints_nothing(block.content) then
      result[#result + 1] = blank_line(block.content)
    elseif block.t == "Div" then
      block.content = keep_empty_lines(block.content)
      result[#result + 1] = block
    else
      result[#result + 1] = block
    end
  end
  return result
end

function Pandoc(doc)
  doc.blocks = keep_empty_lines(doc.blocks)
  return doc
end
