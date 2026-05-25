-- html_lists.lua
--
-- Suppress the stray marker that pandoc's DOCX writer emits for list items
-- synthesized from orphan <ol>/<ul> nested directly inside another list.
--
-- The companion preprocessor (app/HtmlListsPreProcess.py) rewrites every
-- orphan <ol>/<ul> in the input HTML as
--
--     <li><span class="pandoc-suppress-marker"></span><ol>...</ol></li>
--
-- Pandoc's HTML reader turns that <li> into a list item whose first block is
-- a Plain wrapping the sentinel Span:
--
--     ListItem
--       - [ Plain [ Span ("", ["pandoc-suppress-marker"], []) [] ]
--         , OrderedList [...] ]
--
-- This filter walks every OrderedList and BulletList and, for items whose
-- first block is exactly that sentinel, replaces the sentinel block with an
-- empty `RawBlock("openxml", "")`. Pandoc's DOCX writer treats the empty raw
-- block as the first block of the list item and skips emitting its own
-- numbered paragraph; the nested list that follows is rendered at its
-- intended depth with no leading marker, matching the WeasyPrint/PDF output.
--
-- Items that legitimately contain an empty <li> wrapping a nested list
-- (e.g. authored markdown/HTML where the marker IS wanted) are unaffected,
-- because they don't carry the sentinel Span.

local SUPPRESS_CLASS = "pandoc-suppress-marker"

-- True iff `block` is a Plain containing exactly one empty Span carrying the
-- sentinel class. Anything else — text, multiple inlines, a Span with
-- content — counts as real content and the marker is kept.
local function is_sentinel_block(block)
  if block.t ~= "Plain" then return false end
  if #block.content ~= 1 then return false end
  local first = block.content[1]
  if first.t ~= "Span" then return false end
  if #first.content > 0 then return false end
  for _, c in ipairs(first.classes) do
    if c == SUPPRESS_CLASS then return true end
  end
  return false
end

-- For each list item, if its first block is the sentinel, swap it for an
-- empty RawBlock("openxml", ""). The DOCX writer emits nothing for an empty
-- RawBlock and — crucially — does not synthesize its own marker paragraph
-- when a list item's first block is a raw OOXML block.
local function rewrite_items(items)
  for _, item in ipairs(items) do
    if #item >= 1 and is_sentinel_block(item[1]) then
      item[1] = pandoc.RawBlock("openxml", "")
    end
  end
end

function OrderedList(el)
  rewrite_items(el.content)
  return el
end

function BulletList(el)
  rewrite_items(el.content)
  return el
end
