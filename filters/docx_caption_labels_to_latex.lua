-- docx_caption_labels_to_latex.lua
--
-- Extract table captions from the Table AST node back into plain
-- paragraphs when converting docx → LaTeX/PDF.
--
-- When pandoc reads a docx with a Caption-styled paragraph adjacent to a
-- table, it absorbs the paragraph into the Table's Caption field and
-- emits \caption{...} in LaTeX. This causes problems:
-- 1. LaTeX's \caption counter duplicates the numbering ("Table 1: Table 1 ...")
-- 2. The caption moves inside the table environment (centered, different style)
-- 3. Tables without our captions also get \caption from pandoc, causing
--    counter mismatches across the document
--
-- This filter extracts ALL caption content from the Table AST, clears
-- the Table's caption, and returns the caption as a regular paragraph
-- before the table — matching the expected PDF layout where caption
-- numbering is part of the visible text, not a LaTeX counter.
--
-- Figures have the same problem: a Caption-styled paragraph directly before a
-- paragraph that holds only an image becomes a Figure, and LaTeX renders it as
-- a centered float with its own "Figure N:" label. Word shows the caption and
-- the image as two ordinary paragraphs, so the Figure is split back into them.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx → pdf/latex).

function Table(tbl)
  if not tbl.caption or not tbl.caption.long or #tbl.caption.long == 0 then
    return nil
  end

  -- Collect all inlines from the caption blocks
  local inlines = {}
  for _, block in ipairs(tbl.caption.long) do
    -- Handle Div wrappers (pandoc wraps Caption-styled paragraphs in Div)
    local source_blocks = block.t == "Div" and block.content or { block }
    for _, inner in ipairs(source_blocks) do
      if (inner.t == "Para" or inner.t == "Plain") and inner.content then
        for _, il in ipairs(inner.content) do
          inlines[#inlines + 1] = il
        end
      end
    end
  end

  if #inlines == 0 then
    return nil
  end

  -- Clear the table caption
  tbl.caption = pandoc.Caption()

  -- Return caption as a plain paragraph before the table
  return { pandoc.Para(inlines), tbl }
end

-- Split a Figure back into the caption paragraph and the image paragraph it was
-- built from. Pandoc pairs a caption with the image after it, so the caption
-- comes first. The custom style of the image paragraph is on the Figure itself
-- and goes back onto a Div around the image.
function Figure(fig)
  if not fig.caption or not fig.caption.long or #fig.caption.long == 0 then
    return nil
  end

  local body = {}
  for _, block in ipairs(fig.content) do
    body[#body + 1] = block.t == "Plain" and pandoc.Para(block.content) or block
  end

  local blocks = {}
  for _, block in ipairs(fig.caption.long) do
    blocks[#blocks + 1] = block
  end
  blocks[#blocks + 1] = pandoc.Div(body, fig.attr)
  return blocks
end
