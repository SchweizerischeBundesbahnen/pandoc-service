-- docx_caption_labels_to_latex.lua
--
-- Extract table captions from the Table AST node back into plain
-- paragraphs when converting docx → LaTeX/PDF.
--
-- When pandoc reads a docx with a Caption-styled paragraph adjacent to a
-- table, it absorbs the paragraph into the Table's Caption field and
-- emits \caption{...} in LaTeX. This causes two problems:
-- 1. LaTeX's \caption counter duplicates the numbering ("Table 1: Table 1 ...")
-- 2. The caption moves inside the table environment (centered, different style)
--
-- This filter extracts the caption content from the Table AST, clears
-- the Table's caption, and returns the caption as a regular paragraph
-- before the table — matching the expected PDF layout.
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
