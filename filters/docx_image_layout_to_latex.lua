-- docx_image_layout_to_latex.lua
--
-- Companion to app/docx_image_layout_pre_process.py. The preprocessor puts
-- the vertical offset and the side space of an inline picture in front of its
-- title, the one picture property pandoc's DOCX reader keeps:
--
--     {{PICLAYOUT:<position in half-points>|<left in EMU>|<right in EMU>}}<original title>
--
-- This filter restores the title and sets the picture as Word does:
--
--   position (<w:position>)          -> \raisebox{<pt>}{...}, negative lowers it
--   left/right (<wp:effectExtent>)   -> \hspace{<pt>} before/after it

local EMU_PER_POINT = 12700

local function points(value)
  -- %.4g keeps the LaTeX short and drops a trailing ".0".
  return string.format("%.4gpt", value)
end

function Image(el)
  local position, left, right, title = el.title:match("^{{PICLAYOUT:(%-?%d+)|(%d+)|(%d+)}}(.*)$")
  if not position then return nil end
  el.title = title
  position, left, right = tonumber(position), tonumber(left), tonumber(right)

  local result = {}
  if left > 0 then
    result[#result + 1] = pandoc.RawInline("latex", "\\hspace{" .. points(left / EMU_PER_POINT) .. "}")
  end
  if position ~= 0 then
    result[#result + 1] = pandoc.RawInline("latex", "\\raisebox{" .. points(position / 2) .. "}{")
    result[#result + 1] = el
    result[#result + 1] = pandoc.RawInline("latex", "}")
  else
    result[#result + 1] = el
  end
  if right > 0 then
    result[#result + 1] = pandoc.RawInline("latex", "\\hspace{" .. points(right / EMU_PER_POINT) .. "}")
  end
  return result
end
