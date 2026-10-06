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
-- TeX's largest dimension is just under 16384pt.
local MAX_POINTS = 16000

local function points(value)
  -- Fixed-point: %g writes 1e4 and 1e-5 as exponents, which TeX cannot read.
  value = math.max(-MAX_POINTS, math.min(MAX_POINTS, value))
  local s = string.format("%.2f", value):gsub("%.?0+$", "")
  if s == "" or s == "-" or s == "-0" then s = "0" end
  return s .. "pt"
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
