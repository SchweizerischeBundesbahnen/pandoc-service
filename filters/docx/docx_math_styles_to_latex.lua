-- docx_math_styles_to_latex.lua
--
-- Companion to app/docx_math_style_pre_process.py. texmath reads an upright
-- letter of Office math as italic and a bold italic one as upright bold. The
-- preprocessor wraps such a run in markers, which reach this filter inside the
-- Math inline's TeX string:
--
--   PMSzzzpzzzdzzzPMSENDzzz   ->  \symup{d}
--   PMSzzzbizzzvzzzPMSENDzzz  ->  \symbfit{v}
--
-- \symup and \symbfit come from unicode-math, which the XeTeX template loads.

local MARKER = "PMSzzz(%l%l?)zzz(.-)zzzPMSENDzzz"

local COMMAND = { p = "\\symup", bi = "\\symbfit" }

local filter = {}

function filter.Math(el)
  if not FORMAT:match("latex") then
    return nil
  end
  local replaced, count = el.text:gsub(MARKER, function(style, letters)
    return (COMMAND[style] or "") .. "{" .. letters .. "}"
  end)
  if count == 0 then
    return nil
  end
  el.text = replaced
  return el
end

return filter
