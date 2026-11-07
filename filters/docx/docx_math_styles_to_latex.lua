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
-- texmath writes the other styles of a run as \mathit, \mathbf, \mathsf,
-- \mathtt and \mathrm, which unicode-math sets in the text font. Word sets them
-- in the math font, so they become \symit, \symbf, \symsf, \symtt and \symup.
--
-- The \sym commands come from unicode-math, which the XeTeX template loads.

local MARKER = "PMSzzz(%l%l?)zzz(.-)zzzPMSENDzzz"

local COMMAND = { p = "\\symup", bi = "\\symbfit" }

local ALPHABET = { it = "it", bf = "bf", sf = "sf", tt = "tt", rm = "up" }

local filter = {}

function filter.Math(el)
  if not FORMAT:match("latex") then
    return nil
  end
  local replaced, count = el.text:gsub(MARKER, function(style, letters)
    return (COMMAND[style] or "") .. "{" .. letters .. "}"
  end)
  local alphabets
  replaced, alphabets = replaced:gsub("\\math(%a+)%f[^%a]", function(name)
    return ALPHABET[name] and "\\sym" .. ALPHABET[name] or nil
  end)
  count = count + alphabets
  if count == 0 then
    return nil
  end
  el.text = replaced
  return el
end

return filter
