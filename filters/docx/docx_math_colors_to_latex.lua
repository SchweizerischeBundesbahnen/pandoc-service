-- docx_math_colors_to_latex.lua
--
-- Companion to app/DocxMathColorPreProcess.py. On the DOCX -> LaTeX/PDF path,
-- pandoc reads Office math (<m:oMath>) through texmath, whose AST has no color
-- and whose OMML reader ignores <w:color> on math runs, so an equation's color
-- is lost before any filter sees it. The preprocessor works around that by
-- wrapping each colored math run's text in plain-text markers that survive
-- texmath:
--
--   <m:t>x</m:t>  (color RRGGBB)  ->  <m:t>PMCzzzRRGGBBzzzxzzzPMCENDzzz</m:t>
--
-- which reach this filter inside the Math inline's TeX string. Here we turn each
-- marker pair back into an inline color group:
--
--   PMCzzzRRGGBBzzz<content>zzzPMCENDzzz  ->  {\color[HTML]{RRGGBB} <content>}
--
-- \color[HTML]{...} needs xcolor, which the pipeline already loads (the sibling
-- docx_colors_to_latex.lua emits \textcolor for regular text). The markers are
-- pure alphanumerics, so texmath emits them contiguously and this single gsub
-- recovers every colored segment; non-overlapping matches keep adjacent and
-- nested colored runs independent.

local MARKER = "PMCzzz(%x%x%x%x%x%x)zzz(.-)zzzPMCENDzzz"

-- A group is an ordinary atom to TeX, so a colored operator in one would lose
-- its spacing: "a+b" instead of "a + b". These keep their class.
local BINARY = {}
for op in ("+ - * \\pm \\mp \\times \\div \\cdot \\ast \\circ \\bullet \\cup \\cap \\wedge \\vee \\setminus \\oplus \\otimes"):gmatch("%S+") do
  BINARY[op] = true
end
local RELATION = {}
for op in ("= < > \\leq \\geq \\neq \\ne \\le \\ge \\approx \\equiv \\sim \\simeq \\cong \\propto \\in \\notin \\ni \\subset \\supset \\subseteq \\supseteq \\to \\rightarrow \\leftarrow \\Rightarrow \\Leftarrow \\leftrightarrow \\Leftrightarrow \\mapsto \\perp \\parallel \\mid"):gmatch("%S+") do
  RELATION[op] = true
end

local function colored(color, content)
  local op = content:match("^%s*(.-)%s*$")
  if BINARY[op] then
    return "\\mathbin{\\color[HTML]{" .. color .. "}" .. op .. "}"
  elseif RELATION[op] then
    return "\\mathrel{\\color[HTML]{" .. color .. "}" .. op .. "}"
  end
  return "{\\color[HTML]{" .. color .. "} " .. content .. "}"
end

local filter = {}

function filter.Math(el)
  -- Only the LaTeX/PDF writer consumes raw LaTeX in a Math string; for any other
  -- target this would corrupt the math, so leave it untouched. (This filter is
  -- only wired in on the docx->latex/pdf path, but gate defensively.)
  if not FORMAT:match("latex") then
    return nil
  end
  local replaced, count = el.text:gsub(MARKER, colored)
  if count == 0 then
    return nil
  end
  el.text = replaced
  return el
end

return filter
