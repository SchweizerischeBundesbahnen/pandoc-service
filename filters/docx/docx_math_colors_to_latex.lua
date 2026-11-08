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

-- A colored base of a script is a group to TeX, and with an OpenType math font TeX raises the script of a group
-- from its top, higher than from a letter: Word sets a colored b² as a black one. texmath writes the base in a
-- group of its own, {{\color[HTML]{FF0000} b}}^{...}: a script of colored content moves into the group of the
-- base, at its letter. A script of black content stays out, which the color of the base would reach.
local COLOR_GROUP = "^{\\color%[HTML%]{%x+}"

local function colored_script(script)
  local inner = script:match("^{(%b{})}$")
  return inner ~= nil and inner:match(COLOR_GROUP) ~= nil
end

-- Whether a group which starts at the position given stands on its own, and so can be the base of a script. A
-- group after _ or ^ is a script itself, \int_{{\color[HTML]{008000} 0}}^{...}; a group after a command, or after
-- the } or ] of an argument before it, is an argument, \sqrt{{\color[HTML]{FF0000} x}}^{...}: its script belongs
-- to the command.
local function standalone(tex, position)
  local before = tex:sub(1, position - 1):gsub("%s+$", "")
  local last = before:sub(-1)
  return not (last == "_" or last == "^" or last == "}" or last == "]" or before:match("\\%a+$"))
end

local function scripts_into_bases(tex)
  local function into_base(position, base, ...)
    if not standalone(tex, position) or not base:match(COLOR_GROUP) then
      return nil
    end
    local scripts = { ... }
    local inside = base:sub(1, -2)
    for index = 1, #scripts, 2 do
      if not colored_script(scripts[index + 1]) then
        return nil
      end
      inside = inside .. scripts[index] .. scripts[index + 1]
    end
    return inside .. "}"
  end
  tex = tex:gsub("(){(%b{})}([_^])(%b{})([_^])(%b{})", into_base)
  return (tex:gsub("(){(%b{})}([_^])(%b{})", into_base))
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
  el.text = scripts_into_bases(replaced)
  return el
end

return filter
