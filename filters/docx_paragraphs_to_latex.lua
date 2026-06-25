-- docx_paragraphs_to_latex.lua
--
-- Companion to app/DocxParagraphPreProcess.py. The preprocessor rewrites every
-- paragraph with alignment and/or a left indent in a DOCX as a reference to a
-- synthetic *paragraph* style named
--
--   PandocPara__ALIGN_<align>__IND_<twips>      (each segment optional)
--
-- With `-f docx+styles`, pandoc surfaces those references as Div nodes carrying
-- `custom-style="PandocPara__..."`. This filter rewrites each such Div into its
-- inner blocks wrapped in a TeX group that sets the paragraph indent and
-- alignment, so the formatting that pandoc's docx reader otherwise drops (it
-- discards <w:jc> and collapses <w:ind> into a BlockQuote) survives into the
-- PDF/LaTeX output:
--
--   IND_<twips>   -> \leftskip=<pt>pt   (twips/20 = points)
--   ALIGN_center  -> \centering
--   ALIGN_right   -> \raggedleft
--   ALIGN_left    -> \raggedright
--
-- All are plain TeX primitives (no package needed) and are line-breakable. The
-- group ends with \par so the alignment/indent applies to the wrapped
-- paragraph and is scoped away from the surrounding text. Justified text is the
-- LaTeX default, so the preprocessor never emits an ALIGN segment for it.

local STYLE_PREFIX = "PandocPara"

-- Normalised alignment segment -> LaTeX paragraph-alignment primitive. This is
-- a trust boundary: only values present as a key here ever reach the output,
-- and the emitted string is the table's literal value — attacker-controlled
-- attribute text is never echoed into the LaTeX.
local ALIGN_PRIMITIVE = {
  left = "\\raggedright",
  center = "\\centering",
  right = "\\raggedleft",
}

-- Parse an IND segment value into a LaTeX length in points, or nil if it isn't
-- a clean non-negative integer count of twips within a sane bound. Like the
-- twips validation in inline_styles.lua, this keeps anything that isn't a
-- bounded integer out of the emitted \leftskip length.
local function twips_to_pt(raw)
  local n = tonumber(raw)
  if not n then return nil end
  if n < 0 then return nil end
  if n ~= math.floor(n) then return nil end
  if n > 31680 then return nil end           -- ~200 inches; far beyond any real indent
  -- 1pt = 20 twips. string.format emits only digits and a dot, so the value
  -- cannot break out of the length argument.
  return string.format("%.2f", n / 20)
end

local filter = {}

function filter.Div(el)
  -- The filter only emits LaTeX; for any other writer leave the Div untouched.
  if not FORMAT:match("latex") then return nil end

  local style = el.attributes and el.attributes["custom-style"]
  if not style or style:sub(1, #STYLE_PREFIX) ~= STYLE_PREFIX then
    return nil
  end

  local align_cmd, indent_pt
  -- Segments are "__<KEY>_<value>"; KEY is letters, value has no underscore.
  for key, value in style:gmatch("__(%a+)_([^_]+)") do
    if key == "ALIGN" then
      align_cmd = ALIGN_PRIMITIVE[value:lower()]
    elseif key == "IND" then
      indent_pt = twips_to_pt(value)
    end
  end

  if not align_cmd and not indent_pt then
    return nil
  end

  local open = "{"
  if indent_pt then
    open = open .. "\\leftskip=" .. indent_pt .. "pt "
  end
  if align_cmd then
    open = open .. align_cmd .. " "
  end

  local result = { pandoc.RawBlock("latex", open) }
  for _, block in ipairs(el.content) do
    result[#result + 1] = block
  end
  result[#result + 1] = pandoc.RawBlock("latex", "\\par}")
  return result
end

return filter
