-- docx_body_style_to_latex.lua
--
-- Lay the body paragraphs of a docx out in LaTeX/PDF as Word lays them out.
--
-- LaTeX justifies and hyphenates its paragraphs, and pandoc's template puts
-- 6pt between them. Word follows the Body Text style of the document: for
-- pandoc's own template that is ragged right, unhyphenated and 9pt apart.
--
-- app/docx_body_style.py reads the style and passes it as metadata:
--   docx-body-paragraph-skip  the space between two paragraphs, in points
--   docx-body-justified       true when the style justifies the text
--   docx-body-hyphenated      true when the document turns hyphenation on
-- This filter turns that metadata into LaTeX in header-includes, after the
-- preamble the other docx filters add there, and removes the metadata.
--
-- The font and its size are not set here: the service passes them as the
-- mainfont and fontsize variables of the template.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx -> pdf/latex).

local SKIP = "docx-body-paragraph-skip"
local JUSTIFIED = "docx-body-justified"
local HYPHENATED = "docx-body-hyphenated"

-- A metadata value given with -M arrives as a boolean or as a string.
local function is_true(value)
  if type(value) == "boolean" then
    return value
  end
  return value ~= nil and pandoc.utils.stringify(value) == "true"
end

-- The paragraph skip in points, or nil when it is absent or not a number.
local function points(value)
  if value == nil then
    return nil
  end
  local number = tonumber(pandoc.utils.stringify(value))
  if number == nil or number < 0 then
    return nil
  end
  return number
end

local function body_preamble(meta)
  local lines = {}
  local skip = points(meta[SKIP])
  if skip ~= nil then
    lines[#lines + 1] = string.format("\\setlength{\\parskip}{%gpt}", skip)
  end
  if meta[JUSTIFIED] ~= nil and not is_true(meta[JUSTIFIED]) then
    lines[#lines + 1] = "\\AtBeginDocument{\\raggedright}"
  end
  if meta[HYPHENATED] ~= nil and not is_true(meta[HYPHENATED]) then
    lines[#lines + 1] = "\\hyphenpenalty=10000"
    lines[#lines + 1] = "\\exhyphenpenalty=10000"
  end
  return lines
end

function Meta(meta)
  if not FORMAT:match("latex") then
    return nil
  end
  local lines = body_preamble(meta)
  meta[SKIP] = nil
  meta[JUSTIFIED] = nil
  meta[HYPHENATED] = nil
  if #lines == 0 then
    return meta
  end
  local block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", table.concat(lines, "\n")) })
  local existing = meta["header-includes"]
  if existing == nil then
    meta["header-includes"] = pandoc.MetaList({ block })
  elseif existing.t == "MetaList" then
    table.insert(existing, block)
    meta["header-includes"] = existing
  else
    meta["header-includes"] = pandoc.MetaList({ existing, block })
  end
  return meta
end
