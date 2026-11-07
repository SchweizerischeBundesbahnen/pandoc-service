-- docx_run_spaces_to_latex.lua
--
-- app/docx_run_spaces_pre_process.py writes a space at the edge of a formatted run as U+E030, which
-- pandoc's DOCX reader keeps inside the formatting. This filter, the first of the docx filters, turns
-- each run of them back into a space where it stands: struck through, underlined or as wide as the
-- font of its run, as Word draws it.

local SPACE = "\u{E030}"

function Str(str)
  if not str.text:find(SPACE, 1, true) then
    return nil
  end
  local inlines = pandoc.Inlines({})
  local rest = str.text
  while #rest > 0 do
    -- Plain, as the sentinel is three bytes, which a class of a pattern would read one by one
    local first, last = rest:find(SPACE, 1, true)
    while last and rest:sub(last + 1, last + #SPACE) == SPACE do
      last = last + #SPACE
    end
    if first == nil then
      inlines:insert(pandoc.Str(rest))
      break
    end
    if first > 1 then
      inlines:insert(pandoc.Str(rest:sub(1, first - 1)))
    end
    inlines:insert(pandoc.Space())
    rest = rest:sub(last + 1)
  end
  return inlines
end

-- pandoc reads a run of Verbatim Char as code, from its text as it is: its sentinels are spaces of the code
function Code(code)
  if code.text:find(SPACE, 1, true) then
    code.text = code.text:gsub(SPACE, " ")
    return code
  end
end

CodeBlock = Code
