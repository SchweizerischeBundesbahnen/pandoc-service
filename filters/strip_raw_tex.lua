--[[
Drop the raw TeX a document carries, on the paths which end in a TeX engine.

`--sandbox` keeps pandoc itself from reading an address or a path a document
names, but a PDF is produced by handing the generated LaTeX to tectonic, which
runs outside that sandbox and resolves what the TeX names: `\input{/etc/passwd}`
written in a markdown or latex source reaches the engine and its content lands
in the PDF.

This filter runs first, so it removes what the reader produced from the
document. The raw LaTeX the docx filters emit afterwards is untouched, which is
what keeps the colors, lists and tables of a DOCX export working.
]]

local TEX_FORMATS = {tex = true, latex = true, context = true}

function RawBlock(element)
  if TEX_FORMATS[element.format] then
    return {}
  end
end

function RawInline(element)
  if TEX_FORMATS[element.format] then
    return {}
  end
end
