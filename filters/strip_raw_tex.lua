--[[
Keep the TeX a document carries away from the engine which runs it.

`--sandbox` keeps pandoc itself from reading an address or a path a document
names, but a PDF is produced by handing the generated LaTeX to tectonic, which
runs outside that sandbox and resolves what the TeX names. Three routes reach
it, all verified against this image:

  * raw TeX,          `\input{/etc/passwd}` written as a raw block or inline,
  * math,             the same primitive inside `$...$`, which the writer emits
                      verbatim,
  * an image path,    which becomes `\includegraphics{...}` and puts the file
                      into the PDF.

This filter runs first, so it removes what the reader produced from the
document. The raw LaTeX the docx filters emit afterwards is untouched, which is
what keeps the colors, lists and tables of a DOCX export working.

Math is kept, only the primitives which read a file are removed from it, so a
formula still renders. Images are handled by strip_document_images.lua, which
runs only where the source format lets a document write the path itself.
]]

local TEX_FORMATS = {tex = true, latex = true, context = true}

-- Primitives which make TeX read, write or execute something outside itself.
local FILE_PRIMITIVES = {
  "input", "include", "includegraphics", "openin", "openout", "read", "readline",
  "write", "immediate", "special", "usepackage", "RequirePackage", "lstinputlisting",
  "verbatiminput", "catcode", "csname", "expandafter", "InputIfFileExists",
  "IfFileExists", "subfile", "import", "endinput", "batchmode", "shellescape",
}

local function names_a_primitive(text)
  for _, primitive in ipairs(FILE_PRIMITIVES) do
    if text:find("\\" .. primitive, 1, true) then
      return true
    end
  end
  return false
end

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

function Math(element)
  if names_a_primitive(element.text) then
    return {}
  end
end
