--[[
Drop the images a document points at by path, on the paths which end in a TeX
engine.

An image becomes `\includegraphics{...}` in the generated LaTeX, and tectonic
runs outside the pandoc sandbox: a document naming a file of the container
therefore puts that file into the PDF. Verified against this image.

Only the sources where the document itself writes the address are filtered, and
the Python side decides which those are. Where the image travels inside the
document, as a `data:` URI or extracted from a DOCX by pandoc, it is kept.
]]

function Image(element)
  if not element.src:match("^data:") then
    return {}
  end
end
