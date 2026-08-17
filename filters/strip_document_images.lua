--[[
Drop the images a document points at by address, on the paths which end in a TeX
engine.

An image becomes `\includegraphics{...}` in the generated LaTeX, and tectonic
runs outside the pandoc sandbox: a document naming a file of the container
therefore puts that file into the PDF. Verified against this image.

What travels inside the document is kept: a `data:` URI, and anything pandoc
extracted itself, which it holds in the media bag. Asking the media bag rather
than the source format covers every reader, an EPUB among them, whose XHTML can
name an address of its own just like a markdown source can.
]]

function Image(element)
  if element.src:match("^data:") then
    return nil
  end
  if pandoc.mediabag.lookup(element.src) then
    return nil
  end
  return {}
end
