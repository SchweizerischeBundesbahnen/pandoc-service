function insertRawOpenXML(xml)
  return pandoc.RawBlock("openxml", xml)
end

function handleCommand(cmd)
  if cmd == "\\newpage" then
    return insertRawOpenXML('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
  elseif cmd == "\\pageLandscape" then
    return insertRawOpenXML([[
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:sectPr>
    <w:type w:val="nextPage"/>
    <w:pgSz w:orient="landscape" w:w="16838" w:h="11906"/>
  </w:sectPr></w:pPr>
</w:p>
]])
  elseif cmd == "\\pagePortrait" then
    return insertRawOpenXML([[
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:sectPr>
    <w:type w:val="nextPage"/>
    <w:pgSz w:orient="portrait" w:w="11906" w:h="16838"/>
  </w:sectPr></w:pPr>
</w:p>
]])
  end
end

function Para(el)
  local text = pandoc.utils.stringify(el)
  if text:match("\\pageLandscape") then
    return handleCommand("\\pageLandscape")
  elseif text:match("\\pagePortrait") then
    return handleCommand("\\pagePortrait")
  elseif text:match("\\newpage") then
    return handleCommand("\\newpage")
  end
end

function RawBlock(el)
  return handleCommand(el.text)
end

function RawInline(el)
  return handleCommand(el.text)
end
