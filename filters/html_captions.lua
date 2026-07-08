-- html_captions.lua
--
-- Mark genuine Polarion captions so DOCX post-processing can find them
-- precisely, instead of guessing from the paragraph text.
--
-- Polarion emits every figure/table caption as an ordinary paragraph whose
-- sequence counter is wrapped in a dedicated span:
--
--     <p class="polarion-rte-caption-paragraph">
--        Table <span data-sequence="Table" class="polarion-rte-caption">1</span> My caption
--     </p>
--
-- Pandoc's HTML reader keeps that span (class "polarion-rte-caption" + the
-- "sequence" attribute) in the AST, but the DOCX writer drops it and the
-- paragraph comes out as plain body text. app/DocxReferencesPostProcess.py used
-- to recover captions by checking whether the text starts with "Table"/"Figure"
-- — which also matched headings ("Table test III"), cross-references
-- ("Table 1 shows ...") and labels ("Table 50px"), wrongly restyling them as
-- captions and listing them in the Table of Figures/Tables.
--
-- This filter detects the caption span (the one reliable, language-independent
-- signal) and wraps the paragraph in a Div carrying custom-style="Caption", so
-- pandoc emits the paragraph with the "Caption" style. The post-processor then
-- treats exactly those paragraphs as captions — nothing else.
--
-- Only runs for the DOCX writer (the controller gates it on html->docx; the
-- FORMAT check is defensive).

local CAPTION_SPAN_CLASS = "polarion-rte-caption"
local CAPTION_STYLE = "Caption"

-- True when `inlines` contains (at any depth) a Span carrying the Polarion
-- caption class. Recurses into inline containers (Span/Emph/Strong/Link/...);
-- leaf inlines such as Str expose `.text`, not a `.content` list, so the
-- type guard keeps the walk safe.
local function contains_caption_span(inlines)
  for _, il in ipairs(inlines) do
    if il.t == "Span" and il.classes then
      for _, class in ipairs(il.classes) do
        if class == CAPTION_SPAN_CLASS then
          return true
        end
      end
    end
    if type(il.content) == "table" and contains_caption_span(il.content) then
      return true
    end
  end
  return false
end

-- Wrap a caption paragraph in a Div that applies the "Caption" style. A
-- custom-style Div adds no extra element to the DOCX — pandoc just stamps the
-- style onto the contained paragraph.
local function as_caption(block)
  if not FORMAT:match("docx") then
    return nil
  end
  if contains_caption_span(block.content) then
    return pandoc.Div({ block }, pandoc.Attr("", {}, { ["custom-style"] = CAPTION_STYLE }))
  end
  return nil
end

function Para(el)
  return as_caption(el)
end

function Plain(el)
  return as_caption(el)
end
