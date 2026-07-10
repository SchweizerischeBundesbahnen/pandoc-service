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

-- Escape characters that aren't safe inside an XML element body or attribute.
local function escape_xml(s)
  s = s:gsub("&", "&amp;")
  s = s:gsub("<", "&lt;")
  s = s:gsub(">", "&gt;")
  s = s:gsub('"', "&quot;")
  return s
end

-- Build OOXML for a Word SEQ field: { SEQ <seq_type> \* ARABIC }.
-- Uses the complex field form (begin/separate/end) because pandoc can
-- read it back when converting docx→pdf, whereas fldSimple is lost.
-- The cached display value between "separate" and "end" is set to
-- `number_text` so the field shows the correct number before Word
-- recalculates fields.
local function seq_field_xml(seq_type, number_text)
  return '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    .. '<w:r><w:instrText xml:space="preserve"> SEQ '
    .. escape_xml(seq_type) .. ' \\* ARABIC </w:instrText></w:r>'
    .. '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    .. '<w:r><w:t>' .. escape_xml(number_text) .. '</w:t></w:r>'
    .. '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
end

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

-- Wrap a caption paragraph in a Div that applies the "Caption" style AND
-- replace the Polarion caption-counter span with a proper Word SEQ field.
-- A custom-style Div adds no extra element to the DOCX — pandoc just stamps
-- the style onto the contained paragraph. The SEQ field lets Word renumber
-- captions automatically (and makes cross-references resolve correctly).
local function as_caption(block)
  if not FORMAT:match("docx") then
    return nil
  end
  if not contains_caption_span(block.content) then
    return nil
  end
  -- Replace caption-counter spans with SEQ field raw OOXML.
  local new_block = pandoc.walk_block(block, {
    Span = function(el)
      for _, class in ipairs(el.classes) do
        if class == CAPTION_SPAN_CLASS then
          local seq_type = el.attributes["sequence"] or el.attributes["data-sequence"]
          if seq_type then
            local number_text = pandoc.utils.stringify(el.content)
            return pandoc.RawInline("openxml", seq_field_xml(seq_type, number_text))
          end
          break
        end
      end
    end,
  })
  return pandoc.Div({ new_block }, pandoc.Attr("", {}, { ["custom-style"] = CAPTION_STYLE }))
end

function Para(el)
  return as_caption(el)
end

function Plain(el)
  return as_caption(el)
end
