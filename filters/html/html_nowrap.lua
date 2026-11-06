-- html_nowrap.lua
--
-- Keep the content of an element styled white-space: nowrap (or pre) on one
-- line in Word, as a browser does.
--
-- Pandoc's DOCX writer drops the CSS, so Word breaks a line at any space of
-- such an element, and right before or after an inline picture. Polarion
-- writes the status of a work item as an icon followed by its name:
--
--     <span class="polarion-JSEnumOption"><img src="draft.gif">Draft</span>
--
-- After a wide image Word may put "Draft" on the next line, apart from its
-- icon. Only a non-breaking space (U+00A0) holds them together: Word breaks
-- next to a picture also when a word joiner, a zero-width no-break space, a
-- zero-width joiner or a narrow no-break space stands there.
--
-- Inside a nowrap element this filter turns each Space and SoftBreak into a
-- non-breaking space, and puts one between a picture and the text right
-- next to it, also across inline element boundaries. White-space is
-- inherited, so this holds for a Div or a Span and everything inside it, up
-- to a nested element that sets another value.
--
-- The space between a picture and its text is not in the HTML, so it is
-- written in a 1 pt font. Word still does not break there, and the gap is
-- too narrow to see.
--
-- Runs after html_whitespace, which collapses the spaces, and before
-- inline_styles, which turns styled spans into raw OOXML. Only runs for HTML
-- sources converted to DOCX (the controller gates it).

local NBSP = "\u{00A0}"

-- A non-breaking space in a 1 pt font (w:sz counts half-points).
local GLUE = pandoc.RawInline("openxml",
  '<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
  .. '<w:t xml:space="preserve">' .. NBSP .. "</w:t></w:r>")

-- Inlines whose content is a list of inlines. A Note holds blocks, so it is
-- not one of them: its text is not on the line it is anchored to.
local INLINE_CONTAINERS = {
  Cite = true, Emph = true, Link = true, Quoted = true, SmallCaps = true, Span = true,
  Strikeout = true, Strong = true, Subscript = true, Superscript = true, Underline = true,
}

-- Inlines that show text on the line. A picture sits next to one of them
-- without a break only when a non-breaking space stands between them.
local TEXT_INLINES = { Code = true, Math = true, Str = true }

-- white-space values that keep a line from breaking, and values that take the
-- parent's: nil for them, as for an element that does not set white-space.
local NO_WRAP_VALUES = { nowrap = true, pre = true }
local INHERITED_VALUES = { inherit = true, unset = true, revert = true, ["revert-layer"] = true }

-- true when the style of `el` keeps its content on one line, false when it
-- lets it break, nil when it does not say and the parent's applies. The last
-- white-space declaration wins, as in a browser.
local function nowrap_of(el)
  local attributes = el.attributes
  local style = attributes and attributes["style"]
  if not style then
    return nil
  end
  local value = nil
  for declaration in style:lower():gmatch("[^;]+") do
    local found = declaration:match("^%s*white%-space%s*:%s*([%w%-]+)")
    if found then
      value = found
    end
  end
  if value == nil or INHERITED_VALUES[value] then
    return nil
  end
  return NO_WRAP_VALUES[value] == true
end

-- Walk `inlines` in reading order. `state.last` is the kind of the inline seen
-- last on the line ("text", "image" or "space"), and `state.last_nowrap`
-- whether it was inside a nowrap element. A picture and text next to each other
-- get a non-breaking space between them when both are inside a nowrap
-- element, in a 1 pt font. Returns the new list.
local function walk_inlines(inlines, state, nowrap)
  local result = pandoc.Inlines({})
  for _, inline in ipairs(inlines) do
    local kind = nil
    if inline.t == "Space" or inline.t == "SoftBreak" then
      kind = "space"
      if nowrap then
        inline = pandoc.Str(NBSP)
      end
    elseif inline.t == "LineBreak" then
      kind = "space"
    elseif inline.t == "Image" then
      kind = "image"
    elseif TEXT_INLINES[inline.t] then
      kind = "text"
      if nowrap and inline.t == "Code" then
        inline.text = inline.text:gsub(" ", NBSP)
      end
    elseif INLINE_CONTAINERS[inline.t] then
      local inner_nowrap = nowrap_of(inline)
      if inner_nowrap == nil then
        inner_nowrap = nowrap
      end
      inline.content = walk_inlines(inline.content, state, inner_nowrap)
    end
    if kind ~= nil then
      local glued = (kind == "image" and state.last == "text") or (kind == "text" and state.last == "image")
      if glued and nowrap and state.last_nowrap then
        result:insert(GLUE:clone())
      end
      state.last = kind
      state.last_nowrap = nowrap
    end
    result:insert(inline)
  end
  return result
end

-- The filter for blocks inside (`nowrap`) or outside a nowrap element. A Div
-- that changes it has its content walked with the other one. `:walk` honors
-- the top-down traversal, unlike pandoc.walk_block, so a paragraph is reached
-- only through the Divs around it.
local function blocks_filter(nowrap)
  local function inlines_block(block)
    block.content = walk_inlines(block.content, {}, nowrap)
    return block
  end
  return {
    traverse = "topdown",
    Div = function(div)
      local inner_nowrap = nowrap_of(div)
      if inner_nowrap == nil or inner_nowrap == nowrap then
        return nil
      end
      return div:walk(blocks_filter(inner_nowrap)), false
    end,
    Para = inlines_block,
    Plain = inlines_block,
    Header = inlines_block,
  }
end

return { blocks_filter(false) }
