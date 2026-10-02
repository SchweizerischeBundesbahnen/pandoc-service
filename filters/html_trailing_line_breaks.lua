-- html_trailing_line_breaks.lua
--
-- Drop the <br/> that ends a line of text, as a browser does.
--
-- A browser ends the line box at a <br/> and opens a new one only when more
-- content follows, so a <br/> at the end of a block shows nothing. Polarion
-- writes text directly followed by a list in exactly that shape:
--
--     Numbered list:
--     <br/>
--     <ol><li>Item 1</li></ol>
--
-- Pandoc's HTML reader turns the text into a paragraph that ends with a
-- LineBreak, and the DOCX writer emits it as a <w:br/>. Word renders that as
-- an empty line between the text and the list, which the HTML does not show.
--
-- This filter removes one trailing LineBreak from each Para and Plain, along
-- with the spaces around it. It looks into a trailing inline container (Span,
-- Strong, Link, ...) too, since a <br/> inside <span>text<br/></span> ends the
-- line the same way. A paragraph that holds nothing but line breaks is left
-- alone: <p><br/></p> is how Polarion writes an empty line, and the break is
-- what keeps it. A <br/> after a closed block (<p>text</p><br/>) is such a
-- paragraph too, so it stays an empty line.
--
-- Only runs for HTML sources converted to DOCX (the controller gates it).

-- Inlines whose content is a list of inlines. A Note holds blocks, so it is
-- not one of them: its text is not on the line it is anchored to.
local INLINE_CONTAINERS = {
  Cite = true, Emph = true, Link = true, Quoted = true, SmallCaps = true, Span = true,
  Strikeout = true, Strong = true, Subscript = true, Superscript = true, Underline = true,
}

local function is_space(inline)
  return inline.t == "Space" or inline.t == "SoftBreak"
end

-- True when `inlines` holds something other than spaces and line breaks, at
-- any depth.
local function has_content(inlines)
  for _, inline in ipairs(inlines) do
    if INLINE_CONTAINERS[inline.t] then
      if has_content(inline.content) then
        return true
      end
    elseif not is_space(inline) and inline.t ~= "LineBreak" then
      return true
    end
  end
  return false
end

-- Remove the LineBreak that ends `inlines`, if any, together with the spaces
-- around it. Changes `inlines` in place and returns true when one was removed.
-- A container's content is read as a copy, so a changed one is assigned back.
local function strip_trailing_break(inlines)
  local last = #inlines
  while last > 0 and is_space(inlines[last]) do
    last = last - 1
  end
  if last == 0 then
    return false
  end
  local inline = inlines[last]
  if inline.t == "LineBreak" then
    for i = #inlines, last, -1 do
      inlines:remove(i)
    end
    while #inlines > 0 and is_space(inlines[#inlines]) do
      inlines:remove(#inlines)
    end
    return true
  end
  if INLINE_CONTAINERS[inline.t] then
    local content = inline.content
    if strip_trailing_break(content) then
      inline.content = content
      inlines[last] = inline
      return true
    end
  end
  return false
end

local function strip(block)
  local content = block.content
  if not has_content(content) or not strip_trailing_break(content) then
    return nil
  end
  block.content = content
  return block
end

function Para(el)
  return strip(el)
end

function Plain(el)
  return strip(el)
end
