-- html_whitespace.lua
--
-- Collapse whitespace across inline elements, as a browser does.
--
-- CSS turns a run of whitespace into one space, also when the run crosses an
-- element boundary, and drops the whitespace at the start and at the end of a
-- line. Pandoc's HTML reader collapses whitespace within one text node only.
-- Polarion writes its markup indented, with whitespace on both sides of a
-- closing tag:
--
--     <span style="display:inline-block">
--      <span>implements</span>
--     </span>
--     : <a href="...">EL-101</a>
--
-- Pandoc reads the whitespace before </span> and the one after it as two
-- spaces, so Word shows "implements  :", and the whitespace after the opening
-- <span> as a space at the start of the line.
--
-- In each Para, Plain and Header, this filter removes a Space or SoftBreak
-- that follows another one, or that starts or ends a line: at the start or
-- end of the block, or next to a LineBreak. It looks into inline containers
-- (Span, Strong, Link, ...), so a run that crosses their boundaries collapses
-- too. A non-breaking space is part of a Str, so it stays, as in a browser.
-- A container styled white-space: pre, pre-wrap or break-spaces keeps its
-- spaces, since a browser keeps them too.
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

-- True for a container whose style keeps whitespace as written.
local function preserves_whitespace(inline)
  local attributes = inline.attributes
  if not attributes then
    return false
  end
  local style = attributes["style"]
  if not style then
    return false
  end
  local value = style:lower():match("white%-space%s*:%s*([%w%-]+)")
  return value == "pre" or value == "pre-wrap" or value == "break-spaces"
end

-- Walk `inlines` in reading order (`step` 1) or backwards (`step` -1) and drop
-- each space that comes right after the start of a line or after a space
-- already kept. `state.at_edge` carries that across container boundaries.
-- Walked forwards, this collapses runs and drops spaces at line starts.
-- Walked backwards afterwards, when no two spaces are adjacent any more, it
-- drops spaces at line ends. Returns the new list.
local function drop_spaces(inlines, step, state)
  local first, last = 1, #inlines
  if step < 0 then
    first, last = last, first
  end
  local kept = {}
  for i = first, last, step do
    local inline = inlines[i]
    local keep = true
    if is_space(inline) then
      keep = not state.at_edge
      state.at_edge = true
    elseif inline.t == "LineBreak" then
      state.at_edge = true
    elseif INLINE_CONTAINERS[inline.t] and not preserves_whitespace(inline) then
      -- an empty container leaves the state as it is, as in a browser
      inline.content = drop_spaces(inline.content, step, state)
    else
      state.at_edge = false
    end
    if keep then
      kept[#kept + 1] = inline
    end
  end
  local result = pandoc.Inlines({})
  local from, to = 1, #kept
  if step < 0 then
    from, to = to, from
  end
  for i = from, to, step do
    result:insert(kept[i])
  end
  return result
end

local function collapse(block)
  local content = drop_spaces(block.content, 1, { at_edge = true })
  block.content = drop_spaces(content, -1, { at_edge = true })
  return block
end

function Para(el)
  return collapse(el)
end

function Plain(el)
  return collapse(el)
end

function Header(el)
  return collapse(el)
end
