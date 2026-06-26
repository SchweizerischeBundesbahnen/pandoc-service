-- docx_text_decorations.lua
--
-- Runs BEFORE docx_colors_to_latex.lua on the docx->latex path.
--
-- Pandoc renders Underline/Strikeout as soul's \ul/\st. soul hyphenates (good),
-- but reconstructs its argument character-by-character and aborts
-- ("Reconstruction failed") on ANY embedded macro, and cannot live inside the
-- LR boxes that \textsuperscript/\textsubscript build. Real Polarion exports
-- routinely combine underline/strikeout with colour, highlight, other
-- decorations, super/subscripts and line breaks — all of which crash soul.
--
-- ulem's \uline/\sout tolerate nested macros, but have their own constraints,
-- so for a decoration that carries anything beyond plain text we render with
-- ulem and additionally:
--   * apply the decorations in a FIXED order — \uline OUTSIDE \sout. ulem draws
--     a nested \uline inside \sout at the wrong height (it looks like a second
--     strike-through line), whereas \sout inside \uline renders a correct
--     strike + underline. Pandoc's nesting order is incidental, so we peel the
--     combined decorations off and re-emit them outer-underline / inner-strike.
--   * split at LineBreaks so no \\ sits inside the leaders.
--   * render any highlight inside the decoration with \colorbox (box-safe, so
--     it survives inside ulem) instead of soul \hl (whose leaders ulem breaks).
--
-- Plain-text underline/strikeout is left untouched: pandoc emits soul \ul/\st
-- and it still hyphenates (the common, readability-relevant case). Highlight
-- OUTSIDE any decoration is likewise left to docx_colors_to_latex's
-- line-breakable \hl; only highlight trapped inside a decoration is boxified.

local STYLE_PREFIX = "PandocColor"
local ULEM_MACRO = { Underline = "\\uline", Strikeout = "\\sout" }
-- Application order, outermost first: underline must enclose strikeout.
local DECO_ORDER = { "Underline", "Strikeout" }

-- Plain == only characters/spaces. LineBreak is deliberately excluded: a run
-- with breaks is handled by the splitting ulem path, never by soul.
local function is_plain(inlines)
  for _, inline in ipairs(inlines) do
    local t = inline.t
    if not (t == "Str" or t == "Space" or t == "SoftBreak") then
      return false
    end
  end
  return true
end

-- Replace any highlighted PandocColor span inside `inlines` with a \colorbox
-- (box-safe inside ulem) and strip the __BG_/__HL_ segments so the colour
-- filter emits no soul \hl there. Foreground colour (__FG_) is kept. Mutates /
-- returns a rebuilt list. Only __BG_ (an explicit hex) becomes a box; a bare
-- Word highlight palette name (__HL_) is just dropped (rare inside a decoration).
local function boxify_highlight(inlines)
  local out = {}
  for _, inline in ipairs(inlines) do
    if inline.content then
      inline.content = boxify_highlight(inline.content)
    end
    local cs = inline.t == "Span" and inline.attributes and inline.attributes["custom-style"]
    if cs and cs:sub(1, #STYLE_PREFIX) == STYLE_PREFIX then
      local hex = cs:match("__BG_(%x%x%x%x%x%x)")
      inline.attributes["custom-style"] = cs:gsub("__BG_[^_]+", ""):gsub("__HL_[^_]+", "")
      if hex then
        -- \fboxsep=0pt so the box hugs the text (no padding around the
        -- highlight); wrapped in a group so the setting is local.
        out[#out + 1] = pandoc.RawInline("latex", "{\\setlength{\\fboxsep}{0pt}\\colorbox[HTML]{" .. hex .. "}{")
        out[#out + 1] = inline
        out[#out + 1] = pandoc.RawInline("latex", "}}")
      else
        out[#out + 1] = inline
      end
    else
      out[#out + 1] = inline
    end
  end
  return out
end

-- Hoist LineBreaks to the top level of `inlines`: a wrapper enclosing a break
-- is split into copies around the break, recursively. Afterwards no wrapper in
-- the list encloses a LineBreak, so a ulem command built around any segment can
-- never contain \\.
local function hoist_breaks(inlines)
  local out = {}
  for _, inline in ipairs(inlines) do
    if inline.t == "LineBreak" then
      out[#out + 1] = inline
    elseif inline.content then
      local parts = hoist_breaks(inline.content)
      local seg = {}
      local function flush()
        local copy = inline:clone()
        copy.content = seg
        out[#out + 1] = copy
        seg = {}
      end
      for _, x in ipairs(parts) do
        if x.t == "LineBreak" then
          if #seg > 0 then flush() end
          out[#out + 1] = x
        else
          seg[#seg + 1] = x
        end
      end
      if #seg > 0 then flush() end
    else
      out[#out + 1] = inline
    end
  end
  return out
end

-- Peel a chain of single-child Underline/Strikeout wrappers into the SET of
-- decorations present plus the innermost content. This collapses both
-- Underline[Strikeout[x]] and Strikeout[Underline[x]] to {U,S} + x so we can
-- re-emit them in the fixed order regardless of how pandoc nested them.
local function peel(node)
  local set = {}
  while node.t == "Underline" or node.t == "Strikeout" do
    set[node.t] = true
    if #node.content == 1 and (node.content[1].t == "Underline" or node.content[1].t == "Strikeout") then
      node = node.content[1]
    else
      break
    end
  end
  return set, node.content
end

local render_decorations -- forward declaration

-- Convert every Underline/Strikeout in a (break-free) inline list to ulem,
-- recursing into other wrappers so nested decorations are converted too.
local function ulemize(inlines)
  local out = {}
  for _, inline in ipairs(inlines) do
    if inline.t == "Underline" or inline.t == "Strikeout" then
      local set, inner = peel(inline)
      for _, x in ipairs(render_decorations(inner, set)) do
        out[#out + 1] = x
      end
    else
      if inline.content then
        inline.content = ulemize(inline.content)
      end
      out[#out + 1] = inline
    end
  end
  return out
end

-- Render `content` wrapped in the requested decoration set (fixed order:
-- \uline outside \sout), splitting at line breaks so the leaders never span \\.
render_decorations = function(content, set)
  local inner = ulemize(hoist_breaks(content))
  local open, close = "", ""
  for _, t in ipairs(DECO_ORDER) do
    if set[t] then
      open = open .. ULEM_MACRO[t] .. "{"
      close = "}" .. close
    end
  end
  local result, segment = {}, {}
  local function flush()
    if #segment > 0 then
      result[#result + 1] = pandoc.RawInline("latex", open)
      for _, x in ipairs(segment) do
        result[#result + 1] = x
      end
      result[#result + 1] = pandoc.RawInline("latex", close)
      segment = {}
    end
  end
  for _, x in ipairs(inner) do
    if x.t == "LineBreak" then
      flush()
      result[#result + 1] = x
    else
      segment[#segment + 1] = x
    end
  end
  flush()
  return result
end

local function decorate(el)
  if not FORMAT:match("latex") then
    return nil
  end
  if is_plain(el.content) then
    return nil -- leave to pandoc's soul \ul/\st (hyphenates)
  end
  local set, inner = peel(el)
  return render_decorations(boxify_highlight(inner), set)
end

local filter = {}

-- Top-down: the outermost decoration must be reached first so peel() sees the
-- whole Underline/Strikeout chain as AST nodes (bottom-up would convert the
-- inner one to raw ulem before the outer could reorder it).
filter.traverse = "topdown"

function filter.Underline(el)
  return decorate(el)
end

function filter.Strikeout(el)
  return decorate(el)
end

return filter
