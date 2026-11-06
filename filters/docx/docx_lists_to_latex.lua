-- docx_lists_to_latex.lua
--
-- Companion to app/DocxListLevelPreProcess.py. Pandoc's docx reader flattens
-- "irregular" lists — where a deeper level is nested directly inside a
-- shallower one (Polarion allows this) — so a level-3 item with no level-2
-- item before it collapses to the same nesting depth, and therefore the same
-- indentation, as level 2 in the LaTeX/PDF output.
--
-- The preprocessor prepends each list paragraph's true <w:ilvl> as a sentinel
-- that survives into the AST as a leading Str on the item:
--
--     Str "\u{E000}<level>\u{E001}<first word>"
--
-- This filter walks the lists while tracking the structural depth pandoc gave
-- them, reads each list's true level from the sentinel, and pushes any list
-- that is shallower than its true level into extra *marker-less* wrapper
-- levels (\begin{enumerate}\item[] ... or \begin{itemize}\item[] ...). The
-- empty \item[] adds one level of indentation without drawing a marker and
-- without renumbering the sibling list, which reproduces Word's absolute-level
-- indentation. Finally every residual sentinel is stripped from the text.
--
-- Why the markers still come out right: pandoc's depth-based ordered styles
-- (decimal / lower-alpha / lower-roman) and LaTeX's enumerate defaults
-- (arabic / alph / roman), and likewise the itemize bullet defaults, follow
-- the same per-depth cycle. A pushed list lands at the runtime depth matching
-- its style, so the default label for that depth is the intended one.

local OPEN = "\u{E000}"
local CLOSE = "\u{E001}"
local SENTINEL_PATTERN = OPEN .. "(%d+)" .. CLOSE

-- LaTeX's enumerate/itemize nest at most 4 deep; never wrap past that.
local MAX_DEPTH = 4

-- Read a list's true level (0-based) from the first sentinel-tagged Str in any
-- of its items, or nil when none is tagged. One flattened sublist maps to one
-- source numbering level, so any tagged item gives the list's level.
local function read_level(list)
  for _, item in ipairs(list.content) do
    for _, blk in ipairs(item) do
      if (blk.t == "Para" or blk.t == "Plain") and #blk.content >= 1 then
        local first = blk.content[1]
        if first.t == "Str" then
          local n = first.text:match("^" .. SENTINEL_PATTERN)
          if n then return tonumber(n) end
        end
      end
    end
  end
  return nil
end

local process_blocks  -- forward declaration (mutual recursion with handle_list)

local function handle_list(list, depth)
  -- Fix nested lists first: each item's blocks live one structural level deeper.
  for i, item in ipairs(list.content) do
    list.content[i] = process_blocks(item, depth + 1)
  end

  local level = read_level(list)
  local push = level and (level - (depth - 1)) or 0
  if push < 0 then push = 0 end
  if depth + push > MAX_DEPTH then push = MAX_DEPTH - depth end
  if push <= 0 then
    return { list }
  end

  local kind = (list.t == "OrderedList") and "enumerate" or "itemize"
  local open, close = "", ""
  for _ = 1, push do
    open = open .. "\\begin{" .. kind .. "}\\item[] "
    close = "\\end{" .. kind .. "}" .. close
  end
  return { pandoc.RawBlock("latex", open), list, pandoc.RawBlock("latex", close) }
end

-- Walk a block list at the given structural depth (1-based; a top-level list
-- is depth 1). Lists are re-nested; container blocks are descended into so a
-- list inside a Div/BlockQuote still starts a fresh depth-1 context.
process_blocks = function(blocks, depth)
  local out = {}
  for _, block in ipairs(blocks) do
    local t = block.t
    if t == "OrderedList" or t == "BulletList" then
      for _, b in ipairs(handle_list(block, depth)) do
        out[#out + 1] = b
      end
    elseif t == "Div" or t == "BlockQuote" then
      block.content = process_blocks(block.content, 1)
      out[#out + 1] = block
    else
      out[#out + 1] = block
    end
  end
  return out
end

function Pandoc(doc)
  -- Only emit the raw-LaTeX re-nesting for LaTeX targets; the filter is gated
  -- on docx->latex in the controller, so this is defensive.
  if FORMAT:match("latex") then
    doc = pandoc.Pandoc(process_blocks(doc.blocks, 1), doc.meta)
  end
  -- Strip every residual sentinel so it never leaks into the output, even from
  -- a list the re-nesting walk did not reach.
  return doc:walk({
    Str = function(s)
      local cleaned = s.text:gsub(SENTINEL_PATTERN, "")
      if cleaned ~= s.text then
        return pandoc.Str(cleaned)
      end
      return nil
    end,
  })
end
