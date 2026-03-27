
local MAX_HEADING_LEVEL = 9

function Div(el)
  -- Check each class on the div
  for _, class in ipairs(el.classes) do
    -- Match heading-N pattern where N > 6
    local level = class:match("^heading%-(%d+)$")
    if level then
      level = tonumber(level)
      -- Only process if level > 6 (h1-h6 are handled natively)
      if level and level > 6 then
        -- Cap level at 9 for Word compatibility
        if level > MAX_HEADING_LEVEL then
          level = MAX_HEADING_LEVEL
        end

        -- Extract inlines from the div's block content to use as header content.
        -- Div content is always a list of Block elements in Pandoc's AST.
        local content = el.content
        local header_content = {}

        -- First, try to extract inlines from Plain or Para blocks
        for _, block in ipairs(content) do
          if block.t == "Plain" or block.t == "Para" then
            for _, inline in ipairs(block.content) do
              table.insert(header_content, inline)
            end
          end
        end

        -- Fall back to blocks_to_inlines for other block types (e.g., nested divs)
        if #header_content == 0 then
          header_content = pandoc.utils.blocks_to_inlines(content)
        end

        -- Clone attributes and strip the heading-N class to avoid style conflicts
        local attr = el.attr:clone()
        attr.classes = {}

        -- Create a Header element with the extracted level (capped at 9)
        return pandoc.Header(level, header_content, attr)
      end
    end
  end
  -- Return unchanged if not a heading div
  return el
end

-- Cap native Header elements at level 9
function Header(el)
  if el.level > MAX_HEADING_LEVEL then
    el.level = MAX_HEADING_LEVEL
    return el
  end
  return nil
end
