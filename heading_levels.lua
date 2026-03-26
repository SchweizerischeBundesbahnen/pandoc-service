
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

        -- Extract the content from the div to use as header content
        local content = el.content

        -- If the div contains inline content directly, use it
        -- Otherwise, try to extract text from nested elements
        local header_content = {}

        for _, block in ipairs(content) do
          if block.t == "Plain" or block.t == "Para" then
            for _, inline in ipairs(block.content) do
              table.insert(header_content, inline)
            end
          end
        end

        -- If we couldn't extract inline content, flatten the blocks
        if #header_content == 0 then
          header_content = pandoc.utils.blocks_to_inlines(content)
        end

        -- Create a Header element with the extracted level (capped at 9)
        return pandoc.Header(level, header_content, el.attr)
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
