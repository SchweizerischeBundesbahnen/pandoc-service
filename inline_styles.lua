-- inline_styles.lua
--
-- Convert inline CSS carried on HTML <span style="..."> elements into AST
-- nodes that pandoc's DOCX writer renders correctly. Pandoc's HTML reader
-- preserves the style attribute on Span but the writer ignores it, so
-- formatting expressed only via inline CSS would otherwise be dropped.
--
-- Mappings:
--   font-weight: bold|bolder|>=600         -> Strong
--   font-style:  italic|oblique            -> Emph
--   text-decoration: underline             -> Underline
--   text-decoration: line-through          -> Strikeout
--   color: <hex|rgb()>                     -> Span{custom-style = "Color-<HEX>"}
--   background-color: <hex|rgb()>          -> Span{custom-style = "Highlight-<HEX>"}
--
-- Color and highlight resolve via character styles in the reference.docx
-- (e.g. a "Color-FF0000" character style with red text). Pandoc auto-creates
-- the style if absent, so the run is at minimum tagged for follow-up styling
-- in Word; to get real colors out of the box, define the styles in the
-- reference.docx passed via --reference-doc.

local function parse_style(style)
  local props = {}
  for decl in style:gmatch("([^;]+)") do
    local k, v = decl:match("^%s*([%w%-]+)%s*:%s*(.-)%s*$")
    if k and v then
      props[k:lower()] = v:lower()
    end
  end
  return props
end

local function normalize_color(value)
  if not value then return nil end
  local s = value:gsub("%s", "")
  local hex6 = s:match("^#?(%x%x%x%x%x%x)$")
  if hex6 then return hex6:upper() end
  local a, b, c = s:match("^#?(%x)(%x)(%x)$")
  if a then return (a .. a .. b .. b .. c .. c):upper() end
  local r, g, bl = s:match("^rgb%((%d+),(%d+),(%d+)%)$")
  if r then
    return string.format("%02X%02X%02X", tonumber(r), tonumber(g), tonumber(bl))
  end
  return nil
end

local function has_token(value, token)
  if not value then return false end
  for w in value:gmatch("[^%s,]+") do
    if w == token then return true end
  end
  return false
end

local function is_bold(weight)
  if not weight then return false end
  if weight == "bold" or weight == "bolder" then return true end
  local n = tonumber(weight)
  return n ~= nil and n >= 600
end

function Span(el)
  local style = el.attributes.style
  if not style then return nil end

  local props = parse_style(style)
  local content = el.content

  if is_bold(props["font-weight"]) then
    content = { pandoc.Strong(content) }
  end
  if props["font-style"] == "italic" or props["font-style"] == "oblique" then
    content = { pandoc.Emph(content) }
  end

  local td = props["text-decoration"] or props["text-decoration-line"]
  if has_token(td, "underline") then
    content = { pandoc.Underline(content) }
  end
  if has_token(td, "line-through") then
    content = { pandoc.Strikeout(content) }
  end

  local fg = normalize_color(props["color"])
  if fg then
    content = { pandoc.Span(content, pandoc.Attr("", {}, {{"custom-style", "Color-" .. fg}})) }
  end
  local bg = normalize_color(props["background-color"])
  if bg then
    content = { pandoc.Span(content, pandoc.Attr("", {}, {{"custom-style", "Highlight-" .. bg}})) }
  end

  el.attributes.style = nil

  local has_other_attrs = false
  for _ in pairs(el.attributes) do has_other_attrs = true; break end
  if el.identifier == "" and #el.classes == 0 and not has_other_attrs then
    return content
  end

  el.content = content
  return el
end
