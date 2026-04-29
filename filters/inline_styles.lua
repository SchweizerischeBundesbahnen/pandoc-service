-- inline_styles.lua
--
-- Convert inline CSS carried on HTML <span style="..."> elements into raw
-- OOXML runs that the DOCX writer renders correctly. Pandoc's HTML reader
-- preserves the style attribute on Span but the writer ignores it, so
-- formatting expressed only via inline CSS would otherwise be dropped.
--
-- Mappings (all baked directly into <w:rPr>; no reference.docx needed):
--   font-weight: bold|bolder|>=600         -> <w:b/>
--   font-style:  italic|oblique            -> <w:i/>
--   text-decoration: underline             -> <w:u w:val="single"/>
--   text-decoration: line-through          -> <w:strike/>
--   color: <hex|rgb()>                     -> <w:color w:val="RRGGBB"/>
--   background-color: <hex|rgb()>          -> <w:shd w:val="clear" w:color="auto" w:fill="RRGGBB"/>
--   font-size: <Nunit|keyword>             -> <w:sz w:val="..."/>  (units: pt, px, pc, in, cm, mm, em, rem, %; keywords: xx-small..xx-large, smaller, larger)
--   font-family: <name>, ...               -> <w:rFonts w:ascii="name" w:hAnsi="name"/>
--
-- Traversal is top-down: the outermost styled span consumes its full subtree
-- in a single pass and emits a flat list of <w:r> runs. CSS on a nested span
-- overrides inherited properties only for the keys it sets.
--
-- Lua quick primer (relevant to this file):
--   * `local x` declares a block-scoped variable. Without `local`, names
--     would be global.
--   * `s:method(...)` is sugar for `string.method(s, ...)` (method call).
--   * Lua patterns look like regex but are simpler: %w = word char,
--     %x = hex digit, %s = whitespace, %d = digit, %a = letter,
--     %-/%./%( etc. escape literals. Parens (...) capture groups.
--     `*` / `+` / `?` are greedy; `-` is the non-greedy quantifier.
--     `^` and `$` anchor to start/end of the string.
--   * Tables `{}` are Lua's only compound type — used as both arrays
--     (1-indexed) and dicts. `#t` is array length. `ipairs` iterates
--     array part in order; `pairs` iterates everything.
--   * `nil` and `false` are the only falsy values; everything else (0,
--     "", {}) is truthy.
--   * Functions can return multiple values and accept variadic args.

-- Split a CSS declaration list ("k1: v1; k2: v2; ...") into a Lua table
-- keyed by lowercased property name with lowercased value.
local function parse_style(style)
  local props = {}
  -- gmatch returns an iterator over all substrings that match the pattern.
  -- The pattern "([^;]+)" captures any run of characters that isn't a ';'.
  for decl in style:gmatch("([^;]+)") do
    -- match captures (key, value) around the first colon, trimming spaces.
    -- "%s*" eats whitespace; "(.-)%s*$" is non-greedy so trailing space is dropped.
    local k, v = decl:match("^%s*([%w%-]+)%s*:%s*(.-)%s*$")
    if k and v then
      props[k:lower()] = v:lower()
    end
  end
  return props
end

-- Accept "#RRGGBB", "RRGGBB", "#RGB", "RGB", or "rgb(r,g,b)" and return
-- the canonical 6-char uppercase hex string Word expects ("FF0000"), or
-- nil when the value can't be parsed.
local function normalize_color(value)
  if not value then return nil end
  -- gsub returns (newstring, count); we ignore the count by assigning to one var.
  -- "%s" matches whitespace; this strips all spaces inside, e.g. "rgb( 255, 0, 0)".
  local s = value:gsub("%s", "")
  -- Try 6-digit hex first.
  local hex6 = s:match("^#?(%x%x%x%x%x%x)$")
  if hex6 then return hex6:upper() end
  -- 3-digit shorthand (#F00 -> FF0000): match returns three captures.
  local a, b, c = s:match("^#?(%x)(%x)(%x)$")
  if a then return (a .. a .. b .. b .. c .. c):upper() end
  -- rgb(r,g,b) form. tonumber converts the captured digits to integers.
  local r, g, bl = s:match("^rgb%((%d+),(%d+),(%d+)%)$")
  if r then
    -- string.format here packs three ints into 2-digit hex pairs.
    return string.format("%02X%02X%02X", tonumber(r), tonumber(g), tonumber(bl))
  end
  return nil
end

-- True if `value` is a whitespace/comma-separated list containing `token`.
-- Used for CSS `text-decoration: underline line-through` and similar.
local function has_token(value, token)
  if not value then return false end
  for w in value:gmatch("[^%s,]+") do
    if w == token then return true end
  end
  return false
end

-- Decide whether a CSS font-weight value means "bold". Accepts the keywords
-- "bold"/"bolder" and any numeric weight >= 600 (matches CSS conventions).
local function is_bold(weight)
  if not weight then return false end
  if weight == "bold" or weight == "bolder" then return true end
  -- tonumber returns nil if the string isn't numeric.
  local n = tonumber(weight)
  return n ~= nil and n >= 600
end

-- font-family generic names we want to skip — they're not real fonts and
-- would confuse Word if emitted as <w:rFonts w:ascii="sans-serif">.
local GENERIC_FAMILIES = {
  ["serif"] = true, ["sans-serif"] = true, ["monospace"] = true,
  ["cursive"] = true, ["fantasy"] = true, ["system-ui"] = true,
}

-- CSS allows a fallback list ("'Segoe UI', Selawik, sans-serif"). DOCX
-- expects a single concrete face name. Pick the first non-generic entry,
-- stripping surrounding quotes.
local function first_font_family(value)
  if not value then return nil end
  for raw in value:gmatch("([^,]+)") do
    -- Trim leading/trailing whitespace.
    local name = raw:gsub("^%s+", ""):gsub("%s+$", "")
    -- Strip a single matching pair of single or double quotes.
    -- "%1" in the replacement is the captured group from the pattern.
    name = name:gsub("^['\"](.*)['\"]$", "%1")
    if name ~= "" and not GENERIC_FAMILIES[name:lower()] then
      return name
    end
  end
  return nil
end

-- Word's <w:sz w:val="..."> takes a font size in *half-points*. Browsers
-- emit font-size in many units, so we normalize all of them here.
--   pt   1pt  = 2 half-points
--   px   1px  = 0.75pt = 1.5 half-points (CSS reference pixel: 1/96 in)
--   pc   1pc  = 12pt   = 24 half-points
--   in   1in  = 72pt   = 144 half-points
--   cm   1cm  ≈ 28.35pt
--   mm   1mm  ≈ 2.835pt
--   em   relative to the parent size; needs `parent_hp` (parent half-points)
--   rem  same as em — we have no document-root context, so treat as em
--   %    relative to parent size
-- A bare number (no unit) is treated as pt — some HTML emitters drop the unit.
-- Absolute keywords (medium, large, ...) and the relative keywords
-- (smaller, larger) are mapped to standard CSS table values; smaller/larger
-- scale the parent size by ~0.83 / ~1.2.
-- When the value is unparseable we return nil so the caller leaves the
-- inherited size in place.
local DEFAULT_HALF_POINTS = 24  -- 12pt — Word's body default
local FONT_SIZE_KEYWORDS = {
  ["xx-small"] = 18,  -- 9pt
  ["x-small"]  = 20,  -- 10pt
  ["small"]    = 22,  -- 11pt
  ["medium"]   = 24,  -- 12pt
  ["large"]    = 28,  -- 14pt
  ["x-large"]  = 36,  -- 18pt
  ["xx-large"] = 48,  -- 24pt
}

local function round_to_string(x)
  -- Round half-up and stringify as an integer.
  return tostring(math.floor(x + 0.5))
end

local function font_size_half_points(value, parent_hp)
  if not value then return nil end
  -- Trim incidental whitespace.
  local v = value:gsub("^%s+", ""):gsub("%s+$", "")

  -- Absolute keywords map to fixed half-point values.
  local kw = FONT_SIZE_KEYWORDS[v]
  if kw then return tostring(kw) end

  -- Relative keywords scale the inherited size; fall back to default if none.
  local base = (parent_hp and tonumber(parent_hp)) or DEFAULT_HALF_POINTS
  if v == "smaller" then return round_to_string(base * 0.83) end
  if v == "larger" then return round_to_string(base * 1.2) end

  -- Numeric form: digits/decimals followed by an optional unit.
  -- Lua patterns don't have alternation inside a single match, so we try
  -- letter units, then "%", then bare-number, in turn.
  local n, unit = v:match("^([%d%.]+)%s*(%a+)$")
  if not n then
    n = v:match("^([%d%.]+)%s*%%$")
    if n then unit = "%" end
  end
  if not n then
    n = v:match("^([%d%.]+)$")
    if n then unit = "" end
  end
  if not n then return nil end

  local num = tonumber(n)
  if not num then return nil end

  if unit == "" or unit == "pt" then
    return round_to_string(num * 2)
  elseif unit == "px" then
    return round_to_string(num * 1.5)
  elseif unit == "pc" then
    return round_to_string(num * 24)
  elseif unit == "in" then
    return round_to_string(num * 144)
  elseif unit == "cm" then
    return round_to_string(num * 56.6929134)
  elseif unit == "mm" then
    return round_to_string(num * 5.66929134)
  elseif unit == "em" or unit == "rem" then
    return round_to_string(num * base)
  elseif unit == "%" then
    return round_to_string(num * base / 100)
  end
  return nil
end

-- Shallow copy of a table. `pairs` iterates every key (string or numeric).
local function clone(t)
  local c = {}
  for k, v in pairs(t) do c[k] = v end
  return c
end

-- Combine an existing run-properties table (`parent`, what we inherit) with
-- a freshly parsed CSS block (`css`, the styles on the current span). Keys
-- the CSS doesn't mention pass through unchanged; keys it does mention
-- override (this is how nested spans cascade).
local function merge_css(parent, css)
  local p = clone(parent)
  if css["font-weight"] then p.bold = is_bold(css["font-weight"]) end
  if css["font-style"] then
    p.italic = (css["font-style"] == "italic" or css["font-style"] == "oblique")
  end
  -- text-decoration is the legacy shorthand; text-decoration-line is the
  -- modern long-hand. Either one fully replaces inherited decorations.
  local td = css["text-decoration"] or css["text-decoration-line"]
  if td then
    if td == "none" then
      p.underline = false
      p.strikeout = false
    else
      p.underline = has_token(td, "underline")
      p.strikeout = has_token(td, "line-through")
    end
  end
  if css.color then
    local fg = normalize_color(css.color)
    if fg then p.fg = fg end
  end
  if css["background-color"] then
    local bg = normalize_color(css["background-color"])
    if bg then p.bg = bg end
  end
  if css["font-size"] then
    -- Pass the inherited size (in half-points) so em/rem/% can resolve
    -- against the actual parent rather than always falling back to default.
    local sz = font_size_half_points(css["font-size"], parent.size)
    if sz then p.size = sz end
  end
  if css["font-family"] then
    local ff = first_font_family(css["font-family"])
    if ff then p.font = ff end
  end
  return p
end

-- Escape characters that aren't safe inside an XML element body or
-- attribute value. We embed literal user text into raw OOXML strings, so
-- these substitutions are mandatory.
local function escape_xml(s)
  s = s:gsub("&", "&amp;")
  s = s:gsub("<", "&lt;")
  s = s:gsub(">", "&gt;")
  s = s:gsub('"', "&quot;")
  return s
end

local function escape_attr(s)
  return escape_xml(s)
end

-- Build the <w:rPr> (run properties) sub-element from a `props` table and
-- an optional vert_align ("superscript" / "subscript"). Returns "" when
-- there's nothing to set, so callers can concatenate unconditionally.
--
-- Element order follows the OOXML CT_RPr schema sequence (ECMA-376 Part 1
-- §17.3.2.28). Word's tolerance for out-of-order children is version-
-- dependent (some builds re-order silently, others log a recovery warning,
-- LibreOffice may drop the misplaced child entirely), so we emit them in
-- the canonical order. The schema position of each child we emit is
-- annotated below for future maintainers.
local function rpr_xml(props, vert_align)
  local parts = {}
  if props.font then
    local f = escape_attr(props.font)
    -- `parts[#parts + 1] = x` is the idiomatic Lua "append to array".
    parts[#parts + 1] = '<w:rFonts w:ascii="' .. f .. '" w:hAnsi="' .. f .. '" w:cs="' .. f .. '"/>'  -- pos 2
  end
  if props.bold then parts[#parts + 1] = "<w:b/>" end                                                  -- pos 3
  if props.italic then parts[#parts + 1] = "<w:i/>" end                                                -- pos 5
  if props.strikeout then parts[#parts + 1] = "<w:strike/>" end                                        -- pos 9
  if props.fg then parts[#parts + 1] = '<w:color w:val="' .. props.fg .. '"/>' end                     -- pos 19
  if props.size then
    parts[#parts + 1] = '<w:sz w:val="' .. props.size .. '"/>'                                         -- pos 24
    parts[#parts + 1] = '<w:szCs w:val="' .. props.size .. '"/>'                                       -- pos 25
  end
  if props.underline then parts[#parts + 1] = '<w:u w:val="single"/>' end                              -- pos 27
  if props.bg then
    parts[#parts + 1] = '<w:shd w:val="clear" w:color="auto" w:fill="' .. props.bg .. '"/>'            -- pos 30
  end
  if vert_align == "superscript" then
    parts[#parts + 1] = '<w:vertAlign w:val="superscript"/>'                                           -- pos 32
  elseif vert_align == "subscript" then
    parts[#parts + 1] = '<w:vertAlign w:val="subscript"/>'                                             -- pos 32
  end
  if #parts == 0 then return "" end
  -- table.concat joins array items with the given separator (empty here).
  return "<w:rPr>" .. table.concat(parts) .. "</w:rPr>"
end

-- Wrap a piece of literal text into a single OOXML run with the supplied
-- formatting. xml:space="preserve" keeps leading/trailing whitespace
-- (Word would otherwise collapse it).
local function emit_run(text, props, vert_align)
  return pandoc.RawInline("openxml",
    "<w:r>" .. rpr_xml(props, vert_align)
    .. '<w:t xml:space="preserve">' .. escape_xml(text) .. "</w:t></w:r>")
end

-- In-place "extend a Lua array with another array" — Lua has no built-in.
local function append_all(target, items)
  for _, item in ipairs(items) do target[#target + 1] = item end
end

-- Forward declaration: walk and walk_with_flag recurse into each other,
-- and Lua resolves `local`s top-to-bottom, so we declare `walk` first
-- and assign it later.
local walk

-- Tiny convenience: clone props, set one boolean flag (bold/italic/...),
-- then recurse into nested content. Used for native AST nodes like Strong
-- that simply add one property on top of inherited formatting.
local function walk_with_flag(content, props, vert_align, key)
  local p = clone(props)
  p[key] = true
  return walk(content, p, vert_align)
end

-- Recursive content walker. Takes a list of Pandoc inline AST nodes plus
-- the run properties currently inherited from the surrounding spans, and
-- returns a flat list of pandoc.RawInline("openxml", ...) runs. The
-- caller (filter.Span below) splices that list back into the document.
walk = function(inlines, props, vert_align)
  local result = {}
  for _, inline in ipairs(inlines) do
    -- Every Pandoc AST node carries a `.t` tag identifying its type.
    local t = inline.t
    if t == "Str" then
      -- Plain text leaf — emit one run with current formatting.
      result[#result + 1] = emit_run(inline.text, props, vert_align)
    elseif t == "Space" or t == "SoftBreak" then
      -- Whitespace between words. SoftBreak (raw newline in source) is
      -- treated like a regular space for inline rendering.
      result[#result + 1] = emit_run(" ", props, vert_align)
    elseif t == "LineBreak" then
      -- Hard line break (<br/>). OOXML uses <w:br/> inside a run.
      result[#result + 1] = pandoc.RawInline("openxml", "<w:r><w:br/></w:r>")
    elseif t == "Strong" then
      append_all(result, walk_with_flag(inline.content, props, vert_align, "bold"))
    elseif t == "Emph" then
      append_all(result, walk_with_flag(inline.content, props, vert_align, "italic"))
    elseif t == "Underline" then
      append_all(result, walk_with_flag(inline.content, props, vert_align, "underline"))
    elseif t == "Strikeout" then
      append_all(result, walk_with_flag(inline.content, props, vert_align, "strikeout"))
    elseif t == "Superscript" then
      append_all(result, walk(inline.content, props, "superscript"))
    elseif t == "Subscript" then
      append_all(result, walk(inline.content, props, "subscript"))
    elseif t == "Span" then
      -- Nested span: parse its style (if any) on top of inherited props,
      -- then descend with the merged set. `attributes` may be nil when the
      -- span has no key/value attributes at all, hence the `inline.attributes and ...`
      -- guard (Lua's `and` short-circuits — if the left side is falsy
      -- the whole expression is that falsy value, never an indexing error).
      local style = inline.attributes and inline.attributes.style
      local p = props
      if style then p = merge_css(props, parse_style(style)) end
      append_all(result, walk(inline.content, p, vert_align))
    elseif t == "RawInline" then
      -- Already-OOXML content (e.g. produced by another filter or a previous
      -- pass): pass through verbatim. Format-mismatched RawInlines (latex/html
      -- in a docx output) would normally be dropped by the writer; we keep
      -- them in case downstream tooling expects them.
      result[#result + 1] = inline
    elseif t == "Code" then
      -- Inline code spans carry their text in `.text` (not in `.content`).
      -- We treat them as ordinary runs; the monospace styling that pandoc
      -- normally applies via the "VerbatimChar" style is intentionally
      -- dropped here in favour of the inherited inline-CSS formatting.
      result[#result + 1] = emit_run(inline.text, props, vert_align)
    elseif t == "Link" then
      -- Hyperlinks would need a relationship entry in word/_rels — we can't
      -- create those from a Lua filter. Keep the link text and lose the URL.
      append_all(result, walk(inline.content, props, vert_align))
    elseif t == "Image" then
      -- Images are embedded by Pandoc's DOCX writer as <w:drawing> with a
      -- relationship in word/_rels/document.xml.rels — that pipeline isn't
      -- accessible from a Lua filter, so we MUST pass the node through
      -- untouched. Stringifying would silently drop the image (alt text
      -- only) and is exactly the bug we just fixed.
      result[#result + 1] = inline
    else
      -- Anything else (Note, Cite, Math, Quoted, SmallCaps, ...) needs
      -- writer-level handling we can't replicate inline. Pass through; the
      -- DOCX writer will emit it correctly. The trade-off is that the
      -- surrounding span's color/font won't propagate into these nodes —
      -- which is the right call (a footnote marker shouldn't inherit a
      -- highlight, an Image has no text color, etc.).
      result[#result + 1] = inline
    end
  end
  return result
end

-- Pandoc Lua filters are tables of "ElementType -> function" entries.
-- Returning the table from the script registers all entries at once.
local filter = {}

-- Tell Pandoc to walk top-down (parents before children). With the default
-- bottom-up order, an inner colored span would be converted to RawInline
-- first, and then a wrapping Strong would silently drop its bold (the DOCX
-- writer emits RawInline verbatim, ignoring the surrounding AST).
filter.traverse = "topdown"

-- The actual entry point: any Span carrying a `style` attribute consumes
-- its entire subtree and is replaced by a flat list of OOXML runs.
-- Returning a list (instead of a single inline) splices in place. Returning
-- nil means "leave this node alone" — pandoc then continues normal traversal.
function filter.Span(el)
  if not el.attributes.style then return nil end
  local props = merge_css({}, parse_style(el.attributes.style))
  return walk(el.content, props, nil)
end

return filter
