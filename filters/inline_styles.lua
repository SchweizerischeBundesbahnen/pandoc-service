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

-- Validate a CSS image dimension (the value of `width:`/`height:` in an
-- <img style="...">). Returns the value unchanged when it is a number with a
-- unit pandoc's image-size parser understands (px, in, cm, mm, pt, pc, % — or
-- unitless, which pandoc treats as px), else nil. Font-relative/viewport units
-- (em, ex, vw, vh, ...) are rejected: pandoc can't turn them into a concrete
-- <wp:extent>, so we leave the image unsized rather than emit garbage. Only a
-- validated value ever reaches the node attribute (never raw style text).
local SUPPORTED_IMG_UNITS = { [""] = true, px = true, ["in"] = true, cm = true, mm = true, pt = true, pc = true, ["%"] = true }
local function image_dim(value)
  if not value then return nil end
  -- A number (optional decimals) followed by an optional unit: either letters
  -- (px, in, cm, ...) or a literal "%". Surrounding whitespace is tolerated.
  local num, unit = value:match("^%s*(%d+%.?%d*)%s*([%a%%]*)%s*$")
  if num and SUPPORTED_IMG_UNITS[unit] then
    return num .. unit
  end
  return nil
end

-- Pandoc's HTML reader leaves <img style="width:..;height:.."> as an opaque
-- `style` attribute, and the DOCX writer sizes images only from the node's
-- `width`/`height` attributes — so CSS-sized images come out at native size.
-- Copy validated CSS width/height onto those attributes (a node-attribute tweak,
-- not the writer's unreachable embedding pipeline). Never overwrite an existing
-- width/height attribute (an <img width=..> or the value app/svg_processor.py
-- sets for rasterised SVGs); set only the side(s) given so the writer keeps the
-- aspect ratio when just one is present. Mutates the Image in place.
local function apply_image_style_dimensions(img)
  local style = img.attributes and img.attributes.style
  if not style then return end
  local props = parse_style(style)
  for _, dim in ipairs({ "width", "height" }) do
    local cur = img.attributes[dim]
    if not cur or cur == "" then
      local v = image_dim(props[dim])
      if v then img.attributes[dim] = v end
    end
  end
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
  -- modern long-hand. Child decorations are additive over inherited ones —
  -- in CSS, an ancestor's decoration draws through descendants regardless
  -- of what the child sets, so a nested span with `line-through` must not
  -- silently strip the inherited underline (and vice versa). The explicit
  -- `none` keyword is the only way to clear inherited decorations.
  local td = css["text-decoration"] or css["text-decoration-line"]
  if td then
    if td == "none" then
      p.underline = false
      p.strikeout = false
    else
      if has_token(td, "underline") then p.underline = true end
      if has_token(td, "line-through") then p.strikeout = true end
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
      -- Check if this is Polarion's caption-counter span. If so, emit a
      -- proper Word SEQ field instead of plain text so Word can renumber
      -- captions and resolve cross-references.
      local is_caption_span = false
      if inline.classes then
        for _, cls in ipairs(inline.classes) do
          if cls == "polarion-rte-caption" then
            is_caption_span = true
            break
          end
        end
      end
      if is_caption_span then
        -- Emit the SEQ field as raw OOXML so inlines_to_openxml succeeds
        -- and build_para_w_p can apply paragraph alignment. The Caption
        -- style is added by contains_caption_span() in build_para_w_p,
        -- and as a fallback by html_captions.lua for non-aligned captions.
        local seq_type = inline.attributes and (inline.attributes["sequence"] or inline.attributes["data-sequence"])
        if seq_type then
          local number_text = pandoc.utils.stringify(inline.content)
          result[#result + 1] = pandoc.RawInline("openxml",
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            .. '<w:r><w:instrText xml:space="preserve"> SEQ '
            .. escape_attr(seq_type) .. ' \\* ARABIC </w:instrText></w:r>'
            .. '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            .. '<w:r><w:t>' .. escape_xml(number_text) .. '</w:t></w:r>'
            .. '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        else
          result[#result + 1] = inline
        end
      else
        -- Nested span: parse its style (if any) on top of inherited props,
        -- then descend with the merged set.
        local style = inline.attributes and inline.attributes.style
        local p = props
        if style then p = merge_css(props, parse_style(style)) end
        append_all(result, walk(inline.content, p, vert_align))
      end
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
      -- Walk the link's content with the surrounding props so the styled
      -- span's color/font/highlight apply to the link text too, then wrap
      -- the walked runs back in a Link carrying the original target/title/
      -- attr. Pandoc's DOCX writer turns that Link into <w:hyperlink> and
      -- registers the relationship in word/_rels/document.xml.rels — we
      -- get the relationship side-effect for free as long as we hand it a
      -- Link node. Discarding the wrapper (the previous behaviour) kept
      -- the visible text but dropped the click target entirely.
      local walked = walk(inline.content, props, vert_align)
      result[#result + 1] = pandoc.Link(walked, inline.target, inline.title, inline.attr)
    elseif t == "Image" then
      -- Images are embedded by Pandoc's DOCX writer as <w:drawing> with a
      -- relationship in word/_rels/document.xml.rels — that pipeline isn't
      -- accessible from a Lua filter, so we MUST pass the node through
      -- untouched. Stringifying would silently drop the image (alt text
      -- only) and is exactly the bug we just fixed.
      --
      -- ...but we still copy CSS width/height onto the node's width/height
      -- attributes so the writer sizes it (see apply_image_style_dimensions).
      -- filter.Image does this for standalone images; doing it here too covers
      -- images that sit inside a styled span (idempotent — it never overwrites
      -- an already-set dimension).
      apply_image_style_dimensions(inline)
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

-- Module-level flag set by the Meta pass. The Table handler only activates
-- when callers pass `-M preserve_table_styles=true` on the command line.
local preserve_table_styles = false

-- ---- Pass 1: read metadata ----
local meta_pass = {}

function meta_pass.Meta(meta)
  if meta.preserve_table_styles then
    preserve_table_styles = true
  end
end

-- ---- Pass 2: rewrite AST elements ----
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

-- Standalone images (the common case — an <img style="width:.."> not wrapped in
-- a styled span) never enter walk(), so size them here from their CSS style.
function filter.Image(el)
  apply_image_style_dimensions(el)
  return el
end

-- True if any descendant inline is a Span carrying a `style` attribute.
local function has_styled_span(inlines)
  for _, inline in ipairs(inlines) do
    if inline.t == "Span" and inline.attributes and inline.attributes.style then
      return true
    end
    if inline.content and has_styled_span(inline.content) then
      return true
    end
  end
  return false
end

-- Native inline-formatting wrappers (Superscript/Subscript/Emph/Strong/...)
-- that ENCLOSE a styled <span>. Topdown traversal reaches such a wrapper
-- before the span, but without these handlers pandoc would just descend and
-- filter.Span would consume the span subtree with no knowledge of the
-- surrounding formatting — so the wrapper's superscript/bold/etc. is dropped
-- (the DOCX writer ignores the AST wrapper around the RawInline runs the span
-- produced). When the wrapper contains a styled span we therefore walk its
-- whole subtree ourselves, seeding the run properties with the wrapper's own
-- formatting; nested wrappers compose because walk() handles them in turn.
-- When there is no styled span inside, we return nil and let pandoc render the
-- wrapper natively (the common case — no inline CSS involved).
local function wrap_native(el, props, vert_align)
  if not has_styled_span(el.content) then return nil end
  return walk(el.content, props, vert_align)
end

function filter.Superscript(el) return wrap_native(el, {}, "superscript") end
function filter.Subscript(el) return wrap_native(el, {}, "subscript") end
function filter.Emph(el) return wrap_native(el, { italic = true }, nil) end
function filter.Strong(el) return wrap_native(el, { bold = true }, nil) end
function filter.Underline(el) return wrap_native(el, { underline = true }, nil) end
function filter.Strikeout(el) return wrap_native(el, { strikeout = true }, nil) end

-- Paragraph-level formatting (margin-left indent + text-align). Pandoc's HTML
-- reader drops the style attribute from <p> outright, so the formatting has to
-- ride into the AST on a wrapping Div that app/HtmlParagraphPreProcess.py
-- inserts. The contract:
--
--     <div class="pandoc-para" data-indent-twips="N" data-text-align="center"><p>...</p></div>
--
-- becomes Div("", ["pandoc-para"], [("indent-twips","N"),("text-align","center")])
-- [Para [...]] — pandoc strips the data- prefix off both attributes (the align
-- attribute is data-text-align rather than data-align precisely so this
-- stripping happens; "align" is a reserved HTML attribute pandoc would keep
-- prefixed). Each data-* attribute is optional. We rewrite each inner Para
-- into a raw OOXML <w:p> whose <w:pPr> carries <w:ind w:left="N"/> and/or
-- <w:jc w:val="..."/>. The Para's inlines are run through the same `walk()`
-- used by filter.Span, so nested <strong>/<em>/styled <span>s in the source
-- still render correctly inside the formatted paragraph.
--
-- Graceful degradation: if walk() returns anything that can't be embedded as
-- raw OOXML (a Link, Image, footnote, etc. — these need writer-level rels/
-- drawing handling that a raw <w:p> can't reproduce), we drop the paragraph
-- formatting rather than corrupt its content. The Para passes through with its
-- semantics intact, just without the indent/alignment applied.

-- Concatenate a list of inlines into a single OOXML string, or return nil
-- if any element can't be safely flattened (Link, Image, Note, ...).
local function inlines_to_openxml(inlines)
  local runs = walk(inlines, {}, nil)
  local parts = {}
  for _, r in ipairs(runs) do
    if r.t ~= "RawInline" or r.format ~= "openxml" then
      return nil
    end
    parts[#parts + 1] = r.text
  end
  return table.concat(parts)
end

-- True when `inlines` contains (at any depth) Polarion's caption counter span
-- (<span data-sequence=... class="polarion-rte-caption">). This is the reliable
-- signal that a paragraph is a real figure/table caption. A caption <p> also
-- carries text-align, so it reaches this filter as a pandoc-para wrapper and
-- would otherwise lose any paragraph style in the raw-OOXML rewrite below —
-- leaving app/DocxReferencesPostProcess.py unable to recognise it structurally.
-- (Captions without an alignment are handled separately by
-- filters/html_captions.lua, which runs on the un-wrapped Para.)
local function contains_caption_span(inlines)
  for _, il in ipairs(inlines) do
    if il.t == "Span" and il.classes then
      for _, class in ipairs(il.classes) do
        if class == "polarion-rte-caption" then return true end
      end
    end
    if type(il.content) == "table" and contains_caption_span(il.content) then
      return true
    end
  end
  return false
end

-- Build the single OOXML <w:p> for a formatted paragraph, or nil to signal
-- "fall back to the original Para". Both `twips` and `jc` are trusted here:
-- callers MUST pass an already-validated non-negative integer (Lua number, not
-- string) and an allowlisted justification literal. See filter.Div below for
-- the validation that establishes that trust. At least one of them must be
-- set — an empty <w:pPr> would be pointless, so we return nil in that case and
-- let the caller keep pandoc's native Para.
local function build_para_w_p(inlines, twips, jc)
  local body = inlines_to_openxml(inlines)
  if not body then return nil end
  -- <w:pPr> children must follow the CT_PPr schema sequence (ECMA-376 Part 1
  -- §17.3.1.26): <w:ind> precedes <w:jc>. Word/LibreOffice are version-
  -- dependent about out-of-order children (silent reorder, recovery warning,
  -- or dropping the child), so we emit them in canonical order.
  local ppr = {}
  -- <w:pStyle> comes first in the CT_PPr sequence. Stamp "Caption" on genuine
  -- captions so DocxReferencesPostProcess recognises them by style, not text.
  if contains_caption_span(inlines) then
    ppr[#ppr + 1] = '<w:pStyle w:val="Caption"/>'
  end
  -- string.format("%d", n) round-trips a validated integer through Lua's
  -- printf, which emits only decimal digits (and an optional leading '-' we
  -- already ruled out). Concatenating the raw attribute string here instead
  -- — as an earlier version did — would have let HTML input like
  --   <div class="pandoc-para" data-indent-twips='1"/><w:r>...</w:r><w:p w:x="'>
  -- close the <w:ind ...> early and splice arbitrary OOXML into the
  -- document. The validation in filter.Div + this %d format together close
  -- that injection path. `jc` is one of a fixed set of literal strings (never
  -- echoed input), so it carries no injection risk.
  if twips then
    ppr[#ppr + 1] = '<w:ind w:left="' .. string.format("%d", twips) .. '"/>'
  end
  if jc then
    ppr[#ppr + 1] = '<w:jc w:val="' .. jc .. '"/>'
  end
  if #ppr == 0 then return nil end
  return pandoc.RawBlock("openxml",
    "<w:p><w:pPr>" .. table.concat(ppr) .. "</w:pPr>" .. body .. "</w:p>")
end

local function has_class(el, name)
  for _, c in ipairs(el.classes) do
    if c == name then return true end
  end
  return false
end

-- Parse `data-indent-twips` into a non-negative Lua integer, or return nil
-- if the attribute is missing, non-numeric, negative, or fractional. This is
-- the trust boundary between attacker-controlled HTML and the raw OOXML we
-- splice into the document — every byte of `twips` ends up inside a
-- <w:ind w:left="..."/> attribute, so anything that isn't a clean integer
-- could escape the attribute and inject arbitrary XML.
local function parse_twips(raw)
  if not raw then return nil end
  local n = tonumber(raw)
  if not n then return nil end                  -- not numeric at all
  if n < 0 then return nil end                  -- negative — Word indents are >= 0
  if n ~= math.floor(n) then return nil end     -- fractional — would render as "600.5"
  -- Cap at a Word-sane upper bound. <w:ind w:left> is documented as a 32-bit
  -- signed integer in EMUs/twips; a value of ~200 inches is already far
  -- beyond any plausible paragraph indent and prevents stringly-large
  -- integers from showing up in the output.
  if n > 31680 then return nil end
  return math.floor(n)
end

-- Map the canonical text-align token (written by HtmlParagraphPreProcess) to
-- the OOXML <w:jc w:val> value, or nil for anything unrecognized. This is a
-- trust boundary: only values present as a *key* in this fixed table are ever
-- emitted, and the emitted string is the table's literal *value* — attacker
-- input is never echoed into the OOXML, so there is no attribute-injection
-- path the way there is for the free-form twips attribute.
local ALIGN_TO_JC = {
  left = "left",
  center = "center",
  right = "right",
  justify = "both",   -- CSS "justify" -> OOXML "both"
  both = "both",      -- accept the OOXML spelling too, defensively
}

local function parse_align(raw)
  if not raw then return nil end
  return ALIGN_TO_JC[raw]
end

function filter.Div(el)
  if not has_class(el, "pandoc-para") then return nil end
  local twips = parse_twips(el.attributes["indent-twips"])
  local jc = parse_align(el.attributes["text-align"])
  -- Nothing valid to apply — leave the Div for pandoc's normal handling
  -- rather than emit an empty/ malformed <w:pPr>.
  if not twips and not jc then return nil end

  local result = {}
  for _, block in ipairs(el.content) do
    if block.t == "Para" or block.t == "Plain" then
      local rb = build_para_w_p(block.content, twips, jc)
      result[#result + 1] = rb or block
    else
      -- Anything else in the wrapper (nested lists, code blocks, ...) keeps
      -- its normal writer treatment. Indent/alignment aren't applied to
      -- non-Para blocks — they have their own pPr semantics we shouldn't trample.
      result[#result + 1] = block
    end
  end
  return result
end

-- ======================== TABLE CELL STYLING ========================
--
-- Pandoc's DOCX writer ignores cell.attr on <td>/<th> elements, so CSS
-- properties like background-color and border on table cells are dropped.
-- This section detects styled tables and rebuilds them as raw OOXML
-- (<w:tbl>), preserving cell backgrounds, borders, and merges.
--
-- When no cell in a table carries a style attribute the handler returns
-- nil and lets the default DOCX writer emit the table normally.

-- CSS border-style keyword → OOXML w:val
local CSS_BORDER_TO_OOXML = {
  solid = "single", dashed = "dashed", dotted = "dotted",
  double = "double", groove = "single", ridge = "single",
  inset = "single", outset = "single", none = "nil", hidden = "nil",
}

-- Named CSS colors frequently seen in HTML tables.
local NAMED_COLORS = {
  black = "000000", white = "FFFFFF", red = "FF0000", green = "008000",
  blue = "0000FF", yellow = "FFFF00", gray = "808080", grey = "808080",
  silver = "C0C0C0", maroon = "800000", olive = "808000", navy = "000080",
  purple = "800080", teal = "008080", fuchsia = "FF00FF", aqua = "00FFFF",
  orange = "FFA500", windowtext = "000000",
}

-- Parse a CSS border-width value (e.g. "1.5pt", "1px") into OOXML eighths
-- of a point (w:sz units). Returns nil on failure.
local function border_width_eighths(val)
  if not val then return nil end
  local n, unit = val:lower():match("^([%d%.]+)(%a+)$")
  if not n then
    -- bare number → treat as pt
    n = tonumber(val)
    if n then return math.floor(n * 8 + 0.5) end
    return nil
  end
  n = tonumber(n)
  if not n then return nil end
  if unit == "pt" then return math.floor(n * 8 + 0.5)
  elseif unit == "px" then return math.floor(n * 6 + 0.5)
  elseif unit == "cm" then return math.floor(n * 226.77 + 0.5)
  elseif unit == "mm" then return math.floor(n * 22.677 + 0.5)
  elseif unit == "in" then return math.floor(n * 576 + 0.5)
  end
  return nil
end

-- Resolve a color token into uppercase 6-hex. Tries normalize_color first
-- (handles #RGB, #RRGGBB, rgb()), then falls back to named colors.
local function resolve_border_color(token)
  if not token then return nil end
  local c = normalize_color(token)
  if c then return c end
  return NAMED_COLORS[token:lower()]
end

-- Parse a CSS border shorthand ("1.5pt solid black") into
-- { sz = <eighths>, val = <ooxml_style>, color = <6hex> }, or nil.
local function parse_border(value)
  if not value then return nil end
  -- Collapse spaces inside rgb()/hsl() so the tokeniser below doesn't
  -- split "rgb(255, 0, 0)" into separate fragments.
  value = value:gsub("(%a+)(%b())", function(fn, parens)
    return fn .. parens:gsub("%s", "")
  end)
  local sz, val, color
  for p in value:gmatch("%S+") do
    local lp = p:lower()
    if CSS_BORDER_TO_OOXML[lp] then
      val = CSS_BORDER_TO_OOXML[lp]
    elseif lp:match("^[%d%.]+%a*$") then
      sz = sz or border_width_eighths(lp)
    else
      color = color or resolve_border_color(p)
    end
  end
  if not val and not sz and not color then return nil end
  return {
    sz    = sz    or 4,        -- default 0.5pt
    val   = val   or "single",
    color = color or "000000",
  }
end

-- Emit a single <w:XYZ .../> border element, or "" if nil/none.
local function border_element_xml(side, b)
  if not b or b.val == "nil" then return "" end
  return string.format('<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="%s"/>',
    side, b.val, b.sz, b.color)
end

-- Build <w:tcBorders>...</w:tcBorders> from a parsed CSS table, or "".
local function build_tc_borders_xml(css)
  local top    = parse_border(css["border-top"])    or parse_border(css["border"])
  local left   = parse_border(css["border-left"])   or parse_border(css["border"])
  local bottom = parse_border(css["border-bottom"]) or parse_border(css["border"])
  local right  = parse_border(css["border-right"])  or parse_border(css["border"])
  if not (top or left or bottom or right) then return "" end
  return "<w:tcBorders>"
    .. border_element_xml("top",    top)
    .. border_element_xml("left",   left)
    .. border_element_xml("bottom", bottom)
    .. border_element_xml("right",  right)
    .. "</w:tcBorders>"
end

-- Build <w:tcPr> for a cell. Follows CT_TcPr schema order (§17.4.70).
--   css            — parsed CSS table (nil = no styling)
--   col_span       — integer (1 = no span)
--   row_span       — integer (1 = no span)
--   vmerge_cont    — true for rowspan continuation cells
local function build_tc_pr_xml(css, col_span, row_span, vmerge_cont)
  local parts = {}
  -- pos 3: gridSpan
  if col_span and col_span > 1 then
    parts[#parts + 1] = string.format('<w:gridSpan w:val="%d"/>', col_span)
  end
  -- pos 5: vMerge
  if vmerge_cont then
    parts[#parts + 1] = "<w:vMerge/>"
  elseif row_span and row_span > 1 then
    parts[#parts + 1] = '<w:vMerge w:val="restart"/>'
  end
  if css then
    -- pos 6: tcBorders
    parts[#parts + 1] = build_tc_borders_xml(css)
    -- pos 7: shd (background-color)
    local bg = resolve_border_color(css["background-color"])
    if bg then
      parts[#parts + 1] = string.format(
        '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>', bg)
    end
    -- pos 12: vAlign
    local va = css["vertical-align"]
    if va == "top" then         parts[#parts + 1] = '<w:vAlign w:val="top"/>'
    elseif va == "middle" then  parts[#parts + 1] = '<w:vAlign w:val="center"/>'
    elseif va == "bottom" then  parts[#parts + 1] = '<w:vAlign w:val="bottom"/>'
    end
  end
  if #parts == 0 then return "<w:tcPr/>" end
  return "<w:tcPr>" .. table.concat(parts) .. "</w:tcPr>"
end

-- Convert a Pandoc Alignment value to a w:jc val, or nil for default.
local function alignment_to_jc(align)
  if align == pandoc.AlignLeft   then return "left"   end
  if align == pandoc.AlignCenter then return "center" end
  if align == pandoc.AlignRight  then return "right"  end
  return nil
end

-- Convert a single Block to a <w:p> OOXML string.
local function block_to_ooxml(block, jc_val)
  local ppr = ""
  if jc_val then
    ppr = '<w:pPr><w:jc w:val="' .. jc_val .. '"/></w:pPr>'
  end

  if block.t == "Para" or block.t == "Plain" then
    local runs = walk(block.content, {}, nil)
    local run_parts = {}
    for _, r in ipairs(runs) do
      if r.t == "RawInline" and r.format == "openxml" then
        run_parts[#run_parts + 1] = r.text
      elseif r.t == "Image" and r.src and r.src ~= "" then
        -- Images need writer-level relationship handling. Emit a
        -- placeholder for the Python post-processor.
        run_parts[#run_parts + 1] = "<w:r><w:t xml:space=\"preserve\">"
          .. "{{IMG:" .. escape_xml(r.src) .. "}}"
          .. "</w:t></w:r>"
      elseif r.t == "Link" then
        -- Links need writer-level .rels entries for the hyperlink target.
        -- Emit a <w:hyperlink> with a placeholder tooltip that encodes the
        -- URL. The Python post-processor registers the real relationship.
        -- Walk the link content, then replace rPr with just Hyperlink rStyle
        -- so the link renders blue/underlined. Inline CSS colors would
        -- override the Hyperlink style, so we strip them.
        local link_inlines = walk(r.content, {}, nil)
        local link_runs = {}
        for _, lr in ipairs(link_inlines) do
          if lr.t == "RawInline" and lr.format == "openxml" then
            -- Replace existing <w:rPr> with Hyperlink rStyle, and add rPr
            -- to bare <w:r> runs that don't have one (a single lr.text may
            -- contain multiple <w:r> elements).
            local text = lr.text
            text = text:gsub("<w:rPr>.-</w:rPr>", '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>')
            text = text:gsub("<w:r>(<w:t)", '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>%1')
            link_runs[#link_runs + 1] = text
          elseif lr.t == "Image" and lr.src and lr.src ~= "" then
            link_runs[#link_runs + 1] = "<w:r><w:t xml:space=\"preserve\">"
              .. "{{IMG:" .. escape_xml(lr.src) .. "}}"
              .. "</w:t></w:r>"
          else
            link_runs[#link_runs + 1] = '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
              .. '<w:t xml:space="preserve">'
              .. escape_xml(pandoc.utils.stringify(lr)) .. "</w:t></w:r>"
          end
        end
        run_parts[#run_parts + 1] = '<w:hyperlink w:tooltip="{{HREF:'
          .. escape_attr(r.target) .. '}}">'
          .. table.concat(link_runs) .. "</w:hyperlink>"
      else
        -- Nested tables, lists, etc. — fall back to plain text.
        run_parts[#run_parts + 1] = "<w:r><w:t xml:space=\"preserve\">"
          .. escape_xml(pandoc.utils.stringify(r))
          .. "</w:t></w:r>"
      end
    end
    return "<w:p>" .. ppr .. table.concat(run_parts) .. "</w:p>"
  end

  if block.t == "RawBlock" and block.format == "openxml" then
    return block.text
  end

  -- Fallback: extract plain text
  return "<w:p>" .. ppr .. "<w:r><w:t xml:space=\"preserve\">"
    .. escape_xml(pandoc.utils.stringify(block)) .. "</w:t></w:r></w:p>"
end

-- Convert all blocks in a cell to OOXML. Always returns at least "<w:p/>".
local function cell_blocks_to_ooxml(blocks, jc_val)
  if #blocks == 0 then return "<w:p/>" end
  local paras = {}
  for _, block in ipairs(blocks) do
    paras[#paras + 1] = block_to_ooxml(block, jc_val)
  end
  return table.concat(paras)
end

-- Check if any cell in a list of Rows carries a style attribute.
local function rows_have_styles(rows)
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      if cell.attributes and cell.attributes.style then return true end
    end
  end
  return false
end

-- Check the whole table for any styled cell.
local function table_has_cell_styles(tbl)
  if rows_have_styles(tbl.head.rows) then return true end
  for _, body in ipairs(tbl.bodies) do
    if rows_have_styles(body.body) then return true end
    if body.head and rows_have_styles(body.head) then return true end
  end
  if rows_have_styles(tbl.foot.rows) then return true end
  return false
end

-- Rebuild a styled Table as a raw OOXML <w:tbl> block.
-- Gated by the pandoc metadata variable `preserve_table_styles` — callers
-- pass `-M preserve_table_styles=true` on the command line to activate.
-- The flag is read in meta_pass.Meta (pass 1) and stored in the module-level
-- `preserve_table_styles` variable.
function filter.Table(tbl)
  -- Only emit raw OOXML when the target writer is docx.
  if FORMAT ~= "docx" then return nil end
  -- Check the opt-in metadata flag (set by meta_pass.Meta).
  if not preserve_table_styles then return nil end
  if not table_has_cell_styles(tbl) then return nil end

  local num_cols = #tbl.colspecs
  if num_cols == 0 then return nil end

  -- Flatten all rows in document order (head → body → foot).
  local all_rows = {}
  for _, r in ipairs(tbl.head.rows) do all_rows[#all_rows + 1] = r end
  for _, body in ipairs(tbl.bodies) do
    if body.head then
      for _, r in ipairs(body.head) do all_rows[#all_rows + 1] = r end
    end
    for _, r in ipairs(body.body) do all_rows[#all_rows + 1] = r end
  end
  for _, r in ipairs(tbl.foot.rows) do all_rows[#all_rows + 1] = r end
  if #all_rows == 0 then return nil end

  -- ---- Phase 1: build a logical grid with rowspan tracking ----
  -- coverage[row][col] = { col_span, css } for vMerge continuation cells.
  local coverage = {}
  local grid = {}

  for ri, row in ipairs(all_rows) do
    local cells_info = {}
    local gc = 1   -- grid column (1-based)
    local ci = 1   -- index into row.cells

    while gc <= num_cols do
      if coverage[ri] and coverage[ri][gc] then
        local cov = coverage[ri][gc]
        cells_info[#cells_info + 1] = {
          grid_col        = gc,
          col_span        = cov.col_span,
          is_continuation = true,
          css             = cov.css,
        }
        gc = gc + cov.col_span
      else
        if ci > #row.cells then break end
        local cell = row.cells[ci]; ci = ci + 1
        local cs = cell.col_span or 1
        local rs = cell.row_span or 1
        local style = cell.attributes and cell.attributes.style
        local css = style and parse_style(style) or nil

        if rs > 1 then
          for r2 = ri + 1, ri + rs - 1 do
            if not coverage[r2] then coverage[r2] = {} end
            coverage[r2][gc] = { col_span = cs, css = css }
          end
        end

        -- Resolve paragraph alignment: cell overrides colspec.
        local jc = nil
        if cell.alignment and cell.alignment ~= pandoc.AlignDefault then
          jc = alignment_to_jc(cell.alignment)
        elseif tbl.colspecs[gc] then
          jc = alignment_to_jc(tbl.colspecs[gc][1])
        end

        cells_info[#cells_info + 1] = {
          cell            = cell,
          grid_col        = gc,
          col_span        = cs,
          row_span        = rs,
          is_continuation = false,
          css             = css,
          jc              = jc,
        }
        gc = gc + cs
      end
    end

    grid[ri] = cells_info
  end

  -- ---- Phase 2: emit OOXML ----
  local xml = {}
  xml[#xml + 1] = "<w:tbl>"

  -- Determine table width from the HTML style attribute.
  -- - Percentage (e.g. "100%") → pct (fiftieths of a percent, so 100% = 5000)
  -- - Pixel value (e.g. "738px") → dxa (twips, 1px ≈ 15 twips at 96 dpi)
  -- - No width / unrecognised → default to 100 % (5000 pct)
  local tbl_style = tbl.attr and tbl.attr.attributes and tbl.attr.attributes.style
  local tbl_w_val = "5000"
  local tbl_w_type = "pct"
  local tbl_total_twips = nil  -- set when we know a fixed width in twips

  if tbl_style then
    local pct = tbl_style:match("width:%s*(%d+)%%")
    local px = tbl_style:match("width:%s*(%d+)px")
    if pct then
      tbl_w_val = tostring(math.floor(tonumber(pct) * 50))
      tbl_w_type = "pct"
    elseif px then
      tbl_total_twips = math.floor(tonumber(px) * 15)  -- 1px ≈ 15 twips
      tbl_w_val = tostring(tbl_total_twips)
      tbl_w_type = "dxa"
    end
  end

  -- Table properties with default single-line borders matching Pandoc's
  -- default DOCX writer output. Without these, tables that only set
  -- background-color (no per-cell border) render with no gridlines.
  xml[#xml + 1] = "<w:tblPr>"
  xml[#xml + 1] = '<w:tblW w:w="' .. tbl_w_val .. '" w:type="' .. tbl_w_type .. '"/>'
  xml[#xml + 1] = "<w:tblBorders>"
  xml[#xml + 1] = '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  xml[#xml + 1] = '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  xml[#xml + 1] = '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  xml[#xml + 1] = '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  xml[#xml + 1] = '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  xml[#xml + 1] = '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
  xml[#xml + 1] = "</w:tblBorders>"
  if tbl_total_twips then
    xml[#xml + 1] = '<w:tblLayout w:type="fixed"/>'
  else
    xml[#xml + 1] = '<w:tblLayout w:type="autofit"/>'
  end
  xml[#xml + 1] = "</w:tblPr>"

  -- Grid columns. When colspecs provide fractional widths, use them;
  -- otherwise distribute the total width evenly across columns.
  xml[#xml + 1] = "<w:tblGrid>"
  for i = 1, num_cols do
    local cw = tbl.colspecs[i] and tbl.colspecs[i][2]
    if cw and tbl_total_twips then
      xml[#xml + 1] = '<w:gridCol w:w="' .. math.floor(cw * tbl_total_twips) .. '"/>'
    elseif tbl_total_twips then
      xml[#xml + 1] = '<w:gridCol w:w="' .. math.floor(tbl_total_twips / num_cols) .. '"/>'
    else
      xml[#xml + 1] = "<w:gridCol/>"
    end
  end
  xml[#xml + 1] = "</w:tblGrid>"

  -- Rows
  for _, cells_info in ipairs(grid) do
    xml[#xml + 1] = "<w:tr>"
    for _, info in ipairs(cells_info) do
      xml[#xml + 1] = "<w:tc>"
      if info.is_continuation then
        xml[#xml + 1] = build_tc_pr_xml(info.css, info.col_span, nil, true)
        xml[#xml + 1] = "<w:p/>"
      else
        xml[#xml + 1] = build_tc_pr_xml(info.css, info.col_span, info.row_span, false)
        xml[#xml + 1] = cell_blocks_to_ooxml(info.cell.contents, info.jc)
      end
      xml[#xml + 1] = "</w:tc>"
    end
    xml[#xml + 1] = "</w:tr>"
  end

  xml[#xml + 1] = "</w:tbl>"

  return pandoc.RawBlock("openxml", table.concat(xml))
end

-- Return two passes: first reads metadata, second rewrites AST elements.
-- Pandoc executes them in order — the meta_pass sets module-level flags
-- that the main filter pass consults.
return { meta_pass, filter }
