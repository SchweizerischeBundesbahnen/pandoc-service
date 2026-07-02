-- docx_tables_to_latex.lua
--
-- Companion to app/DocxTablePreProcess.py. The preprocessor prepends each
-- styled table cell's properties as a PUA-delimited sentinel to the first
-- paragraph of the cell:
--
--     Str "\u{E010}bg=RRGGBB\u{E011}..."
--
-- This filter intercepts Table nodes, scans each cell for the sentinel,
-- strips it, and injects \cellcolor[HTML]{RRGGBB} raw LaTeX so cell
-- background shading survives into the PDF output.
--
-- The filter also adds \usepackage{colortbl} to header-includes so that
-- \cellcolor is available.
--
-- Two-pass design: the first pass processes Table nodes and injects
-- \cellcolor; the second pass strips any residual sentinel text from Str
-- nodes.  This ordering is essential because pandoc's default bottom-up
-- traversal visits Str nodes before Table nodes — if both handlers were
-- in one pass, the Str cleanup would strip sentinels before the Table
-- handler saw them.
--
-- Only runs for LaTeX/PDF targets (the filter is gated on docx->latex in
-- the controller, but the FORMAT check is defensive).

local OPEN = "\u{E010}"
local CLOSE = "\u{E011}"

-- Pattern to extract the sentinel payload from a Str node's text.
-- The sentinel sits at the very start: \u{E010}<payload>\u{E011}<rest>.
-- We capture the payload and the rest of the text separately.
local SENTINEL_PATTERN = "^" .. OPEN .. "(.-)" .. CLOSE .. "(.*)"

-- Pattern to strip ALL sentinel occurrences from a string (cleanup pass).
local STRIP_PATTERN = OPEN .. ".-" .. CLOSE

-- Valid 6-char uppercase hex colour.
local function valid_hex(s)
  return s and #s == 6 and s:match("^%x%x%x%x%x%x$") ~= nil
end

-- Parse the sentinel payload into a properties table.
-- Currently only "bg=RRGGBB" is supported; additional segments (borders,
-- vAlign) can be added by extending this parser and the preprocessor.
local function parse_payload(payload)
  local props = {}
  for segment in payload:gmatch("[^;]+") do
    local key, value = segment:match("^(%a+)=(.+)$")
    if key == "bg" and valid_hex(value:upper()) then
      props.bg = value:upper()
    end
  end
  return props
end

-- Try to read and consume the sentinel from a cell's content blocks.
-- Returns the parsed properties table (or nil) and mutates the blocks
-- in-place to strip the sentinel text.
local function consume_sentinel(cell_blocks)
  -- Scan all blocks (not just the first) because a cell with a nested table
  -- may have the sentinel in a later paragraph block.
  for _, block in ipairs(cell_blocks) do
    if block.t == "Para" or block.t == "Plain" then
      local inlines = block.content
      if #inlines > 0 and inlines[1].t == "Str" then
        local payload, rest = inlines[1].text:match(SENTINEL_PATTERN)
        if payload then
          local props = parse_payload(payload)
          -- Strip the sentinel from the text.
          if rest == "" then
            table.remove(inlines, 1)
            if #inlines > 0 and inlines[1].t == "Space" then
              table.remove(inlines, 1)
            end
          else
            inlines[1] = pandoc.Str(rest)
          end
          return props
        end
      end
    end
  end
  return nil
end

-- Inject \cellcolor at the very start of the cell's first block.
local function inject_cellcolor(cell_blocks, hex)
  if #cell_blocks == 0 then
    -- Empty cell: insert a Plain block with just the cellcolor command.
    cell_blocks[1] = pandoc.Plain({ pandoc.RawInline("latex", "\\cellcolor[HTML]{" .. hex .. "}") })
    return
  end

  local first_block = cell_blocks[1]
  if first_block.t == "Para" or first_block.t == "Plain" then
    table.insert(first_block.content, 1, pandoc.RawInline("latex", "\\cellcolor[HTML]{" .. hex .. "}"))
  else
    -- Non-inline block (e.g. CodeBlock, Table): prepend a Plain with the
    -- cellcolor followed by a newline so the colour applies to the whole cell.
    table.insert(cell_blocks, 1, pandoc.Plain({ pandoc.RawInline("latex", "\\cellcolor[HTML]{" .. hex .. "}") }))
  end
end

-- Walk all rows in a row-set (head, body, foot) and process sentinels.
-- Returns true when at least one cell was modified.
local function process_rows(rows)
  local modified = false
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      local props = consume_sentinel(cell.contents)
      if props and props.bg then
        inject_cellcolor(cell.contents, props.bg)
        modified = true
      end
    end
  end
  return modified
end

-- ---- Pass 1: Table processing ----

local has_cellcolor = false  -- set when at least one cell gets \cellcolor

local table_pass = {}

function table_pass.Table(tbl)
  if not FORMAT:match("latex") then return nil end

  local modified = false

  if process_rows(tbl.head.rows) then modified = true end
  for _, body in ipairs(tbl.bodies) do
    if process_rows(body.body) then modified = true end
    if process_rows(body.head) then modified = true end
  end
  if process_rows(tbl.foot.rows) then modified = true end

  if not modified then return nil end
  has_cellcolor = true
  return tbl
end

-- ---- Pass 2: Conditional preamble injection ----
-- Only adds \usepackage{colortbl} when at least one cell was coloured in
-- pass 1.  colortbl redefines internal table macros and can interact with
-- longtable or custom preambles, so we avoid loading it unnecessarily.

local PREAMBLE = "\\usepackage{colortbl}"

local preamble_pass = {}

function preamble_pass.Meta(meta)
  if not FORMAT:match("latex") then return nil end
  if not has_cellcolor then return nil end

  local block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", PREAMBLE) })
  local existing = meta["header-includes"]
  if existing == nil then
    meta["header-includes"] = pandoc.MetaList({ block })
  elseif existing.t == "MetaList" then
    table.insert(existing, block)
    meta["header-includes"] = existing
  else
    meta["header-includes"] = pandoc.MetaList({ existing, block })
  end
  return meta
end

-- ---- Pass 3: Global sentinel cleanup ----
-- Strip any residual sentinel text from Str nodes that the Table handler
-- might have missed (e.g. sentinels in cells that pandoc wrapped in
-- unexpected block types).

local cleanup_pass = {}

function cleanup_pass.Str(s)
  local cleaned = s.text:gsub(STRIP_PATTERN, "")
  if cleaned ~= s.text then
    if cleaned == "" then
      -- Return an empty list to remove the node entirely.
      return {}
    end
    return pandoc.Str(cleaned)
  end
  return nil
end

-- Return all three passes in order: tables, conditional preamble, cleanup.
return { table_pass, preamble_pass, cleanup_pass }
