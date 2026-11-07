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
-- The booktabs rules pandoc's writer puts around a table, as Word draws its lines:
-- none for a table of pandoc's Table style without a header row, a thin rule under
-- the header row with one, and LaTeX's own rules for a table with borders.
local RULES = {
  none = "\\def\\toprule{}\\def\\midrule{}\\def\\bottomrule{}",
  -- The line under the header row as Word draws a border of the default width: 2 eighths of a point,
  -- over the rows rather than between them
  head = "\\def\\toprule{}\\def\\bottomrule{}\\def\\midrule{\\noalign{\\kern-0.125pt\\hrule height 0.25pt\\kern-0.125pt}}",
  grid = false,
}

local function valid_hex(s)
  return s and #s == 6 and s:match("^%x%x%x%x%x%x$") ~= nil
end

-- Parse the sentinel payload into a properties table. Supported segments:
--   bg=RRGGBB  cell background colour (per cell)
--   tw=<0..1>  table width as a fraction of the line (table-level, first cell)
--   ta=left|center|right  table horizontal alignment (table-level, first cell)
--   rl=none|head|grid  the lines Word draws (table-level, first cell)
--   cc=1  the cells hold paragraphs of the Compact style (table-level, first cell)
--   lc=RRGGBB, lw=<points>  the colour and width of the lines of a grid
local function parse_payload(payload)
  local props = {}
  for segment in payload:gmatch("[^;]+") do
    local key, value = segment:match("^(%a+)=(.+)$")
    if key == "bg" and valid_hex(value:upper()) then
      props.bg = value:upper()
    elseif key == "tw" then
      local n = tonumber(value)
      if n and n > 0 and n <= 1 then
        props.tw = n
      end
    elseif key == "ta" and (value == "left" or value == "center" or value == "right") then
      props.ta = value
    elseif key == "ha" and value == "justify" then
      props.ha = value
    elseif key == "aw" and value == "1" then
      props.aw = true
    elseif key == "rl" and RULES[value] ~= nil then
      props.rl = value
    elseif key == "cc" and value == "1" then
      props.cc = true
    elseif key == "lc" and valid_hex(value:upper()) then
      props.lc = value:upper()
    elseif key == "lw" then
      local n = tonumber(value)
      if n and n > 0 and n <= 6 then
        props.lw = n
      end
    end
  end
  return props
end

-- Flatten the pure container blocks a cell's content may sit inside.
--
-- Pandoc's DOCX reader turns any indented paragraph into a BlockQuote, and
-- Polarion pads its table cells with a small <w:ind w:left="142"/>. Inside a
-- cell that reading is wrong twice over. LaTeX's quote environment adds about
-- 2.5em of margin on each side of an already narrow column, so the text wraps
-- and hyphenates where there was room for it; and the extra block makes the
-- cell multi-block, which the LaTeX writer renders through a \minipage. The
-- wrapper also buries the sentinel where consume_sentinel cannot see it, so
-- such a cell lost its background colour too — in the tables this was found
-- on, only the one empty (and therefore unindented) cell kept its shading.
--
-- A cell is not a quotation, so unwrapping loses nothing a table cell means.
local function flatten_cell_blocks(blocks)
  local out = {}
  local function add(bs)
    for _, b in ipairs(bs) do
      if b.t == "BlockQuote" or b.t == "Div" then
        add(b.content)
      else
        out[#out + 1] = b
      end
    end
  end
  add(blocks)
  return out
end

-- True when `blocks` holds a container flatten_cell_blocks would remove.
local function has_container_block(blocks)
  for _, b in ipairs(blocks) do
    if b.t == "BlockQuote" or b.t == "Div" then return true end
  end
  return false
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

-- Inject a raw LaTeX command at the very start of the cell's first block.
local function inject_raw_at_start(cell_blocks, latex)
  if #cell_blocks == 0 then
    -- Empty cell: insert a Plain block holding just the command.
    cell_blocks[1] = pandoc.Plain({ pandoc.RawInline("latex", latex) })
    return
  end

  local first_block = cell_blocks[1]
  if first_block.t == "Para" or first_block.t == "Plain" then
    table.insert(first_block.content, 1, pandoc.RawInline("latex", latex))
  else
    -- Non-inline block (e.g. CodeBlock, Table): prepend a Plain with the
    -- command so it applies to the whole cell.
    table.insert(cell_blocks, 1, pandoc.Plain({ pandoc.RawInline("latex", latex) }))
  end
end

local function inject_cellcolor(cell_blocks, hex)
  inject_raw_at_start(cell_blocks, "\\cellcolor[HTML]{" .. hex .. "}")
end

-- A cell's own horizontal alignment, as the declaration that realises it
-- inside a p{} column.
--
-- Word keeps alignment per paragraph, so a Polarion table routinely centres
-- one cell in an otherwise left-aligned column. Pandoc's LaTeX writer reads
-- only the COLUMN's alignment for an ordinary cell: a Cell's own alignment
-- reaches the output solely through \multicolumn, which pandoc emits for a
-- colspan but not for a rowspan and not for a plain cell. Up to pandoc 3.6 the
-- DOCX reader papered over this by folding the cells' alignment into the
-- colspec; 3.11 no longer does, so a centred header row came out flush left.
--
-- The alignment does survive on the Cell, so emit it in the cell itself. That
-- is also the only form that can reach a \multirow cell at all.
-- LaTeX's own \centering / \raggedleft / \raggedright cannot be used here,
-- because each of them also does \let\\\@centercr. What \\ has to mean
-- depends on where the cell content lands, and the filter cannot tell which:
--   * directly in a p{} cell, \\ ends the ROW, so \centering breaks the table
--     with "Extra alignment tab has been changed to \cr" unless it is paired
--     with \arraybackslash;
--   * inside the \minipage pandoc wraps a multi-line cell in, \\ is an
--     ordinary line break, and \arraybackslash turns it into \tabularnewline,
--     which fails with "Extra }, or forgotten \endgroup".
-- So neither \centering nor \centering\arraybackslash is right everywhere.
-- These macros set exactly the glue \centering and friends set and leave \\
-- alone, so it keeps whatever meaning its context gives it.
--
-- The trailing {} terminates the control word: the cell text follows
-- immediately, and "\\pdcCellCenteringNom" would be an undefined control
-- sequence.
--
-- AlignDefault is mapped too, and that is the point rather than a detail.
-- Word has no column alignment: a DOCX only ever carries <w:jc> per paragraph,
-- and a paragraph without one is left-aligned (this pipeline reads documents
-- as left-to-right, as the rest of the codebase does). Pandoc's DOCX reader
-- synthesises the COLUMN alignment from the cells it saw, so a single centred
-- cell can leave a whole column centred, and an unaligned cell in it would
-- silently inherit that and come out centred where Word shows it flush left.
-- Emitting each cell's own alignment makes the synthesised colspec irrelevant.
local ALIGN_LATEX = {
  AlignCenter = "\\pdcCellCentering{}",
  AlignRight = "\\pdcCellRaggedleft{}",
  AlignLeft = "\\pdcCellRaggedright{}",
  AlignDefault = "\\pdcCellRaggedright{}",
}

-- Justification cannot come from the Cell: pandoc's Alignment has no justified
-- value, so <w:jc w:val="both"> arrives as AlignLeft and "distribute" as
-- AlignDefault, and the mapping above would flush both left. It rides in on
-- the sentinel instead (ha=justify, see app/docx_table_pre_process.py) and
-- wins over the Cell's alignment, which for such a cell is pandoc's guess.
local JUSTIFY_LATEX = "\\pdcCellJustify{}"
-- The height of a row strut, for the first line of a cell of the header row
local HEAD_STRUT = "\\pdcHeadStrut{}"

local has_cellcolor = false  -- set when at least one cell gets \cellcolor
local has_align = false      -- set when at least one cell gets an alignment switch
local has_grid = false       -- set when a table is written as raw LaTeX
local NESTED = "pdc-nested"   -- the attribute of a table in a cell of another

-- True when a cell renders nothing, so there is nothing to align.
--
-- Counting blocks is not enough: consume_sentinel strips the sentinel text but
-- keeps the paragraph that held it, so a shaded spacer cell reaches this point
-- as a single empty Plain and would still be given an alignment switch.
local function cell_is_blank(blocks)
  for _, block in ipairs(blocks) do
    if block.t ~= "Para" and block.t ~= "Plain" then return false end
    if #block.content > 0 then return false end
  end
  return true
end

-- The space below the paragraph of a cell without a style, which
-- docx_document_styles_to_latex defines: Word sets the text at the top of
-- the cell and the space of the Normal style below it.
local PLAIN_CELL_SPACE = "\\ifdefined\\pdcPlainCellSpace\\rule[-\\dimexpr\\dp\\strutbox+\\pdcPlainCellSpace\\relax]{0pt}{0pt}\\fi"

-- A cell of Compact paragraphs ends as deep as the row: the last line of a cell
-- of several lines stands as far above the bottom of the row as a single line.
local COMPACT_CELL_DEPTH = "\\ifdefined\\pdcCompactCellDepth\\pdcCompactCellDepth\\fi"

-- ---- A row Word breaks across pages ----
-- Word breaks a row between the lines of a cell; longtable breaks only between rows. A row whose one
-- cell holds a paragraph of line breaks becomes a row per line, without lines between them, each as
-- high as a line of the body: the space of the cell stays above the first and below the last.
local ROW_SPAN = "pdc-row-span"   -- the rows a cell spans, which its \\multirow takes over
local COVERED = "pdc-covered"     -- a cell which only carries the colour of the cell above
local SPLIT = "pdc-split"      -- the attribute of a row: what follows it, see SPLIT_PREAMBLE
local NO_END = "pdc-no-end"    -- the attribute of a cell which gets no space at its end
local has_split = false
local SPLIT_PREAMBLE = table.concat({
  "\\makeatletter",
  "\\newsavebox{\\pdcRowStrut}",
  "\\newcommand{\\pdcSplitStrut}[1]{\\global\\setbox\\@arstrutbox\\hbox{\\vrule height\\ht\\strutbox depth#1 width\\z@}}",
  "\\newcommand{\\pdcSplitSave}{\\global\\setbox\\pdcRowStrut\\copy\\@arstrutbox\\vskip-\\dimexpr\\dp\\@arstrutbox-\\dp\\strutbox\\relax}",
  "\\newcommand{\\pdcSplitFirstMiddle}{\\noalign{\\pdcSplitSave\\pdcSplitStrut{\\dp\\strutbox}}}",
  "\\newcommand{\\pdcSplitFirstLast}{\\noalign{\\pdcSplitSave\\pdcSplitStrut{\\dp\\pdcRowStrut}}}",
  "\\newcommand{\\pdcSplitLast}{\\noalign{\\pdcSplitStrut{\\dp\\pdcRowStrut}}}",
  "\\newcommand{\\pdcSplitEnd}{\\noalign{\\global\\setbox\\@arstrutbox\\copy\\pdcRowStrut}}",
  "\\makeatother",
}, "\n")

-- The lines of the top level of a block, split at its line breaks; nil when a break sits deeper.
local function block_lines(block, breaks)
  local lines, current, found = {}, {}, 0
  for _, inline in ipairs(block.content) do
    if inline.t == "LineBreak" then
      lines[#lines + 1], current, found = current, {}, found + 1
    else
      current[#current + 1] = inline
    end
  end
  lines[#lines + 1] = current
  return found == breaks and lines or nil
end

-- The position of the picture of a line of only a picture, with the switches of the cell, or nil
local function lone_picture(line)
  local picture = nil
  for index, inline in ipairs(line) do
    if inline.t == "Image" or (inline.t == "Link" and #inline.content == 1 and inline.content[1].t == "Image") then
      if picture then return nil end
      picture = index
    elseif inline.t ~= "Space" and inline.t ~= "SoftBreak" and inline.t ~= "RawInline" then
      return nil
    end
  end
  return picture
end

local function hang_down(blocks)
  if #blocks == 0 then
    return blocks
  end
  local open = "\\raisebox{0pt}[\\height][0pt]{\\begin{minipage}[t]{\\linewidth}"
  local first = blocks[1]
  local color = first.content and first.content[1]
  local result = {}
  if color and color.t == "RawInline" and color.text:match("^\\cellcolor") then
    -- colortbl reads the colour of a cell at its top level
    first.content:remove(1)
    open = color.text .. open
  end
  result[1] = pandoc.RawBlock("latex", open)
  for _, block in ipairs(blocks) do
    result[#result + 1] = block
  end
  result[#result + 1] = pandoc.RawBlock("latex", "\\end{minipage}}")
  return result
end

-- Word sets a picture alone on a line of a cell from the top of the line down, the next line right
-- below it: a cell of TeX lines up with the others at the baseline of its first line, the bottom of
-- such a picture. See CELL_PICTURE_PREAMBLE.
local has_cell_picture = false
-- Below a picture taller than the line, the space a row has below the descent of the font of a line:
-- Word sets a picture alone on its line without the descent, and the space of the cell below it
local CELL_PICTURE_PREAMBLE = table.concat({
  "\\newsavebox{\\pdcCellPictureBox}",
  "\\makeatletter\\newcommand{\\pdcCellPicture}[1]{\\sbox{\\pdcCellPictureBox}{#1}\\ifdim\\ht\\pdcCellPictureBox>\\ht\\strutbox"
    .. "\\raisebox{\\dimexpr\\ht\\strutbox-\\ht\\pdcCellPictureBox\\relax}{\\usebox{\\pdcCellPictureBox}}"
    .. "\\ifdim\\dp\\@arstrutbox>\\dp\\strutbox\\rule[\\dimexpr\\ht\\strutbox-\\ht\\pdcCellPictureBox-\\dp\\@arstrutbox+\\dp\\strutbox\\relax]{0pt}{0pt}\\fi"
    .. "\\else\\usebox{\\pdcCellPictureBox}\\fi}\\makeatother",
}, "\n")

local function hang_pictures(rows)
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      for _, block in ipairs(cell.contents) do
        if block.t == "Plain" or block.t == "Para" then
          local content, line = pandoc.Inlines({}), {}
          local function flush()
            local picture = lone_picture(line)
            for index, inline in ipairs(line) do
              if index == picture then
                has_cell_picture = true
                content:insert(pandoc.RawInline("latex", "\\pdcCellPicture{"))
                content:insert(inline)
                content:insert(pandoc.RawInline("latex", "}"))
              else
                content:insert(inline)
              end
            end
            line = {}
          end
          for _, inline in ipairs(block.content) do
            if inline.t == "LineBreak" then
              flush()
              content:insert(inline)
            else
              line[#line + 1] = inline
            end
          end
          flush()
          block.content = content
        end
      end
    end
  end
end

local function split_row(row)
  local multi, breaks = nil, 0
  for index, cell in ipairs(row.cells) do
    if cell.row_span ~= 1 or cell.col_span ~= 1 or cell.attributes[ROW_SPAN] or cell.attributes[COVERED] then return nil end
    local count = 0
    pandoc.Blocks(cell.contents):walk({ LineBreak = function() count = count + 1 end })
    if count > 0 then
      if multi or #cell.contents ~= 1 or (cell.contents[1].t ~= "Plain" and cell.contents[1].t ~= "Para") then return nil end
      multi, breaks = index, count
    end
  end
  local lines = multi and block_lines(row.cells[multi].contents[1], breaks)
  if not lines then return nil end
  local rows = {}
  for number, line in ipairs(lines) do
    local cells = {}
    for index, cell in ipairs(row.cells) do
      local contents = {}
      if index == multi then
        contents = { pandoc.Plain(line) }
      elseif number == 1 then
        -- Its text runs on beside the rows of the other lines, as the cells of a row of Word flow apart
        contents = hang_down(cell.contents)
      else
        -- The background of the cell goes on in the rows of its lines
        local first = cell.contents[1]
        local color = first and first.content and first.content[1]
        if color and color.t == "RawInline" and color.text:match("^\\cellcolor") then
          contents = { pandoc.Plain({ color }) }
        end
      end
      local attr = pandoc.Attr("", {}, (index == multi and number == #lines) and {} or { [NO_END] = "1" })
      cells[index] = pandoc.Cell(contents, cell.alignment, 1, 1, attr)
    end
    local after = number == #lines and "End" or number == 1 and (#lines == 2 and "FirstLast" or "FirstMiddle") or number == #lines - 1 and "Last" or "Middle"
    rows[number] = pandoc.Row(cells, pandoc.Attr("", {}, { [SPLIT] = after }))
  end
  return rows
end

-- ---- A coloured cell merged down ----
-- colortbl colours a cell of \\multirow in its first row only, and pandoc writes nothing in the rows
-- below. A coloured cell merged down sets its own \\multirow, and the rows below get a cell of its colour.

-- The \\cellcolor among the switches at the start of a cell, and its position in the first block
local function cell_color(cell)
  local first = cell.contents[1]
  for index, inline in ipairs(first and first.content or {}) do
    if inline.t ~= "RawInline" then
      return nil
    elseif inline.text:match("^\\cellcolor") then
      return inline, index
    end
  end
  return nil
end

local function color_merged_cells(rows, columns)
  local carry = {}
  for c = 1, columns do carry[c] = nil end
  for _, row in ipairs(rows) do
    local cells, column, index = {}, 1, 1
    while column <= columns do
      local above = carry[column]
      if above and above.left > 0 then
        cells[#cells + 1] = pandoc.Cell({ pandoc.Plain({ above.color }) }, "AlignDefault", 1, 1, pandoc.Attr("", {}, { [COVERED] = "1" }))
        above.left = above.left - 1
        column = column + 1
      else
        local cell = row.cells[index]
        if cell == nil then break end
        index = index + 1
        local color, at = cell_color(cell)
        if cell.row_span > 1 and color and cell.col_span == 1 and #cell.contents == 1 then
          local span = cell.row_span
          local first = cell.contents[1]
          first.content:remove(at)
          first.content:insert(1, pandoc.RawInline("latex", color.text .. "\\multirow[t]{" .. span .. "}{=}{"))
          first.content:insert(pandoc.RawInline("latex", "}"))
          cell.attributes[ROW_SPAN] = tostring(span)
          cell.row_span = 1
          carry[column] = { color = color, left = span - 1 }
        end
        cells[#cells + 1] = cell
        column = column + cell.col_span
      end
    end
    for rest = index, #row.cells do cells[#cells + 1] = row.cells[rest] end
    row.cells = cells
  end
end

-- Each row of a table body a row per line where Word breaks it across pages: see split_row.
local function split_rows(tbl)
  for _, body in ipairs(tbl.bodies) do
    local rows = {}
    for _, row in ipairs(body.body) do
      for _, part in ipairs(split_row(row) or { row }) do
        rows[#rows + 1] = part
      end
    end
    if #rows ~= #body.body then
      has_split = true
      body.body = rows
    end
  end
end

-- After the rows of the lines of a row, the strut of each: see SPLIT_PREAMBLE.
local function split_markers(rows)
  for _, row in ipairs(rows) do
    local after = row.attributes[SPLIT]
    local last = row.cells[#row.cells]
    if after and after ~= "Middle" and last then
      local marker = pandoc.RawInline("latex", "\\pdcSplitRow{" .. after .. "}")
      local block = last.contents[#last.contents]
      if block and (block.t == "Plain" or block.t == "Para") then
        block.content:insert(marker)
      else
        last.contents:insert(pandoc.Plain({ marker }))
      end
    end
  end
end

local function split_latex(latex)
  return (latex:gsub("\\pdcSplitRow{(%a+)}(.-)\\\\", "%2\\\\\\pdcSplit%1\n"))
end

local function add_plain_cell_space(rows, raw)
  raw = raw or PLAIN_CELL_SPACE
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      local skip = cell.attributes[NO_END] ~= nil
      cell.attributes[NO_END] = nil
      local last = cell.contents[#cell.contents]
      if skip then
        last = false
      end
      if last == nil then
        cell.contents = { pandoc.Plain({ pandoc.RawInline("latex", raw) }) }
      elseif last and (last.t == "Plain" or last.t == "Para") then
        last.content:insert(pandoc.RawInline("latex", raw))
      end
    end
  end
end

-- Walk all rows in a row-set (head, body, foot) and process sentinels.
-- Injects cell background colour and per-cell alignment, and captures any
-- table-level width/alignment (carried on the first cell) into `layout`.
-- Records in `flags` which kinds of injection happened, so the preamble pass
-- loads only what the document actually uses. Returns true when anything in
-- the rows was rewritten.
local function process_rows(rows, layout, flags, head)
  local modified = false
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      -- Unwrap before reading the sentinel: it may be inside the wrapper.
      if has_container_block(cell.contents) then
        cell.contents = flatten_cell_blocks(cell.contents)
        modified = true
      end
      local props = consume_sentinel(cell.contents)
      -- Alignment first, background second: both insert at the front, so the
      -- \cellcolor ends up ahead of the declaration, which is where colortbl
      -- wants it.
      -- A blank cell has nothing to align, and on a cell with no blocks at
      -- all the injection would add a paragraph it did not have.
      local align_latex = nil
      if not cell_is_blank(cell.contents) then
        align_latex = (props and props.ha == "justify") and JUSTIFY_LATEX or ALIGN_LATEX[cell.alignment]
      end
      if align_latex then
        -- pandoc sets a cell of the header row on its last line: its first line gets the height of a
        -- row strut, as Word puts the space of the cell above the first line. Only a first block of text
        -- has that line: before a list or a table the strut would be a line of its own
        local first = cell.contents[1]
        local strut = (head and first and (first.t == "Para" or first.t == "Plain")) and HEAD_STRUT or ""
        inject_raw_at_start(cell.contents, align_latex .. strut)
        flags.align = true
        modified = true
      end
      if props then
        if props.bg then
          inject_cellcolor(cell.contents, props.bg)
          flags.cellcolor = true
          modified = true
        end
        if props.tw then layout.tw = props.tw end
        if props.ta then layout.ta = props.ta end
        if props.aw then layout.aw = true end
        if props.rl then layout.rl = props.rl end
        if props.cc then layout.cc = true end
        if props.lc then layout.lc = props.lc end
        if props.lw then layout.lw = props.lw end
      end
    end
  end
  return modified
end

-- ---- A table drawn as a grid ----
--
-- Word draws every side of each cell of such a table. pandoc's writer sets a
-- longtable with booktabs rules and no vertical lines. The table is written to
-- LaTeX here, and its column specification and rows get the lines: a vertical
-- line between and around the columns, and under each row a line across the
-- columns whose cell ends there, so a merged cell keeps no line inside it.

local ROW_END = "\\pdcGridRowEnd"

-- The index of the brace which closes the one at `open`.
local function closing_brace(s, open)
  local depth = 0
  for i = open, #s do
    local ch = s:sub(i, i)
    if ch == "{" then
      depth = depth + 1
    elseif ch == "}" then
      depth = depth - 1
      if depth == 0 then
        return i
      end
    end
  end
  return nil
end

-- The end of the column at `i` of a column specification: >{..}p{..}, p{..} or a letter.
local function column_end(body, i)
  local ch = body:sub(i, i)
  local stop = i
  if ch == ">" then
    stop = closing_brace(body, i + 1)
    if stop == nil then
      return nil
    end
    stop = stop + 1
  end
  if body:sub(stop + 1, stop + 1) == "{" then
    stop = closing_brace(body, stop + 1)
  end
  return stop
end

-- The column specification of pandoc's writer with a line around every column.
local function grid_columns(spec)
  local body = spec:gsub("^%s*@{}", ""):gsub("@{}%s*$", "")
  local columns = {}
  local i = 1
  while i <= #body do
    if body:sub(i, i):match("%s") then
      i = i + 1
    else
      local stop = column_end(body, i)
      if stop == nil then
        return nil
      end
      columns[#columns + 1] = body:sub(i, stop)
      i = stop + 1
    end
  end
  return "|" .. table.concat(columns, "|") .. "|"
end

-- A \multicolumn replaces the lines of the columns it spans: it gets its own.
local function grid_multicolumns(latex)
  local result, position = {}, 1
  while true do
    local start, finish = latex:find("\\multicolumn{%d+}{", position)
    local close = start and closing_brace(latex, finish)
    if close == nil then
      break
    end
    local first = not latex:sub(1, start - 1):match("&%s*$")
    result[#result + 1] = latex:sub(position, finish)
    result[#result + 1] = (first and "|" or "") .. latex:sub(finish + 1, close - 1) .. "|}"
    position = close + 1
  end
  result[#result + 1] = latex:sub(position)
  return table.concat(result)
end

-- The line under each row, in the order of the rows: across the columns whose cell ends in that row.
local function grid_row_lines(rows, columns)
  local carry, lines = {}, {}
  for c = 1, columns do
    carry[c] = 0
  end
  for index, row in ipairs(rows) do
    local occupied, column = {}, 1
    for c = 1, columns do
      occupied[c] = carry[c] > 0
    end
    for _, cell in ipairs(row.cells) do
      if cell.attributes[COVERED] then
        goto next_cell
      end
      while column <= columns and occupied[column] do
        column = column + 1
      end
      for c = column, math.min(column + cell.col_span - 1, columns) do
        carry[c] = tonumber(cell.attributes[ROW_SPAN]) or cell.row_span
      end
      column = column + cell.col_span
      ::next_cell::
    end
    local runs, from = {}, nil
    for c = 1, columns + 1 do
      local ends = c <= columns and carry[c] <= 1
      if ends and from == nil then
        from = c
      elseif not ends and from ~= nil then
        runs[#runs + 1] = { from, c - 1 }
        from = nil
      end
    end
    for c = 1, columns do
      carry[c] = math.max(carry[c] - 1, 0)
    end
    local split = row.attributes[SPLIT]
    if split and split ~= "End" then
      lines[index] = ""
    elseif #runs == 1 and runs[1][1] == 1 and runs[1][2] == columns then
      lines[index] = "\\hline"
    else
      local parts = {}
      for _, run in ipairs(runs) do
        parts[#parts + 1] = string.format("\\cline{%d-%d}", run[1], run[2])
      end
      lines[index] = table.concat(parts)
    end
  end
  return lines
end

-- What a cell of a grid may hold: the template of pandoc loads the packages of
-- a list, a picture or a note only for one it sees, and a grid is raw LaTeX.
local GRID_BLOCKS = { Plain = true, Para = true, RawBlock = true }
local GRID_INLINES = {
  Str = true, Space = true, SoftBreak = true, LineBreak = true, Emph = true, Strong = true,
  Underline = true, Span = true, RawInline = true, Link = true, Math = true, SmallCaps = true,
  Superscript = true, Subscript = true, Quoted = true, Code = true, Image = true, Strikeout = true,
}

local function simple_cells(rows)
  local simple = true
  local check = {
    Block = function(block)
      if not GRID_BLOCKS[block.t] or (block.t == "RawBlock" and block.text:find("\\begin{longtable}", 1, true)) then
        simple = false
      end
    end,
    Inline = function(inline)
      -- A table inside, written as LaTeX already, cannot sit in the LaTeX of another
      if not GRID_INLINES[inline.t] or (inline.t == "RawInline" and inline.text:find("\\begin{longtable}", 1, true)) then
        simple = false
      end
    end,
  }
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      pandoc.Div(cell.contents):walk(check)
    end
  end
  return simple
end

local function all_rows(tbl)
  local rows = {}
  local function add(list)
    for _, row in ipairs(list) do
      rows[#rows + 1] = row
    end
  end
  add(tbl.head.rows)
  for _, body in ipairs(tbl.bodies) do
    add(body.head)
    add(body.body)
  end
  add(tbl.foot.rows)
  return rows
end

-- The LaTeX of a table as the writer of a PDF sets its text: for the engines of Unicode it keeps a
-- curly quote as it is, where the smart extension of the LaTeX writer makes TeX quotes of it, which
-- the fonts without TeX ligatures print as they are
local LATEX = "latex-smart"

-- The LaTeX of a grid as a block. A picture of it stays an Image between the raw
-- LaTeX: pandoc writes out the file of a picture, a GIF as PNG, only for an Image it sees.
local function grid_block(latex, rows)
  local images = {}
  for _, row in ipairs(rows) do
    for _, cell in ipairs(row.cells) do
      pandoc.Div(cell.contents):walk({ Image = function(image) images[#images + 1] = image end })
    end
  end
  if #images == 0 then
    return pandoc.RawBlock("latex", latex)
  end
  local inlines, from = {}, 1
  for _, image in ipairs(images) do
    local written = pandoc.write(pandoc.Pandoc({ pandoc.Plain({ image }) }), LATEX):gsub("%s+$", "")
    local first, last = latex:find(written, from, true)
    if first == nil then
      return nil
    end
    inlines[#inlines + 1] = pandoc.RawInline("latex", latex:sub(from, first - 1))
    inlines[#inlines + 1] = image
    from = last + 1
  end
  inlines[#inlines + 1] = pandoc.RawInline("latex", latex:sub(from))
  return pandoc.Plain(inlines)
end

-- Word sets a table as wide as the line out to the cell margin on both sides, and each column of it
-- holds the margin on both sides of its text: pandoc shares the padding of the columns out by their
-- widths, which makes a narrow column too wide. A table of an absolute width is that wide with its
-- margins; a column of it narrower than a twentieth of the line keeps the width pandoc gives it.
local function word_widths(latex, absolute)
  local function width(fraction)
    if absolute and tonumber(fraction) >= 0.05 then
      return string.format("\\dimexpr%s\\linewidth-2\\tabcolsep\\relax", fraction)
    elseif absolute then
      return nil
    end
    return string.format("\\dimexpr%s\\linewidth+%s\\tabcolsep+%s\\tabcolsep-2\\tabcolsep\\relax", fraction, fraction, fraction)
  end
  -- A merged cell spans its columns with the margins between them: the width of Word holds them already
  latex = latex:gsub("%(\\linewidth %- %d+\\tabcolsep%) %* \\real{([%d.]+)} %+ %d+\\tabcolsep", width)
  return (latex:gsub("%(\\linewidth %- %d+\\tabcolsep%) %* \\real{([%d.]+)}", width))
end

-- A merged cell at the edge of a table drops the margin there (@{}), which the width of Word holds:
-- its text takes that margin back.
local function edge_widths(latex)
  local result, position = {}, 1
  while true do
    local start, finish = latex:find("\\multicolumn{%d+}{", position)
    local close = start and closing_brace(latex, finish)
    if close == nil then
      break
    end
    local spec = latex:sub(finish + 1, close - 1)
    local _, edges = spec:gsub("@{}", "")
    if edges > 0 then
      spec = spec:gsub("%-2\\tabcolsep\\relax}", string.format("-%d\\tabcolsep\\relax}", 2 - edges), 1)
    end
    result[#result + 1] = latex:sub(position, finish) .. spec .. "}"
    position = close + 1
  end
  result[#result + 1] = latex:sub(position)
  return table.concat(result)
end

-- The table with the lines of a grid, or nil when it cannot be set so.
local function grid_table(tbl, absolute)
  local rows = all_rows(tbl)
  if #tbl.colspecs == 0 or not simple_cells(rows) then
    return nil
  end
  local lines = grid_row_lines(rows, #tbl.colspecs)
  for index, row in ipairs(rows) do
    local last = row.cells[#row.cells]
    if last ~= nil then
      table.insert(last.contents, pandoc.Plain({ pandoc.RawInline("latex", string.format("%s{%d}", ROW_END, index)) }))
    end
  end
  local latex = pandoc.write(pandoc.Pandoc({ tbl }), LATEX)
  local _, open = latex:find("\\begin{longtable}[]{", 1, true)
  local close = open and closing_brace(latex, open)
  local columns = close and grid_columns(latex:sub(open + 1, close - 1))
  if columns == nil then
    return nil
  end
  latex = edge_widths(word_widths(grid_multicolumns(latex:sub(1, open) .. columns .. latex:sub(close)), absolute))
  -- Word sets the text of a merged cell at its top
  latex = latex:gsub("\\multirow{", "\\multirow[t]{")
  -- Each row ends with the line under it
  latex = latex:gsub("%s*\\pdcGridRowEnd{(%d+)}(.-)\\\\", function(index, between)
    return between .. "\\\\" .. (lines[tonumber(index)] or "")
  end)
  if latex:find(ROW_END, 1, true) then
    return nil
  end
  return grid_block(split_latex(latex), rows)
end

-- ---- Pass 1: Table processing ----

-- longtable positions itself via the \LTleft/\LTright glue read at \begin;
-- both \fill is centered (pandoc's default). Pinning one side to 0pt flushes
-- the table to that edge. Mirrors filters/html/html_tables_to_latex.lua.
local ALIGN_GLUE = {
  left = "\\setlength{\\LTleft}{0pt}\\setlength{\\LTright}{\\fill}",
  center = "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{\\fill}",
  right = "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{0pt}",
}
local RESET_GLUE = "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{\\fill}"

local table_pass = {}

function table_pass.Table(tbl)
  if not FORMAT:match("latex") then return nil end

  local modified = false
  local layout = {}
  local flags = {}

  if process_rows(tbl.head.rows, layout, flags, true) then modified = true end
  for _, body in ipairs(tbl.bodies) do
    if process_rows(body.body, layout, flags) then modified = true end
    if process_rows(body.head, layout, flags) then modified = true end
  end
  if process_rows(tbl.foot.rows, layout, flags) then modified = true end

  hang_pictures(all_rows(tbl))

  -- A table written as LaTeX, a grid or one of simple cells, breaks a row as Word does
  local grid_lines = layout.rl == "grid" and layout.lc and layout.lw
  if tbl.attributes[NESTED] == nil and #tbl.colspecs > 0 and (grid_lines or (layout.tw and not layout.aw)) and simple_cells(all_rows(tbl)) then
    for _, body in ipairs(tbl.bodies) do
      color_merged_cells(body.body, #tbl.colspecs)
    end
    split_rows(tbl)
    split_markers(all_rows(tbl))
  end

  local cell_end = not layout.cc and PLAIN_CELL_SPACE or COMPACT_CELL_DEPTH
  do
    add_plain_cell_space(tbl.head.rows, cell_end)
    for _, body in ipairs(tbl.bodies) do
      add_plain_cell_space(body.head, cell_end)
      add_plain_cell_space(body.body, cell_end)
    end
    add_plain_cell_space(tbl.foot.rows, cell_end)
    modified = true
  end

  -- Only what was actually emitted: colortbl redefines internal table macros,
  -- so it is not loaded for a document that has no \cellcolor in it.
  if flags.cellcolor then has_cellcolor = true end
  if flags.align then has_align = true end

  -- Table width: pandoc's DOCX reader normalises column widths to sum to 1.0
  -- (discarding the table's <w:tblW> share of the line) OR, when the DOCX
  -- carries no usable column widths, emits none at all — in which case the
  -- LaTeX writer sizes the table to its content and centers it. Force an
  -- explicit width on every column: a column with no width falls back to an
  -- equal share, then all are scaled to the recovered fraction. This makes a
  -- 40% table render at 40% and a 100% table fill the full line (flush-left
  -- via the glue below) instead of floating content-width in the middle.
  if layout.tw then
    local num_cols = #tbl.colspecs
    if num_cols > 0 then
      -- Read each column's relative width (a column with none gets an equal
      -- share) and total them. Normalising by this total — rather than just
      -- multiplying each width by tw — guarantees the columns sum to EXACTLY
      -- tw of the line, so a 100% table fills the full text column even if
      -- pandoc handed us widths that summed to less than 1.0.
      local bases, total = {}, 0
      for i, colspec in ipairs(tbl.colspecs) do
        bases[i] = colspec[2] or (1.0 / num_cols)
        total = total + bases[i]
      end
      if total <= 0 then
        total = 1.0
      end
      for i, colspec in ipairs(tbl.colspecs) do
        tbl.colspecs[i] = { colspec[1], (bases[i] / total) * layout.tw }
      end
    end
  end

  -- Raw-LaTeX prologue/epilogue around the table, resetting afterwards so
  -- nothing leaks into a later table:
  --  * absolute-width (px/pt) tables get a tightened \tabcolsep — LaTeX's fixed
  --    inter-column padding otherwise leaves a small table (e.g. 50px, 3 cols)
  --    noticeably wider than requested; a group restores the padding of the
  --    document after it.
  --  * the recovered horizontal alignment is applied via longtable
  --    \LTleft/\LTright glue (pandoc drops <w:jc>, defaulting to centered).
  local before, after = {}, {}
  local grid = nil
  if layout.rl == "grid" and layout.lc and layout.lw then
    grid = grid_table(tbl, layout.aw)
  end
  if grid then
    -- The lines of the cells in their colour and width, the table out to the cell margin as Word sets it
    has_cellcolor = true
    has_grid = true
    before[#before + 1] = string.format("\\begingroup\\def\\toprule{\\hline}\\def\\midrule{}\\def\\bottomrule{}\\setlength{\\arrayrulewidth}{%gpt}\\arrayrulecolor[HTML]{%s}", layout.lw, layout.lc)
    if layout.ta == nil or layout.ta == "left" then
      before[#before + 1] = "\\setlength{\\LTleft}{-\\tabcolsep}\\setlength{\\LTright}{\\fill}"
      layout.ta = nil
      after[#after + 1] = RESET_GLUE
    end
  end
  -- A group, so the rules come back for the next table
  local rules = layout.rl and RULES[layout.rl]
  if rules then
    before[#before + 1] = "\\begingroup" .. rules
  end
  local rows = all_rows(tbl)
  -- Written as LaTeX, a table gets the columns of Word; pandoc's table its own
  local nested = tbl.attributes[NESTED] ~= nil
  tbl.attributes[NESTED] = nil
  local simple = not grid and not nested and not layout.aw and layout.tw and #tbl.colspecs > 0 and simple_cells(rows)
  if layout.aw and not grid and not simple then
    -- A group: the table after this one gets the padding the document sets, not 6pt
    before[#before + 1] = "\\begingroup\\setlength{\\tabcolsep}{3pt}"
    after[#after + 1] = "\\endgroup"
  end
  if layout.cc then
    -- Rows as high as a paragraph of Compact, which docx_document_styles_to_latex defines
    before[#before + 1] = "\\begingroup\\ifdefined\\pdcCompactCellStretch\\renewcommand{\\arraystretch}{\\pdcCompactCellStretch}\\setlength{\\extrarowheight}{\\pdcCompactCellExtra}\\fi"
  end
  if not grid and layout.tw and layout.tw > 0.999 then
    -- A table as wide as the line starts its text at the margin, whatever its alignment
    layout.ta = "left"
  end
  local glue = layout.ta and ALIGN_GLUE[layout.ta]
  if glue then
    before[#before + 1] = glue
    after[#after + 1] = RESET_GLUE
  end

  if rules then
    after[#after + 1] = "\\endgroup"
  end
  if layout.cc then
    after[#after + 1] = "\\endgroup"
  end
  if grid then
    after[#after + 1] = "\\endgroup"
    return { pandoc.RawBlock("latex", table.concat(before)), grid, pandoc.RawBlock("latex", table.concat(after)) }
  end
  local raw = simple and grid_block(split_latex(edge_widths(word_widths(pandoc.write(pandoc.Pandoc({ tbl }), LATEX), layout.aw))), rows)
  if raw then
    has_grid = true
    return { pandoc.RawBlock("latex", table.concat(before)), raw, pandoc.RawBlock("latex", table.concat(after)) }
  end
  if #before > 0 then
    return { pandoc.RawBlock("latex", table.concat(before)), tbl, pandoc.RawBlock("latex", table.concat(after)) }
  end

  if not modified and not layout.tw then return nil end
  return tbl
end

-- ---- Pass 2: Conditional preamble injection ----
-- Only adds \usepackage{colortbl} when at least one cell was coloured in
-- pass 1.  colortbl redefines internal table macros and can interact with
-- longtable or custom preambles, so we avoid loading it unnecessarily.

local COLORTBL_PREAMBLE = "\\usepackage{colortbl}"

-- The \\ -preserving counterparts of \centering / \raggedleft / \raggedright,
-- copied from the LaTeX kernel definitions minus their \let\\\@centercr.
-- See ALIGN_LATEX above for why that one line has to go.
local ALIGN_PREAMBLE = table.concat({
  "\\makeatletter",
  "\\providecommand{\\pdcCellCentering}{\\rightskip\\@flushglue \\leftskip\\@flushglue \\parindent\\z@ \\parfillskip\\z@skip}",
  "\\providecommand{\\pdcCellRaggedleft}{\\rightskip\\z@skip \\leftskip\\@flushglue \\parindent\\z@ \\parfillskip\\z@skip}",
  "\\providecommand{\\pdcCellRaggedright}{\\@rightskip\\@flushglue \\rightskip\\@rightskip \\leftskip\\z@skip \\parindent\\z@}",
  -- LaTeX justifies by default, so this restores the kernel's own glue: it
  -- exists to override a colspec that set one of the ragged modes.
  "\\providecommand{\\pdcCellJustify}{\\rightskip\\z@skip \\leftskip\\z@skip \\parindent\\z@ \\parfillskip\\@flushglue}",
  "\\providecommand{\\pdcHeadStrut}{\\rule{0pt}{\\ht\\@arstrutbox}}",
  "\\makeatother",
}, "\n")

local preamble_pass = {}

function preamble_pass.Meta(meta)
  if not FORMAT:match("latex") then return nil end
  if has_grid then
    -- The template loads the packages of a table only for a Table it sees, and a grid is raw LaTeX
    meta.tables = true
    meta.multirow = true
  end
  if not (has_cellcolor or has_align or has_split or has_cell_picture) then return meta end

  local parts = {}
  if has_cell_picture then parts[#parts + 1] = CELL_PICTURE_PREAMBLE end
  if has_split then parts[#parts + 1] = SPLIT_PREAMBLE .. "\n\\providecommand{\\pdcSplitRow}[1]{}" end
  if has_cellcolor then parts[#parts + 1] = COLORTBL_PREAMBLE end
  if has_align then parts[#parts + 1] = ALIGN_PREAMBLE end
  local block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", table.concat(parts, "\n")) })
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
-- ---- Pass 0: A caption under a table ----
-- Word puts no more than the space before a paragraph between a table and the
-- paragraph after it. A caption, which app/docx_table_pre_process.py keeps as
-- PandocCaption, takes the space before of its style instead of the paragraph
-- skip; docx_document_styles_to_latex defines it.
local CAPTION_AFTER_TABLE = "\\ifdefined\\pdcCaptionBefore\\vspace{\\dimexpr\\pdcCaptionBefore-\\parskip\\relax}\\fi"
-- A caption above a table: Word puts its space after between them, the table
-- has none before it. longtable puts the paragraph skip there.
local CAPTION_BEFORE_TABLE = "\\ifdefined\\pdcCaptionAfter\\vspace{\\dimexpr\\pdcCaptionAfter-\\parskip\\relax}\\fi"

local caption_pass = {}

-- A table in a cell of another: latex_nested_tables.lua sets it, which needs the Table
function caption_pass.Table(tbl)
  if not FORMAT:match("latex") then return nil end
  local inner = false
  local mark = { Table = function(nested) nested.attributes[NESTED] = "1"; inner = true; return nested end }
  for _, row in ipairs(all_rows(tbl)) do
    for _, cell in ipairs(row.cells) do
      cell.contents = pandoc.Blocks(cell.contents):walk(mark)
    end
  end
  return inner and tbl or nil
end

function caption_pass.Blocks(blocks)
  if not FORMAT:match("latex") then return nil end
  for index = 2, #blocks do
    local block = blocks[index]
    if blocks[index - 1].t == "Table" and block.t == "Div" and block.attributes["custom-style"] == "PandocCaption" then
      table.insert(block.content, 1, pandoc.RawBlock("latex", CAPTION_AFTER_TABLE))
    end
    local previous = blocks[index - 1]
    if block.t == "Table" and previous.t == "Div" and previous.attributes["custom-style"] == "PandocCaption" then
      previous.content:insert(pandoc.RawBlock("latex", CAPTION_BEFORE_TABLE))
    end
  end
  return blocks
end

return { caption_pass, table_pass, preamble_pass, cleanup_pass }
