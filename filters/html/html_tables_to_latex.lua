-- html_tables_to_latex.lua
--
-- Preserve table WIDTH and horizontal ALIGNMENT from an HTML source when the
-- target writer is LaTeX/PDF. Pandoc's HTML reader keeps the table's inline
-- `style` in the Table node's Attr, but the LaTeX writer ignores it: every
-- table comes out content-width and CENTERED (longtable's default LTleft/
-- LTright glue is `\fill` on both sides), so a table authored at `width: 40%`
-- or `margin-left: 0` still renders centered at its natural width.
--
-- This is the LaTeX/PDF counterpart to the DOCX post-processing in
-- app/DocxPostProcess.py + app/HtmlTableLayout.py. It reads the same
-- properties from the same `style` attribute:
--
--   * width  — `N%` sets the table to that fraction of the line width by
--              distributing the fraction across the column widths (pandoc's
--              LaTeX writer renders columns with an explicit width as a
--              fraction of \linewidth, so their sum becomes the table width).
--              An absolute length (`50px`, `pt`, `cm`, ...) is converted to a
--              fraction of a reference text width (Letter 8.5in minus 1in
--              margins each side = 6.5in), matching the Letter assumption in
--              app/DocxPostProcess.py; exact absolute widths are not
--              expressible through pandoc's fraction-only column model, so this
--              is best-effort (same caveat as the DOCX px path).
--   * align  — derived from `margin-left`/`margin-right` the way a browser
--              positions a block: 0/auto -> left, auto/auto -> center,
--              auto/0 -> right. Emitted as longtable \LTleft/\LTright glue.
--
-- Only runs for LaTeX/PDF targets (the controller gates it on html->pdf/latex,
-- and the FORMAT check is defensive).

-- Reference text width in points for converting absolute widths to a line
-- fraction: Letter (8.5in) minus 1in margins on each side = 6.5in = 468pt.
local REFERENCE_TEXT_WIDTH_PT = 468.0

-- CSS unit -> points. 1px = 0.75pt at the 96dpi CSS reference; 1pt = 1pt;
-- 1in = 72pt; 1pc = 12pt; 1cm = 28.3465pt; 1mm = 2.83465pt.
local UNIT_TO_PT = {
  px = 0.75,
  pt = 1.0,
  ["in"] = 72.0,
  pc = 12.0,
  cm = 28.3465,
  mm = 2.83465,
}

-- Split "k1: v1; k2: v2" into a table {k1=v1, k2=v2} with trimmed, lowercased
-- keys. Later declarations win (CSS cascade). `max-width` stays distinct from
-- `width` so only an exact `width` is read as the table width.
local function parse_declarations(style)
  local decls = {}
  for segment in style:gmatch("[^;]+") do
    local key, value = segment:match("^%s*([^:]-)%s*:%s*(.-)%s*$")
    if key then
      decls[key:lower()] = value
    end
  end
  return decls
end

-- Return the line fraction (0..1] for a CSS width value, or nil when it is
-- absent, `auto`, unparseable or non-positive. `N%` -> N/100 (clamped to 1);
-- an absolute length -> length_pt / reference_text_width (clamped to 1).
local function width_fraction(value)
  if not value then return nil end
  local number, unit = value:match("^%s*([%d%.]+)%s*(%a*%%?)%s*$")
  if not number then return nil end
  number = tonumber(number)
  if not number or number <= 0 then return nil end

  if unit == "%" then
    return math.min(number / 100.0, 1.0)
  end

  local factor = UNIT_TO_PT[unit ~= "" and unit or "px"]
  if not factor then return nil end
  return math.min((number * factor) / REFERENCE_TEXT_WIDTH_PT, 1.0)
end

-- Map the margin-left/margin-right pair to a table alignment, mirroring
-- app/HtmlTableLayout.py: an `auto` margin absorbs the free space on that side.
local function resolve_align(margin_left, margin_right)
  local left_auto = margin_left ~= nil and margin_left:lower():match("^%s*auto%s*$") ~= nil
  local right_auto = margin_right ~= nil and margin_right:lower():match("^%s*auto%s*$") ~= nil
  if left_auto and right_auto then return "center" end
  if left_auto then return "right" end
  if right_auto then return "left" end
  return nil
end

-- longtable positions itself with the \LTleft/\LTright glue lengths, read when
-- the environment begins. Both `\fill` is centered (pandoc's default); 0pt on
-- one side pins the table to that edge.
local ALIGN_GLUE = {
  left = "\\setlength{\\LTleft}{0pt}\\setlength{\\LTright}{\\fill}",
  center = "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{\\fill}",
  right = "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{0pt}",
}
-- Reset to longtable's centered default after each table we touched, so the
-- glue never leaks into a later table that carries no alignment of its own.
local RESET_GLUE = "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{\\fill}"

function Table(tbl)
  if not FORMAT:match("latex") then return nil end

  local attrs = tbl.attr and tbl.attr.attributes
  local style = attrs and attrs.style
  if not style then return nil end

  local decls = parse_declarations(style)
  local fraction = width_fraction(decls["width"])
  local align = resolve_align(decls["margin-left"], decls["margin-right"])
  if not fraction and not align then return nil end

  -- Width: distribute the target line fraction evenly across the columns. The
  -- source tables carry no per-column widths, so an even split is faithful and
  -- makes the column widths sum to the requested fraction of \linewidth.
  if fraction then
    local num_cols = #tbl.colspecs
    if num_cols > 0 then
      local per_column = fraction / num_cols
      for i = 1, num_cols do
        tbl.colspecs[i] = { tbl.colspecs[i][1], per_column }
      end
    end
  end

  -- Alignment: wrap the table in the matching \LTleft/\LTright glue and reset
  -- afterwards. When no alignment was recovered we leave pandoc's centered
  -- default untouched.
  if align then
    return {
      pandoc.RawBlock("latex", ALIGN_GLUE[align]),
      tbl,
      pandoc.RawBlock("latex", RESET_GLUE),
    }
  end
  return tbl
end
