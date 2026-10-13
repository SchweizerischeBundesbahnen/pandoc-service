-- Write a table nested in a table cell as a tabular, which LaTeX can set in a cell.
--
-- pandoc writes every table as a longtable, and a nested one inside the
-- minipage of its cell. A longtable cannot stand in a minipage, and the TeX run
-- fails with "TeX capacity exceeded". A nested table cannot break across pages
-- anyway, so it becomes a tabular of the same columns: its header is kept once,
-- and its foot moves to its end.
--
-- Runs on the latex and pdf targets after the other filters, so a nested table
-- carries the colors and cell properties they give it.

-- The cell properties docx_tables_to_latex reads from a sentinel at the start of
-- a cell; written into the tabular they would print as text.
local SENTINEL = "\u{E010}.-\u{E011}"

local function tabular(tbl)
  -- A table of the float class is written as a floating table, which no cell can hold
  tbl.attr.classes = tbl.attr.classes:filter(function(class)
    return class ~= "float"
  end)
  -- A caption needs a longtable or a float: it goes before the table as a paragraph
  local caption = tbl.caption.long
  tbl.caption = pandoc.Caption()
  local latex = pandoc.write(pandoc.Pandoc({ tbl }), "latex")
  latex = latex:gsub(SENTINEL, "")
  -- A header repeated on each page: keep the first one only
  latex = latex:gsub("\\endfirsthead.-\\endhead", "")
  -- The foot, set before the body in a longtable, closes a tabular
  local foot = ""
  latex = latex:gsub("\\endhead(.-)\\endlastfoot", function(rows)
    foot = rows
    return ""
  end)
  latex = latex:gsub("\\endhead", "")
  latex = latex:gsub("\\begin{longtable}%[[^%]]*%]", "\\begin{tabular}")
  latex = latex:gsub("\\begin{longtable}", "\\begin{tabular}")
  latex = latex:gsub("\\end{longtable}", function()
    return foot .. "\\end{tabular}"
  end)
  local blocks = pandoc.Blocks(caption)
  blocks:insert(pandoc.RawBlock("latex", latex))
  return blocks
end

local nested = { Table = tabular }

return {
  {
    traverse = "topdown",
    Table = function(tbl)
      -- The cells of an outer table, where every table is a nested one
      tbl = tbl:walk(nested)
      return tbl, false
    end,
  },
}
