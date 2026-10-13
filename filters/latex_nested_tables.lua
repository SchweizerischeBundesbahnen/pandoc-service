-- Set a table nested in a table cell as a tabular, which LaTeX can set in a cell.
--
-- pandoc writes every table as a longtable, and a nested one inside the
-- minipage of its cell. A longtable cannot stand in a minipage, and the TeX run
-- fails with "TeX capacity exceeded". The nested table stays a table for the
-- writer, which writes its images, merged cells and colors as anywhere else and
-- loads their packages. Only a group around it makes its longtable a tabular:
-- the head and foot markers do nothing, and the bottom rule closes the table.
--
-- Runs on the latex and pdf targets after the other filters.

local AS_TABULAR = table.concat({
  "\\begingroup",
  -- The \LTleft and \LTright glue the table filters set aligns the tabular as it did the longtable
  "\\def\\longtable{\\par\\noindent\\hskip\\LTleft\\relax\\tabular}",
  -- The rule of the document, saved by the outermost nested table only
  "\\ifdefined\\pandocnestedbottomrule\\else\\let\\pandocnestedbottomrule\\bottomrule\\fi",
  "\\def\\endlongtable{\\pandocnestedbottomrule\\endtabular\\hskip\\LTright\\relax\\par}",
  "\\def\\bottomrule{\\noalign{}}",
  "\\def\\endhead{}\\def\\endfirsthead{}\\def\\endfoot{}\\def\\endlastfoot{}",
}, "\n")

local function as_tabular(tbl)
  -- A table of the float class is written as a floating table, which no cell can hold
  tbl.attr.classes = tbl.attr.classes:filter(function(class)
    return class ~= "float"
  end)
  -- Columns of no width are as wide as their content, past the edge of the cell: they share it
  local defaults = 0
  for _, colspec in ipairs(tbl.colspecs) do
    if colspec[2] == nil or colspec[2] == pandoc.ColWidthDefault or colspec[2] == 0 then
      defaults = defaults + 1
    end
  end
  if defaults == #tbl.colspecs and defaults > 0 then
    for index, colspec in ipairs(tbl.colspecs) do
      tbl.colspecs[index] = { colspec[1], 1.0 / defaults }
    end
  end
  -- A caption needs a longtable or a float: it goes before the table as a paragraph
  local blocks = pandoc.Blocks(tbl.caption.long)
  tbl.caption = pandoc.Caption()
  -- An identifier makes pandoc write \caption{}\label{} still: an anchor before the table keeps it
  if tbl.identifier ~= "" then
    blocks:insert(pandoc.Plain({ pandoc.Span({}, { id = tbl.identifier }) }))
    tbl.identifier = ""
  end
  -- A longtable sets its foot at its end, a tabular where it is written, before the body
  if #tbl.foot.rows > 0 and #tbl.bodies > 0 then
    local last = tbl.bodies[#tbl.bodies]
    for _, row in ipairs(tbl.foot.rows) do
      last.body:insert(row)
    end
    tbl.bodies[#tbl.bodies] = last
    tbl.foot = pandoc.TableFoot()
  end
  blocks:insert(pandoc.RawBlock("latex", AS_TABULAR))
  blocks:insert(tbl)
  blocks:insert(pandoc.RawBlock("latex", "\\endgroup"))
  return blocks
end

local nested = { Table = as_tabular }

return {
  {
    traverse = "topdown",
    Table = function(tbl)
      -- The cells of an outer table, where every table is a nested one
      return tbl:walk(nested), false
    end,
  },
}
