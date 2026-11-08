-- docx_toc_to_latex.lua
--
-- Set the tables of contents of a docx in LaTeX/PDF as Word sets them.
--
-- app/docx_toc_pre_process.py rewrites the TOC fields of the docx into
-- paragraphs of synthetic styles, which the reader surfaces as Divs:
--   PandocToc__CONTENTS_<depth>  where Word puts its table of contents;
--                                it becomes LaTeX's, of that depth
--   PandocToc__ENTRY             an entry of a list of figures or tables,
--                                ending in the marker PDCPAGEREF<bookmark>;
--                                the marker becomes a dot leader and the
--                                page LaTeX sets the bookmark on
--
-- Both are set as Word sets its TOC styles: in the font and colour of the
-- body, without a heading of their own, a line apart with 5pt after it, a
-- leader of close dots to the page number at the right tab, which Word puts
-- 10 twips short of the right margin, and each level
-- of the table of contents indented 12pt further, down to the ninth.
-- An entry of a heading numbered by its style starts with the number, and its
-- text follows at the tab Word sets, see \pdcTocNumber.
-- The empty line Word leaves after a table of contents, see
-- app/docx_toc_pre_process.py, stands 5pt below the last entry: Word puts no
-- more than the larger of the space after one paragraph and before the next.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx -> pdf/latex).

local CONTENTS = "PandocToc__CONTENTS_"
local ENTRY = "PandocToc__ENTRY"
local MARKER = "PDCPAGEREF"

-- Word measures the number of an entry in its own default font, Aptos 12pt,
-- whatever the font of the document: the advance widths of Aptos, in 2048ths
-- of its size, of what a number can hold besides its digits, 1094 each.
local APTOS_WIDTHS = {
  ["."] = 585, ["-"] = 697, ["("] = 600, [")"] = 600,
  a = 1088, b = 1149, c = 1075, d = 1148, e = 1079, f = 617, g = 992, h = 1129, i = 489,
  j = 489, k = 997, l = 533, m = 1747, n = 1129, o = 1130, p = 1149, q = 1148, r = 685,
  s = 995, t = 662, u = 1144, v = 926, w = 1476, x = 906, y = 926, z = 897,
  A = 1207, B = 1237, C = 1418, D = 1405, E = 1139, F = 1074, G = 1451, H = 1448, I = 533,
  J = 678, K = 1164, L = 1023, M = 1618, N = 1446, O = 1500, P = 1182, Q = 1500, R = 1241,
  S = 1159, T = 980, U = 1395, V = 1199, W = 1827, X = 1132, Y = 1105, Z = 1056,
}
local APTOS_DIGIT = 1094

local function aptos_widths()
  local keys = {}
  for char in pairs(APTOS_WIDTHS) do
    keys[#keys + 1] = char
  end
  table.sort(keys)
  local defs = {}
  for _, char in ipairs(keys) do
    defs[#defs + 1] = string.format("\\@namedef{pdc@tocw@%s}{%d}", char, APTOS_WIDTHS[char])
  end
  return table.concat(defs)
end

-- \pdcTocNumber{indent}{number}: the number of an entry, in a box to the tab
-- Word sets after it. Word puts the tab at the first step of a grid of 12pt from
-- the margin which leaves 12pt after the number, as wide as Aptos 12pt sets it.
-- It does not measure a space. An empty number, of a level which only counts
-- for the numbers below it, sets nothing.
local function tableofcontents_preamble()
  local levels = { "section", "subsection", "subsubsection", "paragraph", "subparagraph", "pdcHeadingF", "pdcHeadingG", "pdcHeadingH", "pdcHeadingI" }
  local lines = {
    "\\usepackage{titletoc}",
    "\\makeatletter",
    "\\newcommand{\\pdcTocLeader}{\\leaders\\hbox to 0.275em{\\hss.\\hss}\\hfill}",
    "\\newcommand{\\pdcTableOfContents}[1]{\\par\\begingroup\\hypersetup{linkcolor=black}\\setcounter{tocdepth}{#1}\\setlength{\\parskip}{0pt}\\@starttoc{toc}\\endgroup\\vspace{-\\parskip}}",
    "\\newcommand{\\pdcTocEntryBegin}{\\par\\begingroup\\hypersetup{linkcolor=black}\\setlength{\\parskip}{0pt}}",
    "\\newcommand{\\pdcTocEntryEnd}{\\par\\endgroup\\vspace{5pt}}",
    "\\newcommand{\\pdcTocPage}[1]{\\nobreak\\pdcTocLeader\\nobreak\\pageref{#1}\\kern0.5bp}",
    aptos_widths(),
    "\\newcount\\pdc@tocn",
    string.format("\\def\\pdc@tocw#1{\\ifx\\relax#1\\else\\advance\\pdc@tocn\\ifcsname pdc@tocw@#1\\endcsname\\csname pdc@tocw@#1\\endcsname\\else %d\\fi\\relax\\expandafter\\pdc@tocw\\fi}", APTOS_DIGIT),
    "\\newcommand{\\pdcTocNumber}[2]{\\edef\\pdc@tocl{#2}\\ifx\\pdc@tocl\\@empty\\else\\pdc@tocn=0 \\expandafter\\pdc@tocw\\pdc@tocl\\relax"
      .. "\\hbox to \\dimexpr\\numexpr(\\dimexpr#1+12bp*\\pdc@tocn/2048+12bp\\relax+\\dimexpr6bp\\relax-1)/\\dimexpr12bp\\relax\\relax\\dimexpr12bp\\relax-#1\\relax{\\pdc@tocl\\hss}\\fi}",
    "\\makeatother",
    "\\contentsmargin{0pt}",
  }
  for depth, command in ipairs(levels) do
    local indent = string.format("%dpt", 12 * (depth - 1))
    lines[#lines + 1] = string.format("\\titlecontents{%s}[%s]{}{\\pdcTocNumber{%s}{\\thecontentslabel}}{}{\\pdcTocLeader\\thecontentspage\\kern0.5bp}[\\vspace{5pt}]", command, indent, indent)
  end
  return table.concat(lines, "\n") .. "\n"
end

local used = false

local function page_marker(str)
  local bookmark = str.text:match("^" .. MARKER .. "([%w_]+)$")
  if bookmark then
    return pandoc.RawInline("latex", "\\pdcTocPage{" .. bookmark .. "}")
  end
  return nil
end

local function contents(depth)
  used = true
  return pandoc.RawBlock("latex", string.format("\\pdcTableOfContents{%d}", depth))
end

local function entry(div)
  used = true
  local walked = div:walk({ Str = page_marker })
  local blocks = { pandoc.RawBlock("latex", "\\pdcTocEntryBegin") }
  for _, block in ipairs(walked.content) do
    blocks[#blocks + 1] = block
  end
  blocks[#blocks + 1] = pandoc.RawBlock("latex", "\\pdcTocEntryEnd")
  return blocks
end

function Div(div)
  local style = div.attributes["custom-style"]
  if style == nil or not FORMAT:match("latex") then
    return nil
  end
  if style == ENTRY then
    return entry(div)
  end
  local depth = style:sub(1, #CONTENTS) == CONTENTS and tonumber(style:sub(#CONTENTS + 1)) or nil
  if depth ~= nil and depth >= 1 and depth <= 9 then
    return contents(math.floor(depth))
  end
  return nil
end

function Meta(meta)
  if not used then
    return nil
  end
  local block = pandoc.MetaBlocks({ pandoc.RawBlock("latex", tableofcontents_preamble()) })
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

-- Whether the document has a heading for a table of contents to list.
local function has_heading(doc)
  local found = false
  doc:walk({
    Header = function()
      found = true
    end,
    Div = function(div)
      found = found or div.attributes["pdc-style"] == "Heading"
    end,
  })
  return found
end

-- Word fills a table of contents without a heading to list with its message, in
-- the paragraph which ends the field: no empty line stays after it.
local NO_ENTRIES = "No table of contents entries found."

local function is_contents(block)
  local style = block.t == "Div" and block.attributes["custom-style"] or ""
  return style:sub(1, #CONTENTS) == CONTENTS
end

local function is_empty_line(block)
  if block.t ~= "Para" and block.t ~= "Plain" then
    return false
  end
  for _, inline in ipairs(block.content) do
    if inline.t ~= "RawInline" then
      return false
    end
  end
  return true
end

local function empty_contents(doc)
  if not FORMAT:match("latex") or has_heading(doc) then
    return nil
  end
  local blocks, skip = {}, false
  for _, block in ipairs(doc.blocks) do
    if skip and is_empty_line(block) then
      skip = false
    elseif is_contents(block) then
      blocks[#blocks + 1] = pandoc.Para({ pandoc.RawInline("latex", "\\strut{}"), pandoc.Strong({ pandoc.Str(NO_ENTRIES) }), pandoc.RawInline("latex", "\\strut{}") })
      skip = true
    else
      skip = false
      blocks[#blocks + 1] = block
    end
  end
  doc.blocks = blocks
  return doc
end

-- Divs first, so Meta knows whether the document holds a table of contents.
return {
  { Pandoc = empty_contents },
  { Div = Div },
  { Meta = Meta },
}
