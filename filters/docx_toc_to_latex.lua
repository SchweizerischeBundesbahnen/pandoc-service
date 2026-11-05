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
-- leader of close dots to the page number at the right margin, and each level
-- of the table of contents indented 12pt further.
--
-- Only runs for the LaTeX writer (gated in PandocController on
-- docx -> pdf/latex).

local CONTENTS = "PandocToc__CONTENTS_"
local ENTRY = "PandocToc__ENTRY"
local MARKER = "PDCPAGEREF"

local PREAMBLE = [[
\usepackage{titletoc}
\makeatletter
\newcommand{\pdcTocLeader}{\leaders\hbox to 0.275em{\hss.\hss}\hfill}
\newcommand{\pdcTableOfContents}[1]{\par\begingroup\hypersetup{linkcolor=black}\setcounter{tocdepth}{#1}\setlength{\parskip}{0pt}\@starttoc{toc}\endgroup}
\newcommand{\pdcTocEntryBegin}{\par\begingroup\hypersetup{linkcolor=black}\setlength{\parskip}{0pt}}
\newcommand{\pdcTocEntryEnd}{\par\endgroup\vspace{5pt}}
\newcommand{\pdcTocPage}[1]{\nobreak\pdcTocLeader\nobreak\pageref{#1}}
\makeatother
\contentsmargin{0pt}
\titlecontents{section}[0pt]{}{}{}{\pdcTocLeader\thecontentspage}[\vspace{5pt}]
\titlecontents{subsection}[12pt]{}{}{}{\pdcTocLeader\thecontentspage}[\vspace{5pt}]
\titlecontents{subsubsection}[24pt]{}{}{}{\pdcTocLeader\thecontentspage}[\vspace{5pt}]
\titlecontents{paragraph}[36pt]{}{}{}{\pdcTocLeader\thecontentspage}[\vspace{5pt}]
\titlecontents{subparagraph}[48pt]{}{}{}{\pdcTocLeader\thecontentspage}[\vspace{5pt}]
]]

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
  if depth ~= nil and depth >= 1 and depth <= 5 then
    return contents(math.floor(depth))
  end
  return nil
end

function Meta(meta)
  if not used then
    return nil
  end
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

-- Divs first, so Meta knows whether the document holds a table of contents.
return {
  { Div = Div },
  { Meta = Meta },
}
