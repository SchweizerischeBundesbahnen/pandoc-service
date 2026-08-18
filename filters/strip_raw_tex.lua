--[[
Keep the TeX a document carries away from the engine which runs it.

`--sandbox` keeps pandoc itself from reading an address or a path a document
names, but a PDF is produced by handing the generated LaTeX to tectonic, which
runs outside that sandbox and resolves what the TeX names. Three routes reach
it, all verified against this image:

  * raw TeX,          `\input{/etc/passwd}` written as a raw block or inline,
  * math,             the same primitive inside `$...$`, which the writer emits
                      verbatim, however the formula spells it,
  * an image path,    which becomes `\includegraphics{...}` and puts the file
                      into the PDF.

This filter runs first, so it removes what the reader produced from the
document. The raw LaTeX the docx filters emit afterwards is untouched, which is
what keeps the colors, lists and tables of a DOCX export working.

A formula which reaches for such a primitive is dropped whole: editing TeX to
keep the rest of it would be guesswork, and a formula is not where a document
reads a file. Every other formula renders as before. Images are handled by
strip_document_images.lua.
]]

local TEX_FORMATS = {tex = true, latex = true, context = true}

-- Pandoc folds the case of a format, and the markdown raw_attribute parser keeps
-- it as written, so `{=LaTeX}` has to match too.
local function is_tex(format)
  return TEX_FORMATS[format:lower()] == true
end

-- A formula is the one thing of a document which the LaTeX writer emits
-- verbatim, so it is held to a closed rule rather than to a list of the names
-- which have been thought of. Two of the three rules below remove the ways a
-- formula can name something it does not spell out, which is what makes the
-- third one, a list, complete: what is left is the primitives themselves, and
-- a formula cannot make more of those.

-- 1. Names which the engine resolves to a file, and the machinery which builds
--    a name out of parts, aliases one, or reads a character as another. Losing
--    the machinery is what closes the list: `\@@input` was reached through
--    `\makeatletter`, and `\input` spelled `^^5cinput` through the notation.
local DENIED = {}
for _, name in ipairs({
  -- the engine reads, writes or runs something
  "input", "endinput", "include", "includegraphics", "openin", "closein",
  "read", "readline", "openout", "closeout", "write", "immediate", "special",
  "font", "XeTeXpdffile", "XeTeXpicfile", "shellescape", "batchmode",
  -- A name matches as a whole run of letters, so the pdfTeX spellings are
  -- listed rather than reached by their stem. This engine answers "undefined
  -- control sequence" to all of them today, which is checked in the tests, so
  -- the list is what keeps that from being a matter of which engine builds it.
  "pdffiledump", "pdffilesize", "pdffilemoddate", "pdfmdfivesum",
  "pdfshellescape", "pdfximage", "pdfrefximage", "pdfprimitive",
  "usepackage", "RequirePackage", "InputIfFileExists", "IfFileExists",
  "subfile", "import", "lstinputlisting", "verbatiminput",
  -- a name is built, aliased, or a character is read as another
  "csname", "expandafter", "scantokens", "string", "detokenize", "noexpand",
  "primitive", "meaning", "catcode", "endlinechar", "newread", "newwrite",
  "def", "edef", "gdef", "xdef", "let", "futurelet", "newcommand",
  "renewcommand", "providecommand", "DeclareRobustCommand", "makeatletter",
  "uppercase", "lowercase",
}) do
  DENIED[name] = true
end

local function names_a_denied_primitive(text)
  for name in text:gmatch("\\([A-Za-z]+)") do
    if DENIED[name] then
      return true
    end
  end
  return false
end

-- 2. `@` is a letter only while `\makeatletter` is in force, and that is how the
--    kernel keeps its own names, `\@@input` among them, out of a document. A
--    formula which carries the character is asking for that half of the kernel,
--    so it is dropped. The commutative diagrams of amsmath use `@` too and are
--    lost with it, which is the one thing this rule costs.
local function reaches_for_the_kernel(text)
  return text:find("@", 1, true) ~= nil
end

-- 3. TeX reads `^^` as the spelling of a character rather than as a superscript:
--    `^^5c` is a backslash, and the wider forms of the engine tectonic builds on
--    spell the same character as `^^^^005c`. A superscript of a superscript is
--    an error in TeX, so no formula which renders today is lost.
local function spells_a_character(text)
  return text:find("^^", 1, true) ~= nil
end

local function is_safe_math(text)
  return not (names_a_denied_primitive(text) or reaches_for_the_kernel(text) or spells_a_character(text))
end

function RawBlock(element)
  if is_tex(element.format) then
    return {}
  end
end

function RawInline(element)
  if is_tex(element.format) then
    return {}
  end
end

function Math(element)
  if not is_safe_math(element.text) then
    return {}
  end
end
