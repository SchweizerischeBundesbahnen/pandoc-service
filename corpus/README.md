# DOCX corpus

The DOCX documents of the docx-exporter system tests, with the PDFs Word prints from them. The
conversion of DOCX to PDF aims at Word's PDF. `tools/corpus/corpus.py` converts the documents with
pandoc-service and measures how far each PDF is from Word's, line by line.

| Path | What |
|---|---|
| `docx/` | the 49 documents, as Polarion's docx-exporter writes them |
| `docx-free/` | the same documents in open fonts of the image, see below |
| `reference/macos/` | the PDFs of Word for Mac, see below |
| `reference/macos-free/` | the PDFs of Word for Mac of `docx-free/` |
| `reference/windows/` | the PDFs of Word for Windows, see below |
| `baseline.json` | the last accepted score, per reference and per run |

## The references

| Reference | Word | System | Made with |
|---|---|---|---|
| `macos` | Word for Mac 16.114.1 (26102031) | macOS 26.6.2 | `tools/corpus/word-macos/export.sh` |
| `macos-free` | Word for Mac 16.114.1 (26102031) | macOS 26.6.2 | `tools/corpus/word-macos/export.sh` |
| `windows` | Word LTSC 2021 16.0.14334.20974 | Windows 11 Pro 25H2 (26200) | `tools/corpus/word-windows/export.ps1` |

Word asks to update the fields of a document of pandoc-service when it opens it. Both scripts update
them, and the whole table of contents, before the export.

Word for Mac runs in a sandbox: it writes only to a folder it may write to. Export into a folder
under `~/Library/Containers/com.microsoft.Word/Data/`, then copy the PDFs here.

To make the Windows references:

1. Copy `docx/` and `tools/corpus/word-windows/export.ps1` to a Windows machine with Word.
2. Without Microsoft 365, install Microsoft's Aptos fonts for all users. Otherwise Word sets the documents in Calibri.
3. Run `powershell -ExecutionPolicy Bypass -File export.ps1 -Source docx -Target windows`.
4. Copy the PDFs to `reference/windows/` and fill in the row of the table above.
5. Store the score: `python -m tools.corpus.corpus score --reference windows --variant image --out ... --write-baseline`.

The two Words agree on 96 percent of the words. The four documents with comments differ most: their lines
stand about 2 points apart. The page `mac-vs-windows/` shows each difference, see below.

## The two runs

The documents use Aptos, Word's default font, which the image does not hold: it sets them in Metric Sans,
as wide and as high as Aptos. Each reference therefore has two runs in `baseline.json`:

| Run | Fonts | Where |
|---|---|---|
| `word-fonts` | Word's own fonts, from a local folder (`--fonts`) | locally only: the fonts may not be distributed |
| `image` | the fonts of the image | locally and in CI |

`word-fonts` shows how close the layout gets to Word. `image` shows what a user of the image gets.

## The free-font corpus

`docx-free/` holds the same documents with every font of Word replaced by an open font of the image,
and the equations in Latin Modern Math. Word and the image then set them in the same fonts, so
`macos-free` measures the layout alone, in CI as well. Its run is `image`.

1. Make the documents: `uv run python -m tools.corpus.free_fonts corpus/docx corpus/docx-free`.
2. Install the fonts of the image on the Mac, the same files: Liberation, DejaVu, Carlito and `latinmodern-math.otf`.
3. Export them with Word as above, into `reference/macos-free/`.

`tools/corpus/free_fonts.py` lists the open font of each font of Word. The script cannot export
`test_toc_tof_tot_without_gaps_conversion`: it does not see the document again after Word updates its fields.
Open it in Word by hand, answer Yes, then save it as PDF.

## The numbers

A line is a run of text on one baseline, without a gap wider than 0.8 times the font size: the
cells of a table row are lines of their own. The same rule reads both PDFs. A line of an equation
does not count: its letters have no characters in either PDF, so its text matches other equations.

| Number | Meaning |
|---|---|
| Pages | the pages of Word's PDF and of the conversion |
| Same page | Word's lines which the conversion sets on the same page, with the same text |
| Any page | Word's lines which the conversion sets on any page |
| Mean shift | the mean vertical distance of the lines on the same page, in points |
| Words in place | Word's words which the conversion sets on the same page within 2 points across and down |
| Regions matched | the squares of 14 points with ink in either PDF where the ink of each lies within 1.5 points of ink of the other: pictures, borders and shading count as well as text |

The page shows each page of a document side by side, as an overlay, or as its difference: Word in
gray, the conversion in red, and each square which does not match framed in orange.

A line further off than 2 points, or missing on its page, is a problem. The page lists them under the pages of the document.

## Run it

```bash
docker build --tag pandoc-service:corpus .
# As in CI: the fonts of the image
uv run --group test python -m tools.corpus.corpus all --image pandoc-service:corpus --out out/corpus/image
# With Word's fonts, from a folder of your own
uv run --group test python -m tools.corpus.corpus all --image pandoc-service:corpus --fonts ~/word-fonts \
  --variant word-fonts --out out/corpus/word-fonts
# The free-font corpus
uv run --group test python -m tools.corpus.corpus all --image pandoc-service:corpus --docx corpus/docx-free \
  --reference macos-free --out out/corpus/free
```

`all` converts the documents, scores them and builds the comparison page in `out/corpus/<run>/site/`.
Open `index.html` there. Each command runs alone as well: `convert`, `score`, `page`.

`--pdf` scores the PDFs of another folder. The PDFs of a run compare with Word for Windows without a new
conversion, and the PDFs of Word for Mac with those of Word for Windows:

```bash
uv run --group test python -m tools.corpus.corpus score --out out/corpus/windows --pdf out/corpus/image/pdf \
  --reference windows --variant image --against
uv run --group test python -m tools.corpus.corpus page --out out/corpus/mac-vs-windows --pdf corpus/reference/macos \
  --reference windows --site mac-vs-windows
```

| Option of `score` | Effect |
|---|---|
| `--against` | fails when a document got worse than `baseline.json`: fewer lines on the same page, another page count than Word where it had Word's, a mean shift 0.5 points larger, or 2 percent fewer words in place or regions matched |
| `--write-baseline` | stores the score in `baseline.json`, for the run `--variant` |
| `--reference windows` | compares with the PDFs of Word for Windows |
| `--pdf DIR` | scores the PDFs of DIR in place of `OUT/pdf` |

| Option of `page` | Effect |
|---|---|
| `--pdf DIR` | shows the PDFs of DIR in place of `OUT/pdf` |
| `--site NAME` | the page of the site it builds, the reference by default: see below |

A change which improves the score stores the new baseline in the same commit.

The workflow `corpus.yml` runs `image` of both corpora on every push to `main`. It scores the PDFs of
`docx/` against Word for Windows as well, and publishes the pages to GitHub Pages. Each page links the
others:

| Folder | Page |
|---|---|
| the root | pandoc-service against Word for Mac |
| `windows/` | pandoc-service against Word for Windows |
| `free/` | the free-font corpus against Word for Mac |
| `mac-vs-windows/` | Word for Mac against Word for Windows, without a baseline |

On a pull request it reports the score and what got worse in the summary of the run.
