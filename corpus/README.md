# DOCX corpus

The DOCX documents of the docx-exporter system tests, with the PDFs Word prints from them. The
conversion of DOCX to PDF aims at Word's PDF. `tools/corpus/corpus.py` converts the documents with
pandoc-service and measures how far each PDF is from Word's, line by line.

| Path | What |
|---|---|
| `docx/` | the 49 documents, as Polarion's docx-exporter writes them |
| `reference/macos/` | the PDFs of Word for Mac, see below |
| `reference/windows/` | the PDFs of Word for Windows, still to make |
| `baseline.json` | the last accepted score, per reference and per run |

## The references

| Reference | Word | System | Made with |
|---|---|---|---|
| `macos` | Word for Mac 16.114.1 (26102031) | macOS 26.6.2 | `tools/corpus/word-macos/export.sh` |
| `windows` | | | `tools/corpus/word-windows/export.ps1` |

Word asks to update the fields of a document of pandoc-service when it opens it. Both scripts update
them, and the whole table of contents, before the export.

Word for Mac runs in a sandbox: it writes only to a folder it may write to. Export into a folder
under `~/Library/Containers/com.microsoft.Word/Data/`, then copy the PDFs here.

To make the Windows references:

1. Copy `docx/` and `tools/corpus/word-windows/export.ps1` to a Windows machine with Word.
2. Run `powershell -ExecutionPolicy Bypass -File export.ps1 -Source docx -Target windows`.
3. Copy the PDFs to `reference/windows/` and fill in the row of the table above.
4. Store the score: `python -m tools.corpus.corpus score --reference windows --variant image --out ... --write-baseline`.

## The two runs

The documents use Aptos, Word's default font, which the image does not hold: it sets them in Metric Sans,
as wide and as high as Aptos. Each reference therefore has two runs in `baseline.json`:

| Run | Fonts | Where |
|---|---|---|
| `word-fonts` | Word's own fonts, from a local folder (`--fonts`) | locally only: the fonts may not be distributed |
| `image` | the fonts of the image | locally and in CI |

`word-fonts` shows how close the layout gets to Word. `image` shows what a user of the image gets.

## The numbers

A line is a run of text on one baseline, without a gap wider than 1.5 times the font size: the
cells of a table row are lines of their own. The same rule reads both PDFs.

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
```

`all` converts the documents, scores them and builds the comparison page in `out/corpus/<run>/site/`.
Open `index.html` there. Each command runs alone as well: `convert`, `score`, `page`.

| Option of `score` | Effect |
|---|---|
| `--against` | fails when a document got worse than `baseline.json`: fewer lines on the same page, another page count than Word where it had Word's, a mean shift 0.5 points larger, or 2 percent fewer words in place or regions matched |
| `--write-baseline` | stores the score in `baseline.json`, for the run `--variant` |
| `--reference windows` | compares with the PDFs of Word for Windows |

A change which improves the score stores the new baseline in the same commit.

The workflow `corpus.yml` runs `image` on every push to `main` and publishes the page to GitHub
Pages. On a pull request it reports the score and what got worse in the summary of the run.
