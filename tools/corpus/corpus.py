"""Convert the DOCX corpus with pandoc-service and compare the PDFs with the ones Word prints.

    python -m tools.corpus.corpus convert --image IMAGE [--fonts DIR] [--docx DIR] --out DIR
    python -m tools.corpus.corpus score --out DIR [--reference macos] [--variant NAME] [--write-baseline | --against]
    python -m tools.corpus.corpus page --out DIR [--reference macos]
    python -m tools.corpus.corpus all --image IMAGE [--fonts DIR] --out DIR [score options]

See corpus/README.md for what the numbers mean.
"""

from __future__ import annotations

import argparse
import json
import statistics
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict
from pathlib import Path

import httpx
import pypdfium2 as pdfium

from tools.corpus import regions
from tools.corpus.pdf_lines import Line, Word, document_lines, document_words

ROOT = Path(__file__).resolve().parents[2]
CORPUS = ROOT / "corpus"
DOCX = CORPUS / "docx"
REFERENCES = CORPUS / "reference"
BASELINE = CORPUS / "baseline.json"
PAGE_TEMPLATE = Path(__file__).parent / "page.html"
PORT = 9082
# Where the fonts of --fonts appear in the container: fontconfig reads every directory under /usr/share/fonts
FONT_MOUNT = "/usr/share/fonts/truetype/corpus"
# A line further off than this, in points, counts as shifted on the page; a word too
SHIFTED = 2.0
# Worse than the baseline by more than this many percent of the words or the regions of a document fails --against
SHARE_TOLERANCE = 2.0
# The rendering of the pages on the comparison page, in pixels per point
SCALE = 1.5
WORKERS = 4
# Worse than the baseline by more than this fails --against
TOLERANCE = 0.5


def _documents(docx: Path = DOCX) -> list[Path]:
    return sorted(docx.glob("*.docx"))


def _run(*command: str) -> str:
    return subprocess.run(command, check=True, capture_output=True, text=True).stdout.strip()  # noqa: S603 - docker with arguments of this script


def convert_one(client: httpx.Client, base: str, docx: Path, target: Path) -> str | None:
    """Convert one DOCX to its PDF in TARGET; why it failed, or None.

    The PDF of an earlier run goes first: the score would read it and miss that the conversion now fails.
    """
    pdf = target / f"{docx.stem}.pdf"
    pdf.unlink(missing_ok=True)
    with docx.open("rb") as source:
        response = client.post(f"{base}/convert/docx/to/pdf", files={"source": (docx.name, source)})
    if response.status_code != 200 or not response.content.startswith(b"%PDF"):  # noqa: PLR2004 - HTTP OK
        return f"conversion failed: {response.text[:200]}"
    pdf.write_bytes(response.content)
    return None


def convert(image: str, fonts: Path | None, out: Path, docx: Path = DOCX) -> list[str]:
    """Convert every DOCX of the folder DOCX to a PDF in OUT/pdf with a container of IMAGE; the documents which failed."""
    target = out / "pdf"
    target.mkdir(parents=True, exist_ok=True)
    mounts = ["--volume", f"{fonts.resolve()}:{FONT_MOUNT}:ro"] if fonts else []
    container = _run("docker", "run", "--detach", "--rm", "--publish", f"127.0.0.1::{PORT}", *mounts, image)
    try:
        address = _run("docker", "port", container, str(PORT)).splitlines()[0]
        base = f"http://{address}"
        with httpx.Client(timeout=300) as client:
            for _ in range(120):
                try:
                    if client.get(f"{base}/version").status_code == 200:  # noqa: PLR2004 - HTTP OK
                        break
                except httpx.TransportError:
                    # The service is still starting: try again until the time is up
                    pass
                time.sleep(1)

            failed = []
            documents = _documents(docx)
            with ThreadPoolExecutor(WORKERS) as pool:
                for document, error in zip(documents, pool.map(lambda document: convert_one(client, base, document, target), documents), strict=True):
                    print(f"{document.stem}: {error or 'ok'}")  # noqa: T201 - a command line tool
                    if error:
                        failed.append(document.stem)
            return failed
    finally:
        subprocess.run(["docker", "stop", container], check=False, capture_output=True)  # noqa: S603, S607 - docker with arguments of this script


def _match(reference: list[list[Line]], converted: list[list[Line]]) -> tuple[list[dict], int]:
    """Each line of the reference with where the conversion sets it, and how many reference lines the conversion has anywhere."""
    anywhere = {line.text for page in converted for line in page}
    found = 0
    matches = []
    for number, lines in enumerate(reference):
        candidates: dict[str, list[Line]] = {}
        for line in converted[number] if number < len(converted) else []:
            candidates.setdefault(line.text, []).append(line)
        for line in lines:
            found += line.text in anywhere
            same = candidates.get(line.text, [])
            if same:
                nearest = min(same, key=lambda other, line=line: abs(other.y - line.y))
                same.remove(nearest)
                matches.append({"line": asdict(line), "dy": round(nearest.y - line.y, 2), "dx": round(nearest.x - line.x, 2)})
            else:
                matches.append({"line": asdict(line), "dy": None, "dx": None})
    return matches, found


def _words_in_place(reference: list[list[Word]], converted: list[list[Word]]) -> int:
    """How many words of the reference the conversion sets on the same page, within SHIFTED points across and down."""
    found = 0
    for number, words in enumerate(reference):
        candidates: dict[str, list[Word]] = {}
        for word in converted[number] if number < len(converted) else []:
            candidates.setdefault(word.text, []).append(word)
        for word in words:
            same = candidates.get(word.text, [])
            if not same:
                continue
            nearest = min(same, key=lambda other, word=word: abs(other.x - word.x) + abs(other.y - word.y))
            if abs(nearest.x - word.x) <= SHIFTED and abs(nearest.y - word.y) <= SHIFTED:
                same.remove(nearest)
                found += 1
    return found


def score_document(reference_pdf: Path, converted_pdf: Path) -> dict:
    """How the conversion of one document compares with its reference: its lines, its words and the regions of its pages."""
    exists = converted_pdf.exists()
    reference = document_lines(reference_pdf)
    converted = document_lines(converted_pdf) if exists else []
    matches, anywhere = _match(reference, converted)
    shifts = [abs(match["dy"]) for match in matches if match["dy"] is not None]
    reference_words = document_words(reference_pdf)
    pages = regions.compare_documents(reference_pdf, converted_pdf if exists else None)
    return {
        "pages": [len(reference), len(converted)],
        "lines": len(matches),
        "same_page": len(shifts),
        "any_page": anywhere,
        "mean_shift": round(statistics.mean(shifts), 2) if shifts else None,
        "words": sum(len(words) for words in reference_words),
        "words_in_place": _words_in_place(reference_words, document_words(converted_pdf) if exists else []),
        "regions": sum(page.inked for page in pages),
        "regions_matched": sum(page.matched for page in pages),
        "problems": [match for match in matches if match["dy"] is None or abs(match["dy"]) > SHIFTED],
    }


def _total(documents: dict[str, dict]) -> dict:
    lines = sum(result["lines"] for result in documents.values())
    shifts = [result["mean_shift"] for result in documents.values() if result["mean_shift"] is not None]
    return {
        "documents": len(documents),
        "same_page": round(100 * sum(result["same_page"] for result in documents.values()) / max(lines, 1), 1),
        "any_page": round(100 * sum(result["any_page"] for result in documents.values()) / max(lines, 1), 1),
        "pages_as_reference": sum(result["pages"][0] == result["pages"][1] for result in documents.values()),
        "mean_shift": round(statistics.mean(shifts), 2) if shifts else None,
        "words_in_place": round(100 * sum(result["words_in_place"] for result in documents.values()) / max(sum(result["words"] for result in documents.values()), 1), 1),
        "regions_matched": round(100 * sum(result["regions_matched"] for result in documents.values()) / max(sum(result["regions"] for result in documents.values()), 1), 1),
    }


def score(out: Path, reference: str) -> dict:
    """Score every converted PDF of OUT against the reference PDFs; writes OUT/score.json."""
    references = REFERENCES / reference
    documents = {pdf.stem: score_document(pdf, out / "pdf" / pdf.name) for pdf in sorted(references.glob("*.pdf"))}
    result = {"reference": reference, "total": _total(documents), "documents": documents}
    (out / "score.json").write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n")
    return result


def _summary(documents: dict[str, dict]) -> dict[str, dict]:
    keys = ("pages", "lines", "same_page", "any_page", "mean_shift", "words", "words_in_place", "regions", "regions_matched")
    return {name: {key: result[key] for key in keys if key in result} for name, result in documents.items()}


def report(result: dict) -> str:
    """The score as a Markdown table, the documents furthest off first."""
    total = result["total"]
    rows = [
        (
            f"Reference: {result['reference']}. Lines on the same page {total['same_page']}%, on any page {total['any_page']}%, pages as the reference {total['pages_as_reference']}/{total['documents']}, mean shift {total['mean_shift']} pt."
            f" Words in place {total.get('words_in_place')}%, regions matched {total.get('regions_matched')}%."
        ),
        "",
        "| Document | Pages | Same page | Any page | Mean shift | Words in place | Regions matched |",
        "|---|---|---|---|---|---|---|",
    ]
    ordered = sorted(result["documents"].items(), key=lambda item: -(item[1]["mean_shift"] or 1000))
    for name, document in ordered:
        lines = max(document["lines"], 1)
        rows.append(
            f"| {name} | {document['pages'][0]} / {document['pages'][1]} | {100 * document['same_page'] / lines:.1f}% | {100 * document['any_page'] / lines:.1f}% | {document['mean_shift']} "
            f"| {_share(document, 'words_in_place', 'words'):.1f}% | {_share(document, 'regions_matched', 'regions'):.1f}% |"
        )
    return "\n".join(rows) + "\n"


def _share(document: dict, part: str, whole: str) -> float:
    return 100 * document.get(part, 0) / max(document.get(whole, 0), 1)


def write_baseline(result: dict, variant: str) -> None:
    """Store the score of VARIANT against its reference in corpus/baseline.json."""
    baseline = json.loads(BASELINE.read_text()) if BASELINE.exists() else {}
    baseline.setdefault(result["reference"], {})[variant] = {"total": result["total"], "documents": _summary(result["documents"])}
    BASELINE.write_text(json.dumps(baseline, indent=1, ensure_ascii=False, sort_keys=True) + "\n")


def regressions(result: dict, variant: str) -> list[str]:
    """What got worse than corpus/baseline.json, document by document."""
    stored = json.loads(BASELINE.read_text()).get(result["reference"], {}).get(variant) if BASELINE.exists() else None
    if stored is None:
        return [f"no baseline for {result['reference']} {variant}"]
    found = []
    for name, document in result["documents"].items():
        before = stored["documents"].get(name)
        if before is None:
            continue
        if before["pages"][0] == before["pages"][1] and document["pages"][0] != document["pages"][1]:
            found.append(f"{name}: {document['pages'][1]} pages, the reference has {document['pages'][0]}")
        if document["same_page"] < before["same_page"]:
            found.append(f"{name}: {document['same_page']} lines on the same page, was {before['same_page']}")
        if document["mean_shift"] is not None and before["mean_shift"] is not None and document["mean_shift"] > before["mean_shift"] + TOLERANCE:
            found.append(f"{name}: mean shift {document['mean_shift']} pt, was {before['mean_shift']}")
        for part, whole, what in (("words_in_place", "words", "words in place"), ("regions_matched", "regions", "regions matched")):
            if whole in before and _share(document, part, whole) < _share(before, part, whole) - SHARE_TOLERANCE:
                found.append(f"{name}: {_share(document, part, whole):.1f}% {what}, was {_share(before, part, whole):.1f}%")
    return found


def _render(pdf: Path, target: Path, prefix: str) -> tuple[list[str], list[list[float]]]:
    """Render each page of PDF to TARGET/PREFIX-N.webp; the names and the sizes in points, in the order of the pages."""
    target.mkdir(parents=True, exist_ok=True)
    names, sizes = [], []
    document = pdfium.PdfDocument(pdf)
    try:
        for number in range(len(document)):
            name = f"{prefix}-{number + 1}.webp"
            page = document[number]
            page.render(scale=SCALE).to_pil().save(target / name, quality=80)
            names.append(name)
            sizes.append([round(page.get_width(), 2), round(page.get_height(), 2)])
    finally:
        document.close()
    return names, sizes


def page(out: Path, reference: str, docx: Path = DOCX) -> Path:
    """Build OUT/site: the comparison page with the pages of the reference and of the conversion."""
    result = json.loads((out / "score.json").read_text())
    site = out / "site"
    baseline = json.loads(BASELINE.read_text()).get(reference, {}) if BASELINE.exists() else {}
    documents = []
    for name, document in result["documents"].items():
        converted = out / "pdf" / f"{name}.pdf"
        reference_pages, sizes = _render(REFERENCES / reference / f"{name}.pdf", site / "reference", name)
        converted_pages = _render(converted, site / "converted", name)[0] if converted.exists() else []
        regions.compare_documents(REFERENCES / reference / f"{name}.pdf", converted if converted.exists() else None, site / "difference", name)
        difference_pages = [f"{name}-{number + 1}.webp" for number in range(max(len(reference_pages), len(converted_pages)))]
        documents.append({"name": name, **document, "reference_pages": reference_pages, "converted_pages": converted_pages, "difference_pages": difference_pages, "sizes": sizes})
    data = {"reference": reference, "total": result["total"], "baseline": {variant: stored["total"] for variant, stored in baseline.items()}, "scale": SCALE, "shifted": SHIFTED, "documents": documents}
    template = PAGE_TEMPLATE.read_text()
    # The data sits in a script element, which only "</" can end
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    (site / "index.html").write_text(template.replace("__CORPUS_DATA__", payload))
    for document in _documents(docx):
        (site / "docx").mkdir(exist_ok=True)
        (site / "docx" / document.name).write_bytes(document.read_bytes())
    return site / "index.html"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=("convert", "score", "page", "all"))
    parser.add_argument("--out", type=Path, required=True, help="the directory of the PDFs, the score and the page")
    parser.add_argument("--image", help="the image of pandoc-service to convert with")
    parser.add_argument("--fonts", type=Path, help="a directory of fonts the container gets, Word's own for the variant word-fonts")
    parser.add_argument("--docx", type=Path, default=DOCX, help="the folder of the documents: corpus/docx, or corpus/docx-free with --reference macos-free")
    parser.add_argument("--reference", default="macos", help="the reference to compare with: a directory of corpus/reference")
    parser.add_argument("--variant", default="image", help="the name of the run in corpus/baseline.json: image or word-fonts")
    parser.add_argument("--write-baseline", action="store_true", help="store the score in corpus/baseline.json")
    parser.add_argument("--against", action="store_true", help="fail when a document got worse than corpus/baseline.json")
    arguments = parser.parse_args()

    if arguments.command in ("convert", "all"):
        if not arguments.image:
            parser.error("convert needs --image")
        failed = convert(arguments.image, arguments.fonts, arguments.out, arguments.docx)
    else:
        failed = []
    # A conversion which failed fails the run, apart from what got worse
    status = 1 if failed else 0
    if arguments.command in ("score", "all"):
        result = score(arguments.out, arguments.reference)
        text = report(result)
        if arguments.against:
            found = regressions(result, arguments.variant)
            text += "".join(f"\n- worse: {message}" for message in found) + ("\n" if found else "")
            status = 1 if found else status
        text += "".join(f"\n- conversion failed: {name}" for name in failed) + ("\n" if failed else "")
        (arguments.out / "report.md").write_text(text)
        print(text)  # noqa: T201 - a command line tool
        if arguments.write_baseline:
            write_baseline(result, arguments.variant)
    if arguments.command in ("page", "all"):
        print(f"page: {page(arguments.out, arguments.reference, arguments.docx)}")  # noqa: T201 - a command line tool
    return status


if __name__ == "__main__":
    sys.exit(main())
