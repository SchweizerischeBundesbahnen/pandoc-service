"""The service in its image: weasysvg keeps a draw.io diagram vector in DOCX and PDF."""

import base64
import io
import zipfile

from lxml import etree
from pypdf import PdfReader

from tests.test_container import TestParameters

# What draw.io writes: a label in a foreignObject with a text fallback, and a shadow through a filter
DIAGRAM = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 200 100">'
    '<defs><filter id="shadow"><feGaussianBlur in="SourceAlpha" stdDeviation="1.7"/><feOffset dx="3" dy="3" result="blur"/>'
    '<feFlood flood-color="#3D4574" flood-opacity="0.4"/><feComposite in2="blur" operator="in"/><feMerge><feMergeNode/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>'
    '<g filter="url(#shadow)"><rect x="20" y="20" width="160" height="60" rx="9" fill="#dae8fc" stroke="#6c8ebf"/></g>'
    '<switch><foreignObject x="20" y="20" width="160" height="60" requiredFeatures="http://www.w3.org/TR/SVG11/feature#Extensibility">'
    '<div xmlns="http://www.w3.org/1999/xhtml" style="font-family:Helvetica;font-size:16px;text-align:center;padding-top:20px">Vector label</div></foreignObject>'
    '<text x="100" y="55" text-anchor="middle" font-size="16">Vector label</text></switch></svg>'
)
HTML = f'<html><body><p>Before</p><img src="data:image/svg+xml;base64,{base64.b64encode(DIAGRAM.encode()).decode()}"><p>After</p></body></html>'

ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def convert(parameters: TestParameters, target: str) -> bytes:
    response = parameters.request_session.post(f"{parameters.base_url}/convert/html/to/{target}", data=HTML.encode())
    assert response.status_code == 200
    return response.content


def test_version_names_weasysvg(test_parameters: TestParameters) -> None:
    version = test_parameters.request_session.get(f"{test_parameters.base_url}/version").json()
    assert version["svgRenderer"] == "weasysvg"
    assert version["weasysvg"]
    assert "chromium" not in version


def test_docx_gets_the_png_and_its_flat_svg(test_parameters: TestParameters) -> None:
    with zipfile.ZipFile(io.BytesIO(convert(test_parameters, "docx"))) as docx:
        document = etree.fromstring(docx.read("word/document.xml"))
        rels = {rel.get("Id"): rel.get("Target") for rel in etree.fromstring(docx.read("word/_rels/document.xml.rels"))}
        svg_blips = document.findall(f".//{{{ASVG}}}svgBlip")
        assert len(svg_blips) == 1
        svg_target = rels[svg_blips[0].get(f"{{{R}}}embed")]
        png_target = rels[svg_blips[0].getparent().getparent().getparent().get(f"{{{R}}}embed")]
        assert png_target.endswith(".png")
        assert docx.read(f"word/{png_target}").startswith(b"\x89PNG")
        flat = docx.read(f"word/{svg_target}").decode()
        assert b'Extension="svg"' in docx.read("[Content_Types].xml")
    # Only what Word draws: no filter, no mask, no foreignObject
    assert "<svg" in flat
    assert "filter=" not in flat
    assert "<mask" not in flat
    assert "foreignObject" not in flat


def test_pdf_gets_the_svg_as_vector(test_parameters: TestParameters) -> None:
    reader = PdfReader(io.BytesIO(convert(test_parameters, "pdf")))
    page = reader.pages[0]
    # The label is text and the page holds no picture: rsvg-convert, which is weasysvg-convert, drew the SVG
    assert "Vector label" in page.extract_text()
    assert not page.images
