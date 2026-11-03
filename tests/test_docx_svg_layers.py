import hashlib
import io
import zipfile

from lxml import etree

from app import docx_svg_layers

PNG = b"\x89PNG one"
OTHER = b"\x89PNG other"
SVG = b'<svg xmlns="http://www.w3.org/2000/svg"/>'
LAYERS = {hashlib.sha256(PNG).digest(): SVG}

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/></Types>'
)
IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def blip(embed, extra=""):
    return f'<w:r><w:drawing><a:blip r:embed="{embed}">{extra}</a:blip></w:drawing></w:r>'


def docx(blips, relationships, media=None, content_types=CONTENT_TYPES, extra=None):
    document = f'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}" xmlns:a="{A}" xmlns:r="{R}"><w:body>{"".join(blips)}</w:body></w:document>'
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="{rid}" Type="{kind}" Target="{target}"{attributes}/>' for rid, kind, target, attributes in relationships)
        + "</Relationships>"
    )
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as package:
        package.writestr("[Content_Types].xml", content_types)
        package.writestr("word/document.xml", document)
        package.writestr("word/_rels/document.xml.rels", rels)
        for name, data in (media or {"word/media/image1.png": PNG}).items():
            package.writestr(name, data)
        for name, data in (extra or {}).items():
            package.writestr(name, data)
    return out.getvalue()


def read(data):
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return {name: package.read(name) for name in package.namelist()}


def svg_blips(document):
    return etree.fromstring(document).findall(f".//{{{ASVG}}}svgBlip")


def test_attach_adds_the_svg_of_a_known_png():
    out = read(docx_svg_layers.attach(docx([blip("rId1")], [("rId1", IMAGE, "media/image1.png", "")]), LAYERS))
    assert out["word/media/image1.svg"] == SVG
    rels = etree.fromstring(out["word/_rels/document.xml.rels"])
    svg_rel = [rel for rel in rels if rel.get("Target") == "media/image1.svg"]
    assert len(svg_rel) == 1
    assert svg_rel[0].get("Type") == IMAGE
    blips = svg_blips(out["word/document.xml"])
    assert [element.get(f"{{{R}}}embed") for element in blips] == [svg_rel[0].get("Id")]
    assert b'Extension="svg"' in out["[Content_Types].xml"]


def test_attach_shares_one_svg_part_between_pictures_of_one_png():
    package = docx([blip("rId1"), blip("rId2")], [("rId1", IMAGE, "media/image1.png", ""), ("rId2", IMAGE, "media/image2.png", "")], {"word/media/image1.png": PNG, "word/media/image2.png": PNG})
    out = read(docx_svg_layers.attach(package, LAYERS))
    assert [name for name in out if name.endswith(".svg")] == ["word/media/image1.svg"]
    assert len(svg_blips(out["word/document.xml"])) == 2


def test_attach_names_the_svg_apart_from_a_part_that_exists():
    package = docx([blip("rId1")], [("rId1", IMAGE, "media/image1.png", ""), ("rIdSvg1", IMAGE, "media/other.png", "")], extra={"word/media/image1.svg": b"old", "word/media/other.png": OTHER})
    out = read(docx_svg_layers.attach(package, LAYERS))
    assert out["word/media/image1.svg"] == b"old"
    assert out["word/media/image1-1.svg"] == SVG
    assert svg_blips(out["word/document.xml"])[0].get(f"{{{R}}}embed") == "rIdSvg2"


def test_attach_leaves_a_docx_without_a_known_png():
    package = docx([blip("rId1"), blip("rId2")], [("rId1", IMAGE, "media/image1.png", ""), ("rId2", IMAGE, "https://example.com/a.png", ' TargetMode="External"')], {"word/media/image1.png": OTHER})
    assert docx_svg_layers.attach(package, LAYERS) == package


def test_attach_skips_a_missing_png_and_a_part_without_rels():
    package = docx([blip("rId1")], [("rId1", IMAGE, "media/missing.png", "")], extra={"word/_rels/footer1.xml.rels": b"<Relationships/>"})
    assert docx_svg_layers.attach(package, LAYERS) == package


def test_attach_keeps_a_blip_that_has_its_svg():
    existing = f'<a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip xmlns:asvg="{ASVG}" r:embed="rId9"/></a:ext></a:extLst>'
    package = docx([blip("rId1", existing)], [("rId1", IMAGE, "media/image1.png", "")])
    assert docx_svg_layers.attach(package, LAYERS) == package


def test_attach_adds_to_an_extension_list_and_keeps_an_svg_content_type():
    content_types = CONTENT_TYPES.replace("</Types>", '<Default Extension="SVG" ContentType="image/svg+xml"/></Types>')
    package = docx([blip("rId1", '<a:extLst><a:ext uri="{other}"/></a:extLst>')], [("rId1", IMAGE, "media/image1.png", "")], content_types=content_types)
    out = read(docx_svg_layers.attach(package, LAYERS))
    extensions = etree.fromstring(out["word/document.xml"]).findall(f".//{{{A}}}extLst/{{{A}}}ext")
    assert [extension.get("uri") for extension in extensions] == ["{other}", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"]
    assert out["[Content_Types].xml"] == content_types.encode()


def test_attach_without_layers_or_with_a_broken_docx():
    assert docx_svg_layers.attach(b"anything", {}) == b"anything"
    assert docx_svg_layers.attach(b"not a zip", LAYERS) == b"not a zip"
    broken = docx([blip("rId1")], [("rId1", IMAGE, "media/image1.png", "")], content_types="<Types")
    assert docx_svg_layers.attach(broken, LAYERS) == broken
