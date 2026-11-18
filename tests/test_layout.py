import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from format_research_word.document import format_docx
from format_research_word.layout import CT, FONT, M, ScriptStyles, W, normalize_story
from lxml import etree as ET

TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "report-styles.docx"
NS = {"w": W[1:-1], "m": M[1:-1]}


def xml(snippet):
    return ET.fromstring(
        (
            '<root xmlns:w="'
            + NS["w"]
            + '" xmlns:m="'
            + NS["m"]
            + '">'
            + snippet
            + "</root>"
        ).encode()
    )


class Layout(unittest.TestCase):
    def test_nested_and_merged_tables(self):
        tree = xml("""<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="2400"/><w:tblLayout w:type="fixed"/></w:tblPr>
            <w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tblPrEx><w:tblLayout w:type="fixed"/></w:tblPrEx>
            <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>
            <w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>""")
        n, _ = normalize_story(tree, ScriptStyles(None))
        self.assertEqual(n, 2)
        for t in tree.iter(W + "tbl"):
            pr = t.find(W + "tblPr")
            self.assertEqual(pr.find(W + "tblW").get(W + "w"), "5000")
            self.assertEqual(pr.find(W + "tblW").get(W + "type"), "pct")
            self.assertEqual(pr.find(W + "tblLayout").get(W + "type"), "autofit")
        self.assertEqual(tree.find(".//" + W + "tcW").get(W + "type"), "auto")
        self.assertEqual(tree.find(".//" + W + "gridSpan").get(W + "val"), "2")
        self.assertEqual(tree.find(".//" + W + "vMerge").get(W + "val"), "restart")
        self.assertEqual(
            tree.xpath(".//w:gridCol/@w:w", namespaces=NS), ["900", "1500"]
        )
        self.assertEqual(
            tree.xpath(".//w:tblPrEx/w:tblLayout/@w:type", namespaces=NS), ["autofit"]
        )

    def test_font_inheritance_and_unicode(self):
        styles = xml(
            '<w:style w:styleId="super"><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:styleId="ref"><w:basedOn w:val="super"/></w:style>'
        )
        root = xml("""<w:p><w:r><w:t>cm⁻¹、Fe³⁺、Å²、°C™</w:t><w:tab/><w:t>标签ᵃ</w:t></w:r>
          <w:hyperlink><w:r><w:rPr><w:rStyle w:val="ref"/><w:rFonts w:asciiTheme="majorAscii" w:ascii="Arial"/></w:rPr><w:t>[9]</w:t></w:r></w:hyperlink>
          <w:r><w:rPr><w:rStyle w:val="ref"/><w:vertAlign w:val="baseline"/></w:rPr><w:t>B</w:t></w:r>
          <w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:instrText> REF id² </w:instrText></w:r></w:p>""")
        normalize_story(root, ScriptStyles(styles))
        elevated = root.xpath(
            './/w:r[w:rPr/w:vertAlign[@w:val="superscript"]]', namespaces=NS
        )
        self.assertEqual(
            "".join(t.text or "" for r in elevated for t in r.findall(W + "t")),
            "−13+2a[9]",
        )
        self.assertEqual(len(elevated), 6)
        for r in elevated:
            font = r.find(W + "rPr/" + W + "rFonts")
            self.assertTrue(
                all(
                    font.get(W + k) == FONT
                    for k in ("ascii", "eastAsia", "hAnsi", "cs")
                )
            )
            self.assertFalse(any("theme" in k.lower() for k in font.attrib))
        self.assertIn("°C™", "".join(root.itertext()))
        self.assertEqual(len(list(root.iter(W + "tab"))), 1)
        self.assertEqual(root.find(".//" + W + "instrText").text, " REF id² ")
        self.assertEqual(
            root.xpath('.//w:r[w:t="B"]/w:rPr/w:vertAlign/@w:val', namespaces=NS),
            ["baseline"],
        )
        before = ET.tostring(root)
        normalize_story(root, ScriptStyles(styles))
        self.assertEqual(before, ET.tostring(root))

    def test_native_math_superscripts(self):
        root = xml("""<m:oMath><m:r><m:t>x+y²</m:t></m:r><m:sSubSup><m:e><m:r><m:t>E</m:t></m:r></m:e>
        <m:sub><m:r><m:t>g</m:t></m:r></m:sub><m:sup><m:r><m:t>3</m:t></m:r></m:sup></m:sSubSup></m:oMath>""")
        normalize_story(root, ScriptStyles(None))
        self.assertEqual(
            root.xpath(".//m:sSup/m:e/m:r/m:t/text()", namespaces=NS), ["y"]
        )
        self.assertEqual(
            root.xpath(".//m:sup/m:r/m:t/text()", namespaces=NS), ["2", "3"]
        )
        self.assertEqual(
            root.xpath(".//m:sup/m:r/w:rPr/w:rFonts/@w:ascii", namespaces=NS),
            [FONT, FONT],
        )
        self.assertEqual(root.xpath(".//m:sup//w:vertAlign", namespaces=NS), [])
        self.assertEqual(len(root.xpath(".//m:sup/m:r/m:rPr/m:nor", namespaces=NS)), 2)

    def test_split_math_exponent_and_raised_label(self):
        root = xml("""<m:oMath><m:r><m:t>10⁻</m:t></m:r><m:r><m:t>⁵</m:t></m:r></m:oMath>
        <w:p><w:r><w:rPr><w:position w:val="4"/></w:rPr><w:t>[4]</w:t></w:r></w:p>""")
        normalize_story(root, ScriptStyles(None))
        powers = root.findall(".//" + M + "sSup")
        self.assertEqual(len(powers), 1)
        self.assertEqual("".join(powers[0].find(M + "e").itertext()), "10")
        self.assertEqual("".join(powers[0].find(M + "sup").itertext()), "−5")
        props = root.find(".//" + W + "p/" + W + "r/" + W + "rPr")
        self.assertIsNone(props.find(W + "position"))
        self.assertEqual(props.find(W + "vertAlign").get(W + "val"), "superscript")
        self.assertEqual(props.find(W + "rFonts").get(W + "ascii"), FONT)

    def test_layout_only_in_notes_and_header(self):
        with tempfile.TemporaryDirectory() as d:
            source = Path(d) / "src.docx"
            out = Path(d) / "out.docx"
            with ZipFile(TEMPLATE) as a, ZipFile(source, "w") as z:
                types = ET.fromstring(a.read("[Content_Types].xml"))
                for kind, path, body in [
                    (
                        "footnotes",
                        "word/footnotes.xml",
                        '<w:footnote w:id="1"><w:p><w:r><w:t>0.123456 m²</w:t></w:r></w:p></w:footnote>',
                    ),
                    (
                        "header",
                        "word/headerX.xml",
                        "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
                    ),
                ]:
                    ET.SubElement(
                        types,
                        CT + "Override",
                        PartName="/" + path,
                        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml."
                        + kind
                        + "+xml",
                    )
                    name = "hdr" if kind == "header" else kind
                    z.writestr(
                        path,
                        (
                            "<w:"
                            + name
                            + ' xmlns:w="'
                            + NS["w"]
                            + '">'
                            + body
                            + "</w:"
                            + name
                            + ">"
                        ).encode(),
                    )
                for item in a.infolist():
                    z.writestr(
                        item,
                        ET.tostring(types)
                        if item.filename == "[Content_Types].xml"
                        else a.read(item),
                    )
            format_docx(source, out, layout_only=True)
            with ZipFile(out) as z:
                notes = ET.fromstring(z.read("word/footnotes.xml"))
                self.assertEqual("".join(notes.itertext()), "0.123456 m2")
                self.assertEqual(
                    notes.xpath(
                        './/w:r[w:t="2"]/w:rPr/w:rFonts/@w:ascii', namespaces=NS
                    ),
                    [FONT],
                )
                self.assertIn(b"autofit", z.read("word/headerX.xml"))
