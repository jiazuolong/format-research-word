import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from format_research_word import audit
from format_research_word.document import M, W, format_docx
from format_research_word.typography import Char, joined, typography
from lxml import etree as ET

TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "report-styles.docx"


def textformat(text):
    return joined(typography([Char(c, 0) for c in text]))


def words(p, *parts):
    for value in parts:
        r = ET.SubElement(p, W + "r")
        ET.SubElement(r, W + "t").text = value
    return p


def addp(body, *parts):
    return words(ET.SubElement(body, W + "p"), *parts)


class Rules(unittest.TestCase):
    def test_precision(self):
        for old, new in [
            ("−4.497105916", "−4.50"),
            ("9.99999", "10.0"),
            ("0.01870", "0.0187"),
            ("1.23984", "1.24"),
            ("10.9415", "10.9415"),
        ]:
            with self.subTest(old=old):
                self.assertEqual(textformat(old), new)

    def test_units_and_identifiers(self):
        self.assertEqual(
            textformat("约 10μ m 时，温度 300K，变化 63%。"),
            "约10 μm时，温度300 K，变化63 %。",
        )
        self.assertEqual(textformat("结果为0.123456eV。"), "结果为0.123 eV。")
        self.assertEqual(textformat("结构为Fm-3m。"), "结构为Fm-3m。")
        for literal in (
            "Materials Project",
            "https://doi.org/10.12345/1.123456",
            "v1.234567.2",
        ):
            self.assertEqual(textformat(literal), literal)


class WordFiles(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "source.docx"
        self.target = self.folder / "result.docx"

    def make(self, field=False):
        with ZipFile(TEMPLATE) as z, ZipFile(self.source, "w") as new:
            root = ET.fromstring(z.read("word/document.xml"))
            body = root.find(W + "body")
            section = body.find(W + "sectPr")
            for e in list(body):
                body.remove(e)
            addp(body, "测量 Pb₁₋ₓ", "GeₓCdAs₂ 时 x=0.497105，见[3,1]。")
            addp(body, "再次引用[3]，以及[1–2]。")
            t = ET.SubElement(body, W + "tbl")
            cell = ET.SubElement(ET.SubElement(t, W + "tr"), W + "tc")
            addp(cell, "300K，0.00001 eV")
            p = addp(body)
            math = ET.SubElement(p, M + "oMath")
            mr = ET.SubElement(math, M + "r")
            ET.SubElement(mr, M + "t").text = "x≤0.00001"
            if field:
                r = ET.SubElement(p, W + "r")
                ET.SubElement(
                    r, W + "instrText"
                ).text = " ADDIN ZOTERO_ITEM CSL_CITATION "
            addp(body, "参考文献")
            for i, title in enumerate(("Alpha", "Beta", "Gamma", "Delta"), 1):
                p = addp(body)
                r = ET.SubElement(p, W + "r")
                ET.SubElement(r, W + "lastRenderedPageBreak")
                ET.SubElement(r, W + "t").text = str(i)
                ET.SubElement(ET.SubElement(p, W + "r"), W + "tab")
                words(p, title)
            addp(body, "附录A 数据")
            addp(body, "独立引用[2]。")
            body.append(section)
            for info in z.infolist():
                new.writestr(
                    info,
                    ET.tostring(root)
                    if info.filename == "word/document.xml"
                    else z.read(info),
                )

    def test_format_and_renumber(self):
        self.make()
        before = self.source.read_bytes()
        stats = format_docx(self.source, self.target, True)
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(stats["重排引文处"], 4)
        report = audit.run(self.target)
        self.assertEqual(report["首次出现次序"], [1, 2, 3])
        self.assertEqual(report["未见静态引用的原号"], [4])
        with ZipFile(self.target) as out, ZipFile(self.source) as original:
            self.assertEqual(out.namelist(), original.namelist())
            for name in out.namelist():
                if name != "word/document.xml":
                    self.assertEqual(out.read(name), original.read(name))
            root = ET.fromstring(out.read("word/document.xml"))
            body = root.find(W + "body")
            text = "".join(body.itertext())
            self.assertIn("测量Pb1-xGexCdAs2时x=0.497，见[1,2]", text)
            self.assertIn("1Gamma2Alpha3Beta4Delta", text)
            self.assertIn("300 K，1.00×10−5 eV", text)
            ns = {"w": W[1:-1], "m": M[1:-1]}
            self.assertEqual(
                "".join(
                    root.xpath(
                        '//w:r[w:rPr/w:vertAlign[@w:val="superscript"]]/w:t/text()',
                        namespaces=ns,
                    )
                ),
                "−5",
            )
            subs = root.xpath(
                '//w:r[w:rPr/w:vertAlign[@w:val="subscript"]]', namespaces=ns
            )
            self.assertEqual("".join("".join(r.itertext()) for r in subs), "1-xx2")
            self.assertTrue(
                all(
                    r.find(W + "rPr/" + W + "rFonts").get(W + "ascii")
                    == "Times New Roman"
                    for r in subs
                )
            )
            self.assertTrue(
                root.xpath(
                    'boolean(//m:r[m:t="x"]/m:rPr/m:sty[@m:val="p"])', namespaces=ns
                )
            )
            self.assertTrue(
                root.xpath(
                    'boolean(//m:sSup[m:e/m:r/m:t="10" and m:sup/m:r/m:t="−5"])',
                    namespaces=ns,
                )
            )

    def test_existing_destination(self):
        self.make()
        self.target.write_bytes(b"existing")
        with self.assertRaises(FileExistsError):
            format_docx(self.source, self.target)
        self.assertEqual(self.target.read_bytes(), b"existing")

    def test_dynamic_citations(self):
        self.make(field=True)
        with self.assertRaisesRegex(ValueError, "人工判定"):
            format_docx(self.source, self.target, True)
        self.assertFalse(self.target.exists())

    def test_bad_range(self):
        self.assertIsNone(audit.groups("1–100000000", 100000000, {1, 100000000}))


if __name__ == "__main__":
    unittest.main()
