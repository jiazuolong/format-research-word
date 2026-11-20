"""Run the Word integration check on Windows with FORMAT_WORD_TESTS=1."""

import os
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET

from format_research_word.audit import W
from format_research_word.toc import generate_toc

TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "report-styles.docx"


class TemplateContents(unittest.TestCase):
    def test_three_entry_styles_use_the_same_format(self):
        with ZipFile(TEMPLATE) as doc:
            styles = ET.fromstring(doc.read("word/styles.xml"))
        for level in (1, 2, 3):
            style = styles.find(f"{W}style[@{W}styleId='TOC{level}']")
            font = style.find(f"{W}rPr/{W}rFonts")
            self.assertEqual(font.get(W + "eastAsia"), "宋体")
            self.assertEqual(font.get(W + "ascii"), "Times New Roman")
            self.assertEqual(style.find(f"{W}rPr/{W}sz").get(W + "val"), "24")
            self.assertEqual(style.find(f"{W}rPr/{W}color").get(W + "val"), "467886")
            self.assertEqual(style.find(f"{W}rPr/{W}u").get(W + "val"), "single")
            self.assertEqual(style.find(f"{W}pPr/{W}ind").get(W + "leftChars"), "200")
            self.assertEqual(style.find(f"{W}pPr/{W}jc").get(W + "val"), "left")
            self.assertEqual(style.find(f"{W}pPr/{W}spacing").get(W + "line"), "278")


@unittest.skipUnless(
    os.environ.get("FORMAT_WORD_TESTS") == "1", "requires desktop Word"
)
class Contents(unittest.TestCase):
    def test_generate_then_refresh_changed_heading(self):
        with tempfile.TemporaryDirectory() as folder:
            first, revised, final = [
                Path(folder) / f"{n}.docx" for n in ("first", "revised", "final")
            ]
            self.assertEqual(generate_toc(TEMPLATE, first), 3)
            with ZipFile(first) as src, ZipFile(revised, "w") as dst:
                for item in src.infolist():
                    data = src.read(item)
                    if item.filename == "word/document.xml":
                        root = ET.fromstring(data)
                        heading = next(
                            p
                            for p in root.iter(W + "p")
                            if p.find(W + "pPr/" + W + "pStyle") is not None
                            and p.find(W + "pPr/" + W + "pStyle").get(W + "val")
                            == "xx3"
                        )
                        next(heading.iter(W + "t")).text = "1.1 新标题"
                        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    dst.writestr(item, data)
            self.assertEqual(generate_toc(revised, final), 3)
            with ZipFile(final) as doc:
                root = ET.fromstring(doc.read("word/document.xml"))
                codes = [t.text or "" for t in root.iter(W + "instrText")]
                self.assertEqual(sum(c.strip().startswith("TOC ") for c in codes), 1)
                links = [
                    r
                    for r in root.iter(W + "hyperlink")
                    if r.get(W + "anchor", "").startswith("_Toc")
                ]
                self.assertEqual(len(links), 3)
                self.assertIn("新标题", "".join(links[1].itertext()))
            from win32com.client import DispatchEx

            app = DispatchEx("Word.Application")
            try:
                app.DisplayAlerts = 0
                doc = app.Documents.Open(
                    str(final), ReadOnly=True, AddToRecentFiles=False
                )
                toc = doc.TablesOfContents(1)
                toc.Update()
                start = toc.Range.Start
                title = doc.Range(start - 1, start - 1).Paragraphs(1).Range
                self.assertEqual(
                    (
                        title.Characters(1).Font.NameFarEast,
                        title.Characters(1).Font.Size,
                    ),
                    ("黑体", 16),
                )
                for index in (1, 2, 3):
                    row = toc.Range.Paragraphs(index).Range
                    font = row.Characters(1).Font
                    self.assertEqual(
                        (font.NameFarEast, font.NameAscii, font.Size, font.Underline),
                        ("宋体", "Times New Roman", 12, 1),
                    )
                    pf = row.ParagraphFormat
                    self.assertEqual(
                        (
                            pf.Alignment,
                            pf.CharacterUnitLeftIndent,
                            pf.CharacterUnitRightIndent,
                            pf.CharacterUnitFirstLineIndent,
                        ),
                        (0, 2, 0, 0),
                    )
                    self.assertEqual(
                        (pf.SpaceAfter, pf.LineSpacingRule, pf.OutlineLevel), (8, 5, 10)
                    )
                    self.assertAlmostEqual(pf.LineSpacing / 12, 1.16, places=2)
            finally:
                try:
                    app.Documents.Close(0)
                finally:
                    app.Quit(0)
            with self.assertRaises(FileExistsError):
                generate_toc(revised, final)
