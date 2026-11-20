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
            with self.assertRaises(FileExistsError):
                generate_toc(revised, final)
