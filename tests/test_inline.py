import unittest
from lxml import etree as ET

from format_research_word.audit import M, W
from format_research_word.document import field_runs
from format_research_word.inline import normalize_inline
from format_research_word.layout import ScriptStyles

NS = {"w": W[1:-1], "m": M[1:-1]}


def run(s, raised=False):
    node = ET.Element(W + "r")
    if raised:
        prop = ET.SubElement(node, W + "rPr")
        ET.SubElement(prop, W + "vertAlign").set(W + "val", "superscript")
    ET.SubElement(node, W + "t").text = s
    return ET.tostring(node, encoding="unicode")


def para(xml):
    return ET.fromstring(
        f'<w:p xmlns:w="{W[1:-1]}" xmlns:m="{M[1:-1]}">{xml}</w:p>'.encode()
    )


def visible(p):
    return "".join(p.xpath(".//w:t/text() | .//m:t/text()", namespaces=NS))


def format_p(p):
    normalize_inline(p, ScriptStyles(None), field_runs(p), set())


class InlineRules(unittest.TestCase):
    def test_bounds_and_chained_ascii_relations(self):
        p = para(run("<0.1，0<x>1，|x|<3。"))
        format_p(p)
        self.assertEqual(visible(p), "< 0.1，0 < x > 1，|x| < 3。")

    def test_relations_dashes_and_literals(self):
        p = para(
            run("区间0<")
            + run("x≤1，x  >")
            + run("=  0，DFT—")
            + run(
                "—模型，范围1——3，材料——说明。https://example.org/?q=x=3 与<tag>，−2。"
            )
        )
        format_p(p)
        self.assertEqual(
            visible(p),
            "区间0 < x ≤ 1，x >= 0，DFT-模型，范围1-3，材料——说明。https://example.org/?q=x=3 与<tag>，−2。",
        )
        snapshot = ET.tostring(p)
        format_p(p)
        self.assertEqual(snapshot, ET.tostring(p))

    def test_math_structure(self):
        p = para(
            "<m:oMath><m:f><m:num><m:r><m:t>a=b</m:t></m:r></m:num><m:den><m:r><m:t>c</m:t></m:r></m:den></m:f><m:r><m:t>≤</m:t></m:r><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>"
        )
        format_p(p)
        self.assertEqual(
            p.xpath(".//m:t/text()", namespaces=NS), ["a = b", "c", " ≤ ", "x", "2"]
        )
        self.assertEqual(len(p.xpath(".//m:f | .//m:sSup", namespaces=NS)), 2)

    def test_citation_after_punctuation_and_spacing(self):
        p = para(
            run("研究表明。")
            + run("  ")
            + run("[", True)
            + run("1,2]", True)
            + run("后续结果")
            + run("[3]，", True)
            + run("结论；")
            + run("[4]", True)
            + run("[5]", True)
        )
        format_p(p)
        self.assertEqual(visible(p), "研究表明 [1,2]。后续结果 [3]，结论 [4] [5]；")
        raised = p.xpath(
            './/w:r[w:rPr/w:vertAlign[@w:val="superscript"]]/w:t/text()', namespaces=NS
        )
        self.assertEqual("".join(raised), "[1,2][3][4][5]")
        snapshot = ET.tostring(p)
        format_p(p)
        self.assertEqual(snapshot, ET.tostring(p))

    def test_hyperlink_and_field(self):
        link = '<w:hyperlink w:anchor="Ref7">' + run("[7]", True) + "</w:hyperlink>"
        field = (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF Ref8 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            + run("[8]", True)
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        )
        p = para(run("文献，") + link + run("结果。") + field + run("坐标[1,2]。"))
        format_p(p)
        self.assertEqual(visible(p), "文献 [7]，结果 [8]。坐标[1,2]。")
        self.assertEqual(p.find(W + "hyperlink").get(W + "anchor"), "Ref7")
        self.assertEqual(p.find(".//" + W + "instrText").text, " REF Ref8 ")
        self.assertEqual(len(list(p.iter(W + "fldChar"))), 3)

    def test_citation_across_cached_page_breaks(self):
        p = para(
            "<w:r><w:lastRenderedPageBreak/><w:t>结果。</w:t></w:r>"
            "<w:r><w:lastRenderedPageBreak/></w:r>" + run("[4,5]", True)
        )
        format_p(p)
        self.assertEqual(visible(p), "结果 [4,5]。")
        self.assertEqual(len(list(p.iter(W + "lastRenderedPageBreak"))), 2)
