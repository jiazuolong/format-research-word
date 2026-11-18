"""Inspect a DOCX for static numbered references and numeric display candidates."""

from __future__ import annotations

import argparse
import json
import re
from decimal import ROUND_HALF_UP, Decimal, localcontext
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
CITATION = re.compile(r"[\[［]\s*([1-9]\d*(?:\s*[,，、;；–—－-]\s*[1-9]\d*)*)\s*[\]］]")
DECIMAL = re.compile(
    r"(?<![A-Za-z_\d.])[+−-]?(?:\d+\.\d{5,}|\.\d{5,})(?![\d_]|\.\d)(?:(?=e[+-]?\d|meV|eV|[GMk]?Hz|[μµncm]m|K|GPa|MPa|Pa)|(?=[^A-Za-z]|$))"
)
IDENTIFIER = re.compile(
    r"https?://[^\s，。；）]+|\b10\.\d{4,9}/[^\s，。；）]+|\b[vV]?\d+(?:\.\d+){2,}\b|\barXiv:\s*[\w./-]+",
    re.I,
)
BIB_TITLE = re.compile(
    r"^(?:(?:第[一二三四五六七八九十\d]+章|\d+(?:\.\d+)*[、.．]?)\s*)?(?:参考文献|引用文献|references|bibliography)\s*[：:]?$",
    re.I,
)
ENTRY = re.compile(r"^\s*(?:[\[［](\d+)[\]］]|(\d+)(?=[\t .、．)]))")


def visible_text(p):
    chars = []
    for n in p.iter():
        if n.tag in (W + "t", M + "t"):
            chars.append(n.text or "")
        elif n.tag in (W + "tab", W + "br"):
            chars.append("\t")
        elif n.tag in (M + "sub", M + "sup", M + "den"):
            chars.append("\u2063")
    return "".join(chars)


def groups(expr, max_id, available=None):
    values = []
    for word in re.split(r"\s*[,，、;；]\s*", expr):
        edge = re.split(r"\s*[–—－-]\s*", word)
        low, high = int(edge[0]), int(edge[-1])
        if len(edge) > 2 or not 1 <= low <= high <= max_id:
            return None
        if available is not None and high - low + 1 > len(available):
            return None
        values.extend(range(low, high + 1))
    return values


def display_three(token):
    value = Decimal(token.replace("−", "-"))
    if not value:
        return "0.00"
    with localcontext() as ctx:
        ctx.prec = max(30, len(value.as_tuple().digits) + 3)
        places = Decimal(1).scaleb(value.adjusted() - 2)
        rounded = value.quantize(places, rounding=ROUND_HALF_UP)
        # A carry can change which decimal place is significant.
        rounded = rounded.quantize(Decimal(1).scaleb(rounded.adjusted() - 2))
    if rounded.adjusted() < -3 or rounded.adjusted() >= 3:
        return f"{rounded.scaleb(-rounded.adjusted()):.2f}×10^({rounded.adjusted()})"
    return format(rounded, "f")


def run(path):
    with ZipFile(path) as z:
        doc = ET.fromstring(z.read("word/document.xml"))
        styles = (
            ET.fromstring(z.read("word/styles.xml"))
            if "word/styles.xml" in z.namelist()
            else ET.Element(W + "styles")
        )
        notes = [
            s for s in ("word/footnotes.xml", "word/endnotes.xml") if s in z.namelist()
        ]
    body = doc.find(W + "body")
    items = list(body.iter(W + "p"))
    names = {}
    for s in styles.findall(W + "style"):
        name = s.find(W + "name")
        if name is not None:
            names[s.get(W + "styleId")] = name.get(W + "val", "")
    paragraphs = []
    for n, p in enumerate(items, 1):
        st = p.find(W + "pPr/" + W + "pStyle")
        style = names.get(st.get(W + "val"), "") if st is not None else ""
        if re.match(r"^(?:toc\s*\d|目录\s*\d)", style, re.I):
            continue
        paragraphs.append((n, style, visible_text(p).strip(), p))
    sections = [i for i, item in enumerate(paragraphs) if BIB_TITLE.fullmatch(item[2])]
    start = sections[0] if len(sections) == 1 else len(paragraphs)
    entries = {}
    flags = []
    end = start + 1
    if len(sections) != 1:
        flags.append(f"文献表标题数为 {len(sections)}，需结合报告结构定位文献表。")
    for i in range(start + 1, len(paragraphs)):
        n, style, text, p = paragraphs[i]
        if re.match(r"^(?:附录|致谢|附件|appendix\b|acknowledg)", text, re.I) or (
            text and style and style == paragraphs[start][1] and not ENTRY.match(text)
        ):
            break
        end = i + 1
        hit = ENTRY.match(text)
        if hit:
            old = int(hit[1] or hit[2])
            if old in entries:
                flags.append(f"文献表重复原号 {old}，见段落 {n}。")
            entries[old] = text[hit.end() :].strip(" \t.．、)")
    maximum = max(entries, default=0)
    seen = {}
    sites = 0
    numbers = []
    ambiguous = []
    citation_fields = 0
    note_markers = 0
    for i, (n, style, text, p) in enumerate(paragraphs):
        if start <= i < end:
            continue
        instr = "".join(k.text or "" for k in p.iter(W + "instrText"))
        if re.search(r"CITATION|ZOTERO|EN\.CITE|MENDELEY", instr, re.I):
            citation_fields += 1
        if any(
            v.tag in (W + "footnoteReference", W + "endnoteReference") for v in p.iter()
        ):
            note_markers += 1
        for hit in CITATION.finditer(text):
            nums = groups(hit[1], maximum, entries)
            if not nums or not all(v in entries for v in nums):
                ambiguous.append({"段": n, "标记": hit[0]})
                continue
            sites += 1
            for v in nums:
                seen.setdefault(v, len(seen) + 1)
        guards = [h.span() for h in IDENTIFIER.finditer(text)]
        for h in DECIMAL.finditer(text):
            if any(a <= h.start() < b for a, b in guards):
                continue
            numbers.append(
                {
                    "段": n,
                    "原值": h[0],
                    "建议显示": display_three(h[0]),
                    "上下文": text[max(0, h.start() - 15) : h.end() + 15],
                }
            )
    if not entries and sections:
        flags.append(
            "未读到显式条目号。自动编号或动态书目应在 Word 或文献管理器中核对。"
        )
    if citation_fields:
        flags.append(f"发现 {citation_fields} 个段落含文献管理域，请按域来源核对。")
    if notes and note_markers:
        flags.append("存在正文脚注或尾注标记，需将注内引文纳入首次出现顺序。")
    unused = [v for v in entries if v not in seen]
    scheduled = dict(seen)
    for v in unused:
        scheduled[v] = len(scheduled) + 1
    return {
        "文件": str(path),
        "表": sum(1 for _ in body.iter(W + "tbl")),
        "图形": sum(1 for _ in body.iter(W + "drawing")),
        "嵌入对象": sum(1 for _ in body.iter(W + "object")),
        "原生公式": sum(1 for _ in body.iter(M + "oMath")),
        "文献条目": len(entries),
        "静态候选引用处": sites,
        "已引用条目": len(seen),
        "首次出现次序": list(seen),
        "未见静态引用的原号": unused,
        "建议原号至新号": scheduled,
        "长小数候选": numbers,
        "需区分的方括号项": ambiguous,
        "范围提示": flags,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", type=Path, help="可选，输出至新建 JSON 文件")
    args = parser.parse_args()
    result = run(args.docx)
    if args.json:
        with args.json.open("x", encoding="utf-8") as out:
            json.dump(result, out, ensure_ascii=False, indent=2)
    print(
        "表 {表}，图形 {图形}，嵌入对象 {嵌入对象}，原生公式 {原生公式}".format(
            **result
        )
    )
    print(
        "识别文献 {文献条目} 条，候选静态引文 {静态候选引用处} 处，涉及 {已引用条目} 条。".format(
            **result
        )
    )
    print("首次出现的前十个原号", result["首次出现次序"][:10])
    print("未见静态引用的原号", result["未见静态引用的原号"])
    print("长小数候选", len(result["长小数候选"]))
    for hit in result["长小数候选"][:5]:
        print(" ", hit["上下文"], "→", hit["建议显示"])
    for msg in result["范围提示"]:
        print(msg)


if __name__ == "__main__":
    main()
