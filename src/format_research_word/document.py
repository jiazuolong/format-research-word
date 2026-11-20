"""Edit text while retaining the rest of the OOXML package, including media and fields."""

import re
import shutil
import tempfile
from collections import Counter
from copy import deepcopy
from itertools import groupby
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET

from . import audit
from .inline import PROTECTED, normalize_inline
from .layout import RPR_ORDER, ScriptStyles, normalize_story, set_font, story_names
from .typography import Char, joined, round_values, typography

W, M = audit.W, audit.M
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
HEADING = re.compile(
    r"^(?:第[一二三四五六七八九十\d]+章|[图表][A-E]?\d+|\d+(?:\.\d+)+|附录[A-E])\s"
)


def prop(parent, tag, first=False):
    result = parent.find(tag)
    if result is None:
        result = ET.Element(tag)
        parent.insert(0 if first else len(parent), result)
    return result


def attrs(run, ch):
    props = prop(run, W + "rPr", first=True)
    if ch.chemical or ch.script == "superscript":
        set_font(props)
    if ch.script:
        prop(props, W + "vertAlign").set(W + "val", ch.script)
    if ch.upright:
        prop(props, W + "i").set(W + "val", "0")
        prop(props, W + "iCs").set(W + "val", "0")
    # Word expects run properties in schema order.
    ranks = {W + tag: i for i, tag in enumerate(RPR_ORDER)}
    props[:] = sorted(props, key=lambda e: ranks.get(e.tag, len(ranks)))


def rewrite(runs, chars):
    if not runs:
        return
    for _, group in groupby(
        chars, key=lambda c: (c.origin, c.script, c.chemical, c.upright)
    ):
        piece = list(group)
        source = runs[piece[0].origin]
        new = ET.Element(source.tag, dict(source.attrib))
        props = source.find(W + "rPr")
        if props is not None:
            new.append(deepcopy(props))
        if piece[0].script or piece[0].chemical or piece[0].upright:
            attrs(new, piece[0])
        t = ET.SubElement(new, W + "t")
        t.text = joined(piece)
        if t.text[0].isspace() or t.text[-1].isspace():
            t.set(XML_SPACE, "preserve")
        runs[0].addprevious(new)
    for r in runs:
        r.getparent().remove(r)


def field_runs(root):
    marked = set()
    depth = 0
    for r in root.iter(W + "r"):
        fields = list(r.iter(W + "fldChar"))
        if depth or fields:
            marked.add(r)
        for item in fields:
            kind = item.get(W + "fldCharType")
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth = max(0, depth - 1)
    return marked


def chunks(parent, skip):
    chunk = []
    for r in parent:
        if (
            r.tag == W + "r"
            and r not in skip
            and len(r)
            and all(n.tag in (W + "t", W + "rPr") for n in r)
            and r.find(W + "t") is not None
        ):
            chunk.append(r)
        else:
            if chunk:
                yield chunk
                chunk = []
    if chunk:
        yield chunk


def chars_of(runs):
    return [
        Char(ch, i)
        for i, r in enumerate(runs)
        for t in r.findall(W + "t")
        for ch in (t.text or "")
    ]


def find_bibliography(body):
    blocks = list(body)
    starts = [
        i
        for i, p in enumerate(blocks)
        if p.tag == W + "p" and audit.BIB_TITLE.fullmatch(audit.visible_text(p).strip())
    ]
    if len(starts) != 1:
        return starts, []
    stop = starts[0] + 1
    while stop < len(blocks):
        p = blocks[stop]
        if p.tag != W + "p" or re.match(
            r"^(?:附录|致谢|附件|appendix\b|acknowledg)",
            audit.visible_text(p).strip(),
            re.I,
        ):
            break
        stop += 1
    return starts, blocks[starts[0] + 1 : stop]


def renumber(body, bib, source, fields):
    info = audit.run(source)
    if info["范围提示"] or info["需区分的方括号项"]:
        raise ValueError("引文需要人工判定，请运行 audit 查看映射、注释和域提示。")
    for node in body.iter():
        if node.tag in PROTECTED or node.tag in (
            W + "footnoteReference",
            W + "endnoteReference",
        ):
            # Simple page/sequence fields elsewhere do not affect the bibliography.
            if node.tag == W + "fldSimple" and not re.search(
                r"CITATION|BIBLIOGRAPHY|ADDIN", node.get(W + "instr", ""), re.I
            ):
                continue
            raise ValueError("有修订、注释引用或动态引文字段，请用工作流核对编号。")
    mapping = info["建议原号至新号"]
    if not mapping or not info["静态候选引用处"]:
        raise ValueError("未识别到可重排的静态编号文献和正文引文。")
    records = {}
    for b in bib:
        text = audit.visible_text(b).strip()
        if not text:
            continue
        m = audit.ENTRY.match(text)
        if not m:
            raise ValueError("文献表存在续行段或未编号段，请用工作流确认条目边界。")
        old = int(m[1] or m[2])
        if old in records:
            raise ValueError("文后条目有重复编号。")
        records[old] = b
    if set(records) != set(mapping):
        raise ValueError("文献表与正文使用的编号范围需要核对。")
    seen = 0
    for p in body.iter(W + "p"):
        if p in bib:
            continue
        for parent in [p, *p.findall(W + "hyperlink")]:
            for group in list(chunks(parent, fields)):
                chars = chars_of(group)
                text = joined(chars)
                matches = list(audit.CITATION.finditer(text))
                for m in reversed(matches):
                    ids = audit.groups(m[1], max(mapping), mapping)
                    if not ids or any(k not in mapping for k in ids):
                        raise ValueError("遇到含义不明的方括号标记。")
                    values = sorted({mapping[k] for k in ids})
                    val = m[0][0] + ",".join(map(str, values)) + m[0][-1]
                    chars[m.start() : m.end()] = [
                        Char(c, chars[m.start()].origin, script="superscript")
                        for c in val
                    ]
                    seen += 1
                if matches:
                    rewrite(group, chars)
    if seen != info["静态候选引用处"]:
        raise ValueError("引文横跨域、书签或特殊对象，请按审查结果处理。")
    for old, b in records.items():
        # Locate the printed number without dropping tabs, page-break hints or hyperlinks.
        pieces = []
        offset = 0
        for run in b.iter(W + "r"):
            for item in run:
                if item.tag == W + "t":
                    length = len(item.text or "")
                    pieces.append((offset, offset + length, item, run))
                    offset += length
                elif item.tag in (W + "tab", W + "br"):
                    offset += 1
        visible = "".join(
            (item.text or "") if item.tag == W + "t" else "\t"
            for run in b.iter(W + "r")
            for item in run
            if item.tag in (W + "t", W + "tab", W + "br")
        )
        hit = re.match(r"^\s*[\[［]?(\d+)", visible)
        if not hit or int(hit[1]) != old:
            raise ValueError("无法定位文后条目的可编辑编号。")
        a, z = hit.span(1)
        replacement = str(mapping[old])
        for left, right, item, run in pieces:
            if left < z and a < right:
                if run in fields:
                    raise ValueError("文后条目号由域控制。")
                value = item.text or ""
                item.text = (
                    value[: max(0, a - left)]
                    + replacement
                    + value[min(right, z) - left :]
                )
                replacement = ""
    slots = sorted(body.index(b) for b in records.values())
    for p in records.values():
        body.remove(p)
    for slot, old in zip(slots, sorted(records, key=mapping.get)):
        body.insert(slot, records[old])
    return seen


def math_run(run, value, upright):
    new = deepcopy(run)
    nt = new.find(M + "t")
    nt.text = value
    if value[:1].isspace() or value[-1:].isspace():
        nt.set(XML_SPACE, "preserve")
    if upright:
        prop(prop(new, M + "rPr", first=True), M + "sty").set(M + "val", "p")
        rp = new.find(W + "rPr")
        if rp is not None:
            for italic in rp.findall(W + "i") + rp.findall(W + "iCs"):
                italic.set(W + "val", "0")
    return new


def format_math(root):
    changed = 0
    for run in list(root.iter(M + "r")):
        if any(a.tag in PROTECTED for a in run.iterancestors()):
            continue
        t = run.find(M + "t")
        if (
            t is None
            or not t.text
            or any(n.tag not in (M + "rPr", W + "rPr", M + "t") for n in run)
        ):
            continue
        original = [Char(c, 0) for c in t.text]
        chars = original.copy()
        round_values(chars)
        text = joined(chars)
        # Recognize an isolated x, a tensor's xx index, or the x in a differential dx.
        for m in re.finditer(r"(?<![A-Za-z_])d?(x+)(?![A-Za-z_])", text):
            for i in range(*m.span(1)):
                c = chars[i]
                chars[i] = Char(c.text, c.origin, c.script, c.chemical, True)
        if chars == original:
            continue

        pos = 0
        while pos < len(chars):
            c = chars[pos]
            end = pos + 1
            while end < len(chars) and (chars[end].upright, chars[end].script) == (
                c.upright,
                c.script,
            ):
                end += 1
            value = joined(chars[pos:end])
            if end < len(chars) and chars[end].script == "superscript":
                assert value.endswith("×10")
                run.addprevious(math_run(run, value[:-2], c.upright))
                exp_end = end
                while exp_end < len(chars) and chars[exp_end].script == "superscript":
                    exp_end += 1
                power = ET.Element(M + "sSup")
                ET.SubElement(power, M + "e").append(math_run(run, "10", True))
                ET.SubElement(power, M + "sup").append(
                    math_run(run, joined(chars[end:exp_end]), True)
                )
                run.addprevious(power)
                end = exp_end
            else:
                run.addprevious(math_run(run, value, c.upright))
            pos = end
        run.getparent().remove(run)
        changed += 1
    return changed


def format_docx(source, destination, renumber_citations=False, *, layout_only=False):
    if layout_only and renumber_citations:
        raise ValueError("仅排表格和上标与重排引文应分开调用。")
    source, destination = Path(source), Path(destination)
    if source.suffix.lower() != ".docx" or destination.suffix.lower() != ".docx":
        raise ValueError("输入与输出都应使用 .docx 扩展名。")
    if destination.exists():
        raise FileExistsError(f"输出已存在，请指定新文件：{destination}")
    stats = Counter()
    with ZipFile(source) as archive:
        parser = ET.XMLParser(resolve_entities=False, no_network=True)
        root = ET.fromstring(archive.read("word/document.xml"), parser)
        body = root.find(W + "body")
        if body is None:
            raise ValueError("DOCX 中未识别到 Word 主正文。")
        styles = ScriptStyles(
            ET.fromstring(archive.read("word/styles.xml"), parser)
            if "word/styles.xml" in archive.namelist()
            else None
        )
        if not layout_only:
            _, bib = find_bibliography(body)
            fields = field_runs(body)
            if renumber_citations:
                stats["重排引文处"] = renumber(body, bib, source, fields)
            for p in body.iter(W + "p"):
                if p in bib or any(a.tag in PROTECTED for a in p.iterancestors()):
                    continue
                structural = bool(HEADING.match(audit.visible_text(p)))
                for chunk in list(chunks(p, fields)):
                    for raised, runs in groupby(chunk, key=styles.is_raised):
                        group = list(runs)
                        orig = chars_of(group)
                        new = typography(
                            list(orig), structural, recognize_formulas=not raised
                        )
                        if new != orig:
                            rewrite(group, new)
                            stats["处理文本组"] += 1
            stats["处理原生公式文本组"] = format_math(body)
        parts = {}
        for partname in ("word/document.xml", *story_names(archive, parser)):
            story = (
                root
                if partname == "word/document.xml"
                else ET.fromstring(archive.read(partname), parser)
            )
            tables, supers = normalize_story(story, styles)
            inline_changes = 0
            if not layout_only:
                inline_changes = normalize_inline(
                    story,
                    styles,
                    field_runs(story),
                    set(bib) if story is root else set(),
                )
                stats["关系符、破折号及引文排版文本组"] += inline_changes
            stats["自动调整的表格"] += tables
            stats["处理上标文本组"] += supers
            if partname == "word/document.xml" or tables or supers or inline_changes:
                parts[partname] = ET.tostring(
                    story, encoding="UTF-8", xml_declaration=True, standalone=True
                )
        # Build the package before creating the destination. Concurrent runs cannot overwrite one another.
        with tempfile.TemporaryFile() as tmp:
            with ZipFile(tmp, "w") as output:
                for item in archive.infolist():
                    output.writestr(
                        item,
                        parts[item.filename]
                        if item.filename in parts
                        else archive.read(item),
                    )
            tmp.seek(0)
            with destination.open("xb") as out:
                try:
                    shutil.copyfileobj(tmp, out)
                except BaseException:
                    out.close()
                    destination.unlink(missing_ok=True)
                    raise
    return dict(stats)
