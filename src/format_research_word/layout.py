"""Word-native table fitting and superscripts across the document's text stories."""

import re
import unicodedata
from copy import deepcopy
from functools import lru_cache
from itertools import groupby

from lxml import etree as ET

from .audit import M, W

FONT = "Times New Roman"
SPACE = "{http://www.w3.org/XML/1998/namespace}space"
STORY_TYPES = {
    "header",
    "footer",
    "footnotes",
    "endnotes",
    "comments",
    "document.glossary",
}
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
HISTORY = {
    W + n
    for n in (
        "del",
        "moveFrom",
        "tblPrChange",
        "tblPrExChange",
        "tcPrChange",
        "rPrChange",
    )
}
RPR_ORDER = [
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
]
TABLE_ORDER = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
]


def ordered(parent, tag, order):
    item = parent.find(W + tag)
    if item is None:
        item = ET.Element(W + tag)
        later = {W + s for s in order[order.index(tag) + 1 :]}
        pos = next((i for i, e in enumerate(parent) if e.tag in later), len(parent))
        parent.insert(pos, item)
    return item


def set_font(pr):
    font = ordered(pr, "rFonts", RPR_ORDER)
    for key in list(font.attrib):
        if "theme" in key.lower():
            del font.attrib[key]
    for family in ("ascii", "hAnsi", "eastAsia", "cs"):
        font.set(W + family, FONT)


def fit_tables(root):
    count = 0
    for table in root.iter(W + "tbl"):
        if any(a.tag in HISTORY for a in table.iterancestors()):
            continue
        pr = table.find(W + "tblPr")
        if pr is None:
            pr = ET.Element(W + "tblPr")
            table.insert(0, pr)
        width = ordered(pr, "tblW", TABLE_ORDER)
        width.set(W + "type", "pct")
        width.set(W + "w", "5000")
        ordered(pr, "tblLayout", TABLE_ORDER).set(W + "type", "autofit")
        count += 1
    # Remove fixed cell preferences so Word can measure the contents when it lays out
    # each table. Leave grid spans, merge markers and starting grid geometry intact.
    for tcpr in root.iter(W + "tcPr"):
        if any(a.tag in HISTORY for a in tcpr.iterancestors()):
            continue
        width = tcpr.find(W + "tcW")
        if width is not None:
            width.set(W + "type", "auto")
            width.set(W + "w", "0")
    for ex in root.iter(W + "tblPrEx"):
        if any(a.tag in HISTORY for a in ex.iterancestors()):
            continue
        for name, values in (
            ("tblW", {"type": "pct", "w": "5000"}),
            ("tblLayout", {"type": "autofit"}),
        ):
            el = ex.find(W + name)
            if el is not None:
                for key, val in values.items():
                    el.set(W + key, val)
    return count


class ScriptStyles:
    def __init__(self, styles):
        self.by_id = {}
        self.default = {}
        self.doc_align = None
        if styles is not None:
            for s in styles.findall(W + "style"):
                sid = s.get(W + "styleId")
                self.by_id[sid] = s
                if s.get(W + "default") in ("1", "true"):
                    self.default[s.get(W + "type")] = sid
            self.doc_align = styles.find(
                W + "docDefaults/" + W + "rPrDefault/" + W + "rPr/" + W + "vertAlign"
            )

        self.alignments = {sid: self.style_value(sid) for sid in self.by_id}

    def style_value(self, sid):
        visited = set()
        while sid in self.by_id and sid not in visited:
            visited.add(sid)
            s = self.by_id[sid]
            align = s.find(W + "rPr/" + W + "vertAlign")
            if align is not None:
                return align.get(W + "val")
            base = s.find(W + "basedOn")
            sid = base.get(W + "val") if base is not None else None
        return None

    def is_raised(self, r):
        align = r.find(W + "rPr/" + W + "vertAlign")
        if align is not None:
            return align.get(W + "val") == "superscript"
        if any(
            r.find(W + n) is not None
            for n in (
                "footnoteReference",
                "endnoteReference",
                "footnoteRef",
                "endnoteRef",
            )
        ):
            return True
        pos = r.find(W + "rPr/" + W + "position")
        if pos is not None:
            offset = pos.get(W + "val", "")
            label = "".join(t.text or "" for t in r.findall(W + "t")).strip()
            if (
                re.fullmatch(r"[+]?[0-9]+", offset)
                and int(offset) > 0
                and re.fullmatch(r"[0-9A-Za-z+−–=.,，、()［］\[\]\-]+", label)
            ):
                return True
        rs = r.find(W + "rPr/" + W + "rStyle")
        val = self.alignments.get(
            rs.get(W + "val") if rs is not None else self.default.get("character")
        )
        if val is not None:
            return val == "superscript"
        p = next(r.iterancestors(W + "p"), None)
        ps = p.find(W + "pPr/" + W + "pStyle") if p is not None else None
        val = self.alignments.get(
            ps.get(W + "val") if ps is not None else self.default.get("paragraph")
        )
        if val is not None:
            return val == "superscript"
        return (
            self.doc_align is not None
            and self.doc_align.get(W + "val") == "superscript"
        )


@lru_cache(maxsize=1024)
def ordinary(c):
    # Keep semantic symbols such as ™ and °; convert only superscript glyphs.
    name = unicodedata.name(c, "")
    if name.startswith(
        ("SUPERSCRIPT ", "MODIFIER LETTER SMALL ", "MODIFIER LETTER CAPITAL ")
    ) and unicodedata.decomposition(c).startswith("<super>"):
        return unicodedata.normalize("NFKC", c)
    return c


def props(run):
    pr = run.find(W + "rPr")
    if pr is None:
        pr = ET.Element(W + "rPr")
        run.insert(1 if run.find(M + "rPr") is not None else 0, pr)
    return pr


def mark(run, math=False):
    pr = props(run)
    set_font(pr)
    if math:
        mp = run.find(M + "rPr")
        if mp is None:
            mp = ET.Element(M + "rPr")
            run.insert(0, mp)
        if mp.find(M + "nor") is None:
            mp.insert(0, ET.Element(M + "nor"))
    else:
        ordered(pr, "vertAlign", RPR_ORDER).set(W + "val", "superscript")
        pos = pr.find(W + "position")
        if pos is not None:
            pr.remove(pos)


def text_value(el, s):
    el.text = s
    if s[:1].isspace() or s[-1:].isspace():
        el.set(SPACE, "preserve")


def split_glyphs(run, generated_powers):
    """Split only target runs, retaining tabs, breaks, objects and field characters in order."""
    is_math = run.tag == M + "r"
    text_tag = M + "t" if is_math else W + "t"
    for item in list(run):
        if item.tag in (W + "rPr", M + "rPr"):
            continue
        groups = (
            [(False, None)]
            if item.tag != text_tag
            else [
                (raised, "".join(ordinary(c) for c in chars))
                for raised, chars in groupby(
                    item.text or "", key=lambda c: ordinary(c) != c
                )
            ]
        )
        for raised, text in groups:
            new = ET.Element(run.tag, dict(run.attrib))
            new.extend(deepcopy(pr) for pr in run if pr.tag in (W + "rPr", M + "rPr"))
            node = item if text is None else deepcopy(item)
            new.append(node)
            if text is not None:
                text_value(node, text)
            run.addprevious(new)
            if raised:
                mark(new, math=is_math)
                if is_math:
                    # Attach the raised group to the preceding math atom. Use the last
                    # character or number of a multi-atom run, e.g. the y in x+y².
                    prev = new.getprevious()
                    if prev in generated_powers:
                        prev.find(M + "sup").append(new)
                        continue
                    power = ET.Element(M + "sSup")
                    generated_powers.add(power)
                    new.addprevious(power)
                    base = ET.SubElement(power, M + "e")
                    if (
                        prev is not None
                        and prev.tag.startswith(M)
                        and not ET.QName(prev).localname.endswith("Pr")
                    ):
                        last = prev.find(M + "t") if prev.tag == M + "r" else None
                        if last is None:
                            base.append(prev)
                        elif last.text:
                            tail = re.search(r"(?:\d+(?:\.\d+)?|.)$", last.text)
                            if tail and tail.start():
                                atom = deepcopy(prev)
                                text_value(atom.find(M + "t"), tail[0])
                                text_value(last, last.text[: tail.start()])
                                base.append(atom)
                            else:
                                base.append(prev)
                    ET.SubElement(power, M + "sup").append(new)
    run.getparent().remove(run)


def normalize_story(root, styles):
    tables = fit_tables(root)
    raised_count = 0
    generated_powers = set()
    for r in list(root.iter(W + "r", M + "r")):
        if any(a.tag in HISTORY for a in r.iterancestors()):
            continue
        math = r.tag == M + "r"
        tag = M + "t" if math else W + "t"
        data = r.findall(tag)
        in_exp = math and any(a.tag == M + "sup" for a in r.iterancestors())
        raised = in_exp or styles.is_raised(r)
        # Field instructions keep their original code and order; apply formatting to
        # visible field results and note marks only.
        if not data and not any(
            r.find(W + n) is not None
            for n in (
                "footnoteReference",
                "endnoteReference",
                "footnoteRef",
                "endnoteRef",
            )
        ):
            continue
        special = any(any(ordinary(c) != c for c in t.text or "") for t in data)
        if raised:
            for t in data:
                if special:
                    text_value(t, "".join(ordinary(c) for c in t.text or ""))
            mark(r, math=in_exp)
            raised_count += 1
        elif special:
            split_glyphs(r, generated_powers)
            raised_count += 1
    return tables, raised_count


def story_names(archive, parser):
    types = ET.fromstring(archive.read("[Content_Types].xml"), parser)
    for entry in types.findall(CT + "Override"):
        content = entry.get("ContentType", "")
        if (
            content.startswith(
                "application/vnd.openxmlformats-officedocument.wordprocessingml."
            )
            and content.removeprefix(
                "application/vnd.openxmlformats-officedocument.wordprocessingml."
            ).removesuffix("+xml")
            in STORY_TYPES
        ):
            yield entry.get("PartName").lstrip("/")
