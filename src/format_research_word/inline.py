"""Spacing around relations and superscript references, preserving Word objects."""

import re
from copy import deepcopy
from dataclasses import dataclass, replace
from itertools import groupby

from lxml import etree as ET

from .audit import CITATION, M, W
from .layout import RPR_ORDER, ordered, text_value
from .typography import Char, exclusions, joined, overlaps

BLANK = " \u00a0\u2009\u202f\u3000"
WS = "[" + BLANK + "]*"
RELATION = re.compile(r"<=|>=|!=|==|:=|[=<>≤≥≠≈≃≅≡≪≫≲≳≦≧∝∈∉＝＜＞]")
DASH = re.compile(WS + r"(?P<dash>—+|⸺)" + WS)
PUNCT = "。，、；：！？….,;:!?"
CLOSE = "”’」』）》】）"
OBJECT = "\ufffc"
MARKUP = re.compile(r"<(?:/?[A-Za-z][\w:.-]*(?:\s+[^<>]*?)?/?|[^<>\s]+@[^<>\s]+)>")
PROTECTED = {W + s for s in ("del", "ins", "moveTo", "moveFrom", "fldSimple")}


def symbol_spacing(chars, *, math=False):
    """Edit only the adjacent horizontal spaces. Keep multi-character signs intact."""
    text = joined(chars)
    if not RELATION.search(text) and not DASH.search(text):
        return chars
    protected = exclusions(text)
    if not math:
        for tag in MARKUP.finditer(text):
            # The middle of a chained comparison can resemble a short XML tag.
            if (
                tag.start()
                and tag.end() < len(text)
                and re.fullmatch(r"<[A-Za-z]>", tag[0])
                and re.fullmatch(r"[A-Za-z0-9]", text[tag.start() - 1])
                and re.fullmatch(r"[A-Za-z0-9]", text[tag.end()])
            ):
                continue
            protected.append(tag.span())
    edits = []
    for m in DASH.finditer(text):
        left = text[m.start() - 1 : m.start()] if m.start() else ""
        right = text[m.end() : m.end() + 1]
        if re.search(r"[A-Za-z0-9]", left + right) and not overlaps(
            *m.span(), protected
        ):
            edits.append(
                (m.start(), m.end(), [replace(chars[m.start("dash")], text="-")])
            )
    for m in RELATION.finditer(text):
        a, b = m.span()
        if overlaps(a, b, protected):
            continue
        left, right = a, b
        while left and text[left - 1] in BLANK:
            left -= 1
        while right < len(text) and text[right] in BLANK:
            right += 1
        if not math:
            # Check operand boundaries in prose, including standalone bounds such as <5.
            if right == len(text):
                continue
            if left and not (
                text[left - 1].isalnum() or text[left - 1] in ")]}）］′'″%°∞|" + OBJECT
            ):
                continue
            if not (text[right].isalnum() or text[right] in "+−-±(（[.√∞|" + OBJECT):
                continue
        for start, stop, origin in ((left, a, a), (b, right, b - 1)):
            if start == 0 or stop == len(text) or text[start:stop] == " ":
                continue
            edits.append((start, stop, [replace(chars[origin], text=" ")]))
    for start, stop, replacement in sorted(
        edits, key=lambda edit: edit[:2], reverse=True
    ):
        chars[start:stop] = replacement
    return chars


@dataclass
class Piece:
    nodes: list
    text: str
    raised: bool = False
    opaque: bool = False


def citation_spacing(chars, pieces):
    """Move original punctuation past the reference; the separating space is baseline."""
    text = joined(chars)
    found = [
        m
        for m in CITATION.finditer(text)
        if all(pieces[c.origin].raised for c in chars[m.start() : m.end()])
    ]
    if not found:
        return chars
    result = []
    pos = 0
    for m in found:
        result.extend(chars[pos : m.start()])
        while result and result[-1].text in BLANK:
            result.pop()
        end = len(result)
        cut = end
        while cut and result[cut - 1].text in CLOSE:
            cut -= 1
        if cut and result[cut - 1].text in PUNCT:
            while cut and result[cut - 1].text in PUNCT:
                cut -= 1
        else:
            cut = end
        punctuation = [
            replace(c, script="baseline") if pieces[c.origin].raised else c
            for c in result[cut:]
        ]
        del result[cut:]
        while result and result[-1].text in BLANK:
            result.pop()
        if result:
            result.append(Char(" ", -1))
        result.extend(chars[m.start() : m.end()])
        result.extend(punctuation)
        pos = m.end()
        while pos < len(text) and text[pos] in BLANK:
            pos += 1
        if pos < len(text) and text[pos] in PUNCT:
            while pos < len(text) and text[pos] in PUNCT + CLOSE:
                c = chars[pos]
                result.append(
                    replace(c, script="baseline") if pieces[c.origin].raised else c
                )
                pos += 1
        else:
            pos = m.end()
    result.extend(chars[pos:])
    # Do not split or replace correctly formatted runs on a second pass.
    if (
        joined(result) == text
        and not any(c.script for c in result)
        and not any(c.text == " " and pieces[c.origin].raised for c in chars)
    ):
        return chars
    return result


def plain_run(run):
    return (
        run.tag == W + "r"
        and run.find(W + "t") is not None
        and all(n.tag in (W + "rPr", W + "t") for n in run)
    )


def reference_object(nodes, styles):
    visible = [t for node in nodes for t in node.iter(W + "t") if t.text]
    text = "".join(t.text for t in visible)
    if (CITATION.fullmatch(text) or text.isdecimal()) and all(
        styles.is_raised(t.getparent()) for t in visible
    ):
        return Piece(nodes, text, True, True)
    return Piece(nodes, OBJECT, False, True)


def word_groups(p, styles, fields):
    # A cached page boundary can occur inside a text run, including between the
    # last word of a sentence and its reference. It has no visible characters.
    for run in p.findall(W + "r"):
        if (
            run in fields
            or run.find(W + "lastRenderedPageBreak") is None
            or run.find(W + "t") is None
        ):
            continue
        pr = run.find(W + "rPr")
        for child in list(run):
            if child is pr:
                continue
            segment = ET.Element(run.tag, dict(run.attrib))
            if pr is not None:
                segment.append(deepcopy(pr))
            segment.append(child)
            run.addprevious(segment)
        p.remove(run)
    children = list(p)
    group = []
    i = 0
    while i < len(children):
        node = children[i]
        i += 1
        if plain_run(node) and node not in fields:
            group.append(
                Piece(
                    [node],
                    "".join(t.text or "" for t in node.findall(W + "t")),
                    styles.is_raised(node),
                )
            )
            continue
        if node.tag in (W + "hyperlink", W + "fldSimple"):
            group.append(reference_object([node], styles))
            continue
        if node.tag == M + "oMath":
            group.append(Piece([node], OBJECT, False, True))
            continue
        if (
            node.tag == W + "r"
            and node.find(W + "lastRenderedPageBreak") is not None
            and all(n.tag in (W + "rPr", W + "lastRenderedPageBreak") for n in node)
        ):
            group.append(Piece([node], "", False, True))
            continue
        marker = node.find(W + "fldChar")
        if marker is not None and marker.get(W + "fldCharType") == "begin":
            start = i - 1
            depth = 1
            while i < len(children) and depth:
                for flag in children[i].iter(W + "fldChar"):
                    kind = flag.get(W + "fldCharType")
                    depth += (kind == "begin") - (kind == "end")
                i += 1
            if depth == 0:
                group.append(reference_object(children[start:i], styles))
                continue
        if group:
            yield group
            group = []
    if group:
        yield group


def emit_word(group, chars):
    parent = group[0].nodes[0].getparent()
    index = parent.index(group[0].nodes[0])
    for piece in group:
        for node in piece.nodes:
            parent.remove(node)
    output = []
    for (origin, script), part in groupby(chars, key=lambda c: (c.origin, c.script)):
        value = joined(part)
        if origin < 0:
            run = ET.Element(W + "r")
            rp = ET.SubElement(run, W + "rPr")
            ordered(rp, "vertAlign", RPR_ORDER).set(W + "val", "baseline")
        else:
            piece = group[origin]
            if piece.opaque or (value == piece.text and not script):
                output.extend((origin, node) for node in piece.nodes)
                continue
            source = piece.nodes[0]
            run = ET.Element(W + "r", dict(source.attrib))
            pr = source.find(W + "rPr")
            if pr is not None:
                run.append(deepcopy(pr))
            if script:
                rp = run.find(W + "rPr")
                if rp is None:
                    rp = ET.SubElement(run, W + "rPr")
                ordered(rp, "vertAlign", RPR_ORDER).set(W + "val", script)
                pos = rp.find(W + "position")
                if pos is not None:
                    rp.remove(pos)
        text_value(ET.SubElement(run, W + "t"), value)
        output.append((origin, run))
    for origin, piece in enumerate(group):
        if piece.text:
            continue
        at = next(
            (i for i, (owner, _) in enumerate(output) if owner > origin), len(output)
        )
        output[at:at] = [(origin, node) for node in piece.nodes]
    parent[index:index] = [node for _, node in output]


def math_spacing(root, fields):
    changed = 0
    # Each fraction argument, script, or matrix cell is a separate expression.
    for parent in root.iter():
        if (
            not isinstance(parent.tag, str)
            or not parent.tag.startswith(M)
            or parent.find(M + "r") is None
        ):
            continue
        if any(a.tag in PROTECTED or a in fields for a in parent.iterancestors()):
            continue
        nodes = []
        chars = []
        for child in parent:
            if not isinstance(child.tag, str) or ET.QName(child).localname.endswith(
                "Pr"
            ):
                continue
            texts = child.findall(M + "t") if child.tag == M + "r" else []
            if not texts:
                chars.append(Char(OBJECT, -1))
            for t in texts:
                chars.extend(Char(c, len(nodes)) for c in t.text or "")
                nodes.append(t)
        new = symbol_spacing(list(chars), math=True)
        if new != chars:
            values = [[] for _ in nodes]
            for c in new:
                if c.origin >= 0:
                    values[c.origin].append(c.text)
            for node, value in zip(nodes, values):
                text_value(node, "".join(value))
            changed += 1
    return changed


def normalize_inline(root, styles, fields, bibliography):
    changed = math_spacing(root, fields)
    for p in root.iter(W + "p"):
        if any(a.tag in PROTECTED for a in p.iterancestors()):
            continue
        text = "".join(t.text or "" for t in p.iter(W + "t"))
        if not (
            RELATION.search(text) or DASH.search(text) or "[" in text or "［" in text
        ):
            continue
        for group in word_groups(p, styles, fields):
            original = [Char(c, i) for i, piece in enumerate(group) for c in piece.text]
            chars = symbol_spacing(list(original))
            if p not in bibliography:
                chars = citation_spacing(chars, group)
            if chars != original:
                emit_word(group, chars)
                changed += 1
    return changed
