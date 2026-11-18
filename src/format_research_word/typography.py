"""Typography rules applied to editable Word text. Counts and units are context sensitive."""

import re
from dataclasses import dataclass, replace

from .audit import DECIMAL, IDENTIFIER, display_three

ELEMENTS = (
    "H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga "
    "Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd "
    "Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac "
    "Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr Rf Db Sg Bh Hs Mt Ds Rg Cn Nh Fl Mc Lv Ts Og"
).split()
ELEMENT = "(?:" + "|".join(sorted(ELEMENTS, key=lambda t: (-len(t), t))) + ")"
COUNT = r"(?:\d+(?:\.\d+)?(?:\(1[-−–]x\)|[-−–]x|x)?|x)"
ATOM = rf"{ELEMENT}(?:{COUNT})?"
GROUP = rf"(?:{ATOM}|\((?:{ATOM}|,\s*)+\)(?:{COUNT})?)"
FORMULA = re.compile(rf"(?<![a-zA-Z_])(?:{GROUP})+(?![a-zA-Z_0-9])")
COUNT_TOKEN = re.compile(COUNT)
SUB_CHARS = str.maketrans("₀₁₂₃₄₅₆₇₈₉ₓ₋₊₍₎", "0123456789x-+()")
# Protect versions, paths/IDs, and space-group symbols that look like quantities.
LITERAL = re.compile(
    r"(?<![A-Za-z0-9_])(?:F|I|P|R|C)[a-z]?[-−]?\d+[a-z][\w/-]*|(?<![\w/])[\w.-]+\.(?:py|xml|docx|json|csv)\b"
)
WS = r"[ \u00a0\u2009\u202f\u3000]"
CJK = "\u3400-\u9fff，。；：、！？（）《》“”"
LATIN = "a-zA-Z0-9µμÅ℃°%％"
BOUNDARY = re.compile(rf"(?<=[{CJK}]){WS}+(?=[{LATIN}])|(?<=[{LATIN}]){WS}+(?=[{CJK}])")
UNIT = re.compile(
    rf"(?<=[0-9]){WS}*(meV|eV|μm|µm|mm|nm|cm|km|m[eₑ]|THz|GHz|MHz|Hz|GPa|MPa|Pa|GB|MB|TB|Å|℃|°C|°|%|％|K|V|W|m)(?![A-Za-z0-9])"
)
MICRON = re.compile(rf"([μµ]){WS}+m\b|([μµ]){WS}+m(?=[{CJK}])")
ROMAN_X = re.compile(r"(?<![A-Za-z_])x(?![A-Za-z_])")
NON_FORMULAS = {
    "HSE",
    "HSE06",
    "PBE",
    "SOC",
    "DOS",
    "PAW",
    "FCC",
    "HCP",
    "FC2",
    "FC3",
    "ICSD",
    "III",
    "VIII",
    "SCF",
    "CPU",
    "NIST",
}


@dataclass(frozen=True)
class Char:
    text: str
    origin: int
    script: str = ""
    chemical: bool = False
    upright: bool = False


def joined(chars):
    return "".join(c.text for c in chars)


def exclusions(text):
    return [
        m.span() for pattern in (IDENTIFIER, LITERAL) for m in pattern.finditer(text)
    ]


def overlaps(a, b, spans):
    return any(a < end and b > start for start, end in spans)


def formulae(text):
    normal = text.translate(SUB_CHARS)
    for m in FORMULA.finditer(normal):
        term = m[0]
        if term in NON_FORMULAS or not re.search(r"[a-z]|\d", term):
            continue
        symbols = re.findall(ELEMENT, term)
        if len(symbols) < 2 and term not in {
            "H2",
            "N2",
            "O2",
            "F2",
            "Cl2",
            "Br2",
            "I2",
        }:
            continue
        yield m


def round_values(chars):
    text = joined(chars)
    protect = exclusions(text)
    for m in reversed(list(DECIMAL.finditer(text))):
        if overlaps(*m.span(), protect):
            continue
        result = display_three(m[0])
        if m[0].startswith("−"):
            result = result.replace("-", "−", 1)
        if m[0].startswith("+"):
            result = "+" + result
        base, _, exponent = result.partition("×10^(")
        head = chars[m.start()]
        replacement = [replace(head, text=c) for c in base]
        if exponent:
            replacement.extend(replace(head, text=c) for c in "×10")
            replacement.extend(
                replace(head, text=c, script="superscript")
                for c in exponent[:-1].replace("-", "−")
            )
        chars[m.start() : m.end()] = replacement


def typography(chars, structural=False):
    text = joined(chars)
    protect = exclusions(text)
    # Mark the original formula runs before changing the widths of decimal values.
    for m in formulae(text):
        a, b = m.span()
        if overlaps(a, b, protect):
            continue
        for i in range(a, b):
            chars[i] = replace(
                chars[i], text=chars[i].text.translate(SUB_CHARS), chemical=True
            )
        for k in COUNT_TOKEN.finditer(m[0]):
            for i in range(a + k.start(), a + k.end()):
                chars[i] = replace(chars[i], script="subscript", upright=True)
    round_values(chars)
    text = joined(chars)
    protected = exclusions(text)
    for m in ROMAN_X.finditer(text):
        if not overlaps(*m.span(), protected):
            chars[m.start()] = replace(chars[m.start()], upright=True)
    if structural:
        prefix = re.match(
            r"^(?:第[一二三四五六七八九十\d]+章|(?:图|表)?[A-E]?\d+(?:\.\d+)*|附录[A-E])\s+",
            text,
        )
        if prefix:
            protected.append(prefix.span())
    for m in reversed(list(BOUNDARY.finditer(text))):
        if not overlaps(*m.span(), protected):
            del chars[m.start() : m.end()]
    text = joined(chars)
    for m in reversed(list(MICRON.finditer(text))):
        del chars[m.start() + 1 : m.end() - 1]
    text = joined(chars)
    guards = exclusions(text)
    for m in reversed(list(UNIT.finditer(text))):
        a, b = m.span()
        if (
            overlaps(a, b, guards)
            or chars[a - 1].chemical
            or chars[m.start(1)].chemical
        ):
            continue
        chars[a : m.start(1)] = [
            replace(
                chars[m.start(1)], text=" ", script="", chemical=False, upright=False
            )
        ]
    return chars
