"""Generate and paginate the report's editable contents field using desktop Word."""

import os
import re
import shutil
import tempfile
from pathlib import Path

CUSTOM_HEADINGS = ("一级标题xx", "二级标题xx", "三级标题xx")
# WdBuiltinStyle: Heading 1..3 and TOC 1..3 have stable, language-independent IDs.
WORD_HEADINGS = (-2, -3, -4)
TOC_STYLES = (-20, -21, -22)
TITLE = re.compile(r"^目\s*录$")
STATIC_ENTRY = re.compile(r".+(?:\t|[.．…·]{2,})\s*(?:\d+|[IVXLCDM]+)$", re.I)
OTHER_TOC = re.compile(r"\\[caf](?:\s|$)", re.I)


def first_heading(doc, com_error):
    candidates = []
    style_map = []
    native = False
    for level, (custom, builtin) in enumerate(zip(CUSTOM_HEADINGS, WORD_HEADINGS), 1):
        for key in (custom, builtin):
            try:
                style = doc.Styles(key)
            except com_error:
                continue
            if key == custom:
                style_map.extend((custom, str(level)))
            rng = doc.Content.Duplicate
            find = rng.Find
            find.ClearFormatting()
            find.Style = style
            while find.Execute(FindText="", Format=True, Forward=True, Wrap=0):
                paragraph = rng.Paragraphs(1).Range
                if TITLE.fullmatch(paragraph.Text.strip()):
                    if paragraph.End >= doc.Content.End:
                        break
                    rng.SetRange(paragraph.End, doc.Content.End)
                    continue
                candidates.append(paragraph.Start)
                if key == builtin:
                    native = True
                break
    if not candidates:
        raise ValueError("请先给正文应用一级至三级的项目标题样式或 Word 内置标题样式。")
    switches = [r"\h", r"\z"]
    if style_map:
        switches.append(r'\t "' + ",".join(style_map) + '"')
    if native:
        switches.append(r'\o "1-3"')
    return min(candidates), " TOC " + " ".join(switches) + " "


def format_title(para):
    pf = para.Format
    pf.OutlineLevel = 10  # wdOutlineLevelBodyText
    pf.Alignment = 1
    pf.CharacterUnitLeftIndent = 0
    pf.CharacterUnitFirstLineIndent = 0
    pf.SpaceBefore = 0
    pf.SpaceAfter = 8
    pf.LineSpacingRule = 5
    pf.LineSpacing = 13.9
    font = para.Range.Font
    font.NameAscii = font.NameOther = "Times New Roman"
    font.NameFarEast = "黑体"
    font.Size = 16
    font.Bold = font.Italic = False


def format_entries(doc):
    for level, key in enumerate(TOC_STYLES):
        style = doc.Styles(key)
        style.AutomaticallyUpdate = False
        font = style.Font
        font.NameAscii = font.NameOther = font.NameFarEast = "等线"
        font.Size = 11
        font.Bold = font.Italic = False
        pf = style.ParagraphFormat
        pf.Alignment = 1 if level == 0 else 0
        pf.CharacterUnitFirstLineIndent = 0
        pf.CharacterUnitLeftIndent = 2 * level
        pf.SpaceBefore = 0
        pf.SpaceAfter = 8
        pf.LineSpacingRule = 5
        pf.LineSpacing = 13.9


def preceding_title(doc, anchor):
    if anchor == 0:
        return None
    prior = doc.Range(0, anchor)
    # Search only the front matter; an abstract or cover may precede the contents.
    find = prior.Find
    find.ClearFormatting()
    for text in ("目录", "目 录", "目　录"):
        prior.SetRange(0, anchor)
        if find.Execute(FindText=text, Format=False, Forward=False, Wrap=0):
            para = prior.Paragraphs(1)
            if TITLE.fullmatch(para.Range.Text.strip()):
                return para
    return None


def insert_title(doc, pos):
    # Insert a separate paragraph before the heading/field, retaining the following style.
    doc.Range(pos, pos).InsertBefore("目录\r")
    title = doc.Range(pos, pos + 2).Paragraphs(1)
    title.Style = doc.Styles(-1)  # Normal
    title.Format.PageBreakBefore = pos > 0
    return title


def remove_static_entries(doc, title, heading):
    """Only consume identifiable, plain TOC lines. Preserve following front matter."""
    start = title.Range.End
    end = start
    for para in doc.Range(start, heading).Paragraphs:
        rng = para.Range
        if (
            rng.End > heading
            or rng.Tables.Count
            or rng.InlineShapes.Count
            or rng.Fields.Count
        ):
            break
        text = rng.Text.rstrip("\r")
        if any(c in text for c in ("\x07", "\x0c")):
            break
        if not text.strip():
            continue
        if not STATIC_ENTRY.fullmatch(text.strip()):
            break
        end = rng.End
    if end > start:
        doc.Range(start, end).Delete()


def generate(doc, com_error):
    heading, code = first_heading(doc, com_error)
    fields = [
        f for f in doc.Fields if f.Type == 13 and not OTHER_TOC.search(f.Code.Text)
    ]
    if len(fields) > 1 or any(
        re.search(r"\\b(?:\s|$)", f.Code.Text, re.I) for f in fields
    ):
        raise ValueError(
            "此稿含多个章节目录或分范围目录，请按工作流分别核对各目录范围。"
        )
    format_entries(doc)
    anchor = fields[0].Code.Start - 1 if fields else heading
    title = preceding_title(doc, anchor)
    if fields:
        field = fields[0]
        if title is None:
            title = insert_title(doc, anchor)
        field.Code.Text = code
    else:
        if title is None:
            title = insert_title(doc, heading)
            doc.Range(title.Range.End, title.Range.End).Paragraphs(
                1
            ).Format.PageBreakBefore = True
        else:
            remove_static_entries(doc, title, heading)
        pos = title.Range.End
        doc.Range(pos, pos).InsertBefore("\r")
        slot = doc.Range(pos, pos)
        slot.Paragraphs(1).Style = doc.Styles(-1)
        slot.Paragraphs(1).Format.PageBreakBefore = False
        field = doc.Fields.Add(
            Range=slot, Type=-1, Text=code.strip(), PreserveFormatting=False
        )
    if title.Format.OutlineLevel <= 3:
        title.Style = doc.Styles(-1)
    format_title(title)
    field.Locked = False
    field.Update()
    # A document may contain a list of figures in addition to its chapter contents.
    contents = next(
        t
        for t in doc.TablesOfContents
        if t.Range.Start <= field.Result.Start <= t.Range.End
    )
    contents.RightAlignPageNumbers = True
    contents.TabLeader = 1  # wdTabLeaderDots
    contents.Update()
    previous = None
    for _ in range(5):
        doc.Repaginate()
        contents.UpdatePageNumbers()
        current = contents.Range.Text
        if current == previous:
            break
        previous = current
    else:
        raise ValueError("目录页码尚未稳定，请在 Word 中检查分页。")
    if not any("\t" in line for line in current.split("\r")):
        raise ValueError("目录中没有带页码的条目，请核对正文的标题样式。")
    return current


def generate_toc(source, destination):
    source, destination = Path(source).resolve(), Path(destination).resolve()
    if source.suffix.lower() != ".docx" or destination.suffix.lower() != ".docx":
        raise ValueError("输入、输出均需使用 .docx。")
    if destination.exists():
        raise FileExistsError(f"请使用新的输出文件名：{destination}")
    if os.name != "nt":
        raise ValueError("自动目录排页需要 Windows 桌面版 Microsoft Word。")
    try:
        from pywintypes import com_error
        from win32com.client import DispatchEx
    except ImportError as e:
        raise ValueError(
            '请安装 Word 自动化依赖：python -m pip install ".[word]"'
        ) from e

    with tempfile.TemporaryDirectory(dir=destination.parent) as temp:
        staged = Path(temp) / "report.docx"
        shutil.copyfile(source, staged)
        try:
            app = DispatchEx("Word.Application")
            update_links = app.Options.UpdateLinksAtOpen
            try:
                app.Visible = False
                app.DisplayAlerts = 0
                app.AutomationSecurity = 3
                app.Options.UpdateLinksAtOpen = False

                def open_copy(readonly):
                    return app.Documents.Open(
                        str(staged),
                        ReadOnly=readonly,
                        Visible=False,
                        AddToRecentFiles=False,
                    )

                doc = open_copy(False)
                try:
                    result = generate(doc, com_error)
                    doc.Save()
                finally:
                    doc.Close(0)
                doc = open_copy(True)
                try:
                    contents = next(
                        (t for t in doc.TablesOfContents if t.Range.Text == result),
                        None,
                    )
                    if contents is None:
                        raise ValueError(
                            "重新打开成稿后，目录条目或页码与保存前不一致。"
                        )
                    rows = [line for line in result.split("\r") if "\t" in line]
                    for index in {1, len(rows)}:
                        entry = contents.Range.Paragraphs(index).Range
                        target = doc.Bookmarks(entry.Hyperlinks(1).SubAddress).Range
                        target.Collapse(1)
                        printed = rows[index - 1].rsplit("\t", 1)[-1].strip()
                        if printed.isdecimal() and int(printed) != target.Information(
                            1
                        ):
                            raise ValueError(f"目录页码与正文不一致：{rows[index - 1]}")
                finally:
                    doc.Close(0)
            finally:
                try:
                    app.Options.UpdateLinksAtOpen = update_links
                finally:
                    app.Quit(0)
        except com_error as e:
            raise ValueError(f"Word 无法生成目录：{e}") from e
        os.rename(
            staged, destination
        )  # Windows refuses an existing destination atomically.
    return len(rows)
