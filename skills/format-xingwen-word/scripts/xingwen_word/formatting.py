"""Word formatting helpers for the 行文检查版式 formatter."""

from __future__ import annotations

from typing import Any

try:
    from docx.enum.section import WD_SECTION_START
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor
except ModuleNotFoundError:
    WD_SECTION_START = None
    WD_ALIGN_PARAGRAPH = None
    WD_LINE_SPACING = None
    WD_TAB_ALIGNMENT = None
    WD_TAB_LEADER = None
    OxmlElement = None
    qn = None
    Mm = None
    Pt = None
    RGBColor = None


def safe_xpath(element: Any, expression: str) -> list[Any]:
    if element is None:
        return []
    try:
        return list(element.xpath(expression))
    except Exception:
        return []


def safe_xpath_count(element: Any, expression: str) -> int:
    return len(safe_xpath(element, expression))


def xml_child_val(element: Any, path: str) -> str | None:
    matches = safe_xpath(element, path)
    if not matches:
        return None
    try:
        return matches[0].get(qn("w:val"))
    except Exception:
        return None


def first_child(element: Any, path: str) -> Any | None:
    matches = safe_xpath(element, path)
    return matches[0] if matches else None


def get_or_add_child(element: Any, child_tag: str) -> Any:
    existing = first_child(element, f"./{child_tag}")
    if existing is not None:
        return existing
    child = OxmlElement(child_tag)
    element.append(child)
    return child


ALIGNMENTS = (
    {}
    if WD_ALIGN_PARAGRAPH is None
    else {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
)

TAB_ALIGNMENTS = (
    {}
    if WD_TAB_ALIGNMENT is None
    else {
        "LEFT": WD_TAB_ALIGNMENT.LEFT,
        "CENTER": WD_TAB_ALIGNMENT.CENTER,
        "RIGHT": WD_TAB_ALIGNMENT.RIGHT,
        "DECIMAL": WD_TAB_ALIGNMENT.DECIMAL,
        "BAR": WD_TAB_ALIGNMENT.BAR,
        "LIST": WD_TAB_ALIGNMENT.LIST,
        "CLEAR": WD_TAB_ALIGNMENT.CLEAR,
    }
)

TAB_LEADERS = (
    {}
    if WD_TAB_LEADER is None
    else {
        "SPACES": WD_TAB_LEADER.SPACES,
        "DOTS": WD_TAB_LEADER.DOTS,
        "DASHES": WD_TAB_LEADER.DASHES,
        "LINES": WD_TAB_LEADER.LINES,
        "HEAVY": WD_TAB_LEADER.HEAVY,
        "MIDDLE_DOT": WD_TAB_LEADER.MIDDLE_DOT,
    }
)


def set_east_asia_font(run: Any, font_name: str, latin_font_name: str | None = None) -> None:
    latin_font = latin_font_name or font_name
    run.font.name = latin_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), latin_font)
    rfonts.set(qn("w:hAnsi"), latin_font)


def set_outline_level(paragraph: Any, level: Any) -> None:
    ppr = paragraph._element.get_or_add_pPr()
    outline = get_or_add_child(ppr, "w:outlineLvl")
    outline.set(qn("w:val"), str(level))


def apply_tab_stops(paragraph: Any, tab_stops: list[dict[str, Any]]) -> None:
    try:
        stops = paragraph.paragraph_format.tab_stops
        stops.clear_all()
        for tab in tab_stops:
            position = tab.get("position_pt")
            if position is None:
                continue
            alignment = TAB_ALIGNMENTS.get(str(tab.get("alignment") or "LEFT"))
            leader = TAB_LEADERS.get(str(tab.get("leader") or "SPACES"))
            kwargs = {}
            if alignment is not None:
                kwargs["alignment"] = alignment
            if leader is not None:
                kwargs["leader"] = leader
            stops.add_tab_stop(Pt(float(position)), **kwargs)
    except Exception:
        return


def apply_style(paragraph: Any, style: dict[str, Any]) -> None:
    fmt = paragraph.paragraph_format
    align = style.get("align")
    if align in ALIGNMENTS:
        paragraph.alignment = ALIGNMENTS[align]
    if "left_indent" in style:
        fmt.left_indent = Pt(style["left_indent"])
    if "right_indent" in style:
        fmt.right_indent = Pt(style["right_indent"])
    if "first_indent" in style:
        fmt.first_line_indent = Pt(style["first_indent"])
    if "space_before" in style:
        fmt.space_before = Pt(style["space_before"])
    else:
        fmt.space_before = Pt(0)
    if "space_after" in style:
        fmt.space_after = Pt(style["space_after"])
    else:
        fmt.space_after = Pt(0)
    if "line" in style and style["line"]:
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Pt(style["line"])
    if "keep_together" in style:
        fmt.keep_together = bool(style["keep_together"])
    if "keep_with_next" in style:
        fmt.keep_with_next = bool(style["keep_with_next"])
    if "page_break_before" in style:
        fmt.page_break_before = bool(style["page_break_before"])
    if "widow_control" in style:
        fmt.widow_control = bool(style["widow_control"])
    if "outline_level" in style:
        set_outline_level(paragraph, style["outline_level"])
    if style.get("tab_stops"):
        apply_tab_stops(paragraph, style["tab_stops"])

    font = style.get("font", "仿宋_GB2312")
    latin_font = style.get("latin_font")
    size = style.get("size", 16)
    bold = style.get("bold")
    italic = style.get("italic", False)
    underline = style.get("underline", False)
    color = style.get("color", "000000")
    for run in paragraph.runs:
        set_east_asia_font(run, font, latin_font)
        run.font.size = Pt(size)
        if bold is not None:
            run.font.bold = bool(bold)
        if italic is not None:
            run.font.italic = bool(italic)
        if underline is not None:
            run.font.underline = bool(underline)
        if color:
            try:
                run.font.color.rgb = RGBColor.from_string(str(color).replace("#", ""))
            except ValueError:
                pass


def apply_document_grid(section: Any, page_style: dict[str, Any]) -> None:
    if not any(key in page_style for key in ["grid_type", "line_pitch_twips", "char_space_twips"]):
        return
    sect_pr = section._sectPr
    nodes = safe_xpath(sect_pr, "./w:docGrid")
    doc_grid = nodes[0] if nodes else OxmlElement("w:docGrid")
    if not nodes:
        sect_pr.append(doc_grid)
    if "grid_type" in page_style:
        doc_grid.set(qn("w:type"), str(page_style["grid_type"]))
    if "line_pitch_twips" in page_style:
        doc_grid.set(qn("w:linePitch"), str(page_style["line_pitch_twips"]))
    if "char_space_twips" in page_style:
        doc_grid.set(qn("w:charSpace"), str(page_style["char_space_twips"]))


def apply_page_setup(document: Any, page_style: dict[str, Any]) -> None:
    for section in document.sections:
        section.start_type = WD_SECTION_START.NEW_PAGE
        if "width_mm" in page_style:
            section.page_width = Mm(page_style["width_mm"])
        if "height_mm" in page_style:
            section.page_height = Mm(page_style["height_mm"])
        if "top_mm" in page_style:
            section.top_margin = Mm(page_style["top_mm"])
        if "bottom_mm" in page_style:
            section.bottom_margin = Mm(page_style["bottom_mm"])
        if "left_mm" in page_style:
            section.left_margin = Mm(page_style["left_mm"])
        if "right_mm" in page_style:
            section.right_margin = Mm(page_style["right_mm"])
        apply_document_grid(section, page_style)
