"""Lay out the English PDF in the author's original Times-style academic CV format."""
import calendar
import re

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

from cv_records import add_link, patent_role, patent_title, publication_details

FONT = "Times New Roman"
DATE_COLUMN, TEXT_COLUMN = 2320, 7426  # Original A4 table widths in twips.


def _style(doc, name, size, *, bold=False, before=0, after=0, line=15):
    styles = doc.styles
    style = styles[name] if name in styles else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string("000000")
    style.font.bold = bold
    font = style.element.get_or_add_rPr().rFonts
    font.set(qn("w:eastAsia"), "Noto Serif CJK KR")
    fmt = style.paragraph_format
    fmt.line_spacing = Pt(line)
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.keep_together = True
    return style


def _layout(doc):
    section = doc.sections[0]
    section.page_width, section.page_height = Twips(11906), Twips(16838)
    section.left_margin = section.right_margin = Twips(1080)
    section.top_margin = section.bottom_margin = Twips(1440)
    _style(doc, "Normal", 11)
    _style(doc, "CV title", 18, bold=True, after=11, line=23)
    _style(doc, "CV contact", 11, after=2, line=16)
    header = _style(doc, "CV section", 14, bold=True, before=21, after=3, line=17)
    header.paragraph_format.keep_with_next = True
    header.paragraph_format.first_line_indent = Pt(5.4)
    border = OxmlElement("w:pBdr")
    edge = OxmlElement("w:bottom")
    for key, value in {"val": "double", "sz": "4", "space": "1", "color": "000000"}.items():
        edge.set(qn("w:" + key), value)
    border.append(edge)
    header.element.get_or_add_pPr().append(border)
    _style(doc, "CV dated entry", 11, after=3, line=15)
    bib = _style(doc, "CV citation", 11, before=1, after=6, line=15)
    bib.paragraph_format.left_indent = Pt(5.4)
    sub = _style(doc, "CV subheading", 11, bold=True, before=5, after=3)
    sub.paragraph_format.left_indent = Pt(5.4)
    sub.paragraph_format.keep_with_next = True


def _period(value):
    """Convert numeric site dates to the month-year spelling in the English CV."""
    def month(match):
        year, number = match.groups()
        number = int(number)
        return f"{calendar.month_abbr[number]}. {year}" if 1 <= number <= 12 else match.group()
    value = re.sub(r"\b(\d{4})\.(\d{2})\b", month, str(value))
    if value.endswith("–"):
        value += "Present"
    return value.replace("–", " ~ ")


def _date_table(doc):
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for col, width in zip(table.columns, (DATE_COLUMN, TEXT_COLUMN)):
        col.width = Twips(width)
    return table


def _dated_row(table, date, date_note=None):
    row = table.add_row()
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for cell, width in zip(row.cells, (DATE_COLUMN, TEXT_COLUMN)):
        cell.width = Twips(width)
        cell.paragraphs[0].style = "CV dated entry"
    dates = row.cells[0].paragraphs[0]
    dates.add_run(_period(date)).bold = True
    if date_note:
        dates.add_run("\n(" + date_note + ")").italic = True
    return row.cells[1].paragraphs[0]


def _section(doc, title):
    doc.add_paragraph(title, style="CV section")


def _text(row, field):
    return row[field]["en"].replace("\n", ", ")


def _title_and_details(paragraph, row):
    paragraph.add_run(_text(row, "title")).bold = True
    paragraph.add_run(", " + _text(row, "details"))
    for duty in row.get("duties", []):
        paragraph.add_run("\n   ·  ")
        role, separator, scope = duty["en"].partition(", ")
        paragraph.add_run(role).bold = True
        paragraph.add_run(separator + scope)


def _authors(paragraph, authors):
    # Preserve names supplied by the author, including author ordering and compound surnames.
    for fragment in re.split(r"(\b(?:Suncheol|Sun-Cheol|S\.) Kwon\b)", authors, flags=re.I):
        if fragment:
            run = paragraph.add_run(fragment)
            if re.fullmatch(r"(?:Suncheol|Sun-Cheol|S\.) Kwon", fragment, flags=re.I):
                run.bold = True


def _publication(doc, row, number, group, supplement):
    pub = publication_details(row, supplement, group)
    p = doc.add_paragraph(style="CV citation")
    p.add_run(f"[{number:02d}] ")
    if pub["authors"]:
        _authors(p, pub["authors"])
        p.add_run(", ")
    p.add_run(row["title"].strip())
    for part in (pub["venue"], pub["locator"], pub["year"]):
        if part:
            p.add_run(", " + part)
    p.add_run(".")
    notes = []
    if group == "submitted":
        notes.append("submitted")
    if "Early Access" in pub["badges"]:
        notes.append("Early Access")
    if pub["co_corresponding"]:
        notes.append("co-corresponding author")
    elif pub["corresponding"]:
        notes.append("corresponding author")
    if "1st author" in pub["badges"]:
        notes.append("first author")
    notes.extend(b for b in pub["badges"] if "award" in b.lower())
    if notes:
        p.add_run(" (" + ", ".join(notes) + ")")
    if pub["doi"]:
        p.add_run("  ")
        add_link(p, "https://doi.org/" + pub["doi"], "DOI", font=FONT, size=10, color="0000CC", underline=True)


def render(document_data, filename):
    profile = document_data["profile"]
    supplement = document_data["supplement"]
    author = document_data["author"]
    doc = Document()
    _layout(doc)
    name = profile["name"]["en"]
    doc.core_properties.title = "Curriculum Vitae — " + name
    doc.core_properties.author = name
    doc.core_properties.language = "en"
    p = doc.add_paragraph(style="CV title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Curriculum Vitae").underline = True
    p = doc.add_paragraph(name + ", Ph.D.", style="CV title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    site_url = document_data.get("site_url", "https://suncheolkwon.github.io")
    for label, url, display in [
        ("Email", "mailto:" + author["email"], author["email"]),
        ("Website", site_url, site_url.removeprefix("https://").rstrip("/")),
        ("ORCID", author.get("orcid"), str(author.get("orcid", "")).rstrip("/").rsplit("/", 1)[-1]),
    ]:
        if url:
            p = doc.add_paragraph(style="CV contact")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(label + ":  ")
            add_link(p, url, display, font=FONT, size=11, color="0000CC", underline=True)

    _section(doc, "Position")
    table = _date_table(doc)
    for row in profile["positions"]:
        note = row.get("period_label", {}).get("en")
        _title_and_details(_dated_row(table, row["date"], note), row)

    _section(doc, "Major research interest")
    p = doc.add_paragraph(style="CV citation")
    p.add_run(", ".join(interest["en"] for interest in profile["research_interests"]))

    _section(doc, "Education")
    table = _date_table(doc)
    for row in sorted(profile["education"] + profile["visiting"], key=lambda r: str(r["date"]), reverse=True):
        p = _dated_row(table, row["date"])
        p.add_run(_text(row, "title") + ", " + _text(row, "details")).bold = True
        for field, label in (("thesis", "Thesis"), ("advisor", "Advisor")):
            if row.get(field):
                p.add_run("\n   " + label + ": " + _text(row, field))

    _section(doc, "Honors & Awards")
    table = _date_table(doc)
    for row in profile["honors"]:
        _title_and_details(_dated_row(table, row["date"]), row)

    for group, title in (("journals", "International Journal Articles"),
                         ("submitted", "Submitted Manuscripts"),
                         ("conferences", "International Conference Proceedings (including presentations)")):
        rows = document_data["publications"][group]
        if rows:
            _section(doc, title)
            for i, row in enumerate(rows, 1):
                _publication(doc, row, i, group, supplement)

    _section(doc, "Patents")
    doc.add_paragraph("(Granted)", style="CV subheading")
    grants = []
    for region in ("international", "korean"):
        grants.extend(sorted(document_data["patents"][region], key=lambda r: str(r["date"]), reverse=True))
    for i, row in enumerate(grants, 1):
        p = doc.add_paragraph(style="CV citation")
        p.add_run(f"[{i:02d}] {row['country']} {row['number']}, " + patent_title(row, supplement, "en"))
        p.add_run(", " + str(row["date"]) + ".")
        role = patent_role(row, "en")
        if role:
            p.add_run(" (" + role + ")")
    doc.add_paragraph("(Applications)", style="CV subheading")
    for i, row in enumerate(supplement["patent_applications"], 1):
        doc.add_paragraph(f"[{i:02d}] {row['country']} {row['number']}, {row['title']}.", style="CV citation")

    for key, title in (("technology", "Medical Device Development"), ("service", "Professional Service"),
                       ("funding", "Research and Program Funding"), ("domestic_presentations", "Domestic Conference Presentations")):
        _section(doc, title)
        table = _date_table(doc)
        for row in sorted(supplement[key], key=lambda r: str(r["date"]), reverse=True):
            p = _dated_row(table, row["date"])
            if key == "funding":
                p.add_run(_text(row, "title")).bold = True
                p.add_run(", " + _text(row, "sponsor") + f", KRW {row['krw']:,}")
                if row.get("number"):
                    p.add_run(" (" + row["number"] + ")")
            else:
                _title_and_details(p, row)
    _section(doc, "Book Translation")
    row = supplement["translation"]
    p = _dated_row(_date_table(doc), row["date"])
    p.add_run(_text(row, "title")).bold = True
    p.add_run(", " + _text(row, "details") + "\n" + row["contributors"])
    doc.save(filename)
