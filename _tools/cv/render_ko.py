"""Korean CV with the hierarchy, date columns and numbered lists of Kwon's Word CV."""
import re
from datetime import date

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from cv_records import add_link, patent_role, patent_title, publication_details


FONT = "Noto Sans CJK KR"
SELF_NAME = re.compile(r"(?<!\w)(?:Suncheol Kwon|S\.\s*Kwon|Kwon S)(?!\w)", re.I)


def period(value):
    """Translate the stored ISO/year-month format into dotted Korean dates."""
    text = str(value).strip()
    text = re.sub(r"(\d{4})\.(\d{2})", lambda m: f"{m[1]}. {int(m[2])}.", text)
    if text.endswith("–"):
        text += "현재"
    return text.replace("–", " – ")


def full_date(value):
    return date.fromisoformat(str(value)).strftime("%Y. %m. %d.")


def setup(doc):
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.left_margin = section.right_margin = section.bottom_margin = Pt(72)
    section.top_margin = Pt(85)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string("000000")
    fonts = normal.element.get_or_add_rPr().rFonts
    fonts.set(qn("w:eastAsia"), FONT)
    fmt = normal.paragraph_format
    fmt.line_spacing = Pt(16.3)
    fmt.space_before = fmt.space_after = Pt(0)
    fmt.widow_control = True

    for name, size, before, after in (
        ("KO Title", 14, 0, 25), ("KO Byline", 12, 0, 38),
        ("KO Section", 12, 17, 9), ("KO Subheading", 12, 14, 3),
        ("KO Minor heading", 10, 8, 1),
    ):
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = normal
        style.font.bold = True
        style.font.size = Pt(size)
        fmt = style.paragraph_format
        fmt.space_before, fmt.space_after = Pt(before), Pt(after)
        fmt.line_spacing = Pt(size * 1.55)
        fmt.keep_with_next = True
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", "000000")):
        bottom.set(qn("w:" + key), value)
    border.append(bottom)
    doc.styles["KO Section"].element.get_or_add_pPr().append(border)

    style = doc.styles.add_style("KO Citation", WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = normal
    fmt = style.paragraph_format
    fmt.left_indent, fmt.first_line_indent = Pt(40), Pt(-20)
    fmt.tab_stops.add_tab_stop(Pt(40))
    fmt.keep_together = True


def dated(doc, when, text, *, tab=120, lines=(), gap=0):
    p = doc.add_paragraph()
    fmt = p.paragraph_format
    fmt.left_indent, fmt.first_line_indent = Pt(tab), Pt(-tab)
    fmt.tab_stops.add_tab_stop(Pt(tab))
    fmt.space_before = Pt(gap)
    fmt.keep_together = True
    p.add_run(f"{when}\t{text}")
    for line in lines:
        p.add_run("\n" + line)
    return p


def own_name_bold(p, text):
    start = 0
    for match in SELF_NAME.finditer(text):
        p.add_run(text[start:match.start()])
        p.add_run(match.group()).bold = True
        start = match.end()
    p.add_run(text[start:])


def number(doc, index):
    p = doc.add_paragraph(style="KO Citation")
    p.add_run(f"{index}.\t")
    return p


def publications(doc, data, group):
    for index, entry in enumerate(data["publications"].get(group, []), 1):
        details = publication_details(entry, data["supplement"], group)
        p = number(doc, index)
        if details["authors"]:
            own_name_bold(p, details["authors"].rstrip(" .,") + ", ")
        title = entry["title"].strip()
        if details["doi"]:
            add_link(p, "https://doi.org/" + details["doi"], title, font=FONT, size=10)
        else:
            p.add_run(title)
        info = [s for s in (details["venue"], details["year"], details["locator"]) if s]
        if info:
            p.add_run(", " + ", ".join(info))
        notes = []
        if group == "submitted":
            notes.append("투고")
        if "Early Access" in details["badges"]:
            notes.append("Early Access")
        if "1st author" in details["badges"]:
            notes.append("제1저자")
        if details["co_corresponding"]:
            notes.append("공동교신저자")
        elif details["corresponding"]:
            notes.append("교신저자")
        notes += [s for s in details["badges"] if "award" in s.lower()]
        if notes:
            p.add_run(" (" + ", ".join(notes) + ")")
        p.add_run(".")


def profile_section(doc, data):
    profile, author = data["profile"], data["author"]
    doc.add_paragraph("기본 정보 (Personal information)", "KO Section")
    sub = doc.add_paragraph("현재 소속 및 직위", "KO Subheading")
    sub.paragraph_format.space_before = Pt(0)
    for row in profile["positions"]:
        if str(row["date"]).endswith("–"):
            when = period(row["date"])
            if row.get("period_label"):
                when += " 재직"
            dated(doc, when, row["details"]["ko"].replace(" · ", " ") + " " + row["title"]["ko"])

    doc.add_paragraph("개인 정보", "KO Subheading")
    for label, url, display in [
        ("이메일", "mailto:" + author["email"], author["email"]),
        ("홈페이지", data.get("site_url") or "https://suncheolkwon.github.io/", None),
        ("ORCID", author.get("orcid"), None),
    ]:
        if url:
            p = dated(doc, label, "")
            add_link(p, url, display or url, font=FONT, size=10)

    doc.add_paragraph("학력", "KO Subheading")
    education = sorted(profile["education"] + profile["visiting"],
                       key=lambda row: str(row["date"]).split("–")[-1], reverse=True)
    for row in education:
        when = period(row["date"])
        if row.get("start"):
            when = f'{period(row["start"])} – {when}'
        lines = []
        if row.get("thesis"):
            lines.append("학위논문: " + row["thesis"]["ko"])
        if row.get("advisor"):
            advisor = "지도교수: " + row["advisor"]["ko"]
            if lines:
                lines[0] += " (" + advisor + ")"
            else:
                lines.append("(" + advisor + ")")
        dated(doc, when, row["title"]["ko"] + ", " + row["details"]["ko"], tab=113, lines=lines)

    doc.add_paragraph("전문적 경력", "KO Subheading")
    for row in profile["positions"]:
        if not str(row["date"]).endswith("–"):
            dated(doc, period(row["date"]), row["title"]["ko"] + ", " + row["details"]["ko"],
                  lines=[v["ko"] for v in row.get("duties", [])])


def output_section(doc, data):
    supplement = data["supplement"]
    heading = doc.add_paragraph("연구 성과 (Research outcomes)", "KO Section")
    heading.paragraph_format.page_break_before = True
    doc.add_paragraph("저술 (Publications)", "KO Subheading")
    for label, group in (("국제 학술지", "journals"), ("투고 논문", "submitted"), ("국제 학술대회", "conferences")):
        if data["publications"].get(group):
            doc.add_paragraph(label, "KO Minor heading")
            publications(doc, data, group)
    if supplement["domestic_presentations"]:
        doc.add_paragraph("국내 학술대회", "KO Minor heading")
        for i, row in enumerate(supplement["domestic_presentations"], 1):
            number(doc, i).add_run(f'{row["title"]["ko"]}, {row["details"]["ko"]}, {row["date"]}.')
    doc.add_paragraph("책(역서)", "KO Subheading")
    row = supplement["translation"]
    p = number(doc, 1)
    own_name_bold(p, row["contributors"])
    p.add_run(f', {row["title"]["ko"]}, {row["details"]["ko"].rstrip(".")}, {row["date"]}.')

    doc.add_paragraph("특허", "KO Subheading")
    for label, group in (("해외 등록", "international"), ("국내 등록", "korean")):
        doc.add_paragraph(label, "KO Minor heading")
        for i, row in enumerate(sorted(data["patents"][group], key=lambda r: str(r["date"]), reverse=True), 1):
            p = number(doc, i)
            p.add_run(f'{full_date(row["date"])}  {patent_title(row, supplement, "ko")} '
                      f'({row["country"]} {row["number"]}, {patent_role(row, "ko")})')
    doc.add_paragraph("해외 출원 이력", "KO Minor heading")
    for i, row in enumerate(supplement["patent_applications"], 1):
        number(doc, i).add_run(f'{row["title"]} ({row["country"]} {row["number"]})')


def activities(doc, data):
    supplement, profile = data["supplement"], data["profile"]
    heading = doc.add_paragraph("연구비 (Funding)", "KO Section")
    note = heading.add_run("    *연구/사업책임자인 과제를 기재")
    note.font.size = Pt(9)
    note.bold = False
    for row in supplement["funding"]:
        lines = ["(연구비 지원기관) " + row["sponsor"]["ko"]]
        if row.get("number"):
            lines.append("(과제 고유번호) " + row["number"])
        lines.append(f'(연구비) {row["krw"]:,}원')
        dated(doc, period(row["date"]), "(과제명) " + row["title"]["ko"], tab=110, lines=lines, gap=4)

    doc.add_paragraph("연구 활동 (Research activities)", "KO Section")
    doc.add_paragraph("주요 연구분야 및 관심사", "KO Subheading")
    if supplement.get("research_statement_ko"):
        doc.add_paragraph(supplement["research_statement_ko"])
    else:
        doc.add_paragraph(", ".join(row["ko"] for row in profile["research_interests"]))
    doc.add_paragraph("의료기기, 소프트웨어, 앱 개발", "KO Subheading")
    for row in supplement["technology"]:
        dated(doc, str(row["date"]), row["title"]["ko"], tab=80, gap=3,
              lines=row.get("original_details_ko") or [row["details"]["ko"]])

    doc.add_paragraph("전문적 활동 및 리더십 (Professional activities and Leadership)", "KO Section")
    for row in supplement["service"]:
        dated(doc, period(row["date"]), row["title"]["ko"] + ", " + row["details"]["ko"], tab=110)
    doc.add_paragraph("수상 (Honors and awards)", "KO Section")
    for row in profile["honors"]:
        dated(doc, period(row["date"]), row["title"]["ko"] + ", " + row["details"]["ko"], tab=110)


def render(data, filename):
    doc = Document()
    setup(doc)
    name = data["profile"]["name"]["ko"]
    doc.core_properties.title = name + " 이력서"
    doc.core_properties.author = name
    doc.core_properties.language = "ko-KR"
    doc.add_paragraph("이력서 (Curriculum Vitae)", "KO Title").alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(style="KO Byline")
    p.paragraph_format.tab_stops.add_tab_stop(Pt(451), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run("이름:  " + " ".join(name))
    if data.get("as_of"):
        p.add_run("\t최종 수정일:  " + full_date(data["as_of"]))
    profile_section(doc, data)
    output_section(doc, data)
    activities(doc, data)
    doc.save(filename)
