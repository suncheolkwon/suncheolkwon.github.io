"""Read the public website's CV records for both language editions."""
from datetime import datetime
from pathlib import Path
import re
import subprocess
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import yaml


ROOT = Path(__file__).resolve().parents[2]
SEOUL = ZoneInfo("Asia/Seoul")
PROFILE = "_data/cv_profile.yml"
SUPPLEMENT = "_data/cv_supplement.yml"
PATENTS = "_data/patents.yml"


def as_of_date(root):
    today = datetime.now(SEOUL).date()
    if not (root / ".git").exists():
        return today.isoformat()
    inputs = ["_publications", PATENTS, PROFILE, SUPPLEMENT, "_config.yml", "_tools/cv"]
    status = subprocess.run(["git", "status", "--porcelain", "--", *inputs],
                            cwd=root, check=True, capture_output=True, text=True)
    if status.stdout.strip():
        return today.isoformat()
    timestamp = subprocess.check_output(["git", "log", "-1", "--format=%ct", "--", *inputs],
                                        cwd=root, text=True).strip()
    return datetime.fromtimestamp(int(timestamp), SEOUL).date().isoformat() if timestamp else today.isoformat()


def load_data(source=ROOT):
    root = Path(source)

    def read(relative):
        return yaml.safe_load((root / relative).read_text(encoding="utf-8"))

    groups = {"journals": [], "conferences": [], "submitted": []}
    categories = {"manuscripts": "journals", "conferences": "conferences"}
    for path in sorted((root / "_publications").rglob("*")):
        if path.suffix.lower() not in {".md", ".markdown", ".html"} or not path.is_file():
            continue
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        if not lines or lines[0].strip() != "---":
            raise ValueError(f"Missing publication front matter in {path}")
        try:
            end = lines.index("---", 1)
        except ValueError as error:
            raise ValueError(f"Unclosed publication front matter in {path}") from error
        row = yaml.safe_load("\n".join(lines[1:end]))
        if row.get("published") is False:
            continue
        badges = row.get("badges") or []
        row["badges"] = [badges] if isinstance(badges, str) else badges
        submitted = "Submitted" in row["badges"]
        group = "submitted" if submitted else categories[row["category"]]
        groups[group].append(row)
    for rows in groups.values():
        rows.sort(key=lambda row: (str(row.get("date", "")), row["title"]), reverse=True)
    return {"profile": read(PROFILE), "supplement": read(SUPPLEMENT), "publications": groups,
            "patents": read(PATENTS), "author": read("_config.yml")["author"], "as_of": as_of_date(root)}


def publication_details(row, supplement, group):
    doi = str(row.get("doi") or row.get("DOI") or "").strip()
    if not doi:
        link = str(row.get("paperurl") or "").strip()
        if link.startswith("https://doi.org/"):
            doi = link
    doi = doi.removeprefix("https://doi.org/").removeprefix("http://dx.doi.org/")
    if not re.fullmatch(r"10\.\d{4,9}/\S+", doi):
        doi = ""
    authors = row.get("authors_text") or row.get("citation") or ""
    venue = (row.get("venue") or "").strip()
    data = dict(row)
    if not data.get("article_number"):
        data["article_number"] = data.get("article-number")
    fields = (("volume", "vol."), ("number", "no."), ("article_number", "art."), ("pages", "pp."))
    locator = ", ".join(f"{prefix} {data[key]}" for key, prefix in fields if data.get(key))
    year = str(row.get("year") or row.get("date") or "")[:4]
    return {"authors": authors.removeprefix("Recommended citation:").strip(),
            "venue": "" if venue == "N/A" else venue, "year": "" if group == "submitted" else year,
            "doi": doi, "locator": locator, "badges": row.get("badges") or [],
            "corresponding": bool(row.get("corresponding")), "co_corresponding": bool(row.get("co_corresponding"))}


def patent_title(row, supplement, language):
    if language == "ko":
        return row.get("title_ko") or supplement.get("korean_patent_titles", {}).get(row["number"], row["title"])
    return row["title"]


def patent_role(row, language):
    roles = {"lead": ("Lead inventor", "대표발명자"), "included": ("Co-inventor", "공동발명자")}
    if row.get("role") == "two":
        return row.get("inventors", "")
    return roles.get(row.get("role"), ("", ""))[language == "ko"]


def add_link(paragraph, url, text, *, font=None, size=None, color=None, underline=False):
    if urlsplit(url).scheme not in {"https", "mailto"}:
        raise ValueError(f"Unsupported CV link: {url}")
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True))
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    if font:
        fonts = OxmlElement("w:rFonts")
        for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn("w:" + slot), font)
        properties.append(fonts)
    for tag, value in (("color", color), ("sz", round(size * 2) if size else None), ("u", "single" if underline else None)):
        if value:
            item = OxmlElement("w:" + tag)
            item.set(qn("w:val"), str(value))
            properties.append(item)
    run.append(properties)
    label = OxmlElement("w:t")
    label.text = text
    run.append(label)
    link.append(run)
    paragraph._p.append(link)
