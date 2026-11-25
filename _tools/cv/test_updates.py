"""Exercise a website data edit through the public CV generator and PDF renderer."""
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path

import yaml


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]


def normalized(text):
    # PDF line breaks and spacing can change with fonts and page layout.
    return "".join(unicodedata.normalize("NFKC", text).split())


def main():
    paper = {
        "title": "CV update verification paper",
        "collection": "publications",
        "category": "manuscripts",
        "date": "2025-01-02",
        "venue": "Test Journal",
        "citation": "Suncheol Kwon",
    }
    patent = {
        "title": "CV update verification patent",
        "title_ko": "자동 갱신 검증용 특허",
        "country": "KR",
        "number": "10-0000000",
        "date": "2025-01-02",
        "role": "included",
    }
    with tempfile.TemporaryDirectory() as directory:
        source = Path(directory) / "source"
        output = Path(directory) / "rendered"
        source.mkdir()
        for folder in ("_data", "_publications"):
            shutil.copytree(ROOT / folder, source / folder)
        shutil.copy2(ROOT / "_config.yml", source / "_config.yml")

        publication_path = source / "_publications" / "cv-update-verification.md"
        publication_path.write_text(
            "---\n" + yaml.safe_dump(paper, sort_keys=False) + "---\n", encoding="utf-8"
        )
        draft = dict(paper, title="Unpublished CV test draft", published=False)
        (source / "_publications" / "cv-update-draft.md").write_text(
            "---\n" + yaml.safe_dump(draft, sort_keys=False) + "---\n", encoding="utf-8"
        )
        patents_path = source / "_data" / "patents.yml"
        patents = yaml.safe_load(patents_path.read_text(encoding="utf-8"))
        patents["korean"].append(patent)
        patents_path.write_text(
            yaml.safe_dump(patents, allow_unicode=True, sort_keys=False), encoding="utf-8"
        )

        subprocess.run(
            [sys.executable, HERE / "generate.py", "--source", source, "--output", output],
            check=True,
        )
        for filename, title_field in (("cv.pdf", "title"), ("cv-ko.pdf", "title_ko")):
            pdf_text = subprocess.check_output(
                ["pdftotext", "-enc", "UTF-8", output / filename, "-"], encoding="utf-8"
            )
            contents = normalized(pdf_text)
            for value in (paper["title"], patent[title_field], patent["number"]):
                if normalized(value) not in contents:
                    raise AssertionError(f"{filename} is missing the new record: {value}")
            if normalized(draft["title"]) in contents:
                raise AssertionError(f"{filename} contains an unpublished record")
            print(f"{filename}: verified the new publication and localized patent")


if __name__ == "__main__":
    main()
