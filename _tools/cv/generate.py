"""Render both public CV PDFs from the records used to build the website."""
import argparse
from pathlib import Path
import shutil
import subprocess
import tempfile

from cv_records import ROOT, load_data
from render_en import render as render_en
from render_ko import render as render_ko


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=ROOT, help="Directory containing the website's source records")
    parser.add_argument("--output", type=Path, help="PDF destination (defaults to the site's files directory)")
    args = parser.parse_args()
    data = load_data(args.source)
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if not office:
        parser.error("Install LibreOffice Writer to render the CV PDFs.")
    output = args.output or args.source / "files"
    output.mkdir(parents=True, exist_ok=True)

    # Stage on the target filesystem so readers can only open complete PDFs.
    with tempfile.TemporaryDirectory(dir=output) as temporary:
        stage = Path(temporary).resolve()
        documents = []
        for name, render in (("cv", render_en), ("cv-ko", render_ko)):
            path = stage / f"{name}.docx"
            render(data, path)
            documents.append(path)
        subprocess.run([office, f"-env:UserInstallation={(stage / 'office').as_uri()}", "--headless",
                        "--convert-to", "pdf:writer_pdf_Export", "--outdir", str(stage), *map(str, documents)],
                       check=True, timeout=180)
        pdfs = [path.with_suffix(".pdf") for path in documents]
        for pdf in pdfs:
            with pdf.open("rb") as stream:
                if stream.read(5) != b"%PDF-":
                    raise ValueError(f"Invalid PDF output: {pdf.name}")
        for pdf in pdfs:
            pdf.replace(output / pdf.name)


if __name__ == "__main__":
    main()
