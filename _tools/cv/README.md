# Website and CV sources

The website and both PDFs share the publication records in `_publications/`,
patent grants in `_data/patents.yml`, and the bilingual research interests,
appointments, degrees and honors in `_data/cv_profile.yml`. Update profile
records in English and Korean. The CVs also use `_data/cv_supplement.yml`
for project funding, devices, professional service, the book translation,
domestic presentations, historical patent applications and Korean patent
titles. International publication titles and venue names retain their
published wording. A Korean patent title can also be supplied as `title_ko`
in its grant record; otherwise its source English title is used.

To add a publication, add its Markdown file to `_publications/` using the same
front matter as another record in the same category. To add a Korean patent,
append one entry under `korean` in `_data/patents.yml`. Use `title` for the
English name, `title_ko` for the official Korean name, and the existing
`country`, `number`, `date`, and `role` fields. Both PDFs read that one entry.

To render locally on Ubuntu:

```sh
sudo apt-get install libreoffice-writer fonts-liberation fonts-noto-cjk python3-docx python3-yaml poppler-utils
/usr/bin/python3 _tools/cv/generate.py
```

The English PDF is `files/cv.pdf` and the Korean PDF is `files/cv-ko.pdf`.
Use `--output DIRECTORY` to review the generated files in another location.

The website workflow regenerates both files before each Jekyll build and
checks that the built CV page links to valid PDFs. Pull requests build a
preview artifact and run `python3 _tools/cv/test_updates.py`, which adds a
publication and a bilingual patent in a temporary copy of the source data
and checks the rendered PDFs. Merges to `master` regenerate and deploy both
PDFs at the existing addresses. No additional PDF upload or commit is needed.
For this deployment, select **Settings → Pages →
Build and deployment → Source → GitHub Actions** in the repository. Until that
one-time switch, a branch-based Pages build serves the committed PDF snapshots.
