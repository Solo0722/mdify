# Office fixtures

`office.docx`, `office.xlsx`, and `office.pptx` contain synthetic test content
created with the installed Microsoft Office applications by
`generate_office_fixtures.ps1`. No user documents or third-party content were
copied. Account metadata was replaced with generic fixture metadata by
`sanitize_metadata.py`. These files are distributed under the project license.

The committed fixtures run on every supported platform without Microsoft Office.
The PowerShell generator is optional, Windows-only, and refuses to overwrite
existing fixtures. After deliberate regeneration, run:

```powershell
python tests/fixtures/sanitize_metadata.py
python -m pytest tests/test_office_fixtures.py
```

Coverage:

- Word: native heading and list styles, content control.
- Excel: date formatting, cached formula value.
- PowerPoint: reordered slides, hidden and blank slides, inherited nested
  bullets, grouped shapes, merged cells, speaker notes.
