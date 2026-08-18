# Band-Up Exam Suite

A self-hosted mock-test platform (GitHub Pages frontend + Google Sheets/Apps Script backend)
for **IELTS** and **TOEIC** simulations — candidate login, timed test-taking, server-side
grading, score reports, and an admin panel for creating and publishing test sets.

## Modules

| | IELTS | TOEIC |
|---|---|---|
| Question authoring | `admin.html` (built-in) | `toeic-question-maker.html` |
| Test-taking | `simulation.html` | `toeic-simulation.html` |
| Format | Listening + Reading + Writing (+ Speaking), Band 0-9 | Listening (100) + Reading (100), all multiple choice, scaled 10-990 |
| Configs live in | `configs/IELTS/` | `configs/TOEIC/` |

Both modules share the same login (`index.html`), profile/TRF page (`profile.html`), test
listing (`tests.html`), admin panel (`admin.html`), and Google Sheets backend (`Code.gs`).

## Setup

See `SETUP_GITHUB.txt` for the full step-by-step (written in Bisaya) - GitHub Pages hosting,
Google Sheets/Apps Script backend, and how to publish your first test set for each exam type.

## Adding a TOEIC test

1. Open `toeic-question-maker.html`, import your 200 questions via CSV, download the config,
   and upload it to `configs/TOEIC/` on GitHub.
2. In `admin.html` -> Publish panel, set **Exam Type: TOEIC**, keep "TOEIC" in the Test Set
   Name, paste the config's Raw URL, and save.
3. Send the answer key to Google Sheets from the TOEIC Question Maker's Publish tab.

Candidates picking a TOEIC-tagged test set from `tests.html` are automatically routed to
`toeic-simulation.html` instead of the IELTS engine.

See `configs/TOEIC/README.md` for details on the config schema and numbering convention.
