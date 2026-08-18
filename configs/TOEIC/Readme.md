# configs/TOEIC/

Config JSON files for TOEIC test sets go here (same idea as `configs/IELTS/`).

## Files in this folder

- **toeic-demo-config.json** — a full 200-question *structural* demo (Parts 1–7, correctly
  numbered 1–200). All question text, options, images, and audio are placeholders. Use it to
  test the full candidate flow (login → Listening → Reading → submit → scaled score) end-to-end
  before you've written any real content. **Do not use this with real candidates as-is.**

## Adding a real TOEIC test

1. Open `toeic-question-maker.html` (in the repo root) and go to **2. Import Questions (CSV)**.
2. Download the CSV template, fill in your 200 questions (Google Sheets or Excel work fine —
   export as CSV when done), and upload it back in.
3. Check the **3. Review** tab — it confirms all 200 questions are present and every one has an
   answer.
4. In **4. Publish**:
   - Click **Download Config JSON** — this file has the answers stripped out, so it's safe to
     make public. Upload it here, to `configs/TOEIC/`, via GitHub's "Add file → Upload files".
   - Click the uploaded file → **Raw** → copy that URL.
   - Click **Send Answer Key to Sheets** — this needs a Test Set ID. Create the test set first in
     `admin.html` → Publish panel (paste the Raw URL there, set **Exam Type: TOEIC**, and keep the
     word "TOEIC" in the Test Set Name), copy the Set ID it creates, then come back here and
     paste it in when prompted.

Grading always happens server-side against the `AnswerKeys` sheet — the config JSON that ships to
GitHub Pages never contains correct answers, same as the IELTS side of this app.

## Numbering convention

Real TOEIC numbering is used: **Listening 1–100** (Part 1: 1–6, Part 2: 7–31, Part 3: 32–70,
Part 4: 71–100), **Reading 101–200** (Part 5: 101–130, Part 6: 131–146, Part 7: 147–200).
