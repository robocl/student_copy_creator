# Student Copy Creator

Turns a CommonLit **teacher copy** Google Doc into a **student copy** in one step.

A curriculum writer pastes a teacher-copy link into a small internal web page. They get back a link to a new student copy in the same Drive folder. The teacher copy is never modified.

## What it changes

| Teacher copy | Student copy |
| --- | --- |
| Cover pages (Lesson Overview, pacing, notes to teachers): everything before the first **Name / Class** table | Removed |
| `TEACHER COPY: Button, Button` | `Button, Button` |
| `*Answers in blue…*` and blue `Note: To ensure test security…` lines | Removed |
| Blue answers (paragraphs or bullets) | Replaced with blank writing lines |
| Blue answer text at the end of a question line | Removed, question kept |
| A blue (bold) option in a list of black options, e.g. the correct "This claim just restates the prompt." | Kept, set back to plain black text |
| `*A. Find Evidence…` (optional-question asterisk) | `A. Find Evidence…` |
| File name `…TEACHER_COPY…` | `…STUDENT_COPY…` |
| "Different first page" footer (used by the cover) | Turned off |

The result card also lists **warnings** to check by hand, for example blue text that is still there, the word "teacher" still appearing, or no Name table found.

### What writers need to keep doing

The tool depends on the conventions already used in the Ed2.0 docs:

1. Answers are colored exactly **#0000FF** (Google Docs "blue" in the custom color picker). Other blues are left alone.
2. The student-facing part starts with the **Name / Class** table.
3. Teacher-only notes sit in their own paragraph.

## Why Google Apps Script

- It edits the Google Doc directly. There's no export/import round trip, so fonts, tables, images, headers, and footers come through untouched.
- Nothing to host or pay for. It runs under each writer's own Google account, so the copy gets the writer's normal Drive permissions.
- The same code powers the web page and (optionally) a menu item inside Google Docs.

## Setup (one-time, ~10 minutes)

You need a CommonLit Google account.

### Option A: copy and paste (no tools needed)

1. Go to <https://script.google.com> and click **New project**. Name it "Student Copy Creator".
2. Create these files and paste in the contents from `apps-script/`:
   - `Code.gs` ← `apps-script/Code.js`
   - `Converter.gs` ← `apps-script/Converter.js`
   - `Index.html` ← `apps-script/Index.html` (**File > New > HTML**, name it `Index`)
3. **Project Settings** (gear) → check **Show "appsscript.json" manifest file in editor**. Then replace its contents with `apps-script/appsscript.json`.
4. **Services (+)** → add **Google Docs API**. This is used only to turn off the first-page footer.
5. **Deploy > New deployment** → type **Web app**:
   - *Execute as:* **User accessing the web app**
   - *Who has access:* **Anyone within CommonLit**
6. Share the web app URL with the curriculum team. On first use, each person clicks through a Google permission screen.

### Option B: from this repo with `clasp`

```bash
npm install -g @google/clasp
clasp login
clasp create --type standalone --title "Student Copy Creator" --rootDir apps-script
clasp push
clasp deploy --description "v1"
```

Then open the project (`clasp open`) and do step 5 above once so it's deployed as a web app. To ship updates later, run `npm run deploy`, or in the editor use **Deploy > Manage deployments > Edit > New version**.

### Optional: a menu inside Google Docs

`Code.js` also adds **Extensions → Student Copy Creator → Create student copy** when the script runs as a Docs add-on. Two ways to get that:

- **One writer, quick:** open any doc → **Extensions > Apps Script**, paste the same files, save, reload the doc. This only works in that doc.
- **Whole team:** publish it as a private Google Workspace **Editor add-on** for the CommonLit domain. This needs a Google Cloud project with the *Google Workspace Marketplace SDK*, plus a Workspace admin to install it for the curriculum team. Worth doing once the web app has proven the rules.

## Using it

1. Open the web app URL.
2. Paste one or more teacher-copy links, one per line.
3. Click **Create student copy** and open the link it returns. Skim it, especially anything listed under *Please check*.

If the teacher copy is a `.docx` in Drive, open it and use **File > Save as Google Docs** first.

## Tuning

Options live in `DEFAULTS` at the top of `apps-script/Converter.js`:

- `answerColors`: colors treated as answers (default `#0000ff`).
- `teacherOnlyPatterns`: paragraphs matching these are deleted.
- `minAnswerLines` / `maxAnswerLines` / `charsPerAnswerLine`: how much blank writing space replaces an answer.

## Tests

The conversion rules run in Node against a mock of the DocumentApp API (`test/mock-docs.js`):

```bash
npm test
```

You can also run golden tests against real teacher/student pairs. Export both docs as .docx, convert them with `test/docx_to_json.py`, and drop the JSON into `test/fixtures/` (see the README there). Fixtures are git-ignored because teacher copies contain answer keys. "Writing a Strong Claim" (8G Unit 1) matches the published student copy line for line. The only difference is the number of blank writing lines.

The mock only follows Google's documented behavior. Real-Doc quirks will only show up by running the tool on real teacher copies.
