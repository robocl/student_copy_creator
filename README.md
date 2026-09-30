# Student Copy Creator

Turns a CommonLit **teacher copy** into a **student copy**. It follows the "HOW TO: Making Student Copies for a 360 Unit" guide.

The tool comes in two forms. Both use the same rules:

1. **Prototype web page (no setup, try it now).** In Google Docs, use File → Download → Microsoft Word (.docx). Drop the file on the page, save the student copy it gives back, and upload that to Drive. The conversion happens inside your browser, so nothing is uploaded anywhere and nothing is installed in CommonLit's Google account. The page is built from `web/` into `dist/student-copy-creator.html`.
2. **Google Docs version (later, needs approval).** Paste a teacher copy link and the student copy appears in the same Drive folder, with no download or upload step. This one has to be set up inside CommonLit's Google account (see "Setup" below), so check with IT first.

## What it does

**Done for you**

- Makes a copy **without comments** and names it correctly: removes "Copy of", changes TEACHER to STUDENT, and keeps the lesson number and the Ed2.0 ending.
- Saves it in the same folder as the teacher copy.
- Deletes the teacher cover page(s): everything before the **Name / Class** box or the typed "Name: ___ Class: ___" line.
- Header: "Teacher Copy" → "Student Copy" (on page 1 and on the other pages).
- Footer: "© CommonLit, Inc. 2026" → "Unless otherwise noted, this content is licensed under the CC BY-NC-SA 4.0 license." with the link.
- Logos: page 1 gets the large logo and the other pages keep the small one. The web page moves the cover page's large-logo header onto page 1 automatically. The Google Docs version can't yet, so check it by hand there.
- Removes "TEACHER COPY:" from the title on page 1.
- Deletes blue answers and leaves the same amount of blank space the answer took up, so questions stay lined up with their paragraphs as in the teacher copy. A box that held only an answer gets at least 3 lines.
- When the correct choice is marked in blue among black options, it keeps the option but makes it plain black text.
- Deletes the "Answers in blue…" line and the blue "Note: To ensure test security…" notes.
- **Fully deletes optional (\*) questions**, along with their answer choices and answers.
- Deletes "Notes to Teacher" boxes.
- Removes all highlights.

**Still done by a person** (the tool lists these every time)

- Check page 1 has the large logo and the other pages have the small one.
- Check answer boxes are a reasonable size (short response box = length of the page).
- Check no question is split across two pages, and questions still line up with their paragraphs.
- Link it in the Dig Guide (force copy) and the tracker, and turn the box blue.

The tool also lists anything it wasn't sure about under **Please check**, for example blue text it couldn't place, or each optional question it deleted.

**It depends on writers keeping to these habits:** answers are exactly the standard blue (#0000FF), the student part starts with the Name / Class box, and teacher notes sit in their own paragraph or box.

## Setup for the Google Docs version (one time, about 10 minutes, no coding)

One person does this once. Everyone else just uses the link it produces.

1. Go to **script.google.com** while signed in to your CommonLit Google account. Click **New project**.
2. Click **Untitled project** at the top and rename it **Student Copy Creator**.
3. You'll see a file called **Code.gs** with a few lines in it. Delete those lines. Paste in everything from [`paste-into-google/Code.gs`](paste-into-google/Code.gs). On GitHub, the "Copy raw file" button copies it all.
4. Click the **+** next to "Files" → **HTML**. Name it `Index` (Google adds the `.html`). Delete what's in it and paste in everything from [`paste-into-google/Index.html`](paste-into-google/Index.html).
5. Click the 💾 **Save** icon.
6. Click the blue **Deploy** button (top right) → **New deployment**. Click the ⚙️ next to "Select type" → **Web app**. Fill in:
   - **Execute as:** *User accessing the web app*
   - **Who has access:** *Anyone within CommonLit*

   Then click **Deploy**.
7. Google asks you to **Authorize access**. Pick your account. You may see a "Google hasn't verified this app" screen, which is normal for internal tools. Click **Advanced → Go to Student Copy Creator**, then **Allow**. The tool needs to read and copy Docs in your Drive.
8. Copy the **Web app URL** it shows you. That's the tool. Bookmark it and share it with the curriculum team. Each person sees the same permission screen the first time they use it.

**Try it first on a test copy.** Make a copy of a teacher copy (like Katie's test copy) and run the tool on that before using real lessons.

### Updating it later

If the rules change, paste the new `Code.gs` over the old one and save. Then go to **Deploy → Manage deployments → ✏️ (edit) → Version: New version → Deploy**. The link stays the same.

### Optional: a button inside Google Docs

The same code can add a menu inside every Google Doc: **Extensions → Student Copy Creator → Create student copy**. For the whole team, it has to be published as a private "Google Workspace add-on", which a Google Workspace admin at CommonLit needs to help with. Worth doing once the web page version has proven itself.

## Using it

1. Resolve comments and make sure copy-editing feedback is in (step 1 of the how-to). Suggested edits should be accepted or rejected first.
2. Open the tool's link, paste one or more teacher copy links (one per line), and click **Create student copy**.
3. Open the new doc and work through **Please check** and **Still do by hand**.

If the teacher copy is a Word file (.docx) in Drive, open it and use **File → Save as Google Docs** first.

---

## For developers

- `apps-script/`: the source. `Converter.js` holds the conversion rules (settings are in `DEFAULTS` at the top, e.g. `optionalQuestions: 'unmark'` keeps optional questions but drops the `*`). `Code.js` holds the web app, Drive copy, and header/footer handling. `Index.html` is the page.
- `web/`: the prototype page. `docx-adapter.js` lets the same `Converter.js` rules edit a Word file directly, and `page.html` is the page.
- `paste-into-google/` and `dist/`: built files (the Google paste bundle and the web page). Regenerate them with `npm run build`. A test fails if you forget.
- `npm test` runs the rules against a mock of Google's DocumentApp (`test/mock-docs.js`), and end to end on .docx files (`test/docx.test.js`). For .docx golden tests, drop `<name>.teacher.docx` and `<name>.student.docx` into `test/fixtures/`. You can add golden tests from real teacher/student pairs with `test/docx_to_json.py` (see `test/fixtures/README.md`). Fixtures are git-ignored because teacher copies contain answer keys.
- You can also deploy with [clasp](https://github.com/google/clasp) (`.clasp.json.example`, `npm run deploy`) instead of copying and pasting.

The mock only follows Google's documentation. Real-Doc quirks will only show up by running the tool on real teacher copies.
