# Generating CommonLit documents with an AI system

Practical guidance for AI systems (and the people prompting them) that create or edit CommonLit curriculum documents. It comes from building and testing the Student Copy Creator on real teacher copies: these are the things that actually went wrong, and what worked.

## Contents
1. Pick the right route to the file
2. Start from a template, not a blank page
3. Formatting the way Google Docs and Word understand it
4. Writing the content
5. Making the student copy
6. Checking your work
7. Common AI failure modes
8. Prompt starters

## 1. Pick the right route to the file

| Route | Good for | Watch out for |
| --- | --- | --- |
| **Edit a copy of a template Google Doc** (Docs API / Apps Script / a person) | Production lessons. Logos, headers, footers and section setup come for free. | Needs access to the template and the person's Drive. Apps Script can't edit section breaks or first-page headers directly; the Docs API (`updateSectionStyle`, `useFirstPageHeaderFooter`) can. |
| **Generate a .docx from a template .docx** (python-docx, docx-js, editing the XML) | Prototypes, batch work, tools that run anywhere. Imports into Google Docs well. | Round-tripping Google Doc → .docx → Google Doc can shift spacing slightly. Compare against the original. |
| **Generate a .docx from scratch** | Quick drafts only. | You won't have the real logo, header setup or styles. Say so, and leave logo placeholders to a person. |
| **HTML/PDF preview** | Reviewing a design idea. | Not an editable deliverable for writers. |

Writers work in Google Docs, so the end product almost always needs to be a Google Doc. If you produce a .docx, tell the person to upload it to Drive and use **File → Save as Google Docs**.

## 2. Start from a template, not a blank page

- Ask for the current template or a recent teacher copy from the same unit and grade, and build inside a copy of it. Templates carry the large and small logo headers, the footer with its page field, the section that numbers pages from 1, and the styles.
- Keep the template's structure (section breaks, the Name/Class line, box tables) and replace the content.
- Copy existing boxes rather than drawing new ones. A copied box keeps the fill, border, padding and fonts.
- Never recreate the logo from text, an icon font or an image search. If there's no logo file, leave the header to a person.

## 3. Formatting the way Google Docs and Word understand it

- **Use styles.** Title for page titles, Heading 2 for parts, Normal for body. Don't fake a heading with big bold Normal text: the outline, accessibility and conversion tools all read styles.
- **Pick weights by font name.** Use "Open Sans Medium" or "Open Sans Light" as the font. Don't apply Bold to Open Sans; that produces the 700 weight the brand forbids.
- **Colors as exact hex.** Answers `#0000FF`, text `#000000`/`#1F1727`, box fill `#E3EEF0`, box border `#C3DDE4`. Approximate colors break the conversion (section 5).
- **Lists as real lists.** Choices `A.`–`D.` and bullets should be auto-numbered or bulleted lists, not typed characters. The exception: lettered question parts inside a question box (`A. Find Evidence:`) are typed, and an optional one is typed `*A.`.
- **Alignment with tab stops, not spaces.** Page numbers, `Name: ____ Class: ____`, and any right-aligned text use a tab stop. Rows of spaces or tabs break when text changes.
- **Allow table rows to break across pages** for reading tables and long answer boxes.
- **Set the paragraph's own formatting** on empty answer lines (font size, color black). An empty line keeps the formatting of the text that was deleted from it, including blue.
- **Section breaks:** the cover is its own section, with "Different first page" on for the large logo. The lesson section restarts page numbering at 1.

## 4. Writing the content

- Follow the anatomy in `page-anatomy.md` for the lesson type. Teachers and students rely on the same parts appearing in the same order every lesson.
- Write student directions as short imperatives with timing: `Directions: Review each text and complete the table. 5 minutes`.
- Put the paragraph range before each during-reading question (`Paragraphs 26–34`) and use the standard labels (Find Evidence, Think and Share, Turn and Talk, Write, Poll the Class).
- Write teacher answers the way a strong student would, at realistic length. The length decides the writing space in the student copy. For open questions, add a model plus `Student responses will vary.`
- Put standards at the end in brackets (`[RL.3]`). If you don't know the right standard, write `[STANDARD]`. Never guess one.
- Don't invent text excerpts, paragraph numbers, author facts or permission lines. Use placeholders and list them for the person.
- Keep real assessment answers out of anything public or shared outside the team.

## 5. Making the student copy

Make it from the finished teacher copy, never write it separately. Two documents written separately drift apart.

- Easiest: run the teacher copy through the **Student Copy Creator** (download the Google Doc as .docx, drop it on the page, upload the result).
- If you're converting in code, follow `teacher-student-pairing.md` section 2 exactly. The lessons learned the hard way:
  - The student part may start with a Name/Class **table** (older template) or a typed **Name: ___ Class: ___ line** (newer). Handle both.
  - Removing the cover also removes its section, and with it the large-logo first-page header. Carry that header onto the new first page.
  - Replace each answer with the **same height** of blank space it took. Adding more space pushes reading-table rows to the next page and strands the "Whole Class Reading" header. Adding less shifts everything up.
  - Clear answer blue and highlights from **paragraph marks and footnote numbers**, not just from visible text.
  - After swapping the footer's copyright line for the longer CC line, keep the page number right-aligned with a tab stop.
  - Match `Answers in blue` and `Answers are in blue`. Wording varies between series.

## 6. Checking your work

- Run `scripts/check_docx.py <file> --mode teacher|student` on .docx output. It checks fonts, colors, highlights, leftover blue and teacher text, the footer line and the page number setup.
- Go through `qa-checklist.md`.
- If you can render pages (LibreOffice → PDF, or ask the person to open it), compare the teacher and student copies page by page. Each question should be on the same page in both.
- Report honestly. Say what you checked by code, what you checked visually, and what you couldn't check (for example "I couldn't open it in Google Docs").

## 7. Common AI failure modes

| Failure | Fix |
| --- | --- |
| Bold Open Sans everywhere | Use Open Sans Medium for labels and headings; emphasis is rare |
| ALL CAPS headings | Sentence case; caps only in Source Code Pro metadata |
| Colored headings or yellow text | Text is ink; color goes in fills |
| Emoji or icons in a handout | Remove; words and boxes only |
| Invented standards, quotes or paragraph numbers | Placeholders + a list for the person |
| Answers in "a blue" (`#1155CC`, `#4285F4`) | Exactly `#0000FF` |
| "(Answer: B)" written into the question | Color the correct option instead |
| Student copy written from scratch | Derive it from the teacher copy |
| Extra blank lines "for writing" everywhere | Match the teacher copy's spacing; boxes get at least 3 lines |
| Page number centered or pushed by spaces | Right-aligned tab stop |
| Logo recreated as text | Use the template's header; otherwise leave it to a person |
| Claims it "looks right" without seeing it | Say what was and wasn't verified |

## 8. Prompt starters

For a person asking an AI system to draft a lesson:

> Using the CommonLit curriculum docs skill, draft a **teacher copy** for a Grade 8 reading lesson on "[text]" by [author], Unit [n] "[unit name]". Skill focus: [skill]. Use the current Thematic Mini Unit template structure: Lesson Overview cover (skill focus, materials, How to Facilitate table with timing), Name/Class line, About this Text, a reading table with Whole Class / Partner / Independent sections and During Reading Questions beside the paragraphs, a discussion, and independent practice. Put answers in #0000FF. Leave [STANDARD] and [Paragraph #] placeholders where you aren't sure, and list them at the end.

For converting:

> Here is a teacher copy (.docx). Make the student copy following the CommonLit teacher → student conversion. Then run the checker in student mode and tell me what you changed, what the checker found, and what I still need to check by hand.

For reviewing:

> Review this handout against the CommonLit curriculum brand and template rules. List problems by page with the rule each one breaks, most important first. Don't rewrite it yet.
