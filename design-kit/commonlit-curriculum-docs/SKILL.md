---
name: commonlit-curriculum-docs
description: Design and generate CommonLit curriculum documents (lesson handouts, teacher copies, student copies, answer keys, writing lessons, independent practice) that match CommonLit branding and the CommonLit 360 template conventions. Use this whenever someone asks to create, draft, format, restyle, review or convert a CommonLit lesson, handout, worksheet, teacher copy or student copy, or asks for anything "on brand" for CommonLit curriculum, even if they don't mention branding or templates. Also use it when turning a teacher copy into a student copy, or when building a tool, prompt or pipeline that produces CommonLit documents.
---

# CommonLit curriculum documents

This skill helps you make CommonLit lesson documents that look like they came from the curriculum team: on-brand, accessible, and built so a **teacher copy** and its **student copy** stay in sync.

Most CommonLit lessons ship as a pair:

- **Teacher copy**: a cover page (Lesson Overview), the full student-facing lesson, and answers in blue.
- **Student copy**: the same lesson with the cover, answers and teacher-only notes removed, and the same layout, so a question sits at the same spot on the same page in both.

The student copy is *derived* from the teacher copy. Design the teacher copy so the student copy can be produced mechanically: by a person following the how-to, or by a tool like the Student Copy Creator. Nearly every rule below exists to keep that conversion safe.

## Before you start

1. **Find out the output.** A Google Doc (most common for writers), a .docx, or a PDF/HTML preview? Curriculum writers work in Google Docs. If you generate a file, .docx imports into Google Docs cleanly.
2. **Start from a real template when one exists.** Ask for the current CommonLit template or a recent teacher copy from the same unit, and copy it. Templates carry the header logos, footers, section setup and styles that are hard to recreate by hand. Never redraw or approximate the logo.
3. **Learn the lesson type.** Reading lesson, writing lesson, discussion, mini unit, or assessment. Each has its own anatomy (see `references/page-anatomy.md`).

## The rules that matter most

These are the ones AI-generated documents most often get wrong. The reason for each is in the reference files.

**Brand** (details: `references/brand.md`)
- Type: **Open Sans** for everything in a handout (Light for body, Medium for labels and headings). Never use Open Sans Bold (700). Crimson Pro is for display and marketing, never below 20px. Source Code Pro is for small metadata, and is the only place ALL CAPS is allowed.
- Text is **ink** (`#1F1727`, or black in curriculum docs). Accent colors are for fills, never for text. The one exception in curriculum documents is **answer blue `#0000FF`** in teacher copies.
- Boxes and callouts use the soft **Blue-Gray** fills (`#E3EEF0`, border `#C3DDE4`). Yellow is for action items in product and marketing, not lesson content.
- Voice: plain, warm, teacher to teacher. Student directions start with a verb ("Read the prompt. Then answer the question.") and include the time ("5 minutes").
- No emoji, no gradients, no decorative icons in handouts.

**Teacher/student pairing** (details: `references/teacher-student-pairing.md`)
- Every answer is in exactly `#0000FF`, and only answers use that blue.
- The cover ends with a page break or section break, and the student part starts with the **Name / Class** line.
- The header says `Teacher Copy • <Unit name>`. In Ed2.0 documents the page-1 title also reads `TEACHER COPY: <Lesson title>`; newer templates use the plain title.
- The `*Answers in blue…*` notice, and any teacher-only note, sits in its **own paragraph or box**, never mixed into a student line.
- Optional questions start with `*` (e.g. `*A. Find Evidence: …`). They appear only on the teacher copy.
- In a list of choices, mark the correct one by making it blue, never by adding text such as "(correct)".
- Size answer space by its purpose. A blank answer area in the student copy keeps the same height as the answer did in the teacher copy, so pages line up.
- Footers: the teacher copy carries `© CommonLit, Inc. <year>`. The student copy carries `Unless otherwise noted, this content is licensed under the CC BY-NC-SA 4.0 license.` with a link. The page number is right-aligned.

## Workflow for generating a lesson pair

1. **Outline** the lesson from the request: lesson type, title, unit, text(s), skill focus with standard codes (e.g. `[RL.8.3]`), parts with timing, and question types.
2. **Write the teacher copy first**, in this order: cover page → Name/Class line → title and author → goals or About this Text → parts → independent practice. Put answers in blue directly under or beside each question, with the same layout the student will see.
3. **Run the checklist** in `references/qa-checklist.md`. If you produced a .docx, run `scripts/check_docx.py <file> --mode teacher` and fix what it reports.
4. **Derive the student copy** by applying the conversion in `references/teacher-student-pairing.md`, or by running the teacher copy through the Student Copy Creator. Then run `scripts/check_docx.py <file> --mode student`.
5. **Tell the person what still needs eyes.** Logos on page 1, answer box sizes, questions split across pages, and anything you weren't sure about. Don't claim a visual check you didn't do.

## Content you must not invent

Standard codes, Lexile levels, text excerpts, paragraph numbers, author bios, copyright/permission lines, and answer keys for real assessment items all need a source. If the person didn't give it to you, leave a clearly marked placeholder such as `[STANDARD]` or `[Paragraph #]` and list the placeholders at the end. Assessment answers are confidential: don't put real assessment answer keys anywhere public. The teacher copy notice exists because of this.

## Reference files

Read the one you need; you don't need all of them for every task.

- `references/brand.md`: colors, type, voice, logo, accessibility, and how curriculum documents differ from marketing.
- `references/page-anatomy.md`: page-by-page layout of teacher and student copies, by lesson type, with the exact styles the current template uses.
- `references/teacher-student-pairing.md`: the authoring conventions and the full teacher → student conversion.
- `references/ai-generation-tips.md`: how to produce these files with an AI system (Google Docs vs .docx, using templates, prompting, common failure modes).
- `references/qa-checklist.md`: the final checklist before a document goes to a person.
- `assets/style-spec.json`: the same specs as data, for scripts and tools.
- `scripts/check_docx.py`: a checker for .docx files (fonts, colors, blue answers, highlights, footer line, page numbers, leftover teacher text).
