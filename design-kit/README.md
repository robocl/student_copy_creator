# CommonLit curriculum docs: design kit

A package for making CommonLit teacher copies and student copies that look right and convert cleanly. It works for people and for AI systems.

It's written as a **Claude skill**: a folder of instructions that Claude loads automatically when a task needs it. Every file is plain text, so you can also read it yourself or paste parts into another AI tool.

## What's inside

| File | What it's for |
| --- | --- |
| `commonlit-curriculum-docs/SKILL.md` | The starting point: the most important rules and the step-by-step workflow |
| `references/brand.md` | Colors, fonts, voice, logo and accessibility, and where curriculum docs differ from marketing |
| `references/page-anatomy.md` | What goes on each page of a teacher and student copy, by lesson type, with the template's exact fonts and sizes |
| `references/teacher-student-pairing.md` | How to write a teacher copy so the student copy can be made automatically, plus the full conversion steps |
| `references/ai-generation-tips.md` | Tips for AI systems: templates, formatting, what goes wrong, and prompts to start from |
| `references/qa-checklist.md` | The checklist before a document goes out |
| `assets/style-spec.json` | The same specs as data, for tools |
| `scripts/check_docx.py` | Checks a .docx teacher or student copy and lists problems |

## How to use it

- **In Claude (claude.ai):** open `commonlit-curriculum-docs.skill` and click **Save skill**, or upload the folder as a zip under Settings → Capabilities → Skills. After that, Claude uses it whenever you ask for CommonLit lessons, handouts or copies.
- **In Claude Code:** copy the `commonlit-curriculum-docs` folder into `.claude/skills/` in your project.
- **With another AI tool:** paste `SKILL.md` into its instructions, and attach the reference files it needs.
- **As a person:** start with `SKILL.md`, then `page-anatomy.md`.

To check a file yourself (needs Python 3):

```
python3 commonlit-curriculum-docs/scripts/check_docx.py "My Lesson TEACHER COPY.docx" --mode teacher
python3 commonlit-curriculum-docs/scripts/check_docx.py "My Lesson STUDENT COPY.docx" --mode student
```

## Where this came from

- The CommonLit Design System (from the 2025 Brand Guidelines).
- "HOW TO: Making Student Copies for a 360 Unit".
- Real lessons: "Button, Button" and "Writing a Strong Claim" (8G Unit 1, Ed2.0), and "The Gift of the Magi" and "Planning Your Literary Analysis Essay" (8G Thematic Mini Unit 1).
- What we learned building the Student Copy Creator.

Some details come from only those four lessons, such as font sizes and box colors. Check them against your project's template, and update `page-anatomy.md` and `style-spec.json` if they differ.
