# Teacher and student copies: authoring for a clean conversion

The student copy is made *from* the teacher copy, by a person following "HOW TO: Making Student Copies for a 360 Unit", or by a tool (the Student Copy Creator does the steps marked **auto**). A teacher copy written with these conventions converts with almost no hand work. One that breaks them needs a person to hunt for leftovers. Leftovers are how answer keys leak to students.

## Contents
1. Authoring conventions
2. The conversion, step by step
3. File naming and process
4. Edge cases

## 1. Authoring conventions

| Convention | Why |
| --- | --- |
| Answers in exactly `#0000FF`, and nothing else in that color | The conversion finds answers by color. A near-blue (`#1155CC` link blue, `#4A86E8`) won't be found and will leak. A blue heading will be deleted. |
| The cover ends with a page or section break, and the next line is the Name/Class line | Everything before Name/Class is deleted. Nothing student-facing may sit above it. |
| Title is `TEACHER COPY: <title>` | The prefix is removed and the rest stays. |
| Teacher-only notes each in their own paragraph or box ("Answers in blue…", "Note: To ensure test security…", "Notes to Teacher") | Whole paragraphs and boxes are deleted. A note glued to a student line can't be removed cleanly. |
| Optional questions start with `*` + label (`*A. Find Evidence:`) | The question, its options and its answers are deleted from the student copy. |
| Correct choice marked by blue, not by words | The option is turned back to black and the rest stay as is. "(correct)" would survive. |
| Answer text goes where the student will write, at a realistic length | The blank space left behind is the student's writing space. A one-word answer leaves one line. |
| Rows in reading tables can break across pages | Removing or changing answer text can't then push a whole row to the next page. |
| Highlights only for digitizer notes | All highlights are removed from the student copy. |
| Footer uses a right-aligned tab stop for the page number | The footer text changes length when the copyright line is swapped. |
| Comments are resolved before conversion (except notes to digitizers) | Copies are made without comments, and suggested edits should be accepted or rejected first. |

## 2. The conversion, step by step

**auto** = done by the Student Copy Creator web page. **person** = still done by hand.

1. **Copy** the teacher copy without comments, into the same Drive folder. *(auto in the Google version; the web page works on a downloaded .docx)*
2. **Rename**: remove "Copy of", change TEACHER to STUDENT, keep the lesson number and the edition (`04. 8G U1 Writing a Strong Claim STUDENT COPY Ed2.0`). *(auto)*
3. **Delete the cover**: everything before the Name/Class line or table, including the cover's section break. *(auto)*
4. **Logos**: page 1 gets the large logo and later pages the small one. When the cover's section is deleted, its "different first page" header has to move onto the student copy's first page. *(auto in the web page)*
5. **Header** `Teacher Copy` → `Student Copy` on every header, including the first-page header. *(auto)*
6. **Footer** `© CommonLit, Inc. <year>` → the CC BY-NC-SA 4.0 line with link, page number still right-aligned and starting at 1. *(auto)*
7. **Title** `TEACHER COPY: X` → `X`. *(auto)*
8. **Teacher notes**: delete "Answers in blue", test-security notes, and Notes to Teacher boxes. *(auto)*
9. **Optional questions**: delete `*` questions with their options and answers. *(auto)* Note: some older published student copies kept them and only dropped the `*`. The how-to says delete. Confirm which rule applies to your series.
10. **Answers**:
    - Blue answer paragraphs → the same number of empty lines, same size. *(auto)*
    - Blue text at the end of a question line → removed, wrapped lines kept as space. *(auto)*
    - Blue correct choice among black choices → black, not bold. *(auto)*
    - Answer-only boxes → empty, at least 3 lines. *(auto)*
    - Blue left on empty line formatting → cleared, so student typing comes out black. *(auto)*
11. **Highlights**: remove all, including on footnote numbers. *(auto)*
12. **Check by hand** *(person)*:
    - Answer boxes are a reasonable size (the short response box runs the length of the page).
    - No question is split across two pages, and questions still line up with their paragraphs. Compare page by page with the teacher copy.
    - Nothing teacher-only is left (search the copy for "teacher", "answer", "note", and for blue).
13. **Publish** *(person)*: link in the Dig Guide (set to "force copy") and in the tracker, and turn the tracker box blue.

For lessons with independent practice, an **Internal Answer Key** is made before the student copy: a copy of the independent practice *with* comments (MCQ answer comments), named `Internal Answer Key <lesson> TEACHER COPY Ed2.0`, saved in the Answer Keys (INTERNAL) folder.

## 3. File naming and process

- Teacher copy: `<NN>. <Grade>G <Unit> <Lesson title> TEACHER COPY Ed2.0` (e.g. `04. 8G Unit 1 Writing a Strong Claim TEACHER COPY Ed2.0`).
- Student copy: the same with `STUDENT COPY`.
- Both live in the unit's FINAL folder.
- Make the student copy only after copy editing is complete (the tracker cell has gone blue).

## 4. Edge cases

| Situation | What to do |
| --- | --- |
| An example or model answer the *student* should see (e.g. "Example: Prompt … Strong claim …") | Keep it black. It's lesson content, not an answer. |
| A teacher wants to show where answers are without giving them (e.g. "Student responses will vary.") | Write it in blue. It's removed from the student copy. |
| A table where the left column is labels and the right is all answers | The right cells become empty writing boxes, at least 3 lines. Make the teacher answers realistic in length so the box is the right size. |
| A question whose answer is a choice the student circles | Make the correct choice blue. Don't add an arrow, a check mark or a highlight. |
| A picture or chart that contains an answer | The conversion can't see inside images. Put the answer as blue text, or flag it for a person. |
| Footnote glosses (vocabulary at the foot of the page) | Student-facing. Keep them. Remove only their highlights. |
| Links | Link blue `#035FE6`/`#1155CC` is not answer blue and is kept. Never use `#0000FF` for a link. |
