/**
 * Helpers shared by the Google Apps Script version (Code.js) and the
 * .docx web page (web/). No Google APIs in here.
 */

/** "Copy of 04. X TEACHER COPY Ed2.0" -> "04. X STUDENT COPY Ed2.0" */
function studentCopyName(name) {
  var n = name.replace(/^Copy of\s+/i, '');
  var out = n.replace(/TEACHER([ _-])COPY/i, function (m, sep) {
    return m === m.toUpperCase() ? 'STUDENT' + sep + 'COPY' : 'Student' + sep + 'Copy';
  });
  return out === n ? n + ' (Student Copy)' : out;
}

/** Steps from "HOW TO: Making Student Copies" that still need a person. */
var MANUAL_CHECKS = [
  'Page 1 has the large CommonLit logo and the other pages have the small one.',
  'Answer boxes are a reasonable size (the short response box should run the length of the page).',
  'No question is split across two pages, and questions still line up with their paragraphs.',
  'Link the student copy in the Dig Guide (set to "force copy") and in the tracker, then turn the box blue.'
];

/** Plain-language list of what changed, for the result card. */
function summarize(r) {
  var out = [];
  var add = function (n, one, many) { if (n) out.push(n === 1 ? one : many.replace('#', n)); };
  if (r.coverElementsRemoved) out.push('Teacher cover page removed');
  add(r.titlePrefixesRemoved, '"TEACHER COPY" removed from the title', '"TEACHER COPY" removed from # places');
  add(r.teacherNotesRemoved, '1 teacher-only note removed ("Answers in blue", test security)', '# teacher-only notes removed ("Answers in blue", test security)');
  add(r.answersBlanked, '1 blue answer erased and replaced with blank lines', '# blue answers erased and replaced with blank lines');
  add(r.inlineAnswersRemoved, '1 blue answer erased from the end of a question line', '# blue answers erased from the end of question lines');
  add(r.choicesUnmarked, '1 correct choice changed from blue to black', '# correct choices changed from blue to black');
  add(r.optionalMarkersRemoved, '1 optional-question "*" removed', '# optional-question "*" marks removed');
  add(r.optionalQuestionsRemoved, '1 optional (*) question deleted', '# optional (*) questions deleted');
  add(r.teacherBoxesRemoved, '1 "Notes to Teacher" box deleted', '# "Notes to Teacher" boxes deleted');
  if (r.highlightsRemoved) out.push('Highlights removed');
  if (r.headerLabelsChanged) out.push('Header changed from "Teacher Copy" to "Student Copy"');
  if (r.footerLinesReplaced) out.push('Footer copyright line changed to the CC BY-NC-SA 4.0 line');
  if (!out.length) out.push('No changes were needed.');
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = { studentCopyName: studentCopyName, summarize: summarize, MANUAL_CHECKS: MANUAL_CHECKS };
}
