'use strict';
// Run with: node --test test/
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const mock = require('./mock-docs');

global.DocumentApp = mock.DocumentApp;
const StudentCopy = require('../apps-script/Converter.js');

const BLUE = '#0000ff';
const p = (text, extra = {}) => ({ type: 'PARAGRAPH', runs: text ? [{ text, color: extra.color || null, bold: !!extra.bold }] : [], ...extra });
const li = (text, level, extra = {}) => ({ ...p(text, extra), type: 'LIST_ITEM', level, listId: 'L1' });
const runs = (...rs) => ({ type: 'PARAGRAPH', runs: rs.map(([text, color]) => ({ text, color: color || null })) });
const table = (...rows) => ({ type: 'TABLE', rows });
const nameTable = table([[p('Name')], [p('')], [p('Class')], [p('')]]);

function convert(json) {
  const body = mock.bodyFromJson(json);
  const report = StudentCopy.convertBody(body);
  return { body, report, lines: mock.outline(body) };
}

test('removes cover pages before the Name table and keeps one empty paragraph', () => {
  const { lines, report } = convert([
    p('LESSON OVERVIEW: BUTTON'), p('(This page does not appear on the student copy.)'),
    table([[p('Part')], [p('Time')]]), p('', { pageBreak: true }),
    nameTable, p('TEACHER COPY: Button, Button')
  ]);
  assert.deepStrictEqual(lines.map(l => l.text), ['', '[0,0] Name', '[0,1] ', '[0,2] Class', '[0,3] ', 'Button, Button']);
  assert.strictEqual(report.coverElementsRemoved, 3);
  assert.strictEqual(report.warnings.length, 0);
});

test('warns when there is no Name table', () => {
  const { report } = convert([p('Something')]);
  assert.match(report.warnings[0], /Name \/ Class/);
});

test('removes teacher-only notes', () => {
  const { lines } = convert([
    p(''), nameTable,
    p('*Answers in blue. To help us ensure assessment security, please do not post or circulate these answers online.*', { color: BLUE }),
    p('Directions: Answer the multiple choice questions.'),
    runs(['Note:', BLUE], [' To ensure test security, the following items are viewable only on commonlit.org.', BLUE]),
    p('1. What is the meaning of "restrained"?')
  ]);
  assert.deepStrictEqual(lines.slice(5).map(l => l.text), ['Directions: Answer the multiple choice questions.', '1. What is the meaning of "restrained"?']);
});

test('un-highlights a blue choice among black list options', () => {
  const { lines, report } = convert([
    p(''), nameTable,
    li('Claim 1: In "Button, Button," Richard Matheson builds suspense.', 0),
    li('This claim summarizes too much.', 1),
    li('This claim just restates the prompt.', 1, { color: BLUE, bold: true }),
    li('This claim skips right to the evidence.', 1)
  ]);
  const choice = lines.find(l => l.text === 'This claim just restates the prompt.');
  assert.ok(choice, 'choice text kept');
  assert.strictEqual(choice.answerColored, false);
  assert.strictEqual(choice.bold, false);
  assert.strictEqual(report.choicesUnmarked, 1);
  assert.strictEqual(report.answersBlanked, 0);
});

test('blanks bulleted answers in a table cell and leaves writing space', () => {
  const { lines, report } = convert([
    p(''), nameTable,
    table([
      [p('A. Find Evidence: Highlight three details that show Norma is not interested.')],
      [li('"Norma repressed a smile. She was sure now it was a sales pitch." (10)', 0, { color: BLUE }),
       li('"I\'m rather busy" (12)', 0, { color: BLUE }),
       p(''),
       li('"I\'ll get you your whatchamacallit" (12)', 0, { color: BLUE })]
    ])
  ]);
  const cell = lines.filter(l => l.text.startsWith('[0,1]'));
  assert.ok(cell.every(l => l.empty && !l.list), 'answers replaced by plain empty lines');
  assert.ok(cell.length >= StudentCopy.DEFAULTS.minAnswerLines);
  assert.strictEqual(report.answersBlanked, 1, 'one block, blank line in the middle absorbed');
  assert.ok(!lines.some(l => l.answerColored));
});

test('removes inline blue answers but keeps the question', () => {
  const { lines } = convert([
    p(''), nameTable,
    runs(['How do their reactions impact the reader? ', null], ['They build tension.', BLUE])
  ]);
  assert.strictEqual(lines[lines.length - 1].text, 'How do their reactions impact the reader? ');
});

test('deletes optional (*) questions with their options and answers, but not *** dividers', () => {
  const { lines, report } = convert([
    p(''), nameTable,
    p('***'),
    p('*D. Poll the Class: Do you think Norma will push the button?'),
    li('Yes', 0), li('No', 0),
    p(''),
    p('E. Why does Norma want the money?'),
    table([[p('*A. Find Evidence: Highlight three details'), li('"She was sure now it was a sales pitch." (10)', 0, { color: BLUE }),
            p('B. Analyze: What does this show?'), li('She is busy.', 0, { color: BLUE })]]),
    p('[12] Story text that must stay.')
  ]);
  const texts = lines.map(l => l.text);
  assert.ok(texts.includes('***'));
  assert.ok(!texts.some(t => /Poll the Class|Find Evidence|^Yes$|^No$|sales pitch/.test(t)), texts.join('\n'));
  assert.ok(texts.includes('E. Why does Norma want the money?'));
  assert.ok(texts.includes('[0,0] B. Analyze: What does this show?'));
  assert.ok(texts.includes('[12] Story text that must stay.'));
  assert.strictEqual(report.optionalQuestionsRemoved, 2);
});

test('optional questions can be kept with just the * removed', () => {
  const body = mock.bodyFromJson([p(''), nameTable, p('*A. Find Evidence: Highlight three details')]);
  const report = StudentCopy.convertBody(body, { optionalQuestions: 'unmark' });
  assert.strictEqual(mock.outline(body).pop().text, 'A. Find Evidence: Highlight three details');
  assert.strictEqual(report.optionalMarkersRemoved, 1);
});

test('deletes "Notes to Teacher" boxes', () => {
  const { lines, report } = convert([
    p(''), nameTable,
    table([[p('NOTES TO TEACHER'), p('Pause here and model the first answer.')]]),
    p(''),
    p('1. What happens first?')
  ]);
  assert.ok(!lines.some(l => /Pause here/.test(l.text)));
  assert.strictEqual(report.teacherBoxesRemoved, 1);
});

test('removes highlights', () => {
  const { lines, report } = convert([
    p(''), nameTable, runs(['Plain ', null], ['digitizer highlight', null])
  ].map((el, i) => i === 2 ? { ...el, runs: [{ text: 'Plain ' }, { text: 'digitizer highlight', bg: '#ffff00' }] } : el));
  assert.ok(!lines.some(l => l.highlighted));
  assert.strictEqual(report.highlightsRemoved, 1);
});

test('header says Student Copy and footer copyright becomes the CC line', () => {
  const header = mock.footerFromJson([p('CommonLit 360 | Teacher Copy')]);
  const footer = mock.footerFromJson([p('Unit 1: The Art of Suspense    © CommonLit, Inc. 2026')]);
  const report = StudentCopy.newReport();
  StudentCopy.convertHeaderFooter(header, report);
  StudentCopy.convertHeaderFooter(footer, report);
  assert.strictEqual(header.getText(), 'CommonLit 360 | Student Copy');
  assert.strictEqual(footer.getText(), 'Unit 1: The Art of Suspense    Unless otherwise noted, this content is licensed under the CC BY-NC-SA 4.0 license.');
  const para = footer.getChild(0);
  const linked = para.chars.filter(c => c.link).map(c => c.ch).join('');
  assert.strictEqual(linked, 'CC BY-NC-SA 4.0');
  assert.strictEqual(report.headerLabelsChanged, 1);
  assert.strictEqual(report.footerLinesReplaced, 1);
});

test('leaves an already-student copy alone', () => {
  const json = [p(''), nameTable, p('Button, Button'), li('A. Yes', 0), li('B. No', 0)];
  const { lines, report } = convert(json);
  assert.deepStrictEqual(lines.map(l => l.text), mock.outline(mock.bodyFromJson(json)).map(l => l.text));
  assert.strictEqual(report.answersBlanked + report.choicesUnmarked + report.coverElementsRemoved, 0);
});

// Golden tests: drop <name>.teacher.json and <name>.student.json (made with
// test/docx_to_json.py) into test/fixtures/. Fixtures are git-ignored because
// teacher copies contain answer keys.
const fixtures = path.join(__dirname, 'fixtures');
const pairs = fs.existsSync(fixtures)
  ? fs.readdirSync(fixtures).filter(f => f.endsWith('.teacher.json')).map(f => f.replace('.teacher.json', ''))
  : [];

// Blank-line counts are a judgment call, so collapse runs of empty lines.
const normalize = lines => lines
  .map(l => ({ ...l, text: l.text.replace(/\s+$/, '') }))
  .filter((l, i, all) => !(l.empty && i > 0 && all[i - 1].empty && l.text === all[i - 1].text));

for (const name of pairs) {
  test(`golden: ${name}`, () => {
    const teacher = JSON.parse(fs.readFileSync(path.join(fixtures, `${name}.teacher.json`), 'utf8'));
    const student = JSON.parse(fs.readFileSync(path.join(fixtures, `${name}.student.json`), 'utf8'));
    const got = normalize(convert(teacher).lines);
    const want = normalize(mock.outline(mock.bodyFromJson(student)));
    assert.deepStrictEqual(got.map(l => l.text), want.map(l => l.text));
    assert.deepStrictEqual(got.map(l => l.bold), want.map(l => l.bold), 'bold matches');
    assert.ok(!got.some(l => l.answerColored), 'no blue answer text left');
  });
}

test('paste-into-google/ is up to date (run `npm run build` if this fails)', () => {
  const { bundle, out } = require('../build.js');
  for (const [name, text] of Object.entries(bundle())) {
    assert.strictEqual(fs.readFileSync(path.join(out, name), 'utf8'), text, name);
  }
});
