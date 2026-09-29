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

test('removes the optional-question asterisk but not *** dividers', () => {
  const { lines, report } = convert([
    p(''), nameTable, p('***'), p('*A. Find Evidence: Highlight three details'), p('*D. Poll the Class: Do you think Norma will push the button?')
  ]);
  assert.deepStrictEqual(lines.slice(5).map(l => l.text), ['***', 'A. Find Evidence: Highlight three details', 'D. Poll the Class: Do you think Norma will push the button?']);
  assert.strictEqual(report.optionalMarkersRemoved, 2);
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
