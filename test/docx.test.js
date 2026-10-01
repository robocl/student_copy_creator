'use strict';
// End-to-end: .docx in -> .docx out, using the same code as the web page.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const StudentCopy = require('../apps-script/Converter.js');
const shared = require('../apps-script/Shared.js');
const Docx = require('../web/docx-adapter.js');
const env = () => ({ JSZip, DOMParser, XMLSerializer, StudentCopy, shared });

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function bodyOf(data) {
  const zip = await JSZip.loadAsync(data);
  const xml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  return { zip, body: new Docx.Container(body, { xml }) };
}

/** One line per paragraph, through the DocumentApp-style API. */
function outline(container, prefix = '') {
  const lines = [];
  for (let i = 0; i < container.getNumChildren(); i++) {
    const el = container.getChild(i);
    if (el.getType() === 'TABLE') {
      for (let r = 0; r < el.getNumRows(); r++) {
        const row = el.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) lines.push(...outline(row.getCell(c), `${prefix}[${r},${c}] `));
      }
    } else {
      const t = el.editAsText(), s = t.getText();
      let blue = false, bold = false;
      for (let k = 0; k < s.length; k++) {
        if (!/\S/.test(s[k])) continue;
        if (t.getForegroundColor(k) === '#0000ff') blue = true;
        if (t.isBold(k)) bold = true;
      }
      lines.push({ text: (prefix + s).replace(/\s+$/, ''), empty: !/\S/.test(s), blue, bold });
    }
  }
  return lines;
}
const collapse = lines => lines.filter((l, i) => !(l.empty && i && lines[i - 1].empty && lines[i - 1].text === l.text));

const fixtures = path.join(__dirname, 'fixtures');
const pairs = fs.existsSync(fixtures)
  ? fs.readdirSync(fixtures).filter(f => f.endsWith('.teacher.docx')).map(f => f.replace('.teacher.docx', ''))
  : [];

for (const name of pairs) {
  test(`docx golden: ${name}`, async () => {
    const teacher = fs.readFileSync(path.join(fixtures, `${name}.teacher.docx`));
    const result = await Docx.convert(teacher, `Copy of 04._8G_Unit_1_${name}_TEACHER_COPY_Ed2.0.docx`, env());
    fs.writeFileSync(path.join(fixtures, `${name}.converted.docx`), result.data);

    const got = collapse(outline((await bodyOf(result.data)).body));
    const want = collapse(outline((await bodyOf(fs.readFileSync(path.join(fixtures, `${name}.student.docx`)))).body));
    assert.deepStrictEqual(got.map(l => l.text), want.map(l => l.text));
    assert.deepStrictEqual(got.map(l => l.bold), want.map(l => l.bold), 'bold matches');
    assert.ok(!got.some(l => l.blue), 'no blue answer text left');
    assert.strictEqual(result.name, `04._8G_Unit_1_${name}_STUDENT_COPY_Ed2.0.docx`);
  });
}

test('docx: footer copyright becomes the CC line with a real link, header says Student Copy', async () => {
  const zip = new JSZip();
  const doc = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${inner}<w:sectPr/></w:body></w:document>`;
  const run = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('word/document.xml', doc(
    `<w:p/><w:tbl><w:tr><w:tc><w:p>${run('Name')}</w:p></w:tc></w:tr></w:tbl>` +
    `<w:p>${run('TEACHER COPY: Button, Button', '<w:b/>')}</w:p>` +
    `<w:p>${run('*A. Find Evidence: Highlight three details')}</w:p>` +
    `<w:p>${run('Keep ')}${run('highlighted', '<w:highlight w:val="yellow"/>')}</w:p>`));
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`);
  zip.file('word/footer1.xml', `<w:ftr xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p>${run('Unit 1: The Art of Suspense     ')}${run('© CommonLit, Inc. 2026', '<w:color w:val="333333"/>')}</w:p></w:ftr>`);
  zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p>${run('Teacher Copy')}</w:p></w:hdr>`);

  const result = await Docx.convert(await zip.generateAsync({ type: 'uint8array' }), 'X TEACHER COPY Ed2.0.docx', env());
  const out = await JSZip.loadAsync(result.data);
  const footer = await out.file('word/footer1.xml').async('string');
  const rels = await out.file('word/_rels/footer1.xml.rels').async('string');
  const header = await out.file('word/header1.xml').async('string');
  const docXml = await out.file('word/document.xml').async('string');

  assert.match(footer, /Unless otherwise noted, this content is licensed under the /);
  assert.doesNotMatch(footer, /©/);
  const rid = footer.match(/<w:hyperlink r:id="([^"]+)"/)[1];
  assert.match(rels, new RegExp(`Id="${rid}"[^>]*Target="https://creativecommons.org/licenses/by-nc-sa/4.0/"`));
  assert.match(header, /Student Copy/);
  assert.match(docXml, />Button, Button</);
  assert.doesNotMatch(docXml, /Find Evidence|w:highlight/);
  assert.strictEqual(result.name, 'X STUDENT COPY Ed2.0.docx');
});

test('docx: removes highlights on footnote numbers and paragraph marks, and blue on blank lines', async () => {
  const zip = new JSZip();
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
    `<w:p><w:r><w:t>Name: ________   Class: ________</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">[1] One dollar and eighty-seven cents.</w:t></w:r>` +
    `<w:r><w:rPr><w:highlight w:val="magenta"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>` +
    `<w:p><w:pPr><w:rPr><w:color w:val="0000ff"/><w:highlight w:val="yellow"/></w:rPr></w:pPr></w:p>` +
    `<w:sectPr/></w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`);
  zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:rPr><w:highlight w:val="magenta"/></w:rPr><w:t>a sum of money</w:t></w:r></w:p></w:footnote></w:footnotes>`);

  const result = await Docx.convert(await zip.generateAsync({ type: 'uint8array' }), 'Magi TEACHER COPY.docx', env());
  const out = await JSZip.loadAsync(result.data);
  const docXml = await out.file('word/document.xml').async('string');
  const notes = await out.file('word/footnotes.xml').async('string');
  assert.doesNotMatch(docXml, /w:highlight/);
  assert.doesNotMatch(docXml, /0000ff/i);
  assert.match(docXml, /footnoteReference/, 'footnote number itself is kept');
  assert.doesNotMatch(notes, /w:highlight/);
  assert.ok(!result.manualChecks.some(c => /logo/i.test(c)), 'logo step is not left for people to do');
  assert.ok(result.warnings.some(w => /logo/i.test(w)), 'but they are told when page 1 may have the wrong logo');
});

test('docx: page number stays right-aligned after the CC line replaces the copyright', async () => {
  const zip = new JSZip();
  const run = t => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
  const tab = '<w:r><w:tab/></w:r>';
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p>${run('Name: ______  Class: ______')}</w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="720" w:right="720" w:top="360" w:bottom="0"/></w:sectPr></w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer3.xml"/></Relationships>`);
  zip.file('word/footer3.xml', `<w:ftr xmlns:w="${W}"><w:p>${run('© CommonLit, Inc. 2026')}${tab.repeat(12)}${run('Page ')}` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`);
  const result = await Docx.convert(await zip.generateAsync({ type: 'uint8array' }), 'X TEACHER COPY.docx', env());
  const footer = await (await JSZip.loadAsync(result.data)).file('word/footer3.xml').async('string');
  assert.match(footer, /Unless otherwise noted/);
  assert.strictEqual((footer.match(/<w:tab\/>/g) || []).length, 1, 'one tab before the page number');
  assert.match(footer, /<w:tab w:val="right" w:pos="10800"\/>/, 'right tab stop at the right margin');
});
