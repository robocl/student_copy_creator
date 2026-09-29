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
