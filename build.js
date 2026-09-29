// Builds the two shareable forms of the tool from apps-script/ and web/:
//   paste-into-google/  Code.gs + Index.html to paste into script.google.com
//   dist/student-copy-creator.html  the drop-a-.docx web page
// Run: npm run build
'use strict';
const fs = require('fs');
const path = require('path');

const read = f => fs.readFileSync(path.join(__dirname, f), 'utf8');
const stripExport = s => s.replace(/\nif \(typeof module !== 'undefined'\)[\s\S]*$/, '\n');

function bundle() {
  const header = '// Student Copy Creator. Paste this whole file into Code.gs.\n' +
    '// Generated from apps-script/ by `npm run build`; edit those files, not this one.\n\n';
  const rules = ['apps-script/Shared.js', 'apps-script/Converter.js'].map(f => stripExport(read(f))).join('\n');
  const inline = rules + '\n' + stripExport(read('web/docx-adapter.js'));
  if (/<\/script/i.test(inline)) throw new Error('Inlined code contains </script>');
  return {
    'paste-into-google/Code.gs': header + read('apps-script/Code.js') + '\n' + rules,
    'paste-into-google/Index.html': read('apps-script/Index.html'),
    'dist/student-copy-creator.html': read('web/page.html').replace('/*INLINE_SCRIPTS*/', () => inline)
  };
}

if (require.main === module) {
  for (const [name, text] of Object.entries(bundle())) {
    fs.mkdirSync(path.dirname(path.join(__dirname, name)), { recursive: true });
    fs.writeFileSync(path.join(__dirname, name), text);
    console.log('Wrote ' + name);
  }
}

module.exports = { bundle };
