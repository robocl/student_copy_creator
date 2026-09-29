// Bundles apps-script/ into the two files people paste into script.google.com.
// Run: npm run build
'use strict';
const fs = require('fs');
const path = require('path');

const src = f => fs.readFileSync(path.join(__dirname, 'apps-script', f), 'utf8');
const out = path.join(__dirname, 'paste-into-google');

function bundle() {
  const header = '// Student Copy Creator. Paste this whole file into Code.gs.\n' +
    '// Generated from apps-script/ by `npm run build`; edit those files, not this one.\n\n';
  return {
    'Code.gs': header + src('Code.js') + '\n' + src('Converter.js').replace(/\nif \(typeof module[^\n]*\n?$/, '\n'),
    'Index.html': src('Index.html')
  };
}

if (require.main === module) {
  fs.mkdirSync(out, { recursive: true });
  for (const [name, text] of Object.entries(bundle())) fs.writeFileSync(path.join(out, name), text);
  console.log('Wrote paste-into-google/Code.gs and paste-into-google/Index.html');
}

module.exports = { bundle, out };
