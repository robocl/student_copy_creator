/**
 * Entry points: a web app (paste a Google Doc link) and a Docs menu item.
 * Both make a copy of the teacher copy next to it in Drive, convert the
 * copy, and never touch the original.
 */

// ------------------------------------------------------------- web app

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Copy Creator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called from Index.html. Accepts a Google Doc URL or file ID. */
function convertFromUrl(urlOrId) {
  var id = extractDocId(urlOrId);
  if (!id) throw new Error('That does not look like a Google Doc link.');
  return createStudentCopy(id);
}

function extractDocId(s) {
  s = String(s || '').trim();
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(s) ? s : null;
}

// ---------------------------------------------------------- docs menu

function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Create student copy', 'menuCreateStudentCopy')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function menuCreateStudentCopy() {
  var ui = DocumentApp.getUi();
  var result = createStudentCopy(DocumentApp.getActiveDocument().getId());
  var html = HtmlService.createHtmlOutput(resultHtml(result)).setWidth(460).setHeight(440);
  ui.showModalDialog(html, 'Student copy created');
}

function resultHtml(r) {
  var esc = function (s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  };
  var h = '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p><a href="' + esc(r.url) + '" target="_blank">Open ' + esc(r.name) + '</a></p>' +
    '<ul>' + r.summary.map(function (s) { return '<li>' + esc(s) + '</li>'; }).join('') + '</ul>';
  if (r.warnings.length) {
    h += '<p><b>Please check:</b></p><ul>' +
      r.warnings.map(function (s) { return '<li>' + esc(s) + '</li>'; }).join('') + '</ul>';
  }
  h += '<p><b>Still do by hand:</b></p><ul>' +
    r.manualChecks.map(function (s) { return '<li>' + esc(s) + '</li>'; }).join('') + '</ul>';
  return h + '</div>';
}

// --------------------------------------------------------------- core

function createStudentCopy(teacherDocId) {
  var src = DriveApp.getFileById(teacherDocId);
  if (src.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('"' + src.getName() + '" is not a Google Doc. If it is a .docx, open it in Google Docs and use File > Save as Google Docs first.');
  }
  var parents = src.getParents();
  var name = studentCopyName(src.getName());
  var copy = parents.hasNext() ? src.makeCopy(name, parents.next()) : src.makeCopy(name);

  var doc = DocumentApp.openById(copy.getId());
  var report = StudentCopy.newReport();
  var bodies = allBodies(doc);
  for (var i = 0; i < bodies.length; i++) StudentCopy.convertBody(bodies[i], null, report);
  var sections = headerFooterSections(doc);
  for (var h = 0; h < sections.length; h++) StudentCopy.convertHeaderFooter(sections[h], report);
  doc.saveAndClose();

  return {
    id: copy.getId(),
    url: copy.getUrl(),
    name: name,
    summary: summarize(report),
    warnings: report.warnings,
    manualChecks: MANUAL_CHECKS
  };
}

/** Steps from "HOW TO: Making Student Copies" that still need a person. */
var MANUAL_CHECKS = [
  'Page 1 has the large CommonLit logo and the other pages have the small one.',
  'Answer boxes are a reasonable size (the short response box should run the length of the page).',
  'No question is split across two pages, and questions still line up with their paragraphs.',
  'Link the student copy in the Dig Guide (set to "force copy") and in the tracker, then turn the box blue.'
];

function studentCopyName(name) {
  var n = name.replace(/^Copy of\s+/i, '');
  var out = n.replace(/TEACHER([ _-])COPY/i, function (m, sep) {
    return m === m.toUpperCase() ? 'STUDENT' + sep + 'COPY' : 'Student' + sep + 'Copy';
  });
  return out === n ? n + ' (Student Copy)' : out;
}

/** The main body, plus every tab's body for docs that use tabs. */
function allBodies(doc) {
  if (typeof doc.getTabs !== 'function') return [doc.getBody()];
  var bodies = [];
  var visit = function (tabs) {
    for (var i = 0; i < tabs.length; i++) {
      if (tabs[i].getType() === DocumentApp.TabType.DOCUMENT_TAB) bodies.push(tabs[i].asDocumentTab().getBody());
      visit(tabs[i].getChildTabs());
    }
  };
  visit(doc.getTabs());
  return bodies.length ? bodies : [doc.getBody()];
}

/**
 * Every header and footer, including the separate first-page ones (which
 * doc.getHeader()/getFooter() don't return).
 */
function headerFooterSections(doc) {
  var out = [];
  var root = doc.getBody().getParent();
  for (var i = 0; i < root.getNumChildren(); i++) {
    var child = root.getChild(i);
    var type = child.getType();
    if (type === DocumentApp.ElementType.HEADER_SECTION) out.push(child.asHeaderSection());
    if (type === DocumentApp.ElementType.FOOTER_SECTION) out.push(child.asFooterSection());
  }
  return out;
}

function summarize(r) {
  var out = [];
  var add = function (n, one, many) { if (n) out.push(n + ' ' + (n === 1 ? one : many)); };
  add(r.coverElementsRemoved, 'cover-page element removed', 'cover-page elements removed');
  add(r.teacherNotesRemoved, 'teacher note removed', 'teacher notes removed');
  add(r.titlePrefixesRemoved, '"TEACHER COPY" label removed', '"TEACHER COPY" labels removed');
  add(r.answersBlanked, 'answer replaced with writing space', 'answers replaced with writing space');
  add(r.inlineAnswersRemoved, 'inline answer removed', 'inline answers removed');
  add(r.choicesUnmarked, 'highlighted choice un-highlighted', 'highlighted choices un-highlighted');
  add(r.optionalMarkersRemoved, 'optional-question "*" removed', 'optional-question "*" markers removed');
  add(r.optionalQuestionsRemoved, 'optional (*) question deleted', 'optional (*) questions deleted');
  add(r.teacherBoxesRemoved, '"Notes to Teacher" box deleted', '"Notes to Teacher" boxes deleted');
  add(r.highlightsRemoved, 'highlight removed', 'highlights removed');
  add(r.headerLabelsChanged, 'header changed to "Student Copy"', 'headers changed to "Student Copy"');
  add(r.footerLinesReplaced, 'footer copyright line changed to the CC BY-NC-SA line', 'footer copyright lines changed to the CC BY-NC-SA line');
  if (!out.length) out.push('No changes were needed.');
  return out;
}
