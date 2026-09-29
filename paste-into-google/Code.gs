// Student Copy Creator. Paste this whole file into Code.gs.
// Generated from apps-script/ by `npm run build`; edit those files, not this one.

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

/**
 * Core teacher-copy -> student-copy conversion rules.
 *
 * Everything here works on a DocumentApp Body (or anything that looks like
 * one), so it can be unit tested outside of Google with test/mock-docs.js.
 *
 * Conventions this relies on (from CommonLit Ed2.0 teacher copies):
 *   - Teacher cover pages come before the first "Name | Class" table.
 *   - The student-facing title reads "TEACHER COPY: <title>".
 *   - Answers are in blue (#0000ff).
 *   - In multiple choice / "pick one" lists, the correct option is blue (and
 *     usually bold); the other options are black.
 *   - Optional questions are marked with a leading "*", e.g. "*A. Find
 *     Evidence". Per the "Making Student Copies" how-to they only appear on
 *     the teacher copy, so they are deleted.
 *   - Teacher-only notes ("Answers in blue...", "Note: To ensure test
 *     security...", "Notes to Teacher" boxes) sit in their own paragraph or
 *     table.
 *   - Highlights are for digitizers and are removed.
 *   - Headers say "Teacher Copy"; footers of docs with answers carry a
 *     "(c) CommonLit, Inc. <year>" line that becomes the CC BY-NC-SA line.
 */
var StudentCopy = (function () {
  var DEFAULTS = {
    // Text in these colors is treated as teacher answers.
    answerColors: ['#0000ff'],
    // Whole paragraphs matching any of these are deleted.
    teacherOnlyPatterns: [
      /answers in blue/i,
      /to ensure (test|assessment) security/i,
      /this page does not appear on the student copy/i
    ],
    // Tables (e.g. orange boxes) whose text starts like this are deleted.
    teacherBoxPattern: /^\s*(notes? to (the )?teachers?|teacher notes?)\b/i,
    // Removed wherever it appears (RE2 syntax, used with replaceText).
    titlePrefixPattern: '(?i)TEACHER COPY:?\\s*',
    // Optional questions such as "*A. Find Evidence".
    optionalMarkerPattern: /^\*\s*(?=[A-Z0-9]{1,2}[.)]\s)/,
    // 'delete' removes optional questions (the how-to); 'unmark' just drops the "*".
    optionalQuestions: 'delete',
    removeHighlights: true,
    // Writing space left in place of a blanked answer.
    charsPerAnswerLine: 60,
    minAnswerLines: 5,
    maxAnswerLines: 12
  };

  var T = function () { return DocumentApp.ElementType; };

  function newReport() {
    return {
      coverElementsRemoved: 0,
      teacherNotesRemoved: 0,
      titlePrefixesRemoved: 0,
      answersBlanked: 0,
      inlineAnswersRemoved: 0,
      choicesUnmarked: 0,
      optionalMarkersRemoved: 0,
      optionalQuestionsRemoved: 0,
      teacherBoxesRemoved: 0,
      highlightsRemoved: 0,
      headerLabelsChanged: 0,
      footerLinesReplaced: 0,
      warnings: []
    };
  }

  function withDefaults(options) {
    var o = {};
    for (var k in DEFAULTS) o[k] = DEFAULTS[k];
    for (var j in (options || {})) o[j] = options[j];
    o.answerColors = o.answerColors.map(function (c) { return String(c).toLowerCase(); });
    return o;
  }

  /** Converts one Body in place. Returns a report of what changed. */
  function convertBody(body, options, report) {
    var o = withDefaults(options);
    report = report || newReport();
    removeCoverPages(body, report);
    removeTeacherBoxes(body, o, report);
    removeTeacherOnlyParagraphs(body, o, report);
    removeTitlePrefix(body, o, report);
    if (o.optionalQuestions === 'delete') removeOptionalQuestions(body, o, report);
    else removeOptionalMarkers(body, o, report);
    processAnswers(body, o, report);
    if (o.removeHighlights) removeAllHighlights(body, report);
    checkLeftovers(body, o, report);
    return report;
  }

  // ---------------------------------------------------------------- cover

  function isNameTable(el) {
    if (el.getType() !== T().TABLE) return false;
    var t = el.asTable ? el.asTable() : el;
    if (t.getNumRows() < 1 || t.getRow(0).getNumCells() < 1) return false;
    return /^\s*name\b/i.test(t.getCell(0, 0).getText());
  }

  function removeCoverPages(body, report) {
    var n = body.getNumChildren();
    var idx = -1;
    for (var i = 0; i < n; i++) {
      if (isNameTable(body.getChild(i))) { idx = i; break; }
    }
    if (idx < 0) {
      report.warnings.push('Could not find the "Name / Class" table, so no cover pages were removed.');
      return;
    }
    if (idx === 0) return;
    // Google Docs needs a paragraph before a leading table, so keep the
    // element directly before the Name table but empty it (it usually holds
    // the page break that ended the cover).
    for (var j = idx - 2; j >= 0; j--) {
      body.removeChild(body.getChild(j));
      report.coverElementsRemoved++;
    }
    var keep = body.getChild(0);
    var kt = keep.getType();
    if (kt === T().PARAGRAPH || kt === T().LIST_ITEM) {
      keep.clear();
    } else {
      try { body.removeChild(keep); report.coverElementsRemoved++; } catch (e) { /* leave it */ }
    }
    report.removedCover = true;
  }

  // -------------------------------------------------------- teacher notes

  function removeTeacherOnlyParagraphs(body, o, report) {
    var paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      var text = paras[i].getText();
      if (!text || text.length > 400) continue;
      for (var p = 0; p < o.teacherOnlyPatterns.length; p++) {
        if (o.teacherOnlyPatterns[p].test(text)) {
          safeRemove(paras[i]);
          report.teacherNotesRemoved++;
          break;
        }
      }
    }
  }

  function removeTeacherBoxes(container, o, report) {
    for (var i = container.getNumChildren() - 1; i >= 0; i--) {
      var el = container.getChild(i);
      if (el.getType() !== T().TABLE) continue;
      var table = el.asTable ? el.asTable() : el;
      if (o.teacherBoxPattern.test(table.getText())) {
        safeRemove(table);
        report.teacherBoxesRemoved++;
        continue;
      }
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        for (var c = 0; c < row.getNumCells(); c++) removeTeacherBoxes(row.getCell(c), o, report);
      }
    }
  }

  function removeTitlePrefix(body, o, report) {
    var before = countMatches(body, /teacher copy/gi);
    body.replaceText(o.titlePrefixPattern, '');
    report.titlePrefixesRemoved += before - countMatches(body, /teacher copy/gi);
  }

  function countMatches(body, re) {
    var m = body.getText().match(re);
    return m ? m.length : 0;
  }

  // -------------------------------------------------------------- answers

  /** Splits a paragraph's text into runs of {start, end, answer} (end inclusive). */
  function answerRuns(text, o) {
    var s = text.getText();
    if (!s.length) return [];
    var idx = text.getTextAttributeIndices();
    var runs = [];
    for (var i = 0; i < idx.length; i++) {
      var start = idx[i];
      var end = (i + 1 < idx.length ? idx[i + 1] : s.length) - 1;
      if (end < start) continue;
      var color = text.getForegroundColor(start);
      runs.push({
        start: start,
        end: end,
        answer: !!color && o.answerColors.indexOf(String(color).toLowerCase()) >= 0,
        blank: !/\S/.test(s.substring(start, end + 1))
      });
    }
    return runs;
  }

  /** 'none' | 'full' | 'partial' — ignores whitespace-only runs. */
  function answerKind(para, o) {
    var runs = answerRuns(para.editAsText(), o);
    var ans = false, other = false;
    for (var i = 0; i < runs.length; i++) {
      if (runs[i].blank) continue;
      if (runs[i].answer) ans = true; else other = true;
    }
    return !ans ? 'none' : (other ? 'partial' : 'full');
  }

  function processAnswers(body, o, report) {
    processContainer(body, o, report);
  }

  /**
   * Works child-by-child inside one container (the body or a table cell) so
   * neighbours can be compared by index. Mutations run back to front so
   * earlier indices stay valid.
   */
  function processContainer(container, o, report) {
    var n = container.getNumChildren();
    var kids = [];
    for (var i = 0; i < n; i++) {
      var el = container.getChild(i);
      var type = el.getType();
      if (type === T().TABLE) {
        var table = el.asTable ? el.asTable() : el;
        for (var r = 0; r < table.getNumRows(); r++) {
          var row = table.getRow(r);
          for (var c = 0; c < row.getNumCells(); c++) processContainer(row.getCell(c), o, report);
        }
        kids.push(null);
      } else if (type === T().PARAGRAPH || type === T().LIST_ITEM) {
        var para = type === T().LIST_ITEM ? (el.asListItem ? el.asListItem() : el)
                                          : (el.asParagraph ? el.asParagraph() : el);
        kids.push({ el: para, list: type === T().LIST_ITEM, kind: answerKind(para, o), text: para.getText() });
      } else {
        kids.push(null);
      }
    }

    for (var a = 0; a < kids.length; a++) {
      var k = kids[a];
      if (!k) continue;
      if (k.kind === 'partial') {
        removeInlineAnswers(k.el, o);
        report.inlineAnswersRemoved++;
      } else if (k.kind === 'full') {
        var sib = choiceSibling(kids, a);
        if (sib) {
          matchStyle(k.el, sib.el);
          k.kind = 'choice';
          report.choicesUnmarked++;
        }
      }
    }

    // Group runs of answers; blank lines between answers join the block.
    var blocks = [];
    var cur = null, pendingBlank = [];
    for (var b = 0; b < kids.length; b++) {
      var kb = kids[b];
      if (kb && kb.kind === 'full') {
        if (!cur) { cur = [b]; } else { cur = cur.concat(pendingBlank, [b]); }
        pendingBlank = [];
      } else if (cur && kb && kb.kind === 'none' && !/\S/.test(kb.text)) {
        pendingBlank.push(b);
      } else {
        if (cur) blocks.push(cur);
        cur = null;
        pendingBlank = [];
      }
    }
    if (cur) blocks.push(cur);

    for (var x = blocks.length - 1; x >= 0; x--) {
      blankAnswerBlock(container, blocks[x], kids, o);
      report.answersBlanked++;
    }
  }

  function removeInlineAnswers(para, o) {
    var text = para.editAsText();
    var runs = answerRuns(text, o);
    for (var i = runs.length - 1; i >= 0; i--) {
      if (runs[i].answer && !runs[i].blank) text.deleteText(runs[i].start, runs[i].end);
    }
  }

  var CHOICE_LABEL = /^\s*[A-H][.)]\s/;

  /**
   * If kids[i] is one option in a list of options (e.g. the correct answer
   * in blue among black alternatives), returns a black sibling option.
   */
  function choiceSibling(kids, i) {
    var me = kids[i];
    if (!me.list && !CHOICE_LABEL.test(me.text)) return null;
    var level = me.list ? me.el.getNestingLevel() : 0;
    for (var dir = -1; dir <= 1; dir += 2) {
      for (var j = i + dir; j >= 0 && j < kids.length; j += dir) {
        var k = kids[j];
        if (!k) break;
        if (me.list) {
          if (!k.list) break;
          var lv = k.el.getNestingLevel();
          if (lv < level) break;
          if (lv > level) continue;
        } else if (!CHOICE_LABEL.test(k.text)) {
          break;
        }
        if (/\S/.test(k.text) && k.kind === 'none') return k;
      }
    }
    return null;
  }

  function matchStyle(para, sibling) {
    var text = para.editAsText();
    var len = text.getText().length;
    if (!len) return;
    var st = sibling.editAsText();
    text.setForegroundColor(0, len - 1, st.getForegroundColor(0) || '#000000');
    text.setBold(0, len - 1, !!st.isBold(0));
  }

  /** Swaps a block of answer paragraphs for empty writing lines. */
  function blankAnswerBlock(container, indices, kids, o) {
    var lines = 0;
    for (var i = 0; i < indices.length; i++) {
      var len = kids[indices[i]].text.length;
      lines += Math.max(1, Math.ceil(len / o.charsPerAnswerLine));
    }
    lines = Math.min(o.maxAnswerLines, Math.max(o.minAnswerLines, lines));

    var at = indices[0];
    for (var n = 0; n < lines; n++) container.insertParagraph(at, '');
    for (var j = indices.length - 1; j >= 0; j--) container.removeChild(container.getChild(indices[j] + lines));
  }

  // ------------------------------------------------------------ markers

  var TYPED_LABEL = /^\s*\*?\s*[A-Z0-9]{1,2}[.)]\s/;

  /**
   * Deletes each "*"-marked question along with what belongs to it: blank
   * lines, blue answers and auto-numbered options (e.g. "A. Yes / B. No").
   * Stops at the next typed question label or any other content, so story
   * text is never swept up.
   */
  function removeOptionalQuestions(container, o, report) {
    for (var i = container.getNumChildren() - 1; i >= 0; i--) {
      var el = container.getChild(i);
      var type = el.getType();
      if (type === T().TABLE) {
        var table = el.asTable ? el.asTable() : el;
        for (var r = 0; r < table.getNumRows(); r++) {
          var row = table.getRow(r);
          for (var c = 0; c < row.getNumCells(); c++) removeOptionalQuestions(row.getCell(c), o, report);
        }
        continue;
      }
      if (type !== T().PARAGRAPH && type !== T().LIST_ITEM) continue;
      var text = el.getText();
      if (!o.optionalMarkerPattern.test(text)) continue;

      var end = i;
      for (var j = i + 1; j < container.getNumChildren(); j++) {
        var next = container.getChild(j);
        var nt = next.getType();
        if (nt !== T().PARAGRAPH && nt !== T().LIST_ITEM) break;
        var ntext = next.getText();
        var belongs = !/\S/.test(ntext) ||
          (!TYPED_LABEL.test(ntext) && (nt === T().LIST_ITEM || answerKind(next, o) === 'full'));
        if (!belongs) break;
        end = j;
      }
      // Leave trailing blank lines alone so spacing before the next item stays.
      while (end > i && !/\S/.test(container.getChild(end).getText())) end--;
      for (var k = end; k >= i; k--) safeRemove(container.getChild(k));
      report.optionalQuestionsRemoved++;
      report.warnings.push('Deleted optional question "' + text.substring(0, 50) +
        '…". Check the questions still line up with their paragraphs.');
    }
  }

  function removeOptionalMarkers(body, o, report) {
    var paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      var m = paras[i].getText().match(o.optionalMarkerPattern);
      if (m && m[0].length) {
        paras[i].editAsText().deleteText(0, m[0].length - 1);
        report.optionalMarkersRemoved++;
      }
    }
  }

  // ---------------------------------------------------------- highlights

  function removeAllHighlights(body, report) {
    var paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      var text = paras[i].editAsText();
      var len = text.getText().length;
      if (!len) continue;
      var idx = text.getTextAttributeIndices();
      for (var j = idx.length - 1; j >= 0; j--) {
        var start = idx[j];
        var end = (j + 1 < idx.length ? idx[j + 1] : len) - 1;
        if (end >= start && text.getBackgroundColor(start)) {
          text.setBackgroundColor(start, end, null);
          report.highlightsRemoved++;
        }
      }
    }
  }

  // ----------------------------------------------------- headers/footers

  var LICENSE_LINK_TEXT = 'CC BY-NC-SA 4.0';
  var LICENSE_TEXT = 'Unless otherwise noted, this content is licensed under the ' + LICENSE_LINK_TEXT + ' license.';
  var LICENSE_URL = 'https://creativecommons.org/licenses/by-nc-sa/4.0/';
  var COPYRIGHT_PATTERN = '(©|\\(c\\))\\s*CommonLit,?\\s*Inc\\.?,?\\s*\\d{4}';

  /**
   * For a header or footer section: "Teacher Copy" -> "Student Copy", and the
   * "(c) CommonLit, Inc. 2026" line -> the Creative Commons line.
   */
  function convertHeaderFooter(section, report) {
    report = report || newReport();
    var before = (section.getText().match(/teacher copy/gi) || []).length;
    section.replaceText('TEACHER COPY', 'STUDENT COPY');
    section.replaceText('(?i)teacher copy', 'Student Copy');
    report.headerLabelsChanged += before;

    var found;
    while ((found = section.findText(COPYRIGHT_PATTERN))) {
      var t = found.getElement().asText();
      var s = found.getStartOffset(), e = found.getEndOffsetInclusive();
      t.insertText(e + 1, LICENSE_TEXT);
      t.deleteText(s, e);
      var ls = s + LICENSE_TEXT.indexOf(LICENSE_LINK_TEXT);
      t.setLinkUrl(ls, ls + LICENSE_LINK_TEXT.length - 1, LICENSE_URL);
      report.footerLinesReplaced++;
    }
    return report;
  }

  // ------------------------------------------------------------ helpers

  /**
   * Removes a paragraph or table. Google Docs won't let a container lose its
   * last paragraph or let two tables touch, so in those cases an empty
   * paragraph is left in its place.
   */
  function safeRemove(el) {
    var parent = el.getParent();
    var idx = parent.getChildIndex(el);
    var n = parent.getNumChildren();
    var prev = idx > 0 ? parent.getChild(idx - 1) : null;
    var next = idx < n - 1 ? parent.getChild(idx + 1) : null;
    var isTable = function (x) { return x && x.getType() === T().TABLE; };
    var needsPlaceholder = !next || (isTable(next) && (!prev || isTable(prev))) ||
      (el.getType() === T().TABLE && isTable(prev) && isTable(next));
    if (needsPlaceholder) parent.insertParagraph(idx + 1, '');
    parent.removeChild(el);
  }

  function checkLeftovers(body, o, report) {
    var paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      if (answerKind(paras[i], o) !== 'none') {
        report.warnings.push('Blue text is still present: "' + paras[i].getText().substring(0, 60) + '"');
      }
    }
    for (var n = 0; n < paras.length; n++) {
      if (o.teacherBoxPattern.test(paras[n].getText())) {
        report.warnings.push('A "Notes to Teacher" section is still there (not in its own box), so it was left alone. Delete it by hand.');
      }
    }
    if (/teacher/i.test(body.getText())) {
      report.warnings.push('The word "teacher" still appears in the document. Worth a quick look.');
    }
  }

  return {
    DEFAULTS: DEFAULTS,
    newReport: newReport,
    convertBody: convertBody,
    convertHeaderFooter: convertHeaderFooter
  };
})();

