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
      /answers (are )?in blue/i,
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
    // Blanked answers keep the teacher copy's spacing: each answer becomes as
    // many empty lines as it filled. A box that held only an answer gets at
    // least minAnswerLines. charsPerAnswerLine is used when the line width
    // can't be read from the document.
    charsPerAnswerLine: 60,
    minAnswerLines: 3
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

  // "Name: ________   Class: ________" typed as a line (newer templates).
  function isNameLine(el) {
    if (el.getType() !== T().PARAGRAPH && el.getType() !== T().LIST_ITEM) return false;
    var text = el.getText();
    return /^\s*name\s*:?\s*_{3,}/i.test(text) || /^\s*name\s*:.*\bclass\s*:/i.test(text);
  }

  function removeCoverPages(body, report) {
    var n = body.getNumChildren();
    var idx = -1;
    var line = false;
    for (var i = 0; i < n; i++) {
      var child = body.getChild(i);
      if (isNameTable(child)) { idx = i; break; }
      if (isNameLine(child)) { idx = i; line = true; break; }
    }
    if (idx < 0) {
      report.warnings.push('Could not find the "Name / Class" line where the student part starts, so the cover page was not removed. Delete it by hand.');
      return;
    }
    if (idx === 0) return;
    if (line) {
      for (var k = idx - 1; k >= 0; k--) {
        body.removeChild(body.getChild(k));
        report.coverElementsRemoved++;
      }
      report.removedCover = true;
      return;
    }
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

    // Edits that add or remove paragraphs run afterwards, back to front.
    var ops = [];
    var width = null, widthRead = false;
    var getWidth = function () { if (!widthRead) { width = lineWidth(container); widthRead = true; } return width; };
    for (var a = 0; a < kids.length; a++) {
      var k = kids[a];
      if (!k) continue;
      if (k.kind === 'partial') {
        var fmtP = lineFormat(k.el);
        var before = estimateLines(k.text, fmtP, getWidth(), o);
        removeInlineAnswers(k.el, o);
        var extra = before - estimateLines(k.el.getText(), fmtP, getWidth(), o);
        if (extra > 0) ops.push({ at: a, extra: extra, fmt: fmtP });
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

    blocks.forEach(function (block) { ops.push({ at: block[0], block: block }); report.answersBlanked++; });
    ops.sort(function (p, q) { return q.at - p.at; });
    ops.forEach(function (op) {
      if (op.block) {
        blankAnswerBlock(container, op.block, kids, o);
      } else {
        // Keep the lines the erased answer used to wrap onto.
        for (var n = 0; n < op.extra; n++) applyFormat(container.insertParagraph(op.at + 1, ''), op.fmt);
      }
    });
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

  /**
   * Swaps each answer paragraph for the same number of empty lines, in the
   * same font size and spacing, so the layout matches the teacher copy (for
   * example, questions stay lined up beside their story paragraphs). A table
   * cell that held nothing but the answer is a writing box and gets at least
   * minAnswerLines lines.
   */
  function blankAnswerBlock(container, indices, kids, o) {
    var isBox = container.getType() === T().TABLE_CELL;
    for (var c = 0; c < kids.length && isBox; c++) {
      if (kids[c] && indices.indexOf(c) < 0 && /\S/.test(kids[c].text)) isBox = false;
      if (!kids[c]) isBox = false;
    }
    var width = lineWidth(container);
    var fmt = null;
    for (var j = indices.length - 1; j >= 0; j--) {
      var i = indices[j], k = kids[i];
      if (!/\S/.test(k.text)) continue; // blank lines inside the block stay as they are
      fmt = lineFormat(k.el);
      var lines = estimateLines(k.text, fmt, width, o);
      for (var n = 0; n < lines; n++) applyFormat(container.insertParagraph(i + 1, ''), fmt);
      container.removeChild(k.el);
    }
    if (isBox && fmt) {
      var empty = 0;
      for (var e = 0; e < container.getNumChildren(); e++) {
        var el = container.getChild(e);
        if (el.getType() === T().PARAGRAPH && !/\S/.test(el.getText())) empty++;
      }
      for (; empty < o.minAnswerLines; empty++) {
        applyFormat(container.insertParagraph(container.getNumChildren(), ''), fmt);
      }
    }
  }

  var ATTRS = ['FONT_SIZE', 'SPACING_BEFORE', 'SPACING_AFTER', 'LINE_SPACING'];

  /** Font size, spacing and indent of a paragraph, for sizing blank lines. */
  function lineFormat(para) {
    var A = (DocumentApp.Attribute || {});
    var attrs = para.getAttributes ? para.getAttributes() || {} : {};
    var fmt = {};
    ATTRS.forEach(function (name) { if (attrs[A[name] || name] != null) fmt[name] = attrs[A[name] || name]; });
    var text = para.editAsText();
    if (fmt.FONT_SIZE == null && text.getFontSize && text.getText().length) fmt.FONT_SIZE = text.getFontSize(0);
    fmt.indent = attrs[A.INDENT_START || 'INDENT_START'] || 0;
    return fmt;
  }

  function applyFormat(para, fmt) {
    if (!para.setAttributes) return;
    var A = (DocumentApp.Attribute || {});
    var attrs = {};
    ATTRS.forEach(function (name) { if (fmt[name] != null) attrs[A[name] || name] = fmt[name]; });
    para.setAttributes(attrs);
  }

  /** Usable line width in points, or null if the document doesn't say. */
  function lineWidth(container) {
    try {
      if (container.getType() === T().TABLE_CELL && container.getWidth) {
        var w = container.getWidth();
        if (w) return w - (container.getPaddingLeft ? (container.getPaddingLeft() || 0) + (container.getPaddingRight() || 0) : 10);
      }
      if (container.getPageWidth) {
        var pw = container.getPageWidth();
        if (pw) return pw - (container.getMarginLeft() || 0) - (container.getMarginRight() || 0);
      }
    } catch (e) { /* fall back to charsPerAnswerLine */ }
    return null;
  }

  /** How many lines the answer text filled in the teacher copy. */
  function estimateLines(text, fmt, width, o) {
    var perLine = o.charsPerAnswerLine;
    if (width) {
      var charWidth = (fmt.FONT_SIZE || 11) * 0.5; // average for Arial / Open Sans
      perLine = Math.max(10, Math.floor((width - (fmt.indent || 0)) / charWidth));
    }
    return Math.max(1, Math.ceil(text.length / perLine));
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

if (typeof module !== 'undefined') module.exports = StudentCopy;
