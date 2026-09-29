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
 *   - Optional questions are marked with a leading "*", e.g. "*A. Find Evidence".
 *   - Teacher-only notes ("Answers in blue...", "Note: To ensure test
 *     security...") sit in their own paragraphs.
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
    // Removed wherever it appears (RE2 syntax, used with replaceText).
    titlePrefixPattern: '(?i)TEACHER COPY:?\\s*',
    // Leading "*" on optional questions such as "*A. Find Evidence".
    optionalMarkerPattern: /^\*\s*(?=[A-Z0-9]{1,2}[.)]\s)/,
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
    removeTeacherOnlyParagraphs(body, o, report);
    removeTitlePrefix(body, o, report);
    processAnswers(body, o, report);
    removeOptionalMarkers(body, o, report);
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

  // ------------------------------------------------------------ helpers

  function safeRemove(para) {
    var parent = para.getParent();
    var idx = parent.getChildIndex(para);
    var last = idx === parent.getNumChildren() - 1;
    var prev = idx > 0 ? parent.getChild(idx - 1) : null;
    var next = last ? null : parent.getChild(idx + 1);
    var betweenTables = next && next.getType() === T().TABLE && (!prev || prev.getType() === T().TABLE);
    if (last || betweenTables || parent.getNumChildren() <= 1) {
      para.clear();
    } else {
      para.removeFromParent();
    }
  }

  function checkLeftovers(body, o, report) {
    var paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      if (answerKind(paras[i], o) !== 'none') {
        report.warnings.push('Blue text is still present: "' + paras[i].getText().substring(0, 60) + '"');
      }
    }
    if (/teacher/i.test(body.getText())) {
      report.warnings.push('The word "teacher" still appears in the document. Worth a quick look.');
    }
  }

  return {
    DEFAULTS: DEFAULTS,
    newReport: newReport,
    convertBody: convertBody
  };
})();

if (typeof module !== 'undefined') module.exports = StudentCopy;
