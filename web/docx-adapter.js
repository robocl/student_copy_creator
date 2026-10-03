/**
 * Lets the Apps Script conversion rules (apps-script/Converter.js) run on a
 * .docx file, entirely in the browser.
 *
 * It wraps the Word XML in objects with the same methods Converter.js calls
 * on Google's DocumentApp (Body, Paragraph, ListItem, Table, Text, ...), so
 * both versions share one set of rules.
 *
 * Needs: StudentCopy (Converter.js), Shared.js helpers, JSZip, and a
 * DOMParser/XMLSerializer (browser globals, or @xmldom/xmldom in Node).
 */
var DocxStudentCopy = (function () {
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var HYPERLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

  var ET = {
    BODY_SECTION: 'BODY_SECTION', HEADER_SECTION: 'HEADER_SECTION', FOOTER_SECTION: 'FOOTER_SECTION',
    PARAGRAPH: 'PARAGRAPH', LIST_ITEM: 'LIST_ITEM', TABLE: 'TABLE', TABLE_ROW: 'TABLE_ROW', TABLE_CELL: 'TABLE_CELL'
  };

  var ATTRIBUTE = { FONT_SIZE: 'FONT_SIZE', SPACING_BEFORE: 'SPACING_BEFORE', SPACING_AFTER: 'SPACING_AFTER',
    LINE_SPACING: 'LINE_SPACING', INDENT_START: 'INDENT_START' };

  // Schema order of <w:rPr> children; Word rejects files that break it.
  var RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
    'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
    'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];

  // ------------------------------------------------------------ xml utils

  function isW(node, name) {
    return node && node.nodeType === 1 && node.namespaceURI === W && (!name || node.localName === name);
  }
  function kidsW(node, name) {
    var out = [];
    for (var c = node.firstChild; c; c = c.nextSibling) if (isW(c, name)) out.push(c);
    return out;
  }
  function firstW(node, name) { return kidsW(node, name)[0] || null; }
  function wAttr(node, name) {
    if (!node) return null;
    var v = node.getAttributeNS(W, name);
    return v === '' || v == null ? (node.getAttribute('w:' + name) || null) : v;
  }
  function el(ctx, name, attrs) {
    var e = ctx.xml.createElementNS(W, 'w:' + name);
    for (var k in (attrs || {})) e.setAttributeNS(W, 'w:' + k, attrs[k]);
    return e;
  }
  function descendantsW(node, name) {
    var list = node.getElementsByTagNameNS(W, name), out = [];
    for (var i = 0; i < list.length; i++) out.push(list[i]);
    return out;
  }
  function hasAncestor(node, stop, names) {
    for (var n = node.parentNode; n && n !== stop; n = n.parentNode) {
      if (isW(n) && names.indexOf(n.localName) >= 0) return true;
    }
    return false;
  }

  /** Splits every run so it holds at most one piece of content (one <w:t>, one tab, one drawing...). */
  function splitRuns(root) {
    descendantsW(root, 'r').forEach(function (r) {
      var rPr = firstW(r, 'rPr');
      var content = [];
      for (var c = r.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c !== rPr) content.push(c);
      var after = r.nextSibling;
      for (var i = 1; i < content.length; i++) {
        var nr = r.cloneNode(false);
        if (rPr) nr.appendChild(rPr.cloneNode(true));
        nr.appendChild(content[i]);
        r.parentNode.insertBefore(nr, after);
      }
    });
  }

  function runText(r) {
    for (var c = r.firstChild; c; c = c.nextSibling) {
      if (isW(c, 't')) return c.textContent || '';
      if (isW(c, 'tab')) return '\t';
    }
    return '';
  }

  function rPrOf(r, create, ctx) {
    var rPr = firstW(r, 'rPr');
    if (!rPr && create) {
      rPr = el(ctx, 'rPr');
      r.insertBefore(rPr, r.firstChild);
    }
    return rPr;
  }
  function setRprChild(ctx, r, name, attrs) {
    var rPr = rPrOf(r, true, ctx);
    kidsW(rPr, name).forEach(function (x) { rPr.removeChild(x); });
    var node = el(ctx, name, attrs);
    var rank = RPR_ORDER.indexOf(name);
    var before = null;
    for (var c = rPr.firstChild; c; c = c.nextSibling) {
      if (isW(c) && RPR_ORDER.indexOf(c.localName) > rank) { before = c; break; }
    }
    rPr.insertBefore(node, before);
  }
  function removeRprChild(r, name) {
    var rPr = firstW(r, 'rPr');
    if (rPr) kidsW(rPr, name).forEach(function (x) { rPr.removeChild(x); });
  }
  function onOff(node) {
    if (!node) return false;
    var v = wAttr(node, 'val');
    return v == null || (v !== '0' && v !== 'false' && v !== 'none');
  }

  // --------------------------------------------------------------- wrappers

  function wrap(node, ctx) {
    if (isW(node, 'p')) return new Para(node, ctx);
    if (isW(node, 'tbl')) return new Table(node, ctx);
    return null;
  }

  var CONTAINER_TYPES = { body: ET.BODY_SECTION, tc: ET.TABLE_CELL, hdr: ET.HEADER_SECTION, ftr: ET.FOOTER_SECTION, txbxContent: ET.TABLE_CELL };
  function wrapContainer(node, ctx) {
    return node && isW(node) && CONTAINER_TYPES[node.localName] ? new Container(node, ctx) : null;
  }

  function Base() {}
  Base.prototype.getParent = function () {
    var p = this.node.parentNode;
    return isW(p, 'tr') ? null : wrapContainer(p, this.ctx);
  };
  Base.prototype.removeFromParent = function () { this.node.parentNode.removeChild(this.node); return this; };
  Base.prototype.siblingAt = function (d) {
    var parent = this.getParent();
    if (!parent) return null;
    var i = parent.getChildIndex(this) + d;
    return i >= 0 && i < parent.getNumChildren() ? parent.getChild(i) : null;
  };
  Base.prototype.getPreviousSibling = function () { return this.siblingAt(-1); };
  Base.prototype.getNextSibling = function () { return this.siblingAt(1); };
  Base.prototype.asParagraph = Base.prototype.asListItem = Base.prototype.asTable = function () { return this; };

  // Container: body, table cell, header, footer
  function Container(node, ctx) { this.node = node; this.ctx = ctx; }
  Container.prototype = Object.create(Base.prototype);
  Container.prototype.getType = function () { return CONTAINER_TYPES[this.node.localName]; };
  Container.prototype.kids = function () {
    var out = [];
    for (var c = this.node.firstChild; c; c = c.nextSibling) if (isW(c, 'p') || isW(c, 'tbl')) out.push(c);
    return out;
  };
  Container.prototype.getNumChildren = function () { return this.kids().length; };
  Container.prototype.getChild = function (i) { return wrap(this.kids()[i], this.ctx); };
  Container.prototype.getChildIndex = function (child) { return this.kids().indexOf(child.node); };
  Container.prototype.removeChild = function (child) {
    noteSectionBreak(child.node, this.ctx);
    this.node.removeChild(child.node);
    return this;
  };
  // Width and padding in points, like DocumentApp's TableCell / Body.
  Container.prototype.getWidth = function () {
    if (!isW(this.node, 'tc')) return null;
    var tcPr = firstW(this.node, 'tcPr'), tcW = tcPr && firstW(tcPr, 'tcW');
    var w = tcW && wAttr(tcW, 'type') !== 'pct' ? parseFloat(wAttr(tcW, 'w')) : NaN;
    if (!isNaN(w) && w > 0) return w / 20;
    // Fall back to the table grid column.
    var tr = this.node.parentNode, tbl = tr && tr.parentNode;
    var grid = tbl && firstW(tbl, 'tblGrid');
    var col = grid && kidsW(grid, 'gridCol')[kidsW(tr, 'tc').indexOf(this.node)];
    var g = col ? parseFloat(wAttr(col, 'w')) : NaN;
    return !isNaN(g) && g > 0 ? g / 20 : null;
  };
  Container.prototype.getPaddingLeft = Container.prototype.getPaddingRight = function () { return 5.4; };
  Container.prototype.pageSetting = function (tag, attr) {
    var sect = descendantsW(this.node, 'sectPr').pop();
    var v = sect && wAttr(firstW(sect, tag), attr);
    return v ? parseFloat(v) / 20 : null;
  };
  Container.prototype.getPageWidth = function () { return isW(this.node, 'body') ? this.pageSetting('pgSz', 'w') : null; };
  Container.prototype.getMarginLeft = function () { return this.pageSetting('pgMar', 'left') || 0; };
  Container.prototype.getMarginRight = function () { return this.pageSetting('pgMar', 'right') || 0; };
  Container.prototype.insertParagraph = function (i, text) {
    var p = el(this.ctx, 'p');
    if (text) {
      var r = el(this.ctx, 'r'), t = el(this.ctx, 't');
      t.setAttribute('xml:space', 'preserve');
      t.textContent = text;
      r.appendChild(t);
      p.appendChild(r);
    }
    // Appending keeps a trailing <w:sectPr> last.
    var ref = this.kids()[i] || firstW(this.node, 'sectPr');
    this.node.insertBefore(p, ref);
    return new Para(p, this.ctx);
  };
  Container.prototype.getParagraphs = function () {
    var self = this;
    return descendantsW(this.node, 'p')
      .filter(function (p) { return wrapContainer(p.parentNode, self.ctx); })
      .map(function (p) { return new Para(p, self.ctx); });
  };
  Container.prototype.getText = function () {
    var ctx = this.ctx;
    return this.kids().map(function (k) { return wrap(k, ctx).getText(); }).join('\n');
  };
  Container.prototype.replaceText = function (pattern, repl) {
    var re = toRegExp(pattern, 'g');
    this.getParagraphs().forEach(function (p) { p.replaceText(re, repl); });
    return this;
  };
  Container.prototype.findText = function (pattern) {
    var re = toRegExp(pattern, '');
    var paras = this.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      var m = paras[i].getText().match(re);
      if (m) {
        var text = paras[i].editAsText();
        return {
          getElement: function () { return { asText: function () { return text; } }; },
          getStartOffset: function () { return m.index; },
          getEndOffsetInclusive: function () { return m.index + m[0].length - 1; }
        };
      }
    }
    return null;
  };

  function toRegExp(pattern, flags) {
    if (pattern.indexOf('(?i)') === 0) { pattern = pattern.slice(4); flags += 'i'; }
    return new RegExp(pattern, flags);
  }

  // Paragraph / list item
  function Para(node, ctx) { this.node = node; this.ctx = ctx; }
  Para.prototype = Object.create(Base.prototype);
  Para.prototype.numPr = function () {
    var pPr = firstW(this.node, 'pPr');
    return pPr ? firstW(pPr, 'numPr') : null;
  };
  Para.prototype.getType = function () {
    var num = this.numPr();
    return num && wAttr(firstW(num, 'numId'), 'val') !== '0' ? ET.LIST_ITEM : ET.PARAGRAPH;
  };
  Para.prototype.getNestingLevel = function () {
    var num = this.numPr();
    return num ? parseInt(wAttr(firstW(num, 'ilvl'), 'val') || '0', 10) : 0;
  };
  Para.prototype.runs = function () {
    var node = this.node;
    // Skip deleted suggestions and runs of paragraphs nested in text boxes.
    return descendantsW(node, 'r').filter(function (r) {
      return closestP(r) === node && !hasAncestor(r, node, ['del', 'moveFrom']);
    });
  };
  function closestP(r) {
    for (var n = r.parentNode; n; n = n.parentNode) if (isW(n, 'p')) return n;
    return null;
  }
  Para.prototype.getText = function () { return this.runs().map(runText).join(''); };
  Para.prototype.editAsText = function () { return new Text(this); };
  Para.prototype.clear = function () {
    // A cleared paragraph also loses its section break (the cover's section).
    var pPr0 = firstW(this.node, 'pPr'), sect = pPr0 && firstW(pPr0, 'sectPr');
    if (sect) { noteSectionBreak(this.node, this.ctx); pPr0.removeChild(sect); }
    var pPr = firstW(this.node, 'pPr');
    var kill = [];
    for (var c = this.node.firstChild; c; c = c.nextSibling) if (c !== pPr) kill.push(c);
    var node = this.node;
    kill.forEach(function (c) { node.removeChild(c); });
    return this;
  };
  /** FONT_SIZE, SPACING_BEFORE/AFTER, LINE_SPACING, INDENT_START (points), like DocumentApp. */
  Para.prototype.getAttributes = function () {
    var pPr = firstW(this.node, 'pPr');
    var attrs = {};
    var sp = pPr && firstW(pPr, 'spacing');
    if (sp) {
      var b = wAttr(sp, 'before'), a = wAttr(sp, 'after'), l = wAttr(sp, 'line'), rule = wAttr(sp, 'lineRule');
      if (b != null) attrs.SPACING_BEFORE = parseFloat(b) / 20;
      if (a != null) attrs.SPACING_AFTER = parseFloat(a) / 20;
      if (l != null && (!rule || rule === 'auto')) attrs.LINE_SPACING = parseFloat(l) / 240;
    }
    var ind = pPr && firstW(pPr, 'ind');
    var left = ind && (wAttr(ind, 'left') || wAttr(ind, 'start'));
    if (left) attrs.INDENT_START = parseFloat(left) / 20;
    var r = this.runs().filter(function (x) { return runText(x).length; })[0];
    var sz = r && firstW(r, 'rPr') && wAttr(firstW(firstW(r, 'rPr'), 'sz'), 'val');
    if (!sz && pPr && firstW(pPr, 'rPr')) sz = wAttr(firstW(firstW(pPr, 'rPr'), 'sz'), 'val');
    if (sz) attrs.FONT_SIZE = parseFloat(sz) / 2;
    return attrs;
  };
  /** Applies the same attributes to this (usually empty) paragraph. */
  Para.prototype.setAttributes = function (attrs) {
    var ctx = this.ctx;
    var pPr = firstW(this.node, 'pPr');
    if (!pPr) { pPr = el(ctx, 'pPr'); this.node.insertBefore(pPr, this.node.firstChild); }
    if (attrs.SPACING_BEFORE != null || attrs.SPACING_AFTER != null || attrs.LINE_SPACING != null) {
      var sp = el(ctx, 'spacing');
      if (attrs.SPACING_BEFORE != null) sp.setAttributeNS(W, 'w:before', String(Math.round(attrs.SPACING_BEFORE * 20)));
      if (attrs.SPACING_AFTER != null) sp.setAttributeNS(W, 'w:after', String(Math.round(attrs.SPACING_AFTER * 20)));
      if (attrs.LINE_SPACING != null) {
        sp.setAttributeNS(W, 'w:line', String(Math.round(attrs.LINE_SPACING * 240)));
        sp.setAttributeNS(W, 'w:lineRule', 'auto');
      }
      kidsW(pPr, 'spacing').forEach(function (x) { pPr.removeChild(x); });
      pPr.insertBefore(sp, firstW(pPr, 'ind') || firstW(pPr, 'jc') || firstW(pPr, 'rPr') || null);
    }
    if (attrs.FONT_SIZE != null) {
      var rPr = firstW(pPr, 'rPr');
      if (!rPr) { rPr = el(ctx, 'rPr'); pPr.appendChild(rPr); }
      var half = String(Math.round(attrs.FONT_SIZE * 2));
      kidsW(rPr, 'sz').concat(kidsW(rPr, 'szCs')).forEach(function (x) { rPr.removeChild(x); });
      rPr.appendChild(el(ctx, 'sz', { val: half }));
      rPr.appendChild(el(ctx, 'szCs', { val: half }));
    }
    return this;
  };
  Para.prototype.replaceText = function (re, repl) {
    var text = this.getText();
    re.lastIndex = 0;
    var matches = [], m;
    if (re.global) { while ((m = re.exec(text))) { matches.push(m); if (!m[0].length) re.lastIndex++; } }
    else if ((m = text.match(re))) matches.push(m);
    var t = this.editAsText();
    for (var i = matches.length - 1; i >= 0; i--) {
      var s = matches[i].index, e = s + matches[i][0].length - 1;
      if (e < s) continue;
      if (repl) t.insertText(e + 1, repl);
      t.deleteText(s, e);
    }
  };

  // Text: character-offset editing on top of runs
  function Text(para) { this.p = para; this.ctx = para.ctx; }
  Text.prototype.segs = function () {
    var pos = 0;
    return this.p.runs().map(function (r) {
      var len = runText(r).length;
      var seg = { run: r, start: pos, len: len };
      pos += len;
      return seg;
    });
  };
  Text.prototype.getText = function () { return this.p.getText(); };
  Text.prototype.runAt = function (off) {
    var segs = this.segs();
    for (var i = 0; i < segs.length; i++) {
      if (segs[i].len && off >= segs[i].start && off < segs[i].start + segs[i].len) return segs[i].run;
    }
    return null;
  };
  Text.prototype.getTextAttributeIndices = function () {
    var out = [];
    this.segs().forEach(function (s) { if (s.len && out.indexOf(s.start) < 0) out.push(s.start); });
    return out;
  };
  Text.prototype.getForegroundColor = function (off) {
    var r = this.runAt(off), rPr = r && firstW(r, 'rPr');
    var v = rPr && wAttr(firstW(rPr, 'color'), 'val');
    return v && v !== 'auto' ? '#' + v.toLowerCase() : null;
  };
  Text.prototype.getFontSize = function (off) {
    var r = this.runAt(off), rPr = r && firstW(r, 'rPr');
    var sz = rPr && wAttr(firstW(rPr, 'sz'), 'val');
    return sz ? parseFloat(sz) / 2 : null;
  };
  Text.prototype.isBold = function (off) {
    var r = this.runAt(off), rPr = r && firstW(r, 'rPr');
    return !!rPr && onOff(firstW(rPr, 'b'));
  };
  Text.prototype.getBackgroundColor = function (off) {
    var r = this.runAt(off), rPr = r && firstW(r, 'rPr');
    if (!rPr) return null;
    var hl = wAttr(firstW(rPr, 'highlight'), 'val');
    if (hl && hl !== 'none') return hl;
    var shd = wAttr(firstW(rPr, 'shd'), 'fill');
    return shd && shd !== 'auto' ? '#' + shd.toLowerCase() : null;
  };
  /** Makes sure a run starts exactly at `off`. */
  Text.prototype.splitAt = function (off) {
    var segs = this.segs();
    for (var i = 0; i < segs.length; i++) {
      var s = segs[i];
      if (off > s.start && off < s.start + s.len) {
        var r = s.run, t = firstW(r, 't');
        var clone = r.cloneNode(true);
        var ct = firstW(clone, 't');
        var full = t.textContent;
        t.textContent = full.slice(0, off - s.start);
        ct.textContent = full.slice(off - s.start);
        t.setAttribute('xml:space', 'preserve');
        ct.setAttribute('xml:space', 'preserve');
        r.parentNode.insertBefore(clone, r.nextSibling);
        return;
      }
    }
  };
  Text.prototype.runsIn = function (s, e) {
    this.splitAt(s);
    this.splitAt(e + 1);
    return this.segs().filter(function (g) { return g.len && g.start >= s && g.start + g.len <= e + 1; })
      .map(function (g) { return g.run; });
  };
  Text.prototype.deleteText = function (s, e) {
    this.runsIn(s, e).forEach(function (r) {
      var parent = r.parentNode;
      parent.removeChild(r);
      if (isW(parent, 'hyperlink') && !descendantsW(parent, 'r').length) parent.parentNode.removeChild(parent);
    });
    return this;
  };
  Text.prototype.setForegroundColor = function (s, e, color) {
    var ctx = this.ctx;
    this.runsIn(s, e).forEach(function (r) {
      if (color) setRprChild(ctx, r, 'color', { val: color.replace('#', '').toUpperCase() });
      else removeRprChild(r, 'color');
    });
    return this;
  };
  Text.prototype.setBold = function (s, e, bold) {
    var ctx = this.ctx;
    this.runsIn(s, e).forEach(function (r) {
      if (bold) { setRprChild(ctx, r, 'b', {}); setRprChild(ctx, r, 'bCs', {}); }
      else { removeRprChild(r, 'b'); removeRprChild(r, 'bCs'); }
    });
    return this;
  };
  Text.prototype.setBackgroundColor = function (s, e, color) {
    this.runsIn(s, e).forEach(function (r) {
      removeRprChild(r, 'highlight');
      removeRprChild(r, 'shd');
    });
    return this;
  };
  Text.prototype.insertText = function (at, str) {
    this.splitAt(at);
    var segs = this.segs();
    var prev = null, next = null;
    for (var i = 0; i < segs.length; i++) {
      if (segs[i].len && segs[i].start + segs[i].len <= at) prev = segs[i].run;
      if (!next && segs[i].len && segs[i].start >= at) next = segs[i].run;
    }
    var model = prev || next;
    var r = el(this.ctx, 'r');
    var mPr = model && firstW(model, 'rPr');
    if (mPr) r.appendChild(mPr.cloneNode(true));
    var t = el(this.ctx, 't');
    t.setAttribute('xml:space', 'preserve');
    t.textContent = str;
    r.appendChild(t);
    if (prev) prev.parentNode.insertBefore(r, prev.nextSibling);
    else if (next) next.parentNode.insertBefore(r, next);
    else this.p.node.appendChild(r);
    return this;
  };
  Text.prototype.setLinkUrl = function (s, e, url) {
    var runs = this.runsIn(s, e);
    if (!runs.length) return this;
    var rid = this.ctx.addHyperlink(url);
    var link = el(this.ctx, 'hyperlink');
    link.setAttributeNS(R, 'r:id', rid);
    runs[0].parentNode.insertBefore(link, runs[0]);
    var ctx = this.ctx;
    runs.forEach(function (r) {
      link.appendChild(r);
      setRprChild(ctx, r, 'color', { val: '1155CC' });
      setRprChild(ctx, r, 'u', { val: 'single' });
    });
    return this;
  };

  // Tables
  function Table(node, ctx) { this.node = node; this.ctx = ctx; }
  Table.prototype = Object.create(Base.prototype);
  Table.prototype.getType = function () { return ET.TABLE; };
  Table.prototype.rowNodes = function () { return kidsW(this.node, 'tr'); };
  Table.prototype.getNumRows = function () { return this.rowNodes().length; };
  Table.prototype.getRow = function (i) { return new Row(this.rowNodes()[i], this.ctx); };
  Table.prototype.getCell = function (r, c) { return this.getRow(r).getCell(c); };
  Table.prototype.getText = function () {
    var out = [];
    for (var r = 0; r < this.getNumRows(); r++) {
      var row = this.getRow(r), cells = [];
      for (var c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText());
      out.push(cells.join('\t'));
    }
    return out.join('\n');
  };
  function Row(node, ctx) { this.node = node; this.ctx = ctx; }
  Row.prototype.getType = function () { return ET.TABLE_ROW; };
  Row.prototype.cellNodes = function () { return kidsW(this.node, 'tc'); };
  Row.prototype.getNumCells = function () { return this.cellNodes().length; };
  Row.prototype.getCell = function (i) { return new Container(this.cellNodes()[i], this.ctx); };

  // ---------------------------------------------------------------- package

  function relsPath(partPath) {
    var i = partPath.lastIndexOf('/');
    return partPath.slice(0, i) + '/_rels/' + partPath.slice(i + 1) + '.rels';
  }

  function Pkg(zip, env) {
    this.zip = zip;
    this.env = env;
    this.parts = {};
  }
  Pkg.prototype.read = function (path) {
    var self = this;
    if (this.parts[path]) return Promise.resolve(this.parts[path]);
    var f = this.zip.file(path);
    if (!f) return Promise.resolve(null);
    return f.async('string').then(function (s) {
      var xml = new self.env.DOMParser().parseFromString(s, 'application/xml');
      self.parts[path] = xml;
      return xml;
    });
  };
  Pkg.prototype.write = function () {
    var ser = new this.env.XMLSerializer();
    for (var path in this.parts) {
      var s = ser.serializeToString(this.parts[path]);
      if (s.indexOf('<?xml') !== 0) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
      this.zip.file(path, s);
    }
  };
  /** Relationships of a part: [{id, type, target, node}]. */
  Pkg.prototype.rels = function (partPath) {
    return this.read(relsPath(partPath)).then(function (xml) {
      if (!xml) return { xml: null, list: [] };
      var nodes = xml.getElementsByTagNameNS(PKG, 'Relationship'), list = [];
      for (var i = 0; i < nodes.length; i++) {
        list.push({ id: nodes[i].getAttribute('Id'), type: nodes[i].getAttribute('Type'), target: nodes[i].getAttribute('Target'), node: nodes[i] });
      }
      return { xml: xml, list: list };
    });
  };
  Pkg.prototype.ensureRels = function (partPath) {
    var self = this, path = relsPath(partPath);
    return this.read(path).then(function (xml) {
      if (xml) return xml;
      xml = new self.env.DOMParser().parseFromString('<Relationships xmlns="' + PKG + '"/>', 'application/xml');
      self.parts[path] = xml;
      return xml;
    });
  };

  /**
   * "Select all → Highlight: none", including what the text-based pass can't
   * reach: footnote numbers and other runs with no text, and paragraph marks.
   * Also clears answer blue left on paragraph marks, so anything a student
   * types on a blank answer line comes out black. Returns how many it cleared.
   */
  function sweepFormatting(root, answerColors, removeHighlights) {
    var n = 0;
    descendantsW(root, 'rPr').forEach(function (rPr) {
      if (removeHighlights) {
        kidsW(rPr, 'highlight').concat(kidsW(rPr, 'shd')).forEach(function (h) { rPr.removeChild(h); n++; });
      }
      var owner = rPr.parentNode;
      if (isW(owner, 'pPr')) {
        kidsW(rPr, 'color').forEach(function (c) {
          var v = (wAttr(c, 'val') || '').toLowerCase();
          if (answerColors.indexOf('#' + v) >= 0) rPr.removeChild(c);
        });
      }
    });
    return n;
  }

  function isTabRun(r) {
    var content = [];
    for (var c = r.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && !isW(c, 'rPr')) content.push(c);
    return content.length === 1 && isW(content[0], 'tab');
  }
  function hasPageField(p) {
    var instr = descendantsW(p, 'instrText').map(function (n) { return n.textContent; }).join(' ');
    var simple = descendantsW(p, 'fldSimple').map(function (n) { return wAttr(n, 'instr') || ''; }).join(' ');
    return /\bPAGE\b/.test(instr + ' ' + simple);
  }
  // <w:pPr> children that must come after <w:tabs>.
  var AFTER_TABS = ['suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
    'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
    'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId',
    'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

  /**
   * Teacher footers push "Page N" to the right edge with a row of tabs after
   * the short copyright line. Once that line becomes the much longer CC line,
   * the same tabs wrap "Page N" toward the middle of the next line. Swap the
   * row of tabs for one right-aligned tab stop at the right margin, so the
   * page number stays right-aligned as on published student copies.
   */
  function rightAlignPageNumber(ftrRoot, textWidthTwips) {
    var fixed = 0;
    descendantsW(ftrRoot, 'p').forEach(function (p) {
      var text = descendantsW(p, 't').map(function (t) { return t.textContent; }).join('');
      if (text.indexOf('Unless otherwise noted') < 0 || !hasPageField(p)) return;
      var tabs = descendantsW(p, 'r').filter(isTabRun);
      if (tabs.length < 2) return;
      tabs.slice(1).forEach(function (r) { r.parentNode.removeChild(r); });
      var xml = p.ownerDocument;
      var pPr = firstW(p, 'pPr');
      if (!pPr) { pPr = xml.createElementNS(W, 'w:pPr'); p.insertBefore(pPr, p.firstChild); }
      kidsW(pPr, 'tabs').forEach(function (t) { pPr.removeChild(t); });
      var tabsEl = xml.createElementNS(W, 'w:tabs');
      var tab = xml.createElementNS(W, 'w:tab');
      tab.setAttributeNS(W, 'w:val', 'right');
      tab.setAttributeNS(W, 'w:pos', String(textWidthTwips));
      tabsEl.appendChild(tab);
      var before = null;
      for (var c = pPr.firstChild; c; c = c.nextSibling) {
        if (isW(c) && AFTER_TABS.indexOf(c.localName) >= 0) { before = c; break; }
      }
      pPr.insertBefore(tabsEl, before);
      fixed++;
    });
    return fixed;
  }

  /** Remembers a section break that goes away with a removed paragraph. */
  function noteSectionBreak(node, ctx) {
    if (!ctx.removedSections) return;
    var list = isW(node, 'sectPr') ? [node] : descendantsW(node, 'sectPr');
    list.forEach(function (s) { if (ctx.removedSections.indexOf(s) < 0) ctx.removedSections.push(s); });
  }

  var SECT_REF_TAGS = ['headerReference', 'footerReference'];
  var SECT_AFTER_TITLEPG = ['textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];

  function refOf(sect, tag, type) {
    return kidsW(sect, tag).filter(function (r) { return (wAttr(r, 'type') || 'default') === type; })[0] || null;
  }

  /**
   * The cover page lives in its own section whose "different first page"
   * header holds the large logo. Deleting the cover deletes that section, so
   * hand its headers to the first remaining section and turn on "different
   * first page" there: page 1 gets the large logo, later pages the small one.
   * Returns true if it changed anything.
   */
  function carryCoverHeaders(body, removedSections, order) {
    if (!removedSections.length) return false;
    var cover = removedSections.slice().sort(function (a, b) { return order.indexOf(b) - order.indexOf(a); })[0];
    var target = descendantsW(body, 'sectPr')[0];
    if (!cover || !target) return false;
    var xml = body.ownerDocument;
    var changed = false;
    var firstNonRef = function () {
      for (var c = target.firstChild; c; c = c.nextSibling) {
        if (isW(c) && SECT_REF_TAGS.indexOf(c.localName) < 0) return c;
      }
      return null;
    };
    [['headerReference', 'default'], ['headerReference', 'first'], ['footerReference', 'default']].forEach(function (pair) {
      var have = refOf(target, pair[0], pair[1]);
      var from = refOf(cover, pair[0], pair[1]);
      if (!have && from) { target.insertBefore(from.cloneNode(true), firstNonRef()); changed = true; }
    });
    if (!refOf(target, 'footerReference', 'first')) {
      // Page 1 uses the same footer (and page number) as the other pages.
      var def = refOf(target, 'footerReference', 'default');
      if (def) {
        var f = def.cloneNode(true);
        f.setAttributeNS(W, 'w:type', 'first');
        target.insertBefore(f, firstNonRef());
        changed = true;
      }
    }
    if (refOf(target, 'headerReference', 'first') && !firstW(target, 'titlePg')) {
      var before = null;
      for (var c = target.firstChild; c; c = c.nextSibling) {
        if (isW(c) && SECT_AFTER_TITLEPG.indexOf(c.localName) >= 0) { before = c; break; }
      }
      target.insertBefore(xml.createElementNS(W, 'w:titlePg'), before);
      changed = true;
    }
    return changed;
  }

  function makeCtx(xml, relsXml) {
    var n = 0;
    return {
      xml: xml,
      addHyperlink: function (url) {
        var id = 'rIdStudentCopy' + (++n);
        var rel = relsXml.createElementNS(PKG, 'Relationship');
        rel.setAttribute('Id', id);
        rel.setAttribute('Type', HYPERLINK_TYPE);
        rel.setAttribute('Target', url);
        rel.setAttribute('TargetMode', 'External');
        relsXml.documentElement.appendChild(rel);
        return id;
      }
    };
  }

  var COMMENT_TYPES = /\/(comments|commentsExtended|commentsIds|commentsExtensible)$/;

  /** Student copies must not carry comments (e.g. MCQ answer comments). */
  function stripComments(pkg, docXml, docRels, report) {
    var found = 0;
    ['commentRangeStart', 'commentRangeEnd'].forEach(function (name) {
      descendantsW(docXml, name).forEach(function (n) { n.parentNode.removeChild(n); found++; });
    });
    descendantsW(docXml, 'commentReference').forEach(function (n) {
      var r = n.parentNode;
      if (isW(r, 'r')) r.parentNode.removeChild(r); else r.removeChild(n);
      found++;
    });
    var removed = [];
    docRels.list.forEach(function (rel) {
      if (COMMENT_TYPES.test(rel.type)) {
        rel.node.parentNode.removeChild(rel.node);
        removed.push('word/' + rel.target.replace(/^\/?word\//, ''));
      }
    });
    removed.forEach(function (p) { pkg.zip.remove(p); delete pkg.parts[p]; });
    return pkg.read('[Content_Types].xml').then(function (ct) {
      if (!ct) return;
      var over = ct.getElementsByTagName('Override');
      for (var i = over.length - 1; i >= 0; i--) {
        var name = over[i].getAttribute('PartName').replace(/^\//, '');
        if (removed.indexOf(name) >= 0) over[i].parentNode.removeChild(over[i]);
      }
      if (found) report.commentsRemoved = found;
    });
  }

  /**
   * Converts a teacher-copy .docx (ArrayBuffer/Uint8Array) to a student copy.
   * Resolves to {data (Uint8Array), name, summary, warnings, manualChecks}.
   */
  function convert(input, fileName, env, options) {
    env = env || {};
    env.JSZip = env.JSZip || (typeof JSZip !== 'undefined' ? JSZip : null);
    env.DOMParser = env.DOMParser || (typeof DOMParser !== 'undefined' ? DOMParser : null);
    env.XMLSerializer = env.XMLSerializer || (typeof XMLSerializer !== 'undefined' ? XMLSerializer : null);
    var names = env.shared || (typeof studentCopyName !== 'undefined'
      ? { studentCopyName: studentCopyName, summarize: summarize, MANUAL_CHECKS: MANUAL_CHECKS } : null);
    var rules = env.StudentCopy || (typeof StudentCopy !== 'undefined' ? StudentCopy : null);

    var pkg, report = rules.newReport();
    return env.JSZip.loadAsync(input).then(function (zip) {
      pkg = new Pkg(zip, env);
      return pkg.read('word/document.xml');
    }).then(function (docXml) {
      if (!docXml || !firstW(docXml.documentElement, 'body')) {
        throw new Error('This does not look like a Word document. In Google Docs use File > Download > Microsoft Word (.docx).');
      }
      return Promise.all([pkg.rels('word/document.xml'), pkg.ensureRels('word/document.xml')]).then(function (res) {
        var docRels = res[0], relsXml = res[1];
        return stripComments(pkg, docXml, docRels, report).then(function () {
          if (descendantsW(docXml, 'ins').length || descendantsW(docXml, 'del').length) {
            report.warnings.push('The file has suggested edits (tracked changes). Accept or reject them in the teacher copy, then download it again.');
          }
          splitRuns(docXml);
          var bodyNode = firstW(docXml.documentElement, 'body');
          var ctx = makeCtx(docXml, relsXml);
          ctx.removedSections = [];
          var sectionOrder = descendantsW(bodyNode, 'sectPr');
          withDocumentApp(function () {
            rules.convertBody(new Container(bodyNode, ctx), options, report);
          });
          if (report.removedCover && carryCoverHeaders(bodyNode, ctx.removedSections, sectionOrder)) {
            report.logoFixed = true;
          }
          var o = options || {};
          var colors = (o.answerColors || rules.DEFAULTS.answerColors).map(function (c) { return String(c).toLowerCase(); });
          var sweepHighlights = o.removeHighlights !== false && rules.DEFAULTS.removeHighlights !== false;
          report.highlightsRemoved += sweepFormatting(docXml.documentElement, colors, sweepHighlights);
          var first = descendantsW(bodyNode, 'sectPr')[0];
          report.firstPageHeaderOk = !!(first && firstW(first, 'titlePg') && refOf(first, 'headerReference', 'first'));
          if (!report.firstPageHeaderOk) {
            report.warnings.push('Page 1 has no separate first-page header, so it may show the small logo. Check that page 1 has the large CommonLit logo.');
          }
          var notes = docRels.list.filter(function (r) { return /\/(footnotes|endnotes)$/.test(r.type); });
          var notesDone = Promise.all(notes.map(function (rel) {
            return pkg.read('word/' + rel.target.replace(/^\/?word\//, '')).then(function (xml) {
              if (xml) report.highlightsRemoved += sweepFormatting(xml.documentElement, colors, sweepHighlights);
            });
          }));
          var hf = docRels.list.filter(function (r) { return /\/(header|footer)$/.test(r.type); });
          return notesDone.then(function () { return Promise.all(hf.map(function (rel) {
            var path = 'word/' + rel.target.replace(/^\/?word\//, '');
            return Promise.all([pkg.read(path), pkg.ensureRels(path)]).then(function (x) {
              if (!x[0]) return;
              splitRuns(x[0]);
              withDocumentApp(function () {
                rules.convertHeaderFooter(new Container(x[0].documentElement, makeCtx(x[0], x[1])), report);
              });
              if (/\/footer$/.test(rel.type)) {
                var page = new Container(bodyNode, ctx);
                var widthTwips = Math.round(((page.getPageWidth() || 612) - page.getMarginLeft() - page.getMarginRight()) * 20);
                rightAlignPageNumber(x[0].documentElement, widthTwips);
              }
            });
          })); });
        });
      });
    }).then(function () {
      pkg.write();
      return pkg.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }).then(function (data) {
      var base = String(fileName || 'Teacher copy').replace(/\.docx$/i, '');
      var summary = names.summarize(report);
      if (report.logoFixed) summary.push('Page 1 set to use the large logo; later pages keep the small one');
      if (report.commentsRemoved) summary.push('Comments removed');
      return {
        data: data,
        name: names.studentCopyName(base) + '.docx',
        summary: summary,
        warnings: report.warnings,
        // Logos are handled here; see the warning below when they couldn't be.
        manualChecks: names.MANUAL_CHECKS.filter(function (c) { return !/logo/i.test(c); })
      };
    });

    // Converter.js reads DocumentApp.ElementType; point it at ours while it runs.
    function withDocumentApp(fn) {
      var g = typeof globalThis !== 'undefined' ? globalThis : window;
      var had = 'DocumentApp' in g, old = g.DocumentApp;
      g.DocumentApp = { ElementType: ET, Attribute: ATTRIBUTE };
      try { fn(); } finally { if (had) g.DocumentApp = old; else delete g.DocumentApp; }
    }
  }

  return { convert: convert, ElementType: ET, Container: Container };
})();

if (typeof module !== 'undefined') module.exports = DocxStudentCopy;
