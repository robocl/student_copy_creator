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
  Container.prototype.removeChild = function (child) { this.node.removeChild(child.node); return this; };
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
    var pPr = firstW(this.node, 'pPr');
    var kill = [];
    for (var c = this.node.firstChild; c; c = c.nextSibling) if (c !== pPr) kill.push(c);
    var node = this.node;
    kill.forEach(function (c) { node.removeChild(c); });
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
          withDocumentApp(function () {
            rules.convertBody(new Container(firstW(docXml.documentElement, 'body'), makeCtx(docXml, relsXml)), options, report);
          });
          var hf = docRels.list.filter(function (r) { return /\/(header|footer)$/.test(r.type); });
          return Promise.all(hf.map(function (rel) {
            var path = 'word/' + rel.target.replace(/^\/?word\//, '');
            return Promise.all([pkg.read(path), pkg.ensureRels(path)]).then(function (x) {
              if (!x[0]) return;
              splitRuns(x[0]);
              withDocumentApp(function () {
                rules.convertHeaderFooter(new Container(x[0].documentElement, makeCtx(x[0], x[1])), report);
              });
            });
          }));
        });
      });
    }).then(function () {
      pkg.write();
      return pkg.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }).then(function (data) {
      var base = String(fileName || 'Teacher copy').replace(/\.docx$/i, '');
      var summary = names.summarize(report);
      if (report.commentsRemoved) summary.push('Comments removed');
      return {
        data: data,
        name: names.studentCopyName(base) + '.docx',
        summary: summary,
        warnings: report.warnings,
        manualChecks: names.MANUAL_CHECKS
      };
    });

    // Converter.js reads DocumentApp.ElementType; point it at ours while it runs.
    function withDocumentApp(fn) {
      var g = typeof globalThis !== 'undefined' ? globalThis : window;
      var had = 'DocumentApp' in g, old = g.DocumentApp;
      g.DocumentApp = { ElementType: ET };
      try { fn(); } finally { if (had) g.DocumentApp = old; else delete g.DocumentApp; }
    }
  }

  return { convert: convert, ElementType: ET, Container: Container };
})();

if (typeof module !== 'undefined') module.exports = DocxStudentCopy;
