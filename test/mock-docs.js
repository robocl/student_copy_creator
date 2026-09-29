/**
 * A small in-memory stand-in for the parts of Apps Script's DocumentApp that
 * Converter.js uses, so the conversion rules can be tested with Node.
 * Behaviour follows the DocumentApp docs; it is not a full emulation.
 */
'use strict';

const ElementType = {
  BODY_SECTION: 'BODY_SECTION', PARAGRAPH: 'PARAGRAPH', LIST_ITEM: 'LIST_ITEM',
  TABLE: 'TABLE', TABLE_ROW: 'TABLE_ROW', TABLE_CELL: 'TABLE_CELL'
};

class Element {
  constructor(type) { this.type = type; this.parent = null; }
  getType() { return this.type; }
  getParent() { return this.parent; }
  getPreviousSibling() { const p = this.parent; return p ? p.children[p.children.indexOf(this) - 1] || null : null; }
  getNextSibling() { const p = this.parent; return p ? p.children[p.children.indexOf(this) + 1] || null : null; }
  removeFromParent() { this.parent.removeChild(this); return this; }
}

class Text {
  constructor(para) { this.p = para; }
  getText() { return this.p.getText(); }
  getTextAttributeIndices() {
    const c = this.p.chars, out = [];
    for (let i = 0; i < c.length; i++) {
      if (i === 0 || c[i].color !== c[i - 1].color || c[i].bold !== c[i - 1].bold) out.push(i);
    }
    return out;
  }
  getForegroundColor(i) { return this.p.chars[i].color; }
  isBold(i) { return this.p.chars[i].bold; }
  deleteText(s, e) { this.p.chars.splice(s, e - s + 1); return this; }
  setForegroundColor(s, e, color) { for (let i = s; i <= e; i++) this.p.chars[i].color = color; return this; }
  setBold(s, e, bold) { for (let i = s; i <= e; i++) this.p.chars[i].bold = bold; return this; }
}

class Paragraph extends Element {
  constructor(json) {
    super(json.type === 'LIST_ITEM' ? ElementType.LIST_ITEM : ElementType.PARAGRAPH);
    this.level = json.level || 0;
    this.listId = json.listId || null;
    this.pageBreak = !!json.pageBreak;
    this.chars = [];
    for (const r of json.runs || []) {
      for (const ch of r.text) this.chars.push({ ch, color: r.color || null, bold: !!r.bold });
    }
  }
  getText() { return this.chars.map(c => c.ch).join(''); }
  editAsText() { return new Text(this); }
  clear() { this.chars = []; this.pageBreak = false; return this; }
  setText(t) { this.chars = [...t].map(ch => ({ ch, color: null, bold: false })); return this; }
  getNestingLevel() { return this.level; }
  getListId() { return this.listId; }
  asParagraph() { return this; }
  asListItem() { return this; }
  replaceText(re, repl) {
    const text = this.getText();
    const matches = [...text.matchAll(re)];
    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i];
      const attrs = this.chars[m.index] || { color: null, bold: false };
      const ins = [...repl].map(ch => ({ ch, color: attrs.color, bold: attrs.bold }));
      this.chars.splice(m.index, m[0].length, ...ins);
    }
  }
}

class Container extends Element {
  constructor(type, jsonChildren) {
    super(type);
    this.children = [];
    for (const j of jsonChildren || []) this._append(j.type === 'TABLE' ? new Table(j) : new Paragraph(j));
  }
  _append(el) { el.parent = this; this.children.push(el); }
  getNumChildren() { return this.children.length; }
  getChild(i) { return this.children[i]; }
  getChildIndex(el) { return this.children.indexOf(el); }
  removeChild(el) {
    const i = this.children.indexOf(el);
    if (i < 0) throw new Error('Element is not a child');
    if (this.children.length === 1) throw new Error("Can't remove the last paragraph in a section");
    if (i === this.children.length - 1 && this.type === ElementType.BODY_SECTION) {
      throw new Error("Can't remove the last paragraph in a document section.");
    }
    this.children.splice(i, 1);
    el.parent = null;
    return this;
  }
  insertParagraph(i, text) {
    const p = new Paragraph({ runs: text ? [{ text }] : [] });
    p.parent = this;
    this.children.splice(i, 0, p);
    return p;
  }
  getParagraphs() {
    const out = [];
    const walk = c => {
      for (const el of c.children) {
        if (el instanceof Paragraph) out.push(el);
        else if (el instanceof Table) el.rows.forEach(r => r.cells.forEach(walk));
      }
    };
    walk(this);
    return out;
  }
  getText() { return this.children.map(c => c.getText()).join('\n'); }
  replaceText(pattern, repl) {
    let flags = 'g';
    if (pattern.startsWith('(?i)')) { pattern = pattern.slice(4); flags += 'i'; }
    const re = new RegExp(pattern, flags);
    this.getParagraphs().forEach(p => p.replaceText(re, repl));
    return this;
  }
}

class TableRow extends Element {
  constructor(cells, table) {
    super(ElementType.TABLE_ROW);
    this.parent = table;
    this.cells = cells.map(c => { const cell = new Container(ElementType.TABLE_CELL, c); cell.parent = this; return cell; });
  }
  getNumCells() { return this.cells.length; }
  getCell(i) { return this.cells[i]; }
}

class Table extends Element {
  constructor(json) {
    super(ElementType.TABLE);
    this.rows = json.rows.map(r => new TableRow(r, this));
  }
  asTable() { return this; }
  getNumRows() { return this.rows.length; }
  getRow(i) { return this.rows[i]; }
  getCell(r, c) { return this.rows[r].cells[c]; }
  getText() { return this.rows.map(r => r.cells.map(c => c.getText()).join('\t')).join('\n'); }
}

function bodyFromJson(json) { return new Container(ElementType.BODY_SECTION, json); }

/** Flat, comparable view of a body: one line per paragraph. */
function outline(container, prefix = '') {
  const lines = [];
  for (const el of container.children) {
    if (el instanceof Table) {
      el.rows.forEach((r, ri) => r.cells.forEach((c, ci) => lines.push(...outline(c, `${prefix}[${ri},${ci}] `))));
    } else {
      lines.push({
        text: prefix + el.getText(),
        empty: !/\S/.test(el.getText()),
        answerColored: el.chars.some(c => c.color === '#0000ff' && /\S/.test(c.ch)),
        bold: el.chars.some(c => c.bold && /\S/.test(c.ch)),
        list: el.type === ElementType.LIST_ITEM
      });
    }
  }
  return lines;
}

module.exports = { DocumentApp: { ElementType }, bodyFromJson, outline, Paragraph, Table };
