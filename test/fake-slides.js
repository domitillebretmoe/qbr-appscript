// Minimal in-memory stand-in for the SlidesApp objects Deck.gs touches (presentation, slides, tables, shapes, text).
'use strict';

// `style` mirrors a template cell's run style: null fields = run-less paragraph (Slides then uses its 18pt default).
class FakeText {
  constructor(text, style) {
    this.text = text;
    this.link = null;
    const s = style || {};
    this.color = s.color || null;
    this.fontSize = s.fontSize || null;
    this.fontFamily = s.fontFamily || null;
    this.bold = s.bold || false;
  }
  asString() { return this.text; }
  setText(text) { this.text = String(text); return this; }
  replaceAllText(from, to) { const n = this.text.split(from).length - 1; this.text = this.text.split(from).join(to); return n; }
  getTextStyle() {
    const style = {
      setLinkUrl: url => { this.link = url; return style; },
      setForegroundColor: color => { this.color = color; return style; },
      setFontSize: size => { this.fontSize = size; return style; },
      setFontFamily: family => { this.fontFamily = family; return style; },
      setBold: bold => { this.bold = bold; return style; },
      getForegroundColor: () => this.color,
      getFontSize: () => this.fontSize,
      getFontFamily: () => this.fontFamily,
      isBold: () => this.bold,
    };
    return style;
  }
}

// Table cell text as the generator writes it: styled runs for non-empty cells, run-less (unstyled) empty ones.
const CELL_STYLE = { fontSize: 8.5, fontFamily: 'Inter', color: '#0F172A' };
function cellText(text, r, c) {
  if (text === '') return new FakeText('');
  if (r === 0) return new FakeText(text, Object.assign({}, CELL_STYLE, { color: '#FFFFFF', bold: true }));
  if (/^\[/.test(text)) return new FakeText(text, Object.assign({}, CELL_STYLE, { color: '#94A3B8' }));
  return new FakeText(text, Object.assign({}, CELL_STYLE, { bold: c === 0 }));
}

let nextId = 1;

class FakeShape {
  constructor(slide, text, box) {
    this.slide = slide;
    this.id = `g${nextId++}`;
    this.textRange = new FakeText(text);
    this.box = box || { left: 10, top: 20, width: 300, height: 200 };
    this.fill = null;
    this.removed = false;
  }
  getObjectId() { return this.id; }
  getPageElementType() { return 'SHAPE'; }
  asShape() { return this; }
  getText() { return this.textRange; }
  getLeft() { return this.box.left; }
  getTop() { return this.box.top; }
  getWidth() { return this.box.width; }
  getHeight() { return this.box.height; }
  getFill() { return { setSolidFill: color => { this.fill = color; } }; }
  remove() { this.removed = true; this.slide.shapes = this.slide.shapes.filter(s => s !== this); }
}

class FakeTable {
  constructor(rows) { this.rows = rows.map((r, ri) => r.map((t, ci) => cellText(t, ri, ci))); }
  getNumRows() { return this.rows.length; }
  getNumColumns() { return this.rows[0].length; }
  getCell(r, c) { return { getText: () => this.rows[r][c] }; }
  appendRow() { this.rows.push(this.rows[this.rows.length - 1].map(cell => new FakeText(cell.text))); }
  getRow(r) { return { remove: () => this.rows.splice(r, 1) }; }
  values() { return this.rows.map(r => r.map(cell => cell.text)); }
}

class FakeSlide {
  constructor() { this.shapes = []; this.tables = []; this.charts = []; }
  shape(text, box) { const s = new FakeShape(this, text, box); this.shapes.push(s); return s; }
  table(rows) { const t = new FakeTable(rows); this.tables.push(t); return t; }
  // Live Slides hands back a Shape handle for every element and only fails on asShape(); mirror that with an
  // opaque element the fill code must skip.
  opaque() { const o = { getPageElementType: () => 'SHAPE', asShape() { throw new Error('Page element is not of type shape.'); } }; this.opaques = (this.opaques || []).concat(o); return o; }
  getPageElements() { return this.shapes.concat(this.opaques || [], this.tables.map(t => ({ getPageElementType: () => 'TABLE', asShape() { throw new Error('Page element is not of type shape.'); } }))); }
  getPageElementById(id) { return this.shapes.filter(s => s.id === id)[0] || null; }
  getShapes() { return this.shapes.slice(); }
  getTables() { return this.tables.slice(); }
  insertSheetsChart(chart, left, top, width, height) { this.charts.push({ chart, left, top, width, height }); }
  allText() { return this.shapes.map(s => s.textRange.text).concat(...this.tables.map(t => t.values())).join('\n'); }
}

class FakePresentation {
  constructor() { this.slides = []; this.replacements = []; }
  slide() { const s = new FakeSlide(); this.slides.push(s); return s; }
  getSlides() { return this.slides.slice(); }
  replaceAllText(from, to) {
    this.replacements.push([from, to]);
    return this.slides.reduce((n, slide) => n
      + slide.shapes.reduce((m, s) => m + s.textRange.replaceAllText(from, to), 0)
      + slide.tables.reduce((m, t) => m + t.rows.reduce((k, row) => k + row.reduce((j, cell) => j + cell.replaceAllText(from, to), 0), 0), 0), 0);
  }
  text() { return this.slides.map(s => s.allText()).join('\n'); }
}

const fakeChart = title => ({ getOptions: () => ({ get: key => (key === 'title' ? title : null) }) });

module.exports = { FakePresentation, FakeSlide, FakeTable, FakeShape, fakeChart };
