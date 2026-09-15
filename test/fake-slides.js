// Minimal in-memory stand-in for the SlidesApp objects Deck.gs touches (presentation, slides, tables, shapes, text).
'use strict';

class FakeText {
  constructor(text) { this.text = text; this.link = null; this.color = null; this.fontSize = null; }
  asString() { return this.text; }
  setText(text) { this.text = String(text); return this; }
  replaceAllText(from, to) { const n = this.text.split(from).length - 1; this.text = this.text.split(from).join(to); return n; }
  getTextStyle() {
    return {
      setLinkUrl: url => { this.link = url; return this; },
      setForegroundColor: color => { this.color = color; return this; },
      setFontSize: size => { this.fontSize = size; return this; },
    };
  }
}

class FakeShape {
  constructor(slide, text, box, description) {
    this.slide = slide;
    this.textRange = new FakeText(text);
    this.box = box || { left: 10, top: 20, width: 300, height: 200 };
    this.description = description || '';
    this.fill = null;
    this.removed = false;
  }
  getText() { return this.textRange; }
  getDescription() { return this.description; }
  getLeft() { return this.box.left; }
  getTop() { return this.box.top; }
  getWidth() { return this.box.width; }
  getHeight() { return this.box.height; }
  getFill() { return { setSolidFill: color => { this.fill = color; } }; }
  remove() { this.removed = true; this.slide.shapes = this.slide.shapes.filter(s => s !== this); }
}

class FakeTable {
  constructor(rows) { this.rows = rows.map(r => r.map(t => new FakeText(t))); }
  getNumRows() { return this.rows.length; }
  getNumColumns() { return this.rows[0].length; }
  getCell(r, c) { return { getText: () => this.rows[r][c] }; }
  appendRow() { this.rows.push(this.rows[this.rows.length - 1].map(cell => new FakeText(cell.text))); }
  getRow(r) { return { remove: () => this.rows.splice(r, 1) }; }
  values() { return this.rows.map(r => r.map(cell => cell.text)); }
}

class FakeSlide {
  constructor() { this.shapes = []; this.tables = []; this.charts = []; }
  shape(text, box, description) { const s = new FakeShape(this, text, box, description); this.shapes.push(s); return s; }
  table(rows) { const t = new FakeTable(rows); this.tables.push(t); return t; }
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
