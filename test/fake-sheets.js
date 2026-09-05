// Minimal in-memory stand-in for the SpreadsheetApp / Charts / Utilities globals used by Render.gs, so the
// renderer can run under Node (tests, and tools/preview-xlsx.py which turns the recorded state into an .xlsx).
'use strict';

const colLetter = c => {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const a1 = (row, col) => `${colLetter(col)}${row}`;

class FakeSheet {
  constructor(name) {
    this.name = name;
    this.cells = {}; // "r,c" -> style + value
    this.merges = [];
    this.borders = [];
    this.charts = [];
    this.rules = [];
    this.rowHeights = {};
    this.colWidths = {};
    this.frozenRows = 0;
    this.hiddenGridlines = false;
  }
  cell(r, c) { return this.cells[`${r},${c}`] || (this.cells[`${r},${c}`] = { row: r, col: c }); }
  getName() { return this.name; }
  getMaxRows() { return 1000; }
  getMaxColumns() { return 40; }
  getLastRow() { return Math.max(0, ...Object.values(this.cells).map(c => c.row)); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
      const col = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      const r1 = Number(m[2]); const c1 = col(m[1]);
      const r2 = m[3] ? Number(m[4]) : r1; const c2 = m[3] ? col(m[3]) : c1;
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, a, b, c == null ? 1 : c, d == null ? 1 : d);
  }
  getCharts() { return this.charts.slice(); }
  removeChart(chart) { this.charts = this.charts.filter(c => c !== chart); }
  insertChart(chart) { this.charts.push(chart); }
  newChart() { return new FakeChartBuilder(this); }
  clear() { this.cells = {}; this.merges = []; this.borders = []; }
  clearConditionalFormatRules() { this.rules = []; }
  getConditionalFormatRules() { return this.rules.slice(); }
  setConditionalFormatRules(rules) { this.rules = rules; }
  setRowHeight(r, h) { this.rowHeights[r] = h; return this; }
  setColumnWidth(c, w) { this.colWidths[c] = w; return this; }
  setColumnWidths(c, n, w) { for (let i = 0; i < n; i++) this.colWidths[c + i] = w; return this; }
  setFrozenRows(n) { this.frozenRows = n; }
  setHiddenGridlines(v) { this.hiddenGridlines = v; }
}

class FakeRange {
  constructor(sheet, row, col, rows, cols) { Object.assign(this, { sheet, row, col, rows, cols }); }
  each(fn) { for (let r = 0; r < this.rows; r++) for (let c = 0; c < this.cols; c++) fn(this.sheet.cell(this.row + r, this.col + c), r, c); return this; }
  style(key, value) { return this.each(cell => { cell[key] = value; }); }
  getA1Notation() { return `${a1(this.row, this.col)}:${a1(this.row + this.rows - 1, this.col + this.cols - 1)}`; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.rows; r++) { out.push([]); for (let c = 0; c < this.cols; c++) out[r].push(this.sheet.cell(this.row + r, this.col + c).value ?? ''); }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(values) { return this.each((cell, r, c) => { cell.value = values[r][c]; }); }
  setValue(v) { return this.style('value', v); }
  setRichTextValue(rt) { return this.each(cell => { cell.value = rt.text; cell.richText = rt.runs; }); }
  setNumberFormat(f) { return this.style('numberFormat', f); }
  setBackground(v) { return this.style('background', v); }
  setFontColor(v) { return this.style('fontColor', v); }
  setFontColors(m) { return this.each((cell, r, c) => { cell.fontColor = m[r][c]; }); }
  setFontWeight(v) { return this.style('fontWeight', v); }
  setFontStyle(v) { return this.style('fontStyle', v); }
  setFontSize(v) { return this.style('fontSize', v); }
  setFontFamily(v) { return this.style('fontFamily', v); }
  setHorizontalAlignment(v) { return this.style('hAlign', v); }
  setVerticalAlignment(v) { return this.style('vAlign', v); }
  setWrap(v) { return this.style('wrap', v); }
  setDataValidation(v) { return this.style('validation', v); }
  merge() { this.sheet.merges.push(this.getA1Notation()); return this; }
  breakApart() { return this; }
  setBorder(top, left, bottom, right, vertical, horizontal, color, style) {
    this.sheet.borders.push({ range: this.getA1Notation(), top, left, bottom, right, vertical, horizontal, color });
    return this;
  }
}

class FakeChartBuilder {
  constructor(sheet) { this.sheet = sheet; this.ranges = []; this.options = {}; }
  setChartType(t) { this.type = t; return this; }
  addRange(range) { this.ranges.push({ a1: range.getA1Notation(), values: range.getValues() }); return this; }
  setNumHeaders(n) { this.numHeaders = n; return this; }
  setOption(k, v) { this.options[k] = v; return this; }
  setPosition(r, c, ox, oy) { this.position = { row: r, col: c, ox, oy }; return this; }
  build() { return { type: this.type, ranges: this.ranges, options: this.options, position: this.position }; }
}

const builder = (state = {}) => new Proxy({}, {
  get: (t, k) => {
    if (k === 'build') return () => state;
    if (k === 'setRanges') return ranges => builder(Object.assign(state, { ranges: ranges.map(r => r.getA1Notation()) }));
    return () => builder(state);
  },
});

const richText = () => {
  const rt = { text: '', runs: [] };
  return {
    setText(t) { rt.text = t; return this; },
    setTextStyle(start, end, style) { rt.runs.push({ start, end, style }); return this; },
    build() { return rt; },
  };
};
const textStyle = () => {
  const s = {};
  return {
    setBold(v) { s.bold = v; return this; },
    setFontSize(v) { s.fontSize = v; return this; },
    setForegroundColor(v) { s.color = v; return this; },
    build() { return s; },
  };
};

function globals() {
  return {
    SpreadsheetApp: {
      newDataValidation: builder,
      newConditionalFormatRule: builder,
      newRichTextValue: richText,
      newTextStyle: textStyle,
      InterpolationType: { NUMBER: 'NUMBER' },
      BorderStyle: { SOLID: 'SOLID' },
    },
    Charts: { ChartType: { COLUMN: 'COLUMN', BAR: 'BAR', PIE: 'PIE', LINE: 'LINE' } },
    Utilities: { formatDate: (d, tz, pattern) => (pattern.length > 10 ? d.toISOString().slice(0, 16).replace('T', ' ') : d.toISOString().slice(0, 10)) },
    Session: { getScriptTimeZone: () => 'UTC' },
  };
}

module.exports = { FakeSheet, globals };
