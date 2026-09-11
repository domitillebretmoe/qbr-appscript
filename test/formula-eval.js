// Evaluates the formulas Render.gs / Formulas.gs write (SUMIFS, COUNTIFS, IF, OR, MAX, cell references, + - * /)
// against in-memory FakeSheets, so tests can check that every formula cell reproduces the computed metric.
'use strict';

const colIndex = letters => letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

function matches(value, criterion) {
  const c = String(criterion);
  const num = v => (typeof v === 'number' ? v : Number(v));
  if (/^(<>|<=|>=|<|>)/.test(c) && c !== '<>') {
    const op = /^(<>|<=|>=|<|>)/.exec(c)[1];
    const rest = c.slice(op.length);
    const target = rest === '' || Number.isNaN(Number(rest)) ? rest : Number(rest);
    if (typeof target === 'string') return op === '<>' ? String(value) !== target : false;
    if (typeof value !== 'number' && value !== '') return op === '<>';
    const v = num(value === '' ? NaN : value);
    return { '<>': v !== target, '<': v < target, '>': v > target, '<=': v <= target, '>=': v >= target }[op];
  }
  if (c.indexOf('*') >= 0) return new RegExp(`^${c.split('*').map(p => p.replace(/[.+?^${}()|[\]\\]/g, '\\$&')).join('.*')}$`, 'i').test(String(value));
  if (typeof value === 'number' && !Number.isNaN(Number(c)) && c !== '') return value === Number(c);
  return String(value).toLowerCase() === c.toLowerCase();
}

// `sheets`: name -> FakeSheet. `current`: the FakeSheet unqualified references resolve against.
function makeEvaluator(sheets, current) {
  const column = (sheetName, letters) => {
    const sheet = sheets[sheetName];
    if (!sheet) throw new Error(`No sheet ${sheetName}`);
    const col = colIndex(letters);
    const rows = [];
    for (let r = 1; r <= sheet.getLastRow(); r++) rows.push(sheet.cell(r, col).value ?? '');
    return rows;
  };
  const cellValue = (sheet, a1) => {
    const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1);
    const v = sheet.cell(Number(m[2]), colIndex(m[1])).value;
    return typeof v === 'string' && v.startsWith('=') ? evaluate(v, sheet) : (v === undefined ? '' : v);
  };
  const SUMIFS = (sumRange, ...pairs) => {
    let total = 0;
    sumRange.forEach((v, i) => {
      for (let p = 0; p < pairs.length; p += 2) if (!matches(pairs[p][i] ?? '', pairs[p + 1])) return;
      total += typeof v === 'number' ? v : 0;
    });
    return total;
  };
  const COUNTIFS = (...pairs) => {
    let n = 0;
    pairs[0].forEach((_, i) => {
      for (let p = 0; p < pairs.length; p += 2) if (!matches(pairs[p][i] ?? '', pairs[p + 1])) return;
      n++;
    });
    return n;
  };
  const fns = { SUMIFS, COUNTIFS, IF: (c, a, b) => (c ? a : b), OR: (...a) => a.some(Boolean), MAX: Math.max, MIN: Math.min };

  function evaluate(formula, sheet = current) {
    if (typeof formula !== 'string' || !formula.startsWith('=')) return formula;
    // Strings and sheet-qualified columns become placeholders first so their text is never parsed as cell references.
    const tokens = [];
    const hold = text => { tokens.push(text); return `__T${tokens.length - 1}_`; };
    let js = formula.slice(1)
      .replace(/"((?:[^"]|"")*)"/g, (m, s) => hold(JSON.stringify(s.replace(/""/g, '"'))))
      .replace(/'([^']+)'!([A-Z]+):[A-Z]+/g, (m, name, letters) => hold(`__COL(${JSON.stringify(name)},${JSON.stringify(letters)})`))
      .replace(/\$?[A-Z]{1,2}\$?\d+\b/g, m => hold(`__CELL(${JSON.stringify(m.replace(/\$/g, ''))})`))
      .replace(/<>/g, '!==').replace(/([^<>=!])=([^=])/g, '$1===$2')
      .replace(/__T(\d+)_/g, (m, i) => tokens[Number(i)]);
    const fn = new Function('__COL', '__CELL', 'SUMIFS', 'COUNTIFS', 'IF', 'OR', 'MAX', 'MIN', `return (${js});`);
    return fn(column, a1 => cellValue(sheet, a1), fns.SUMIFS, fns.COUNTIFS, fns.IF, fns.OR, fns.MAX, fns.MIN);
  }
  return evaluate;
}

module.exports = { makeEvaluator };
