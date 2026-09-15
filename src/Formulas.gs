// Formula-backed metric cells. Every dollar / count metric derived from opportunities is written as a SUMIFS /
// COUNTIFS over the Raw Data tab keyed on the tab's team (B1) and the quarter, goals come from the Goals tab,
// Starting / Ending ARR from the ARR Ledger, and ratios reference the sibling cells of their block. Metrics with
// no sheet source (account counts, quarter elapsed, text) stay as values. The refresh still computes every number
// in Metrics.gs: the formulas let anyone validate a cell against the rows it is built from.
const GOALS_HEADER = ['Tab', 'Quarter', 'Net ARR Goal', 'New Logo Goal'];

const colLetter = c => {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const cellA1 = (row, col) => `${colLetter(col)}${row}`;
const sheetCol = (sheetName, header, name) => {
  const index = header.indexOf(name);
  if (index < 0) throw new Error(`No column "${name}" on ${sheetName}`);
  return `'${sheetName}'!${colLetter(index + 1)}:${colLetter(index + 1)}`;
};
const quote = text => `"${String(text).replace(/"/g, '""')}"`;

// One context per metrics column: `team` and `quarter` are the formula texts of the tab's keys ($B$1 / $B$2 or
// a literal for Q+1 / Q+2), `members` the ARR Ledger teams to sum, `addr` the A1 cell of every metric written so far.
function formulaContext(team, quarter, members, isSelected) {
  return { team: '$B$1', teamName: team, quarter: isSelected ? '$B$2' : quote(quarter), members, addr: {} };
}

// Keyed on the tab and, unless `quarterColumn` is null, on the quarter (the close-date quarter by default).
function rawFormula(fn, sumColumn, ctx, criteria, quarterColumn = 'Quarter') {
  const raw = name => sheetCol(RAW_SHEET, RAW_HEADER, name);
  const parts = [raw('Tab'), ctx.team].concat(quarterColumn ? [raw(quarterColumn), ctx.quarter] : []);
  criteria.forEach(([name, value]) => parts.push(raw(name), typeof value === 'string' && value.charAt(0) === '$' ? value : quote(value)));
  return `${fn}(${(sumColumn ? [raw(sumColumn)] : []).concat(parts).join(',')})`;
}
const rawSum = (column, ctx, ...criteria) => rawFormula('SUMIFS', column, ctx, criteria);
const rawCount = (ctx, ...criteria) => rawFormula('COUNTIFS', null, ctx, criteria);
const rawCountAllQuarters = (ctx, ...criteria) => rawFormula('COUNTIFS', null, ctx, criteria, null);
const perRenewalType = build => RENEWAL_RECORD_TYPES.map(build).join('+');
const perNewLogoType = build => NEW_LOGO_TYPES.map(build).join('+');
const notRenewal = () => RENEWAL_RECORD_TYPES.map(type => ['Record Type', `<>${type}`]);
const goalSum = (column, ctx) => `SUMIFS(${sheetCol(GOALS_SHEET, GOALS_HEADER, column)},${sheetCol(GOALS_SHEET, GOALS_HEADER, 'Tab')},${ctx.team},${sheetCol(GOALS_SHEET, GOALS_HEADER, 'Quarter')},${ctx.quarter})`;
const ledgerSum = (column, ctx) => (ctx.members || []).map(member =>
  `SUMIFS(${sheetCol(LEDGER_SHEET, LEDGER_HEADER, column)},${sheetCol(LEDGER_SHEET, LEDGER_HEADER, 'Team')},${member === ctx.teamName ? ctx.team : quote(member)},${sheetCol(LEDGER_SHEET, LEDGER_HEADER, 'Quarter')},${ctx.quarter})`).join('+');
const safeRatio = (num, den) => `IF(${den}=0,"",${num}/${den})`;
// Sibling references: null when the block has not written the metric (the caller then keeps the value).
const refs = (ctx, ...keys) => (keys.every(k => ctx.addr[k]) ? keys.map(k => ctx.addr[k]) : null);

const wonArr = ctx => rawSum('Delta ARR', ctx, ['Stage', 'Closed Won']);
const fullChurnArr = ctx => rawSum('Delta ARR', ctx, ['Bucket', 'Full churn']);
const downgradeArr = ctx => rawSum('Delta ARR', ctx, ['Bucket', 'Won renewal - downgrade']);
const newLogoCount = (ctx, ...criteria) => perNewLogoType(type => rawCount(ctx, ['Stage', 'Closed Won'], ['Type', type], ...criteria));
const newLogoArr = ctx => perNewLogoType(type => rawSum('Delta ARR', ctx, ['Stage', 'Closed Won'], ['Type', type]));
const netAddedArr = ctx => `${wonArr(ctx)}+${fullChurnArr(ctx)}`;
const forecastArr = ctx => `${perNewLogoType(type => rawSum('Expected Delta ARR', ctx, ['Type', type], ...notRenewal()))}+${rawSum('Expected Delta ARR', ctx, ['Type', 'Expand'], ...notRenewal())}`
  + `+${perRenewalType(type => rawSum('Expected Delta ARR', ctx, ['Record Type', type], ['Expected Delta ARR', '>0']))}`;
const forecastChurnArr = ctx => perRenewalType(type => rawSum('Expected Delta ARR', ctx, ['Record Type', type], ['Expected Delta ARR', '<0']));
const partner = () => ['Group', PARTNER_GROUP];

const METRIC_FORMULAS = {
  revenueGoal: ctx => goalSum('Net ARR Goal', ctx),
  logoGoal: ctx => goalSum('New Logo Goal', ctx),
  netAddedArr,
  attainment: ctx => { const r = refs(ctx, 'netAddedArr', 'revenueGoal'); return r && safeRatio(r[0], r[1]); },
  pace: ctx => { const r = refs(ctx, 'attainment', 'quarterElapsedPct'); return r && `IF(OR(${r[0]}="",${r[1]}=0),"",${r[0]}/${r[1]})`; },
  openPipelineArr: ctx => rawSum('Delta ARR', ctx, ['Bucket', 'Open pipeline']),
  pipelineCoverage: ctx => {
    const r = refs(ctx, 'openPipelineArr', 'revenueGoal', 'netAddedArr');
    if (r) return safeRatio(r[0], `MAX(0,${r[1]}-${r[2]})`);
    const f = refs(ctx, 'pipelineArr', 'revenueGoal');
    return f && safeRatio(f[0], f[1]);
  },
  logoAttainment: ctx => rawSum('Expected Logo Impact', ctx, ['Major Account', 'Yes']),
  logoForecast: ctx => rawSum('Expected Logo Impact', ctx, ['Major Account', 'Yes']),
  logosWonLand: ctx => rawCount(ctx, ['Stage', 'Closed Won'], ['Type', 'Land'], ['Major Account', 'Yes']),
  logosWonMajors: ctx => newLogoCount(ctx, ['Major Account', 'Yes']),
  logoAttainmentPct: ctx => { const r = refs(ctx, 'logoAttainment', 'logoGoal'); return r && safeRatio(`MAX(0,${r[0]})`, r[1]); },
  renewals: ctx => `${rawCount(ctx, ['Bucket', 'Won renewal*'])}+${rawCount(ctx, ['Bucket', 'Full churn'])}`,
  wonRenewals: ctx => rawCount(ctx, ['Bucket', 'Won renewal*']),
  renewalRate: ctx => { const r = refs(ctx, 'wonRenewals', 'renewals'); return r && safeRatio(r[0], r[1]); },
  churnArr: ctx => { const r = refs(ctx, 'downgradeArr', 'fullChurnArr'); return r ? `${r[0]}+${r[1]}` : `${downgradeArr(ctx)}+${fullChurnArr(ctx)}`; },
  churnCustomers: ctx => rawCount(ctx, ['Bucket', 'Full churn']),
  // Future quarters start from the previous quarter's (forecast) Ending ARR, see formulaContext callers.
  startingArr: ctx => (ctx.startingArr ? ctx.startingArr() : ledgerSum('Starting ARR', ctx)),
  endingArr: ctx => ledgerSum('Ending ARR', ctx),
  addedArr: ctx => `${wonArr(ctx)}-${downgradeArr(ctx)}`,
  downgradeArr,
  downgradeCount: ctx => rawCount(ctx, ['Bucket', 'Won renewal - downgrade']),
  fullChurnArr,
  fullChurnCount: ctx => rawCount(ctx, ['Bucket', 'Full churn']),
  grr: ctx => { const r = refs(ctx, 'startingArr', 'churnArr'); return r && safeRatio(`(${r[0]}+${r[1]})`, r[0]); },
  nrr: ctx => { const r = refs(ctx, 'startingArr', 'churnArr', 'addedArr'); return r && safeRatio(`(${r[0]}+${r[1]}+${r[2]}-(${newLogoArr(ctx)}))`, r[0]); },
  conversionRate: ctx => {
    const r = refs(ctx, 'activatedProspects');
    const lands = newLogoCount(ctx);
    return r && safeRatio(`(${lands})`, `(${lands}+${r[0]})`);
  },
  lostPipelineCount: ctx => rawCount(ctx, ['Bucket', 'Lost pipeline']),
  lostPipelineArr: ctx => rawSum('Delta ARR', ctx, ['Bucket', 'Lost pipeline']),
  // Pilots are not tied to the close-date quarter: active = open in the pilot stage in any quarter, completed = by
  // the quarter of the Pilot Actual End Date.
  activePilots: ctx => rawCountAllQuarters(ctx, ['Stage', PILOT_STAGE]),
  pilotsCompleted: ctx => rawFormula('COUNTIFS', null, ctx, [['Pilot Status', PILOT_COMPLETE_STATUS]], 'Pilot End Quarter'),
  netForecastArr: ctx => `${forecastArr(ctx)}+${forecastChurnArr(ctx)}`,
  netForecastPct: ctx => { const r = refs(ctx, 'netForecastArr', 'revenueGoal'); return r && safeRatio(r[0], r[1]); },
  pipelineArr: ctx => rawSum('Delta ARR', ctx, ['Bucket', 'Open pipeline']),
  renewalsDueCount: ctx => perRenewalType(type => rawCount(ctx, ['Bucket', 'Open pipeline'], ['Record Type', type])),
  forecastArr,
  forecastChurnArr,
  forecastChurnCount: ctx => perRenewalType(type => rawCount(ctx, ['Record Type', type], ['Expected Delta ARR', '<0'], ['Expected Logo Impact', '<0'])),
  forecastEndingArr: ctx => { const r = refs(ctx, 'startingArr', 'forecastArr', 'forecastChurnArr'); return r && `${r[0]}+${r[1]}+${r[2]}`; },
  partnerNetAddedArr: ctx => `${rawSum('Delta ARR', ctx, ['Stage', 'Closed Won'], partner())}+${rawSum('Delta ARR', ctx, ['Bucket', 'Full churn'], partner())}`,
  partnerNewLogos: ctx => newLogoCount(ctx, partner()),
  partnerChurnArr: ctx => `${rawSum('Delta ARR', ctx, ['Bucket', 'Won renewal - downgrade'], partner())}+${rawSum('Delta ARR', ctx, ['Bucket', 'Full churn'], partner())}`,
  partnerChurnCustomers: ctx => rawCount(ctx, ['Bucket', 'Full churn'], partner()),
};

// The formula for `key` in this column, or null to write the computed value. Formula-backed cells are also
// written when the value is null so a missing goal shows "" from the formula, like the value would.
function metricFormula(key, ctx) {
  if (!ctx || !METRIC_FORMULAS[key]) return null;
  const formula = METRIC_FORMULAS[key](ctx);
  return formula ? `=${formula}` : null;
}

// "Goals" sheet: Tab | Quarter | Net ARR Goal | New Logo Goal, the summed Goal__c records behind Revenue / Logo Goal.
// Refreshing a tab replaces that tab's rows only.
function writeGoals(tab, goals) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(GOALS_SHEET) || ss.insertSheet(GOALS_SHEET);
  const kept = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, GOALS_HEADER.length).getValues().filter(row => row[0] !== tab)
    : [];
  const rows = kept.concat(Object.keys(goals).map(q => [tab, q, goals[q].revenue || 0, goals[q].logos || 0]))
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])));
  sheet.clear();
  sheet.getRange(1, 1, 1, GOALS_HEADER.length).setValues([GOALS_HEADER]).setFontWeight('bold')
    .setBackground(COLORS.ink).setFontColor('#ffffff');
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, GOALS_HEADER.length).setValues(rows);
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat(FORMATS.money);
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.##');
  }
  sheet.getRange(1, 1, Math.max(2, rows.length + 1), GOALS_HEADER.length).setFontFamily(FONT).setFontSize(9);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, GOALS_HEADER.length, 130);
  return sheet;
}
