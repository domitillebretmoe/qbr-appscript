// Draws one team tab. Layout mirrors the Q3-26 QBR Workbook: B1 team, B2 quarter, then
// PREVIOUS QUARTER / FUTURE QUARTER(S) / PARTNER CONTRIBUTION blocks, charts, and a data area from column T.
const DATA_COL = 20;
const FORMATS = { money: '$#,##0', pct: '0.0%', int: '0', text: '@' };
const COLORS = { title: '#1f3864', header: '#d9e1f2', section: '#f2f2f2', up: '#2e7d32', down: '#c62828' };

// [label, kind, key]. Keys index the metrics objects built in buildView.
const PREVIOUS_ROWS = [
  ['Revenue Goal', 'money', 'revenueGoal'],
  ['Net Added ARR', 'money', 'netAddedArr'],
  ['Attainment (%)', 'pct', 'attainment'],
  ['Logo Goal', 'int', 'logoGoal'],
  ['Attainment', 'int', 'logoAttainment'],
  ['Attainment (%)', 'pct', 'logoAttainmentPct'],
  ['Ending ARR', 'money', 'endingArr'],
  ['# Renewals', 'int', 'renewals'],
  ['# Won Renewals', 'int', 'wonRenewals'],
  ['Renewal Rate', 'pct', 'renewalRate'],
  ['Churn - ARR $', 'money', 'churnArr'],
  ['Churn - Customer #', 'int', 'churnCustomers'],
  ['Top 3 Churns', 'text', 'topChurns'],
  ['Reasons for Churn', 'text', 'churnReasons'],
];
const ARR_ROWS = [
  ['Starting ARR', 'money', 'startingArr'],
  ['Added ARR', 'money', 'addedArr'],
  ['Downgrade $', 'money', 'downgradeArr'],
  ['Downgrade #', 'int', 'downgradeCount'],
  ['Full Churn $', 'money', 'fullChurnArr'],
  ['Full Churn #', 'int', 'fullChurnCount'],
  ['Churn ARR', 'money', 'churnArr'],
  ['Ending ARR', 'money', 'endingArr'],
];
const ACCOUNT_ROWS = [
  ['# Active Customers', 'int', 'activeCustomers'],
  ['# Major Customers', 'int', 'majorCustomers'],
  ['# Enterprise Customers', 'int', 'enterpriseCustomers'],
  ['# Activated Prospects', 'int', 'activatedProspects'],
  ['Conversion Rate Activation:Conversion', 'pct', 'conversionRate'],
  ['# Lost Pipeline', 'int', 'lostPipelineCount'],
  ['$ Lost Pipeline', 'money', 'lostPipelineArr'],
];
const FUTURE_ROWS = [
  ['Revenue Goal', 'money', 'revenueGoal'],
  ['Net Forecast ($)', 'money', 'netForecastArr'],
  ['Net Forecast (%)', 'pct', 'netForecastPct'],
  ['Logo Goal', 'int', 'logoGoal'],
  ['Net Forecast (#)', 'int', 'logoForecast'],
  ['Pipeline', 'money', 'pipelineArr'],
  ['Pipeline coverage', 'pct', 'pipelineCoverage'],
];
const FUTURE_ARR_ROWS = [
  ['Starting ARR', 'money', 'startingArr'],
  ['Forecast ARR', 'money', 'forecastArr'],
  ['Forecast Churn ARR', 'money', 'forecastChurnArr'],
  ['Forecast Churn #', 'int', 'forecastChurnCount'],
  ['Forecast Ending ARR', 'money', 'forecastEndingArr'],
];
const PARTNER_ROWS = [
  ['Net Added ARR', 'money', 'partnerNetAddedArr'],
  ['New Logo', 'int', 'partnerNewLogos'],
  ['Churn - ARR $', 'money', 'partnerChurnArr'],
  ['Churn - Customer #', 'int', 'partnerChurnCustomers'],
];
// Metrics kept per quarter in the data area (drives sparklines and the trend chart).
const TREND_KEYS = ['quarter', 'revenueGoal', 'netAddedArr', 'attainment', 'logoAttainment', 'logoAttainmentPct', 'endingArr',
  'renewals', 'wonRenewals', 'renewalRate', 'churnArr', 'churnCustomers', 'conversionRate', 'lostPipelineCount', 'lostPipelineArr',
  'partnerNetAddedArr', 'partnerNewLogos', 'partnerChurnArr', 'partnerChurnCustomers'];

function renderTeamTab(sheet, view) {
  resetSheet(sheet, view.team, view.quarter);
  const trend = writeTrendData(sheet, view.trend);
  const previous = view.trend.length > 1 ? view.trend[view.trend.length - 2] : null;

  let row = writeTitle(sheet, 4, 'PREVIOUS QUARTER');
  const blocks = [
    writeBlock(sheet, row, 2, ['Metric(s)', view.quarter, 'QoQ', 'Trend'], PREVIOUS_ROWS, [view.current], previous, trend),
    writeBlock(sheet, row, 7, ['Metric', view.quarter], ARR_ROWS, [view.current], null, null),
    writeBlock(sheet, row, 10, ['Metric', view.quarter, 'QoQ', 'Trend'], ACCOUNT_ROWS, [view.current], previous, trend),
  ];
  row = Math.max(...blocks) + 1;
  row = writeLists(sheet, row, [['Logos Won', view.current.logosWon], ['Logos Lost', view.current.logosLost]]) + 1;

  row = writeTitle(sheet, row, 'FUTURE QUARTER(S)');
  const [q1, q2] = view.future;
  row = Math.max(
    writeBlock(sheet, row, 2, ['Metric(s)', `Q+1 (${q1.quarter})`, `Q+2 (${q2.quarter})`], FUTURE_ROWS, [q1, q2], null, null),
    writeBlock(sheet, row, 7, ['Metric', `Q+1 (${q1.quarter})`, `Q+2 (${q2.quarter})`], FUTURE_ARR_ROWS, [q1, q2], null, null),
  ) + 1;

  row = writeTitle(sheet, row, 'PARTNER CONTRIBUTION');
  row = writeBlock(sheet, row, 2, ['Metric(s)', view.quarter, 'QoQ', 'Trend'], PARTNER_ROWS, [view.current], previous, trend) + 1;

  writeCharts(sheet, row, view, trend);
  sheet.setColumnWidths(2, 12, 120);
  [2, 7, 10].forEach(col => sheet.setColumnWidth(col, 230));
}

function resetSheet(sheet, team, quarter) {
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange('A1:B2').setValues([['Team', team], ['Quarter', quarter]]);
  sheet.getRange('A1:A2').setFontWeight('bold');
  sheet.getRange('B1:B2').setBackground('#fff2cc');
  const now = new Date();
  const options = quarterOptions(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  if (options.indexOf(quarter) < 0) options.push(quarter);
  sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).build());
  sheet.getRange('D1').setValue(`Refreshed ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`).setFontColor('#888888');
  sheet.setHiddenGridlines(true);
}

function writeTitle(sheet, row, text) {
  sheet.getRange(row, 2).setValue(text).setFontWeight('bold').setFontSize(13).setFontColor(COLORS.title);
  return row + 2;
}

// Writes header + one line per spec row. `values` holds one metrics object per value column.
// Returns the row after the block.
function writeBlock(sheet, row, col, header, spec, values, previous, trend) {
  const body = spec.map(([label, kind, key]) => {
    const line = [label].concat(values.map(m => cellValue(kind, m[key])));
    if (previous) line.push(qoqText(kind, values[0][key], previous[key]));
    if (trend) line.push(trend[key] ? `=SPARKLINE(${trend[key]},{"charttype","line";"linewidth",2;"color","${COLORS.title}"})` : '');
    return line;
  });
  const width = header.length;
  sheet.getRange(row, col, 1, width).setValues([header]).setFontWeight('bold').setBackground(COLORS.header);
  const range = sheet.getRange(row + 1, col, body.length, width);
  range.setValues(body.map(line => line.concat(Array(width - line.length).fill(''))));
  spec.forEach(([, kind], i) => sheet.getRange(row + 1 + i, col + 1, 1, values.length).setNumberFormat(FORMATS[kind]));
  spec.forEach(([, kind], i) => { if (kind === 'pct') addPercentBar(sheet, row + 1 + i, col + 1, values.length); });
  range.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row + 1, col, body.length, 1).setBackground(COLORS.section).setWrap(true);
  return row + 1 + body.length;
}

function cellValue(kind, value) {
  if (kind === 'text') return value.length ? value.join('\n') : '–';
  return value == null ? '' : value;
}

function qoqText(kind, now, prev) {
  if (kind === 'text' || now == null || prev == null || now === prev) return '–';
  const arrow = now > prev ? '↑ ' : '↓ ';
  const diff = Math.abs(now - prev);
  if (kind === 'pct') return `${arrow}${(diff * 100).toFixed(1)}pp`;
  if (kind === 'money') return arrow + formatMoney(diff);
  return `${arrow}${diff}`;
}

// Data bars on percentage cells, capped at 100%.
function addPercentBar(sheet, row, col, width) {
  const range = sheet.getRange(row, col, 1, width);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#f8cbad', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#a9d08e', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

// Side-by-side lists (Logos Won at column B, Logos Lost at column G). Returns the row after the longest list.
function writeLists(sheet, row, lists) {
  const cols = [2, 7];
  lists.forEach(([title, items], i) => {
    sheet.getRange(row, cols[i]).setValue(title).setFontWeight('bold').setBackground(COLORS.header);
    if (items.length) sheet.getRange(row + 1, cols[i], items.length, 1).setValues(items.map(item => [item]));
  });
  return row + 1 + Math.max(1, ...lists.map(([, items]) => items.length));
}

// Trend table at column T (header row 6, one row per quarter Q1-2026 → selected).
// Returns { key: A1 range of that column's values } plus quarterCount and column(key, withHeader).
function writeTrendData(sheet, trendRows) {
  const body = trendRows.map(m => TREND_KEYS.map(key => (m[key] == null ? '' : m[key])));
  sheet.getRange(6, DATA_COL, 1, TREND_KEYS.length).setValues([TREND_KEYS]).setFontWeight('bold').setFontColor('#888888');
  sheet.getRange(7, DATA_COL, body.length, TREND_KEYS.length).setValues(body).setFontColor('#888888');
  const trend = { quarterCount: body.length };
  trend.column = key => sheet.getRange(6, DATA_COL + TREND_KEYS.indexOf(key), body.length + 1, 1);
  TREND_KEYS.forEach((key, i) => { trend[key] = sheet.getRange(7, DATA_COL + i, body.length, 1).getA1Notation(); });
  return trend;
}

function writeDataBlock(sheet, row, rows) {
  const range = sheet.getRange(row, DATA_COL, rows.length, rows[0].length);
  range.setValues(rows).setFontColor('#888888');
  return range;
}

function writeCharts(sheet, row, view, trend) {
  const m = view.current;
  let dataRow = 8 + trend.quarterCount;

  // Waterfall as stacked columns: transparent base + green up + red down.
  const afterAdded = m.startingArr + m.addedArr;
  const afterDowngrade = afterAdded + m.downgradeArr;
  const waterfall = writeDataBlock(sheet, dataRow, [
    ['Step', 'Base', 'Up', 'Down'],
    ['Starting ARR', 0, m.startingArr, 0],
    ['Added ARR', m.startingArr, m.addedArr, 0],
    ['Downgrade', afterDowngrade, 0, -m.downgradeArr],
    ['Full Churn', m.endingArr, 0, -m.fullChurnArr],
    ['Ending ARR', 0, m.endingArr, 0],
  ]);
  dataRow += 7;
  const attainment = writeDataBlock(sheet, dataRow, [
    ['Metric', 'Attainment'],
    ['Revenue', m.attainment || 0],
    ['Logos', m.logoAttainmentPct || 0],
    ['Renewal rate', m.renewalRate || 0],
  ]);
  dataRow += 5;
  const renewals = writeDataBlock(sheet, dataRow, [
    ['Outcome', 'Renewals'],
    ['Won', m.wonRenewals],
    ['Lost', m.renewals - m.wonRenewals],
  ]);
  dataRow += 4;
  const forecast = writeDataBlock(sheet, dataRow, [['Quarter', 'Goal', 'Net Forecast', 'Pipeline']]
    .concat(view.future.map(f => [f.quarter, f.revenueGoal, f.netForecastArr, f.pipelineArr])));

  const chart = type => sheet.newChart().setChartType(type).setOption('legend', { position: 'bottom' }).setNumHeaders(1);
  const place = (builder, r, c) => sheet.insertChart(builder.setPosition(r, c, 0, 0).setOption('width', 560).setOption('height', 320).build());

  place(chart(Charts.ChartType.COLUMN).addRange(waterfall)
    .setOption('title', `ARR bridge ${view.quarter}`)
    .setOption('isStacked', true)
    .setOption('series', { 0: { color: 'transparent', visibleInLegend: false }, 1: { color: COLORS.up }, 2: { color: COLORS.down } })
    .setOption('vAxis', { format: 'short' }), row, 2);
  place(chart(Charts.ChartType.BAR).addRange(attainment)
    .setOption('title', `Attainment vs goal ${view.quarter}`)
    .setOption('hAxis', { format: 'percent', minValue: 0, maxValue: 1 })
    .setOption('colors', [COLORS.title]), row, 7);
  row += 17;
  place(chart(Charts.ChartType.PIE).addRange(renewals)
    .setOption('title', `Renewal rate ${view.quarter}`)
    .setOption('pieHole', 0.5)
    .setOption('colors', [COLORS.up, COLORS.down]), row, 2);
  place(chart(Charts.ChartType.COLUMN).addRange(forecast)
    .setOption('title', 'Q+1 / Q+2 forecast vs goal')
    .setOption('vAxis', { format: 'short' }), row, 7);
  row += 17;
  place(chart(Charts.ChartType.LINE)
    .addRange(trend.column('quarter')).addRange(trend.column('revenueGoal')).addRange(trend.column('netAddedArr')).addRange(trend.column('churnArr'))
    .setOption('title', 'Net Added ARR vs goal by quarter')
    .setOption('vAxis', { format: 'short' }), row, 2);
  place(chart(Charts.ChartType.LINE)
    .addRange(trend.column('quarter')).addRange(trend.column('endingArr'))
    .setOption('title', 'Ending ARR by quarter')
    .setOption('vAxis', { format: 'short' }), row, 7);
}
