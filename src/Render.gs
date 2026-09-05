// Draws one team tab as a dashboard: inputs (B1/B2), title banner, KPI cards, then the PREVIOUS QUARTER /
// FUTURE QUARTER(S) / PARTNER CONTRIBUTION tables, charts, and the trend data area from column T.
const DATA_COL = 20;
const LAST_COL = 14; // N
const FORMATS = { money: '$#,##0;-$#,##0', pct: '0.0%', int: '0', text: '@' };
const COLORS = {
  navy: '#1f3864', band: '#1f3864', header: '#e9eef6', page: '#f5f7fb', card: '#ffffff', border: '#c9d3e0', line: '#e3e8ef',
  label: '#3c4858', muted: '#8a94a6', up: '#2e7d32', down: '#c62828', barLow: '#f8cbad', barHigh: '#a9d08e',
};
// Metrics where an increase is bad news (colours the QoQ arrow).
const BAD_UP = ['churnCustomers', 'lostPipelineCount', 'downgradeCount', 'fullChurnCount', 'partnerChurnCustomers', 'forecastChurnCount'];

// [label, kind, key]. Keys index the metrics objects built in buildView.
const KPI_CARDS = [
  ['Net Added ARR', 'money', 'netAddedArr'],
  ['Attainment', 'pct', 'attainment'],
  ['Ending ARR', 'money', 'endingArr'],
  ['Renewal rate', 'pct', 'renewalRate'],
  ['Churn ARR', 'money', 'churnArr'],
  ['Active customers', 'int', 'activeCustomers'],
];
const PREVIOUS_ROWS = [
  ['Revenue Goal', 'money', 'revenueGoal'],
  ['Net Added ARR', 'money', 'netAddedArr'],
  ['Attainment (%)', 'pct', 'attainment'],
  ['Logo Goal', 'int', 'logoGoal'],
  ['Logo Attainment', 'int', 'logoAttainment'],
  ['Logo Attainment (%)', 'pct', 'logoAttainmentPct'],
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
  ['Conversion Rate (Activation : Conversion)', 'pct', 'conversionRate'],
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
  'renewals', 'wonRenewals', 'renewalRate', 'churnArr', 'churnCustomers', 'activeCustomers', 'conversionRate', 'lostPipelineCount',
  'lostPipelineArr', 'partnerNetAddedArr', 'partnerNewLogos', 'partnerChurnArr', 'partnerChurnCustomers'];

function renderTeamTab(sheet, view) {
  resetSheet(sheet, view.team, view.quarter);
  const trend = writeTrendData(sheet, view.trend);
  const previous = view.trend.length > 1 ? view.trend[view.trend.length - 2] : null;

  writeBanner(sheet, 3, `${view.team.toUpperCase()}   |   ${view.quarter} QBR`);
  let row = writeKpiCards(sheet, 5, view.current, previous) + 1;

  row = writeSection(sheet, row, 'PREVIOUS QUARTER', `${view.quarter} actuals vs goal, QoQ vs ${previous ? previous.quarter : 'n/a'}, trend from ${FIRST_QUARTER}`);
  row = Math.max(
    writeBlock(sheet, row, 2, ['Metric', view.quarter, 'QoQ', 'Trend'], PREVIOUS_ROWS, [view.current], previous, trend),
    writeBlock(sheet, row, 7, ['ARR bridge', view.quarter, '% of Starting'], ARR_ROWS, [view.current], null, null, view.current.startingArr),
    writeBlock(sheet, row, 11, ['Accounts', view.quarter, 'QoQ', 'Trend'], ACCOUNT_ROWS, [view.current], previous, trend),
  ) + 1;
  row = writeLists(sheet, row, [['Logos Won', view.current.logosWon, 2], ['Logos Lost', view.current.logosLost, 7]]) + 1;

  const [q1, q2] = view.future;
  row = writeSection(sheet, row, 'FUTURE QUARTER(S)', `Forecast for ${q1.quarter} and ${q2.quarter}, pipeline = open opportunities only`);
  row = Math.max(
    writeBlock(sheet, row, 2, ['Metric', `Q+1 (${q1.quarter})`, `Q+2 (${q2.quarter})`], FUTURE_ROWS, [q1, q2], null, null),
    writeBlock(sheet, row, 7, ['ARR forecast', `Q+1 (${q1.quarter})`, `Q+2 (${q2.quarter})`], FUTURE_ARR_ROWS, [q1, q2], null, null),
  ) + 1;

  row = writeSection(sheet, row, 'PARTNER CONTRIBUTION', `Opportunities in the ${PARTNER_GROUP} group for this team`);
  row = writeBlock(sheet, row, 2, ['Metric', view.quarter, 'QoQ', 'Trend'], PARTNER_ROWS, [view.current], previous, trend) + 1;

  row = writeSection(sheet, row, 'CHARTS', `${view.quarter} bridge, attainment, renewals, forecast and trend since ${FIRST_QUARTER}`);
  writeCharts(sheet, row, view, trend);
}

function resetSheet(sheet, team, quarter) {
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 24);
  sheet.setColumnWidths(2, LAST_COL - 1, 105);
  [2, 7, 11].forEach(c => sheet.setColumnWidth(c, 170)); // metric label columns
  sheet.setColumnWidth(LAST_COL + 1, 40);
  sheet.getRange(1, 1, 400, LAST_COL + 1).setBackground(COLORS.page).setFontFamily('Arial').setFontSize(10).setFontColor(COLORS.label);

  sheet.getRange('A1:A2').setValues([['Team'], ['Quarter']]).setFontSize(8).setFontColor(COLORS.muted);
  sheet.getRange('B1:B2').setValues([[team], [quarter]]).setBackground('#fff8dc').setFontWeight('bold')
    .setBorder(true, true, true, true, false, false, '#e0c36a', SpreadsheetApp.BorderStyle.SOLID);
  const now = new Date();
  const options = quarterOptions(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  if (options.indexOf(quarter) < 0) options.push(quarter);
  sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).build());
  sheet.getRange('D1').setValue(`Refreshed ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')} from Salesforce`)
    .setFontSize(8).setFontColor(COLORS.muted).setFontStyle('italic');
  sheet.getRange('D2').setValue('Change B1 (team) or B2 (quarter) and use QBR > Refresh this tab').setFontSize(8).setFontColor(COLORS.muted);
  sheet.setFrozenRows(3);
}

function writeBanner(sheet, row, text) {
  sheet.setRowHeight(row, 46);
  sheet.getRange(row, 2, 1, LAST_COL - 1).merge().setValue(text).setBackground(COLORS.navy).setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignment('left');
}

// Six cards, two columns each: label / big value / QoQ delta. Returns the row after the cards.
function writeKpiCards(sheet, row, current, previous) {
  sheet.setRowHeight(row, 20);
  sheet.setRowHeight(row + 1, 36);
  sheet.setRowHeight(row + 2, 20);
  // First card is the hero (3 columns), the rest 2 columns each, filling B:N.
  KPI_CARDS.forEach(([label, kind, key], i) => {
    const col = i === 0 ? 2 : 3 + i * 2;
    const span = i === 0 ? 3 : 2;
    const card = sheet.getRange(row, col, 3, span);
    card.setBackground(COLORS.card).setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, col, 1, span).merge().setValue(label.toUpperCase()).setFontSize(8).setFontColor(COLORS.muted)
      .setHorizontalAlignment('center').setVerticalAlignment('bottom');
    sheet.getRange(row + 1, col, 1, span).merge().setValue(cellValue(kind, current[key])).setNumberFormat(FORMATS[kind])
      .setFontSize(i === 0 ? 20 : 18).setFontWeight('bold').setFontColor(COLORS.navy).setHorizontalAlignment('center').setVerticalAlignment('middle');
    const delta = previous ? qoqText(kind, current[key], previous[key]) : '';
    sheet.getRange(row + 2, col, 1, span).merge().setValue(delta ? `${delta} QoQ` : '').setFontSize(8)
      .setFontColor(qoqColor(key, current[key], previous ? previous[key] : null)).setHorizontalAlignment('center').setVerticalAlignment('top');
  });
  return row + 3;
}

// Full-width navy band with a title and a muted subtitle. Returns the first row for content.
function writeSection(sheet, row, title, subtitle) {
  sheet.setRowHeight(row, 26);
  const text = `${title}     ${subtitle}`;
  sheet.getRange(row, 2, 1, LAST_COL - 1).merge().setBackground(COLORS.band).setVerticalAlignment('middle')
    .setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text)
      .setTextStyle(0, title.length, SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).setForegroundColor('#ffffff').build())
      .setTextStyle(title.length, text.length, SpreadsheetApp.newTextStyle().setBold(false).setFontSize(9).setForegroundColor('#c9d3e0').build())
      .build());
  return row + 1;
}

// Header + one line per spec row. `values` holds one metrics object per value column. `shareOf` adds a
// "% of <shareOf>" column for money rows. Returns the row after the block.
function writeBlock(sheet, row, col, header, spec, values, previous, trend, shareOf) {
  const width = header.length;
  const body = spec.map(([label, kind, key]) => {
    const line = [label].concat(values.map(m => cellValue(kind, m[key])));
    if (previous) line.push(qoqText(kind, values[0][key], previous[key]));
    if (trend) line.push(trend[key] ? `=SPARKLINE(${trend[key]},{"charttype","line";"linewidth",2;"color","${COLORS.navy}"})` : '');
    if (shareOf != null) line.push(kind === 'money' && shareOf ? values[0][key] / shareOf : '');
    return line.concat(Array(width - line.length).fill(''));
  });
  const colors = spec.map(([, , key]) => {
    const line = Array(width).fill(COLORS.label);
    if (previous) line[1 + values.length] = qoqColor(key, values[0][key], previous[key]);
    return line;
  });

  sheet.getRange(row, col, 1, width).setValues([header]).setFontWeight('bold').setFontColor(COLORS.navy).setBackground(COLORS.header)
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, col).setHorizontalAlignment('left');
  sheet.setRowHeight(row, 24);

  const range = sheet.getRange(row + 1, col, body.length, width);
  range.setValues(body).setBackground(COLORS.card).setFontColors(colors).setVerticalAlignment('middle');
  range.setBorder(null, null, null, null, false, true, COLORS.line, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, col, body.length + 1, width).setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row + 1, col, body.length, 1).setFontWeight('bold').setWrap(true);
  sheet.getRange(row + 1, col + 1, body.length, width - 1).setHorizontalAlignment('right').setWrap(true);
  spec.forEach(([, kind], i) => {
    sheet.getRange(row + 1 + i, col + 1, 1, values.length).setNumberFormat(FORMATS[kind]);
    if (kind === 'text') sheet.getRange(row + 1 + i, col + 1, 1, width - 1).merge().setHorizontalAlignment('left').setVerticalAlignment('top');
    if (kind === 'pct') addPercentBar(sheet, row + 1 + i, col + 1, values.length);
  });
  if (shareOf != null) sheet.getRange(row + 1, col + width - 1, body.length, 1).setNumberFormat(FORMATS.pct).setFontColor(COLORS.muted);
  return row + 1 + body.length;
}

function cellValue(kind, value) {
  if (kind === 'text') return value.length ? value.join('\n') : '-';
  return value == null ? '' : value;
}

function qoqText(kind, now, prev) {
  if (kind === 'text' || now == null || prev == null || now === prev) return '-';
  const arrow = now > prev ? '\u25B2 ' : '\u25BC ';
  const diff = Math.abs(now - prev);
  if (kind === 'pct') return `${arrow}${(diff * 100).toFixed(1)}pp`;
  if (kind === 'money') return arrow + formatMoney(diff);
  return `${arrow}${diff}`;
}

function qoqColor(key, now, prev) {
  if (now == null || prev == null || now === prev || typeof now !== 'number') return COLORS.muted;
  const improved = BAD_UP.indexOf(key) >= 0 ? now < prev : now > prev;
  return improved ? COLORS.up : COLORS.down;
}

// Data bars on percentage cells, capped at 100%.
function addPercentBar(sheet, row, col, width) {
  const range = sheet.getRange(row, col, 1, width);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(COLORS.barLow, SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue(COLORS.barHigh, SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

// Side-by-side lists ([title, items, column]), four columns wide. Returns the row after the longest list.
function writeLists(sheet, row, lists) {
  const height = Math.max(1, ...lists.map(([, items]) => items.length));
  lists.forEach(([title, items, col]) => {
    sheet.getRange(row, col, 1, 4).merge().setValue(`${title} (${items.length})`).setFontWeight('bold').setFontColor(COLORS.navy)
      .setBackground(COLORS.header).setVerticalAlignment('middle');
    sheet.setRowHeight(row, 24);
    const body = sheet.getRange(row + 1, col, height, 4);
    body.setBackground(COLORS.card);
    if (items.length) sheet.getRange(row + 1, col, items.length, 1).setValues(items.map(item => [item]));
    else sheet.getRange(row + 1, col).setValue('-').setFontColor(COLORS.muted);
    sheet.getRange(row, col, height + 1, 4).setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  });
  return row + 1 + height;
}

// Trend table at column T (header row 6, one row per quarter Q1-2026 -> selected).
// Returns { key: A1 range of that column's values } plus quarterCount and column(key, withHeader).
function writeTrendData(sheet, trendRows) {
  const body = trendRows.map(m => TREND_KEYS.map(key => (m[key] == null ? '' : m[key])));
  sheet.getRange(5, DATA_COL).setValue('Data behind the sparklines and charts (do not edit)').setFontColor(COLORS.muted).setFontSize(8);
  sheet.getRange(6, DATA_COL, 1, TREND_KEYS.length).setValues([TREND_KEYS]).setFontWeight('bold').setFontColor(COLORS.muted).setFontSize(8);
  sheet.getRange(7, DATA_COL, body.length, TREND_KEYS.length).setValues(body).setFontColor(COLORS.muted).setFontSize(8);
  const trend = { quarterCount: body.length };
  trend.column = key => sheet.getRange(6, DATA_COL + TREND_KEYS.indexOf(key), body.length + 1, 1);
  TREND_KEYS.forEach((key, i) => { trend[key] = sheet.getRange(7, DATA_COL + i, body.length, 1).getA1Notation(); });
  return trend;
}

function writeDataBlock(sheet, row, rows) {
  const range = sheet.getRange(row, DATA_COL, rows.length, rows[0].length);
  range.setValues(rows).setFontColor(COLORS.muted).setFontSize(8);
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
  // Cut the transparent base so the movements are visible: axis starts just below the lowest bar top.
  const bridgeFloor = Math.max(0, Math.floor(Math.min(m.startingArr, afterAdded, afterDowngrade, m.endingArr) * 0.9 / 1e5) * 1e5);
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

  const chart = type => sheet.newChart().setChartType(type).setNumHeaders(1)
    .setOption('legend', { position: 'bottom', textStyle: { color: COLORS.label, fontSize: 10 } })
    .setOption('titleTextStyle', { color: COLORS.navy, fontSize: 13, bold: true })
    .setOption('backgroundColor', COLORS.card)
    .setOption('chartArea', { left: 60, top: 40, width: '85%', height: '65%' });
  const place = (builder, r, c) => sheet.insertChart(
    builder.setPosition(r, c, 0, 0).setOption('width', 600).setOption('height', 300).build()
  );
  const rowsPerChart = 15;

  place(chart(Charts.ChartType.COLUMN).addRange(waterfall)
    .setOption('title', `ARR bridge ${view.quarter}`)
    .setOption('isStacked', true)
    .setOption('series', { 0: { color: 'transparent', visibleInLegend: false }, 1: { color: COLORS.up }, 2: { color: COLORS.down } })
    .setOption('vAxis', { format: 'short', gridlines: { color: COLORS.line }, viewWindow: { min: bridgeFloor } }), row, 2);
  place(chart(Charts.ChartType.BAR).addRange(attainment)
    .setOption('title', `Attainment vs goal ${view.quarter}`)
    .setOption('hAxis', { format: 'percent', minValue: 0, maxValue: 1, gridlines: { color: COLORS.line } })
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLORS.navy]), row, 8);
  row += rowsPerChart;
  place(chart(Charts.ChartType.PIE).addRange(renewals)
    .setOption('title', `Renewals ${view.quarter}: won vs lost`)
    .setOption('pieHole', 0.55)
    .setOption('colors', [COLORS.up, COLORS.down]), row, 2);
  place(chart(Charts.ChartType.COLUMN).addRange(forecast)
    .setOption('title', 'Q+1 / Q+2 forecast vs goal')
    .setOption('colors', [COLORS.muted, COLORS.navy, '#7f9cc9'])
    .setOption('vAxis', { format: 'short', gridlines: { color: COLORS.line } }), row, 8);
  row += rowsPerChart;
  place(chart(Charts.ChartType.LINE)
    .addRange(trend.column('quarter')).addRange(trend.column('revenueGoal')).addRange(trend.column('netAddedArr')).addRange(trend.column('churnArr'))
    .setOption('title', 'Net Added ARR vs goal by quarter')
    .setOption('colors', [COLORS.muted, COLORS.navy, COLORS.down])
    .setOption('pointSize', 5)
    .setOption('vAxis', { format: 'short', gridlines: { color: COLORS.line } }), row, 2);
  place(chart(Charts.ChartType.LINE)
    .addRange(trend.column('quarter')).addRange(trend.column('endingArr'))
    .setOption('title', 'Ending ARR by quarter')
    .setOption('colors', [COLORS.navy])
    .setOption('pointSize', 5)
    .setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'short', gridlines: { color: COLORS.line } }), row, 8);
}
