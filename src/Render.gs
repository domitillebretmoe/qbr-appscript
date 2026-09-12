// Draws one team tab as a dashboard: inputs (B1/B2), title banner, KPI cards, then the PREVIOUS QUARTER /
// FUTURE QUARTER(S) / PARTNER CONTRIBUTION tables, charts, and the trend data area from column T.
const DATA_COL = 20;
const LAST_COL = 14; // N
const PAGE_ROWS = 600;
const FORMATS = { money: '$#,##0;[Red]($#,##0)', pct: '0.0%', int: '0', mult: '0.0"x"', text: '@', date: 'yyyy-mm-dd' };
// Chart labels and axes pick up the number format of the source cells.
const CHART_FORMATS = { money: '$#,##0.00,,"M";[Red]-$#,##0.00,,"M"', pct: '0%', int: '0' };
// Sparklines need a few quarters before they say anything.
const MIN_TREND_QUARTERS = 4;
// Attainment-style percentages get red / amber / green instead of a data bar.
const RAG_KEYS = ['attainment', 'logoAttainmentPct', 'netForecastPct', 'pace'];
const RAG = { amber: 0.7, green: 1 };
const DEFAULT_TABLE_WIDTH = 6;
const NUMERIC_KINDS = ['money', 'pct', 'int', 'mult'];
// Cognition palette: warm off-white page, near-black ink, electric-blue accent, flat white cards.
const FONT = 'Inter';
const COLORS = {
  ink: '#191919', accent: '#2200ff', accentSoft: '#8f9bff', page: '#f7f6f5', card: '#ffffff', band: '#eeedeb', header: '#f7f6f5',
  border: '#e5e5e5', line: '#efeeec', label: '#191919', muted: '#737373', onDark: '#a3a3a3',
  up: '#15803d', down: '#fa5050', upOnDark: '#4ade80', barLow: '#ffffff', barHigh: '#c5ccff',
  ragRed: '#fee2e2', ragAmber: '#fef3c7', ragGreen: '#dcfce7', amber: '#b45309',
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
  ['Quarter elapsed (%)', 'pct', 'quarterElapsedPct'],
  ['Pace (attainment / elapsed)', 'pct', 'pace'],
  ['Open pipeline (this quarter)', 'money', 'openPipelineArr'],
  ['Pipeline coverage of remaining goal', 'mult', 'pipelineCoverage'],
  ['Logo Goal', 'int', 'logoGoal'],
  ['Logo Attainment (expected, incl. open opps)', 'int', 'logoAttainment'],
  ['Logos Won (Closed Won Land / MSP, Majors)', 'int', 'logosWonMajors'],
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
  ['GRR (%)', 'pct', 'grr'],
  ['NRR (%)', 'pct', 'nrr'],
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
  ['Pipeline coverage of goal', 'mult', 'pipelineCoverage'],
  ['# Renewals due', 'int', 'renewalsDueCount'],
  ['ARR up for renewal', 'money', 'renewalArrDue'],
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
// Linked tables: [header, kind]; kind "link" cells are { text, url }.
const OPP_COLUMNS = [['Account', 'link'], ['Opportunity', 'link'], ['Type', 'text'], ['Close date', 'date'], ['Delta ARR', 'money'], ['Owner', 'text']];
const RENEWAL_COLUMNS = [['Account', 'link'], ['Opportunity', 'link'], ['Record type', 'text'], ['Close date', 'date'], ['Delta ARR', 'money'], ['Owner', 'text']];
const DEAL_COLUMNS = [['Account', 'link'], ['Opportunity', 'link'], ['Stage', 'text'], ['Close date', 'date'], ['Delta ARR', 'money'], ['Owner', 'text']];
const CUSTOMER_COLUMNS = [['#', 'int'], ['Account', 'link'], ['Team', 'text'], ['Current ARR', 'money'], ['% of active ARR', 'pct'], ['Open opp', 'text']];
const RENEWAL_DUE_COLUMNS = [['Account', 'link'], ['Opportunity', 'link'], ['Close date', 'date'], ['Current ARR', 'money'], ['Expected Delta ARR', 'money'], ['Owner', 'text']];
const OWNER_COLUMNS = [['Owner', 'text'], ['Net Added ARR', 'money'], ['# Won', 'int'], ['Churn ARR', 'money'], ['Open pipeline (Q)', 'money'], ['Pipeline Q+1', 'money']];
const QUALITY_COLUMNS = [['Issue', 'text'], ['Account', 'link'], ['Opportunity', 'link'], ['Detail', 'text'], ['Close date', 'date'], ['Owner', 'text']];
// Full-width rep table (13 columns, B..N), same measures as the Salesforce Majors Rep Performance tab.
const REP_COLUMNS = [['Rep', 'link'], ['Months in seat', 'mult'], ['Accts owned', 'int'], ['Rep goal (FY)', 'money'], ['Won ARR (FY)', 'money'],
  ['Attainment', 'pct'], ['Coverage', 'pct'], ['Meetings (30d)', 'int'], ['Activities (30d)', 'int'], ['Acct coverage (30d)', 'pct'],
  ['Pipeline created (Q)', 'money'], ['Stalled >60d', 'money'], ['Renewal risk', 'pct']];

// Hover note on each metric label (same definitions as the Definitions tab, in one line).
const KPI_NOTES = {
  revenueGoal: 'Net ARR goal for the team and quarter (Goals By Quarter - Net ARR report).',
  netAddedArr: 'Delta ARR of Closed Won opportunities + Delta ARR of Closed Lost renewals (full churn), close date in the quarter.',
  attainment: 'Net Added ARR / Revenue Goal. Green >= 100%, amber 70-99%, red < 70%.',
  quarterElapsedPct: 'Calendar days of the quarter elapsed at refresh time / days in the quarter (100% once the quarter is over).',
  pace: 'Attainment / Quarter elapsed: 100% = on a straight-line path to the goal, above = ahead, below = behind.',
  openPipelineArr: 'Delta ARR of open opportunities with a close date in the selected quarter.',
  pipelineCoverage: 'Open pipeline / remaining goal (goal - Net Added ARR); empty once the goal is met. Future quarters: pipeline / goal.',
  logoGoal: 'New Logos goal for the quarter.',
  logoAttainment: 'Sum of Expected Logo Impact of all the quarter\'s opportunities on Major accounts, open ones included (churned logos count -1). A forecast until the quarter closes.',
  logosWonMajors: 'Number of Closed Won "Land" or "MSP" opportunities on Major accounts: logos actually landed so far this quarter.',
  logoAttainmentPct: 'Logo Attainment / Logo Goal (attainment below 0 counts as 0).',
  renewals: 'Won renewals + full churn (Closed Won + Closed Lost renewal opportunities).',
  wonRenewals: 'Closed Won opportunities of record type Renewal or Fed - Renewal.',
  renewalRate: 'Won renewals / (won renewals + full churn), by count.',
  churnArr: 'Downgrade ARR (Closed Won renewals with negative Delta ARR) + Full Churn ARR (Closed Lost renewals).',
  churnCustomers: 'Number of Closed Lost renewal opportunities (full churn).',
  topChurns: 'Three largest churn or downgrade amounts in the quarter.',
  churnReasons: 'Distinct Closed Lost reasons on the churn and downgrade opportunities.',
  startingArr: 'Ending ARR of the previous quarter (ARR Ledger). Roll-up tabs sum their member teams.',
  addedArr: 'Net Added ARR - Churn ARR: the gross ARR added by won opportunities.',
  downgradeArr: 'Delta ARR of Closed Won renewals with negative Delta ARR.',
  downgradeCount: 'Number of Closed Won renewals with negative Delta ARR.',
  fullChurnArr: 'Delta ARR of Closed Lost renewals.',
  fullChurnCount: 'Number of Closed Lost renewals.',
  endingArr: 'Starting ARR + Net Added ARR (seeded quarters keep the FY27 QBR Cockpit values).',
  grr: 'Gross revenue retention: (Starting ARR + Downgrade ARR + Full Churn ARR) / Starting ARR.',
  nrr: 'Net revenue retention: GRR plus expansion from existing customers (won ARR that is not a new logo) / Starting ARR.',
  activeCustomers: 'Accounts with Current ARR > 0.',
  majorCustomers: 'Active customers with the Major admin tag.',
  enterpriseCustomers: 'Active customers without the Major admin tag.',
  activatedProspects: 'Accounts with no ARR and at least one open opportunity.',
  conversionRate: 'New logos won in the quarter / activated prospects.',
  lostPipelineCount: 'Closed Lost opportunities that are not renewals (Land / Expand).',
  lostPipelineArr: 'Delta ARR of Closed Lost non-renewal opportunities.',
  netForecastArr: 'Expected Delta ARR of Land / MSP + Expand + renewals expected to grow, plus forecast churn (renewals expected to shrink).',
  netForecastPct: 'Net Forecast / Revenue Goal.',
  logoForecast: 'Sum of Expected Logo Impact on Major accounts in the quarter.',
  pipelineArr: 'Delta ARR of open opportunities with a close date in the quarter.',
  renewalsDueCount: 'Open renewal opportunities (Renewal / Fed - Renewal) with a close date in the quarter.',
  renewalArrDue: 'Current ARR of the accounts with an open renewal in the quarter (each account counted once).',
  forecastArr: 'Expected Delta ARR of Land / MSP + Expand + renewals expected to grow.',
  forecastChurnArr: 'Expected Delta ARR of renewals expected to shrink.',
  forecastChurnCount: 'Renewals expected to shrink whose Expected Logo Impact is negative (full churn expected).',
  forecastEndingArr: 'Starting ARR + Net Forecast.',
  partnerNetAddedArr: `Net Added ARR of opportunities in the ${PARTNER_GROUP} group.`,
  partnerNewLogos: `New logos won through the ${PARTNER_GROUP} group.`,
  partnerChurnArr: `Churn ARR of opportunities in the ${PARTNER_GROUP} group.`,
  partnerChurnCustomers: `Full churn count in the ${PARTNER_GROUP} group.`,
};

// Metrics kept per quarter in the data area (drives sparklines and the trend chart).
const TREND_KEYS = ['quarter', 'revenueGoal', 'netAddedArr', 'attainment', 'logoAttainment', 'logosWonMajors', 'logoAttainmentPct', 'endingArr',
  'renewals', 'wonRenewals', 'renewalRate', 'churnArr', 'churnCustomers', 'activeCustomers', 'conversionRate', 'lostPipelineCount',
  'lostPipelineArr', 'partnerNetAddedArr', 'partnerNewLogos', 'partnerChurnArr', 'partnerChurnCustomers', 'grr', 'nrr', 'pace',
  'openPipelineArr', 'pipelineCoverage'];

function renderTeamTab(sheet, view) {
  resetSheet(sheet, view.team, view.quarter);
  const trend = writeTrendData(sheet, view.trend);
  const sparklines = trend.quarterCount >= MIN_TREND_QUARTERS ? trend : null;
  const trendHeader = first => [first, view.quarter, 'QoQ'].concat(sparklines ? ['Trend'] : []);
  const previous = view.trend.length > 1 ? view.trend[view.trend.length - 2] : null;

  const members = view.members && view.members.length > 1 ? `   (roll-up of ${view.members.map(teamToken).join(', ')})` : '';
  const status = statusText(view.current);
  writeBanner(sheet, 3, view.team, `${view.quarter} QBR - ${status}${members}`);
  // Metric cells are formulas over Raw Data / Goals / ARR Ledger (see Formulas.gs); the KPI cards are written after
  // the blocks so they can point at the block cells.
  const memberTeams = view.members || [view.team];
  const current = formulaContext(view.team, view.quarter, memberTeams, true);
  const kpiRow = 5;
  let row = kpiRow + 4;

  row = writeSection(sheet, row, `${view.current.status} QUARTER`, `${view.quarter} ${status.toLowerCase()} vs goal, QoQ vs ${previous ? previous.quarter : 'n/a'}, trend from ${FIRST_QUARTER}`);
  row = Math.max(
    writeBlock(sheet, row, 2, trendHeader('Metric'), PREVIOUS_ROWS, [view.current], previous, sparklines, null, [current]),
    writeBlock(sheet, row, 7, ['ARR bridge', view.quarter, '% of Starting'], ARR_ROWS, [view.current], null, null, view.current.startingArr, [current]),
    writeBlock(sheet, row, 11, trendHeader('Accounts'), ACCOUNT_ROWS, [view.current], previous, sparklines, null, [current]),
  ) + 1;
  writeKpiCards(sheet, kpiRow, view.current, previous, current.addr);
  const lists = view.current.lists;
  const oppTable = (title, col, columns, opps, middle) => ({ title: `${title} (${opps.length})`, col, columns, rows: opps.map(o => oppRow(o, middle)) });
  row = writeTables(sheet, row, [
    { title: dealsWonTitle(view.quarter, lists.dealsWon, view.current.wonCount, view.current.wonArr),
      col: 2, columns: OPP_COLUMNS, rows: lists.dealsWon.map(o => oppRow(o, x => x.type)) },
    oppTable('Logos Won', 9, OPP_COLUMNS, lists.logosWon, o => o.type),
  ]) + 1;
  row = writeTables(sheet, row, [
    oppTable('Lost Pipeline', 2, OPP_COLUMNS, lists.lostPipeline, o => o.type),
  ]) + 1;
  row = writeTables(sheet, row, [
    oppTable('Churned Customers', 2, RENEWAL_COLUMNS, lists.churned, o => o.recordType),
    oppTable('Downgrade Customers', 9, RENEWAL_COLUMNS, lists.downgrades, o => o.recordType),
  ]) + 1;
  row = writeTables(sheet, row, [
    oppTable('Renewals Won', 2, RENEWAL_COLUMNS, lists.renewalsWon, o => o.recordType),
    oppTable('Renewals Lost', 9, RENEWAL_COLUMNS, lists.renewalsLost, o => o.recordType),
  ]) + 1;
  const customers = (title, col, accounts, total) => ({
    title: `${title} (${accounts.length} of ${total} active)`, col, columns: CUSTOMER_COLUMNS,
    rows: accounts.map((a, i) => [i + 1, link(a.name, a.url), a.team, a.currentArr, view.current.activeArr ? a.currentArr / view.current.activeArr : '', a.hasOpenOpp ? 'Yes' : 'No']),
  });
  row = writeTables(sheet, row, [
    customers('Top 10 Major Customers', 2, view.current.topMajors, view.current.majorCustomers),
    customers('Top 10 Enterprise Customers', 9, view.current.topEnterprise, view.current.enterpriseCustomers),
  ]) + 1;

  const [q1, q2] = view.future;
  const future1 = formulaContext(view.team, q1.quarter, memberTeams, false);
  const future2 = formulaContext(view.team, q2.quarter, memberTeams, false);
  future1.startingArr = () => METRIC_FORMULAS.endingArr(current);
  future2.startingArr = () => future1.addr.forecastEndingArr;
  row = writeSection(sheet, row, 'FUTURE QUARTER(S)', `Forecast for ${q1.quarter} and ${q2.quarter}, pipeline = open opportunities only`);
  row = Math.max(
    writeBlock(sheet, row, 2, ['Metric', `Q+1 (${q1.quarter})`, `Q+2 (${q2.quarter})`], FUTURE_ROWS, [q1, q2], null, null, null, [future1, future2]),
    writeBlock(sheet, row, 7, ['ARR forecast', `Q+1 (${q1.quarter})`, `Q+2 (${q2.quarter})`], FUTURE_ARR_ROWS, [q1, q2], null, null, null, [future1, future2]),
  ) + 1;
  row = writeTables(sheet, row, [
    oppTable(`Top 10 Deals Q+1 ${q1.quarter}`, 2, DEAL_COLUMNS, q1.topDeals, o => o.stage),
    oppTable(`Top 10 Deals Q+2 ${q2.quarter}`, 9, DEAL_COLUMNS, q2.topDeals, o => o.stage),
  ]) + 1;
  const renewalsDue = (title, col, f) => ({
    title: `${title} (${f.renewalsDueCount}, ${formatMoney(f.renewalArrDue)} up for renewal)`, col, columns: RENEWAL_DUE_COLUMNS,
    rows: f.renewalsDue.map(o => [link(o.account, o.accountUrl), link(o.url ? 'Link' : '-', o.url), o.closeDate || '', o.accountArr, o.expectedDeltaArr, o.owner || '']),
  });
  row = writeTables(sheet, row, [
    renewalsDue(`Renewals due Q+1 ${q1.quarter}`, 2, q1),
    renewalsDue(`Renewals due Q+2 ${q2.quarter}`, 9, q2),
  ]) + 1;
  const predictedChurn = (title, col, f) => ({
    title: `${title} (${f.predictedChurn.length}, ${formatMoney(f.forecastChurnArr)} expected)`, col, columns: RENEWAL_DUE_COLUMNS,
    rows: f.predictedChurn.map(o => [link(o.account, o.accountUrl), link(o.url ? 'Link' : '-', o.url), o.closeDate || '', o.accountArr, o.expectedDeltaArr, o.owner || '']),
  });
  row = writeTables(sheet, row, [
    predictedChurn(`Predicted Churn Q+1 ${q1.quarter}`, 2, q1),
    predictedChurn(`Predicted Churn Q+2 ${q2.quarter}`, 9, q2),
  ]) + 1;

  row = writeSection(sheet, row, 'PARTNER CONTRIBUTION', `Opportunities in the ${PARTNER_GROUP} group for this team`);
  row = writeBlock(sheet, row, 2, trendHeader('Metric'), PARTNER_ROWS, [view.current], previous, sparklines, null, [current]) + 1;

  const reps = view.reps || [];
  row = writeSection(sheet, row, 'REP ACTIVITY & PERFORMANCE', `Global GTM Dashboard > Majors Rep Performance measures for the team's reps: FY${parseQuarter(view.quarter).fy} attainment by opp owner, activity = Gong-synced, last ${REP_ACTIVITY_DAYS} days`);
  row = writeTables(sheet, row, [
    { title: `Reps (${reps.length})`, col: 2, width: LAST_COL - 1, columns: REP_COLUMNS,
      rows: reps.map(r => [link(r.name, r.url), r.monthsInSeat, r.accountsOwned, r.repGoal, r.fyWonArr, r.attainmentPct, r.coveragePct,
        r.meetings30d, r.activities30d, r.activityCoveragePct, r.pipelineCreatedArr, r.stalledArr, r.renewalRiskPct]) },
  ]) + 1;

  row = writeSection(sheet, row, 'OWNERS & DATA QUALITY', `${view.quarter} by opportunity owner; Salesforce hygiene across ${view.quarter}, ${q1.quarter} and ${q2.quarter}`);
  const owners = view.owners || [];
  const issues = view.dataQuality || [];
  row = writeTables(sheet, row, [
    { title: `Owner performance (${owners.length})`, col: 2, columns: OWNER_COLUMNS,
      rows: owners.map(o => [o.owner || '-', o.netAddedArr, o.wonCount, o.churnArr, o.openPipelineArr, o.nextPipelineArr]) },
    { title: `Data quality (${issues.length} ${issues.length === 1 ? 'issue' : 'issues'})`, col: 9, columns: QUALITY_COLUMNS,
      rows: issues.map(i => (i.opp
        ? [i.issue, link(i.opp.account, i.opp.accountUrl), link(i.opp.url ? 'Link' : '-', i.opp.url), i.detail, i.opp.closeDate || '', i.opp.owner || '']
        : [i.issue, link(i.account.name, i.account.url), '-', i.detail, '', i.account.owner || ''])) },
  ]) + 1;

  row = writeSection(sheet, row, 'CHARTS', `${view.quarter} bridge, attainment, renewals, forecast and trend since ${FIRST_QUARTER}`);
  const lastRow = writeCharts(sheet, row, view, trend);
  writePdfLink(sheet, lastRow);
}

// One-click PDF of the dashboard area (landscape, fit to width, no gridlines) through the Sheets export endpoint;
// it opens in the browser with the user's own Google session, so the script needs no Drive scope.
function pdfExportUrl(sheet, lastRow) {
  const params = {
    format: 'pdf', gid: sheet.getSheetId(), size: 'A4', portrait: false, fitw: true, scale: 4,
    gridlines: false, printtitle: false, sheetnames: false, pagenumbers: false, fzr: false, horizontal_alignment: 'CENTER',
    top_margin: 0.3, bottom_margin: 0.3, left_margin: 0.3, right_margin: 0.3,
    r1: 0, c1: 0, r2: lastRow, c2: LAST_COL + 1,
  };
  const query = Object.keys(params).map(k => `${k}=${encodeURIComponent(params[k])}`).join('&');
  return `https://docs.google.com/spreadsheets/d/${sheet.getParent().getId()}/export?${query}`;
}

function writePdfLink(sheet, lastRow) {
  sheet.getRange(2, LAST_COL - 1, 1, 2).merge().setHorizontalAlignment('right').setFontSize(8)
    .setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Download this tab as PDF').setLinkUrl(pdfExportUrl(sheet, lastRow)).build());
}

function resetSheet(sheet, team, quarter) {
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  // A new sheet has 26 columns; the data area behind the charts runs to column DATA_COL + TREND_KEYS.length.
  const neededCols = DATA_COL + TREND_KEYS.length;
  if (sheet.getMaxColumns() < neededCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  if (sheet.getMaxRows() < PAGE_ROWS) sheet.insertRowsAfter(sheet.getMaxRows(), PAGE_ROWS - sheet.getMaxRows());
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(COLORS.accent);
  sheet.setColumnWidth(1, 24);
  sheet.setColumnWidths(2, LAST_COL - 1, 105);
  [2, 9].forEach(c => sheet.setColumnWidth(c, 170)); // first column of the left / right tables
  [7, 11].forEach(c => sheet.setColumnWidth(c, 150)); // metric label columns
  sheet.setColumnWidth(LAST_COL + 1, 40);
  sheet.getRange(1, 1, PAGE_ROWS, LAST_COL + 1).setBackground(COLORS.page).setFontFamily(FONT).setFontSize(10).setFontColor(COLORS.label);

  sheet.getRange('A1:A2').setValues([['Team'], ['Quarter']]).setFontSize(8).setFontColor(COLORS.muted);
  sheet.getRange('B1:B2').setValues([[team], [quarter]]).setBackground(COLORS.card).setFontWeight('bold').setFontColor(COLORS.ink)
    .setBorder(true, true, true, true, false, false, COLORS.accent, SpreadsheetApp.BorderStyle.SOLID);
  const now = new Date();
  const options = quarterOptions(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  if (options.indexOf(quarter) < 0) options.push(quarter);
  sheet.getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).build());
  sheet.getRange('D1').setValue(`Refreshed ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')} from Salesforce`)
    .setFontSize(8).setFontColor(COLORS.muted).setFontStyle('italic');
  sheet.getRange('D2').setValue('Change B1 (team) or B2 (quarter) and use QBR > Refresh this tab').setFontSize(8).setFontColor(COLORS.muted);
  sheet.setFrozenRows(3);
}

// Title row: team name in ink, quarter in the accent colour, ruled off by a thin ink line.
function writeBanner(sheet, row, title, subtitle) {
  sheet.setRowHeight(row, 56);
  const text = `${title}   ${subtitle}`;
  sheet.getRange(row, 2, 1, LAST_COL - 1).merge().setBackground(COLORS.page).setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text)
      .setTextStyle(0, title.length, SpreadsheetApp.newTextStyle().setBold(true).setFontSize(22).setForegroundColor(COLORS.ink).build())
      .setTextStyle(title.length, text.length, SpreadsheetApp.newTextStyle().setBold(false).setFontSize(14).setForegroundColor(COLORS.accent).build())
      .build())
    .setBorder(null, null, true, null, false, false, COLORS.ink, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// Six cards, two columns each: label / big value / QoQ delta. `addr` maps metric keys to the block cells the
// card values point at. Returns the row after the cards.
function writeKpiCards(sheet, row, current, previous, addr) {
  sheet.setRowHeight(row, 20);
  sheet.setRowHeight(row + 1, 36);
  sheet.setRowHeight(row + 2, 20);
  // First card is the hero (3 columns, dark), the rest 2 columns each (white), filling B:N.
  KPI_CARDS.forEach(([label, kind, key], i) => {
    const hero = i === 0;
    const col = hero ? 2 : 3 + i * 2;
    const span = hero ? 3 : 2;
    sheet.getRange(row, col, 3, span).setBackground(hero ? COLORS.ink : COLORS.card)
      .setBorder(true, true, true, true, false, false, hero ? COLORS.ink : COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, col, 1, span).merge().setValue(label.toUpperCase()).setFontSize(8).setFontColor(hero ? COLORS.onDark : COLORS.muted)
      .setHorizontalAlignment('left').setVerticalAlignment('bottom');
    if (KPI_NOTES[key]) sheet.getRange(row, col).setNote(KPI_NOTES[key]);
    const value = sheet.getRange(row + 1, col, 1, span).merge().setValue(addr && addr[key] ? `=${addr[key]}` : cellValue(kind, current[key])).setNumberFormat(FORMATS[kind])
      .setFontSize(hero ? 24 : 18).setFontWeight('bold').setFontColor(hero ? COLORS.card : COLORS.ink).setHorizontalAlignment('left').setVerticalAlignment('middle');
    if (RAG_KEYS.indexOf(key) >= 0) addRag(sheet, value, 'font');
    const delta = previous ? qoqText(kind, current[key], previous[key]) : '';
    let deltaColor = qoqColor(key, current[key], previous ? previous[key] : null);
    if (hero) deltaColor = deltaColor === COLORS.up ? COLORS.upOnDark : deltaColor === COLORS.muted ? COLORS.onDark : deltaColor;
    sheet.getRange(row + 2, col, 1, span).merge().setValue(delta ? `${delta} QoQ` : '').setFontSize(8)
      .setFontColor(deltaColor).setHorizontalAlignment('left').setVerticalAlignment('top');
  });
  return row + 3;
}

// Section header: accent tag with the title in column B, muted subtitle on a light band across the rest.
function writeSection(sheet, row, title, subtitle) {
  sheet.setRowHeight(row, 26);
  sheet.getRange(row, 2).setValue(title).setBackground(COLORS.accent).setFontColor(COLORS.card).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange(row, 3, 1, LAST_COL - 2).merge().setValue(subtitle).setBackground(COLORS.band).setFontColor(COLORS.muted).setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  return row + 1;
}

// Header + one line per spec row. `values` holds one metrics object per value column, `ctxs` the matching formula
// contexts (metric cells with a formula source are written as formulas, others as values). `shareOf` adds a
// "% of <shareOf>" column for money rows. Returns the row after the block.
function writeBlock(sheet, row, col, header, spec, values, previous, trend, shareOf, ctxs) {
  const width = header.length;
  (ctxs || []).forEach((ctx, j) => spec.forEach(([, , key], i) => { ctx.addr[key] = cellA1(row + 1 + i, col + 1 + j); }));
  const body = spec.map(([label, kind, key]) => {
    const line = [label].concat(values.map((m, j) => (ctxs && metricFormula(key, ctxs[j])) || cellValue(kind, m[key])));
    if (previous) line.push(qoqText(kind, values[0][key], previous[key]));
    if (trend) line.push(trend[key] ? `=IFERROR(SPARKLINE(${trend[key]},{"charttype","line";"linewidth",2;"color","${COLORS.accent}"}),"")` : '');
    if (shareOf != null) {
      const start = ctxs && ctxs[0].addr.startingArr;
      line.push(kind !== 'money' ? '' : start ? `=IF(${start}=0,"",${ctxs[0].addr[key]}/${start})` : shareOf ? values[0][key] / shareOf : '');
    }
    return line.concat(Array(width - line.length).fill(''));
  });
  const colors = spec.map(([, , key]) => {
    const line = Array(width).fill(COLORS.label);
    if (previous) line[1 + values.length] = qoqColor(key, values[0][key], previous[key]);
    return line;
  });

  sheet.getRange(row, col, 1, width).setValues([header]).setFontWeight('bold').setFontColor(COLORS.ink).setBackground(COLORS.card).setFontSize(9)
    .setHorizontalAlignment('right').setVerticalAlignment('middle')
    .setBorder(null, null, true, null, false, false, COLORS.ink, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, col).setHorizontalAlignment('left');
  sheet.setRowHeight(row, 24);
  spec.forEach(([, , key], i) => { if (KPI_NOTES[key]) sheet.getRange(row + 1 + i, col).setNote(KPI_NOTES[key]); });

  const range = sheet.getRange(row + 1, col, body.length, width);
  range.setValues(body).setBackground(COLORS.card).setFontColors(colors).setVerticalAlignment('middle');
  range.setBorder(null, null, null, null, false, true, COLORS.line, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, col, body.length + 1, width).setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row + 1, col, body.length, 1).setFontWeight('bold').setWrap(true);
  sheet.getRange(row + 1, col + 1, body.length, width - 1).setHorizontalAlignment('right').setWrap(true);
  spec.forEach(([, kind, key], i) => {
    sheet.getRange(row + 1 + i, col + 1, 1, values.length).setNumberFormat(FORMATS[kind]);
    if (kind === 'text') sheet.getRange(row + 1 + i, col + 1, 1, width - 1).merge().setHorizontalAlignment('left').setVerticalAlignment('top');
    if (kind === 'pct' && RAG_KEYS.indexOf(key) >= 0) addRag(sheet, sheet.getRange(row + 1 + i, col + 1, 1, values.length), 'background');
    else if (kind === 'pct') addPercentBar(sheet, row + 1 + i, col + 1, values.length);
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
  if (kind === 'mult') return `${arrow}${diff.toFixed(1)}x`;
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

// Red / amber / green thresholds (RAG) on attainment-style percentages, as cell background or font colour.
function addRag(sheet, range, mode) {
  const rules = sheet.getConditionalFormatRules();
  const paint = (rule, color) => (mode === 'font' ? rule.setFontColor(color) : rule.setBackground(color));
  const tone = { red: mode === 'font' ? COLORS.down : COLORS.ragRed, amber: mode === 'font' ? COLORS.amber : COLORS.ragAmber, green: mode === 'font' ? COLORS.up : COLORS.ragGreen };
  rules.push(paint(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(RAG.green), tone.green).setRanges([range]).build());
  rules.push(paint(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(RAG.amber, RAG.green - 0.0001), tone.amber).setRanges([range]).build());
  rules.push(paint(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(RAG.amber), tone.red).setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

const link = (text, url) => ({ text: text || '-', url: url || '' });
// The table lists the TOP_N largest wins; when there are more, the title says how much of the quarter's wins it shows.
function dealsWonTitle(quarter, shown, wonCount, wonArr) {
  const shownArr = sum(shown, 'deltaArr');
  return wonCount > shown.length
    ? `Top ${TOP_N} Deals Won ${quarter} (${shown.length} of ${wonCount} Closed Won shown: ${formatMoney(shownArr)} of ${formatMoney(wonArr)} Delta ARR)`
    : `Top ${TOP_N} Deals Won ${quarter} (${wonCount} Closed Won, ${formatMoney(wonArr)} Delta ARR)`;
}
// Opportunity names are long; the cell just says "Link" and points at the opportunity record.
const oppRow = (o, middle) => [link(o.account, o.accountUrl), link(o.url ? 'Link' : '-', o.url), middle(o) || '', o.closeDate || '', o.deltaArr, o.owner || ''];

// Side-by-side tables ({ title, col, columns, rows, width? }), DEFAULT_TABLE_WIDTH columns wide unless `width` is given: title row, header row, one row per
// item (or a single "-" row). Link cells become Salesforce hyperlinks. Returns the row after the tallest table.
function writeTables(sheet, row, tables) {
  const height = Math.max(1, ...tables.map(t => t.rows.length));
  ensureRows(sheet, row + 2 + height);
  tables.forEach(({ title, col, columns, rows, width }) => {
    const tableWidth = width || DEFAULT_TABLE_WIDTH;
    sheet.getRange(row, col, 1, tableWidth).merge().setValue(title).setFontWeight('bold').setFontColor(COLORS.ink).setFontSize(9)
      .setBackground(COLORS.card).setVerticalAlignment('middle')
      .setBorder(null, null, true, null, false, false, COLORS.ink, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(row, 24);
    sheet.getRange(row + 1, col, 1, tableWidth).setBackground(COLORS.card).setFontColor(COLORS.muted).setFontSize(8).setVerticalAlignment('middle');
    sheet.getRange(row + 1, col, 1, columns.length).setValues([columns.map(([header]) => header)])
      .setBorder(null, null, true, null, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row + 2, col, height, tableWidth).setBackground(COLORS.card).setVerticalAlignment('middle');
    if (rows.length) {
      const plain = rows.map(r => r.map(v => (v && typeof v === 'object' ? v.text : v == null ? '' : v)));
      sheet.getRange(row + 2, col, rows.length, columns.length).setValues(plain)
        .setBorder(null, null, null, null, false, true, COLORS.line, SpreadsheetApp.BorderStyle.SOLID);
    } else {
      sheet.getRange(row + 2, col).setValue('-').setFontColor(COLORS.muted);
    }
    columns.forEach(([, kind], i) => {
      const numeric = NUMERIC_KINDS.indexOf(kind) >= 0;
      sheet.getRange(row + 1, col + i, height + 1, 1).setHorizontalAlignment(numeric ? 'right' : 'left');
      if (!rows.length) return;
      const cells = sheet.getRange(row + 2, col + i, rows.length, 1);
      if (FORMATS[kind]) cells.setNumberFormat(FORMATS[kind]);
      if (kind === 'link') cells.setRichTextValues(rows.map(r => [linkValue(r[i])]));
      if (kind === 'text' || kind === 'link') cells.setWrap(true);
    });
    sheet.getRange(row, col, height + 2, tableWidth).setBorder(true, true, true, true, false, false, COLORS.border, SpreadsheetApp.BorderStyle.SOLID);
  });
  return row + 2 + height;
}

// Grows the grid when the dashboard runs past the PAGE_ROWS pre-sized in resetSheet (long owner / data-quality tables),
// styling the new rows like the rest of the page.
function ensureRows(sheet, lastRow) {
  const have = sheet.getMaxRows();
  if (lastRow <= have) return;
  sheet.insertRowsAfter(have, lastRow - have);
  sheet.getRange(have + 1, 1, lastRow - have, LAST_COL + 1).setBackground(COLORS.page).setFontFamily(FONT).setFontSize(10).setFontColor(COLORS.label);
}

function linkValue(cell) {
  const value = SpreadsheetApp.newRichTextValue().setText(String(cell.text));
  if (cell.url) value.setLinkUrl(cell.url);
  return value.build();
}

// Trend table at column T (header row 6, one row per quarter Q1-2026 -> selected).
// Returns { key: A1 range of that column's values } plus quarterCount. A metric with fewer than two numbers over the
// quarters (e.g. Logo Attainment % without a logo goal) gets no range, so no sparkline is drawn for it.
function writeTrendData(sheet, trendRows) {
  const body = trendRows.map(m => TREND_KEYS.map(key => (m[key] == null ? '' : m[key])));
  const numbers = i => body.filter(line => typeof line[i] === 'number').length;
  sheet.getRange(5, DATA_COL).setValue('Data behind the sparklines and charts (do not edit)').setFontColor(COLORS.muted).setFontSize(8);
  sheet.getRange(6, DATA_COL, 1, TREND_KEYS.length).setValues([TREND_KEYS]).setFontWeight('bold').setFontColor(COLORS.muted).setFontSize(8);
  sheet.getRange(7, DATA_COL, body.length, TREND_KEYS.length).setValues(body).setFontColor(COLORS.muted).setFontSize(8);
  const trend = { quarterCount: body.length };
  TREND_KEYS.forEach((key, i) => {
    if (numbers(i) >= 2) trend[key] = sheet.getRange(7, DATA_COL + i, body.length, 1).getA1Notation();
  });
  return trend;
}

// Numeric columns of a data block get a chart-friendly number format (Sheets uses it for labels and axes).
function writeDataBlock(sheet, row, rows, kind) {
  const range = sheet.getRange(row, DATA_COL, rows.length, rows[0].length);
  range.setValues(rows).setFontColor(COLORS.muted).setFontSize(8);
  if (kind && rows.length > 1 && rows[0].length > 1) {
    sheet.getRange(row + 1, DATA_COL + 1, rows.length - 1, rows[0].length - 1).setNumberFormat(CHART_FORMATS[kind]);
  }
  return range;
}

function writeCharts(sheet, row, view, trend) {
  const m = view.current;
  let dataRow = 8 + trend.quarterCount;

  // Waterfall as stacked columns: base drawn in the card colour + green up + red down.
  const afterAdded = m.startingArr + m.addedArr;
  const afterDowngrade = afterAdded + m.downgradeArr;
  const waterfall = writeDataBlock(sheet, dataRow, [
    ['Step', 'Base', 'Up', 'Down'],
    ['Starting ARR', 0, m.startingArr, 0],
    ['Added ARR', m.startingArr, m.addedArr, 0],
    ['Downgrade', afterDowngrade, 0, -m.downgradeArr],
    ['Full Churn', m.endingArr, 0, -m.fullChurnArr],
    ['Ending ARR', 0, m.endingArr, 0],
  ], 'money');
  dataRow += 7;
  const attainmentRows = [
    ['Metric', 'Attainment'],
    ['Revenue', m.attainment || 0],
    ['Logos', m.logoAttainmentPct || 0],
    ['Renewal rate', m.renewalRate || 0],
  ];
  const attainment = writeDataBlock(sheet, dataRow, attainmentRows, 'pct');
  // Axis runs to at least 100% and grows in 25% steps so over-achievement stays visible.
  const attainmentMax = Math.max(1, Math.ceil(Math.max(...attainmentRows.slice(1).map(r => r[1])) * 4) / 4);
  dataRow += 5;
  const lostRenewals = m.renewals - m.wonRenewals;
  const renewals = writeDataBlock(sheet, dataRow, [
    ['Outcome', 'Renewals'],
    ['Won', m.wonRenewals],
    ['Lost', lostRenewals],
  ], 'int');
  dataRow += 4;
  const forecast = writeDataBlock(sheet, dataRow, [['Quarter', 'Goal', 'Net Forecast', 'Pipeline']]
    .concat(view.future.map(f => [f.quarter, f.revenueGoal, f.netForecastArr, f.pipelineArr])), 'money');
  dataRow += view.future.length + 2;
  // Contiguous copies of the trend columns: embedded charts are only reliable on a single block.
  const netTrend = writeDataBlock(sheet, dataRow, [['Quarter', 'Goal', 'Net Added ARR', 'Churn ARR']]
    .concat(view.trend.map(t => [t.quarter, t.revenueGoal || 0, t.netAddedArr, t.churnArr])), 'money');
  dataRow += view.trend.length + 2;
  const arrTrend = writeDataBlock(sheet, dataRow, [['Quarter', 'Ending ARR']].concat(view.trend.map(t => [t.quarter, t.endingArr])), 'money');

  // Only options from the embedded-chart subset (title, legend, colors, axes, stacking, per-series data labels):
  // anything fancier can leave Sheets drawing the title with an empty plot.
  const labelled = (count, extra) => {
    const series = {};
    for (let i = 0; i < count; i++) series[i] = Object.assign({ dataLabel: 'value' }, extra && extra[i]);
    return series;
  };
  const chart = (type, title, seriesCount, extra) => sheet.newChart().setChartType(type).setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('title', title)
    .setOption('legend', { position: 'bottom' })
    .setOption('series', labelled(seriesCount, extra));
  const place = (builder, r, c) => sheet.insertChart(
    builder.setPosition(r, c, 0, 0).setOption('width', 600).setOption('height', 300).build()
  );
  const rowsPerChart = 15;
  ensureRows(sheet, row + 3 * rowsPerChart);

  place(chart(Charts.ChartType.COLUMN, `ARR bridge ${view.quarter}`, 3, { 0: { dataLabel: 'none', visibleInLegend: false } }).addRange(waterfall)
    .setOption('isStacked', true)
    .setOption('colors', [COLORS.card, COLORS.up, COLORS.down])
    .setOption('vAxis', { viewWindow: { min: 0 } }), row, 2);
  place(chart(Charts.ChartType.BAR, `Attainment vs goal ${view.quarter}`, 1).addRange(attainment)
    .setOption('hAxis', { viewWindow: { min: 0, max: attainmentMax } })
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLORS.accent]), row, 8);
  row += rowsPerChart;
  place(chart(Charts.ChartType.PIE, `Renewals ${view.quarter}: ${m.wonRenewals} won / ${lostRenewals} lost (${formatMoney(m.fullChurnArr)} churned)`, 0).addRange(renewals)
    .setOption('pieHole', 0.55)
    .setOption('pieSliceText', 'value')
    .setOption('colors', [COLORS.up, COLORS.down]), row, 2);
  place(chart(Charts.ChartType.COLUMN, 'Q+1 / Q+2 forecast vs goal', 3).addRange(forecast)
    .setOption('colors', [COLORS.muted, COLORS.accent, COLORS.accentSoft])
    .setOption('vAxis', { viewWindow: { min: 0 } }), row, 8);
  row += rowsPerChart;
  place(chart(Charts.ChartType.COLUMN, 'Net Added ARR vs goal by quarter', 3).addRange(netTrend)
    .setOption('colors', [COLORS.muted, COLORS.accent, COLORS.down]), row, 2);
  place(chart(Charts.ChartType.COLUMN, 'Ending ARR by quarter', 1).addRange(arrTrend)
    .setOption('colors', [COLORS.ink])
    .setOption('legend', { position: 'none' })
    .setOption('vAxis', { viewWindow: { min: 0 } }), row, 8);
  return row + rowsPerChart;
}
