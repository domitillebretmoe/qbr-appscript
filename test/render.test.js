// Renders the Europe - DACH sample through the real Render.gs against the in-memory sheet.
// `node test/render.test.js --dump out.json` also writes the recorded sheet for tools/preview-xlsx.py.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');
const { FakeSheet, globals } = require('./fake-sheets');
const { opp, dach, account, dachAccounts, SF, accountUrl } = require('./fixtures');

// Illustrative Q3/Q4-2026 activity layered on the real DACH fixture so the .xlsx preview shows a populated tab.
// SAMPLE DATA ONLY - the live tab is filled from Salesforce.
const previewOpps = dach.concat([
  opp('Q3-2026', 'Closed Won', 'Siemens', 'Land', 'Enterprise', 420000, 1, { major: true, closeDate: '2025-08-29' }),
  opp('Q3-2026', 'Closed Won', 'Lufthansa Group', 'Land', 'Enterprise', 250000, 1, { major: true, closeDate: '2025-09-18', owner: 'Jonas Weber' }),
  opp('Q3-2026', 'Closed Won', 'Allianz', 'Expand', 'Enterprise', 180000, 0, { major: true, closeDate: '2025-10-02' }),
  opp('Q3-2026', 'Closed Won', 'Helaba', 'Renewal', 'Renewal', 30000, 0, { major: true, closeDate: '2025-10-10' }),
  opp('Q3-2026', 'Closed Lost', 'Bosch', 'Renewal', 'Renewal', -120000, -1, { major: true, lostReason: 'Competitor', closeDate: '2025-10-15' }),
  opp('Q3-2026', 'Closed Lost', 'DKB', 'Land', 'Enterprise', 144000, 0, { closeDate: '2025-10-22' }),
  opp('Q4-2026', '3- Proposal', 'SAP SE', 'Land', 'Enterprise', 800000, 0.7, { major: true, expectedDeltaArr: 560000, closeDate: '2025-12-19' }),
  opp('Q4-2026', '3- Proposal', 'Allianz', 'Expand', 'Enterprise', 200000, 0, { major: true, expectedDeltaArr: 150000, closeDate: '2026-01-15', owner: 'Jonas Weber' }),
  opp('Q4-2026', 'R2- Renewal Engagement', 'Zalando', 'Renewal', 'Renewal', -25000, 0, { expectedDeltaArr: -25000, closeDate: '2025-11-28' }),
  opp('Q1-2027', '2- Qualification', 'Deutsche Bank', 'Land', 'Enterprise', 950000, 0.4, { major: true, expectedDeltaArr: 380000, closeDate: '2026-03-12' }),
  opp('Q1-2027', 'R1- Renewal Prep', 'Siemens', 'Renewal', 'Renewal', 60000, 0, { expectedDeltaArr: 60000, closeDate: '2026-04-20' }),
]);
const previewAccounts = dachAccounts.concat([
  account('Siemens', 420000, true), account('Lufthansa Group', 250000, true), account('Allianz', 610000, true, { hasOpenOpp: true }),
  account('Bosch', 0, true), account('Deutsche Bank', 0, true, { hasOpenOpp: true }),
]);

const ctx = vm.createContext(globals());
['Config.gs', 'Metrics.gs', 'Render.gs', 'RawData.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));

// Same shape as buildView() in Main.gs, with the ledger replaced by the seeded DACH values.
function sampleView(quarter = 'Q2-2026', opps = dach, accounts = dachAccounts) {
  const goals = {
    'Q1-2026': { revenue: 2000000, logos: 0 }, 'Q2-2026': { revenue: 6700000, logos: 1 }, 'Q3-2026': { revenue: 7500000, logos: 2 },
    'Q4-2026': { revenue: 8000000, logos: 2 }, 'Q1-2027': { revenue: 8500000, logos: 3 },
  };
  const goalFor = q => goals[q] || { revenue: 0, logos: 0 };
  const seed = vm.runInContext("TEAM_SEEDS.filter(s => s[0] === 'Europe - DACH')[0]", ctx);
  const ledger = { 'Q1-2026': { startingArr: seed[1], endingArr: seed[2] }, 'Q2-2026': { startingArr: seed[2], endingArr: seed[3] } };
  const trend = ctx.quartersBetween(vm.runInContext('FIRST_QUARTER', ctx), quarter).map(q => {
    const m = Object.assign(ctx.quarterMetrics(opps, q, goalFor(q)), ctx.accountMetrics(accounts, opps, q), ctx.partnerMetrics(opps, q));
    const prevEnding = ledger[ctx.shiftQuarter(q, -1)] ? ledger[ctx.shiftQuarter(q, -1)].endingArr : null;
    ledger[q] = ledger[q] || { startingArr: prevEnding, endingArr: prevEnding + m.netAddedArr };
    return ctx.withArr(m, ledger[q]);
  });
  const current = trend[trend.length - 1];
  const next1 = ctx.shiftQuarter(quarter, 1);
  const next2 = ctx.shiftQuarter(quarter, 2);
  const future1 = ctx.forecastMetrics(opps, next1, current.endingArr, goalFor(next1));
  const future2 = ctx.forecastMetrics(opps, next2, future1.forecastEndingArr, goalFor(next2));
  return { team: 'Europe - DACH', members: ['Europe - DACH'], quarter, trend, current, future: [future1, future2], opps };
}

function render(view) {
  const sheet = new FakeSheet(view.team);
  ctx.renderTeamTab(sheet, view);
  return sheet;
}

test('team tab renders every block, six KPI cards and six charts', () => {
  const sheet = render(sampleView());
  const values = Object.values(sheet.cells).map(c => c.value);
  assert.equal(sheet.cell(1, 2).value, 'Europe - DACH');
  assert.equal(sheet.cell(2, 2).value, 'Q2-2026');
  assert.equal(sheet.cell(3, 2).value, 'Europe - DACH   Q2-2026 QBR');
  ['Net Added ARR', 'Starting ARR', 'Forecast Ending ARR', '# Active Customers', 'Top 3 Churns'].forEach(label => assert.ok(values.includes(label), label));
  ['PREVIOUS QUARTER', 'FUTURE QUARTER(S)', 'PARTNER CONTRIBUTION', 'CHARTS'].forEach(label => assert.ok(values.includes(label), label));
  assert.equal(sheet.charts.length, 6);
  assert.equal(sheet.frozenRows, 3);
  assert.ok(sheet.getMaxColumns() >= 20 + vm.runInContext('TREND_KEYS', ctx).length, 'sheet widened for the data area');
  // KPI cards: label row 5, value row 6, QoQ row 7.
  assert.equal(sheet.cell(5, 2).value, 'NET ADDED ARR');
  assert.equal(sheet.cell(6, 2).value, -66000);
  assert.match(String(sheet.cell(7, 2).value), /QoQ$/);
  // Values stay exact dollars and the ARR bridge ties.
  const at = label => Object.values(sheet.cells).find(c => c.value === label && c.col === 7);
  const val = label => sheet.cell(at(label).row, 8).value;
  assert.equal(val('Ending ARR'), 14040851.2);
  assert.equal(val('Starting ARR') + val('Added ARR') + val('Downgrade $') + val('Full Churn $'), val('Ending ARR'));
});

const cellsWhere = (sheet, pred) => Object.values(sheet.cells).filter(pred);
const titleCell = (sheet, prefix) => cellsWhere(sheet, c => typeof c.value === 'string' && c.value.startsWith(prefix))[0];
// Rows of the table titled `prefix`: [[cell, ...], ...] until the first empty first column.
function tableRows(sheet, prefix) {
  const title = titleCell(sheet, prefix);
  assert.ok(title, `table "${prefix}" present`);
  const headers = [];
  for (let c = title.col; sheet.cell(title.row + 1, c).value; c++) headers.push(sheet.cell(title.row + 1, c).value);
  const rows = [];
  for (let r = title.row + 2; sheet.cell(r, title.col).value !== undefined && sheet.cell(r, title.col).value !== ''; r++) {
    rows.push(headers.map((h, i) => sheet.cell(r, title.col + i)));
  }
  return { title: title.value, headers, rows };
}

test('linked tables: churn vs lost pipeline, renewals, top customers, top deals, every name links to Salesforce', () => {
  const sheet = render(sampleView('Q3-2026'));

  const logos = tableRows(sheet, 'Logos Won');
  assert.equal(logos.title, 'Logos Won (1)');
  assert.deepEqual(logos.headers, ['Account', 'Opportunity', 'Type', 'Close date', 'Delta ARR', 'Owner']);
  assert.equal(logos.rows[0][0].value, 'Zalando');
  assert.equal(logos.rows[0][0].link, accountUrl('Zalando'));
  assert.match(logos.rows[0][0].link, /^https:\/\/codeium\.lightning\.force\.com\/lightning\/r\/Account\/001[A-Za-z0-9]{15}\/view$/);
  assert.match(logos.rows[0][1].link, /\/lightning\/r\/Opportunity\/006\d{15}\/view$/);
  assert.equal(logos.rows[0][3].value, '2025-08-20');
  assert.equal(logos.rows[0][4].value, 96000);
  assert.equal(logos.rows[0][4].numberFormat, '$#,##0;-$#,##0');

  // Lost pipeline (Closed Lost non-renewal) is not churn (Closed Lost renewal).
  const lost = tableRows(sheet, 'Lost Pipeline');
  assert.equal(lost.title, 'Lost Pipeline (1)');
  assert.equal(lost.rows[0][0].value, 'Siemens');
  const churned = tableRows(sheet, 'Churned Customers');
  assert.equal(churned.title, 'Churned Customers (1)');
  assert.equal(churned.rows[0][0].value, 'Bolt');
  assert.equal(churned.rows[0][4].value, -84000);
  assert.ok(!cellsWhere(sheet, c => c.value === 'Logos Lost').length, 'no ambiguous "Logos Lost" label');

  const downgrades = tableRows(sheet, 'Downgrade Customers');
  assert.equal(downgrades.title, 'Downgrade Customers (1)');
  assert.equal(downgrades.rows[0][0].value, 'Julius Baer');
  assert.equal(downgrades.rows[0][4].value, -30000);

  const renewalsWon = tableRows(sheet, 'Renewals Won');
  assert.equal(renewalsWon.title, 'Renewals Won (2)');
  assert.deepEqual(renewalsWon.rows.map(r => r[0].value), ['CompuGroup', 'Julius Baer']);
  assert.equal(tableRows(sheet, 'Renewals Lost').rows[0][0].value, 'Bolt');
  // Renewal rate on the tab = won / (won + lost) from those same two tables.
  assert.equal(sheet.cell(5, 9).value, 'RENEWAL RATE');
  assert.equal(sheet.cell(6, 9).value, 2 / 3);

  const majors = tableRows(sheet, 'Top 10 Major Customers');
  assert.equal(majors.title, 'Top 10 Major Customers (3 of 3 active)');
  assert.deepEqual(majors.rows.map(r => r[1].value), ['Deutsche Telekom', 'SAP SE', 'Helaba']);
  assert.equal(majors.rows[0][1].link, accountUrl('Deutsche Telekom'));
  assert.deepEqual(majors.rows.map(r => r[3].value), [1200000, 850000, 640000]);
  const enterprise = tableRows(sheet, 'Top 10 Enterprise Customers');
  assert.deepEqual(enterprise.rows.map(r => r[1].value), ['CompuGroup', 'Zalando', 'Serrala', 'Julius Baer']);
  assert.equal(enterprise.rows[0][4].value, 410000 / (1200000 + 850000 + 640000 + 410000 + 96000 + 90000 + 60000));

  const q1 = tableRows(sheet, 'Top 10 Deals Q+1 Q4-2026');
  assert.deepEqual(q1.headers, ['Account', 'Opportunity', 'Stage', 'Close date', 'Delta ARR', 'Owner']);
  assert.deepEqual(q1.rows.map(r => [r[0].value, r[4].value, r[3].value]), [['BMW Group', 540000, '2026-01-20'], ['DKB', 120000, '2025-11-30'], ['CompuGroup', 0, '2025-12-15']]);
  assert.ok(q1.rows.every(r => r[1].link.indexOf('/lightning/r/Opportunity/') > 0));
  const q2 = tableRows(sheet, 'Top 10 Deals Q+2 Q1-2027');
  assert.deepEqual(q2.rows.map(r => r[0].value), ['Roche', 'Zalando']);

  // Empty table still renders its title and a placeholder.
  const empty = render(sampleView('Q1-2026'));
  assert.equal(titleCell(empty, 'Churned Customers').value, 'Churned Customers (0)');
  assert.equal(empty.cell(titleCell(empty, 'Churned Customers').row + 2, 2).value, '-');
});

test('every chart reads one contiguous block whose header row and data rows are populated', () => {
  const sheet = render(sampleView('Q3-2026'));
  const titles = sheet.charts.map(c => c.options.title);
  assert.deepEqual(titles, ['ARR bridge Q3-2026', 'Attainment vs goal Q3-2026', 'Renewals Q3-2026: won vs lost', 'Q+1 / Q+2 forecast vs goal',
    'Net Added ARR vs goal by quarter', 'Ending ARR by quarter']);
  sheet.charts.forEach(chart => {
    assert.equal(chart.ranges.length, 1, `${chart.options.title}: single range`);
    const { a1, values } = chart.ranges[0];
    assert.match(a1, /^[A-Z]+\d+:[A-Z]+\d+$/);
    assert.ok(values.length >= 2, `${chart.options.title}: header + data`);
    values[0].forEach(h => assert.equal(typeof h, 'string'));
    values.slice(1).forEach(row => {
      assert.equal(typeof row[0], 'string', `${chart.options.title}: domain label`);
      row.slice(1).forEach(v => assert.equal(typeof v, 'number', `${chart.options.title}: numeric series in ${a1}`));
    });
    // Only the embedded-chart option subset.
    Object.keys(chart.options).forEach(k => assert.ok(['title', 'legend', 'colors', 'isStacked', 'vAxis', 'hAxis', 'pieHole', 'pointSize',
      'useFirstColumnAsDomain', 'width', 'height'].includes(k), `${chart.options.title}: option ${k}`));
  });
  const trendChart = sheet.charts[4];
  assert.deepEqual(trendChart.ranges[0].values[0], ['Quarter', 'Goal', 'Net Added ARR', 'Churn ARR']);
  assert.deepEqual(trendChart.ranges[0].values.slice(1).map(r => r[0]), ['Q1-2026', 'Q2-2026', 'Q3-2026']);
  const arr = sheet.charts[5].ranges[0].values;
  assert.equal(sheet.cell(5, 7).value, 'ENDING ARR');
  assert.equal(arr[arr.length - 1][1], sheet.cell(6, 7).value, 'Ending ARR chart ends on the KPI card value');
});

test('Europe roll-up banner names its members', () => {
  const view = Object.assign(sampleView('Q3-2026'), { team: 'Europe', members: ['Europe - Nordics', 'Europe - Benelux', 'Europe - UKI', 'Europe - DACH', 'Europe - South'] });
  const sheet = render(view);
  assert.equal(sheet.cell(1, 2).value, 'Europe');
  assert.equal(sheet.cell(3, 2).value, 'Europe   Q3-2026 QBR   (roll-up of Nordics, Benelux, UKI, DACH, South)');
});

const RAW_HEADER = vm.runInContext('RAW_HEADER', ctx);

test('Raw Data tab lists every opportunity of the tab with its bucket and replaces only that tab\'s rows', () => {
  const raw = new FakeSheet('Raw Data');
  ctx.SpreadsheetApp.getActive = () => ({ getSheetByName: () => raw, insertSheet: () => raw });
  raw.getRange(1, 1, 2, RAW_HEADER.length).setValues([RAW_HEADER, ['Other team', 'Other team', 'Q2-2026', 'Acme'].concat(Array(RAW_HEADER.length - 4).fill(''))]);
  ctx.writeRawData('Europe - DACH', dach);
  const rows = raw.getRange(2, 1, raw.getLastRow() - 1, RAW_HEADER.length).getValues();
  assert.equal(rows.length, dach.length + 1);
  assert.equal(rows.filter(r => r[0] === 'Other team').length, 1);
  const bucket = (quarter, account) => rows.find(r => r[2] === quarter && r[3] === account)[8];
  assert.equal(bucket('Q2-2026', 'Deutsche Telekom'), 'Full churn');
  assert.equal(bucket('Q1-2026', 'Deutsche Telekom'), 'Lost pipeline');
  assert.equal(bucket('Q2-2026', 'Helaba'), 'Closed Won - Land');
  assert.equal(bucket('Q3-2026', 'Julius Baer'), 'Won renewal - downgrade');
  const telekom = rows.find(r => r[2] === 'Q2-2026' && r[3] === 'Deutsche Telekom');
  assert.match(telekom[RAW_HEADER.indexOf('Opportunity URL')], /\/lightning\/r\/Opportunity\/006\d{15}\/view$/);
  assert.equal(telekom[RAW_HEADER.indexOf('Account URL')], accountUrl('Deutsche Telekom'));
  assert.ok(telekom[RAW_HEADER.indexOf('Account URL')].startsWith(SF));
  // Net Added ARR ties to the raw rows: Closed Won Delta ARR + full-churn Delta ARR.
  const q2 = rows.filter(r => r[0] === 'Europe - DACH' && r[2] === 'Q2-2026');
  const netAdded = q2.filter(r => r[5] === 'Closed Won' || r[8] === 'Full churn').reduce((t, r) => t + r[10], 0);
  assert.equal(netAdded, -66000);
  assert.equal(raw.frozenRows, 1);
  assert.ok(raw.filter);
  // Re-running for the same tab does not duplicate.
  ctx.writeRawData('Europe - DACH', dach);
  assert.equal(raw.getLastRow() - 1, dach.length + 1);
});

if (process.argv.includes('--dump')) {
  const quarter = process.argv[process.argv.indexOf('--dump') + 2] || 'Q3-2026';
  const view = sampleView(quarter, previewOpps, previewAccounts);
  const sheet = render(view);
  sheet.getRange('D1').setValue('SAMPLE DATA - illustrative numbers showing the layout Render.gs draws; the live tab is refreshed from Salesforce');
  const raw = new FakeSheet('Raw Data');
  ctx.SpreadsheetApp.getActive = () => ({ getSheetByName: () => raw, insertSheet: () => raw });
  ctx.writeRawData(view.team, view.opps);
  fs.writeFileSync(process.argv[process.argv.indexOf('--dump') + 1], JSON.stringify([sheet, raw], null, 1));
}
