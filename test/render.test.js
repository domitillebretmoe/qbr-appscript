// Renders the Europe - DACH sample through the real Render.gs against the in-memory sheet.
// `node test/render.test.js --dump out.json` also writes the recorded sheet for tools/preview-xlsx.py.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');
const { FakeSheet, globals } = require('./fake-sheets');
const { opp, dach } = require('./fixtures');

// Illustrative Q3/Q4-2026 activity layered on the real DACH fixture so the .xlsx preview shows a populated tab.
// SAMPLE DATA ONLY - the live tab is filled from Salesforce.
const previewOpps = dach.filter(o => o.quarter < 'Q3-2026').concat([
  opp('Q3-2026', 'Closed Won', 'Siemens', 'Land', 'Enterprise', 420000, 1, { major: true }),
  opp('Q3-2026', 'Closed Won', 'Lufthansa Group', 'Land', 'Enterprise', 250000, 1, { major: true }),
  opp('Q3-2026', 'Closed Won', 'Allianz', 'Expand', 'Enterprise', 180000, 0, { major: true }),
  opp('Q3-2026', 'Closed Won', 'Zalando', 'Land', 'Enterprise', 96000, 0, { oppGroup: 'Partnerships', oppTeam: 'Europe - DACH' }),
  opp('Q3-2026', 'Closed Won', 'CompuGroup', 'Renewal', 'Renewal', 0, 0),
  opp('Q3-2026', 'Closed Won', 'Helaba', 'Renewal', 'Renewal', 30000, 0, { major: true }),
  opp('Q3-2026', 'Closed Won', 'Serrala', 'Renewal', 'Renewal', -40000, 0, { lostReason: 'Seat reduction' }),
  opp('Q3-2026', 'Closed Lost', 'Bosch', 'Renewal', 'Renewal', -120000, -1, { major: true, lostReason: 'Competitor' }),
  opp('Q3-2026', 'Closed Lost', 'DKB', 'Land', 'Enterprise', 144000, 0),
  opp('Q3-2026', 'Closed Lost', 'Julius Baer', 'Land', 'Enterprise', 90000, 0),
  opp('Q4-2026', '3- Proposal', 'SAP SE', 'Land', 'Enterprise', 800000, 0.7, { major: true, expectedDeltaArr: 560000 }),
  opp('Q4-2026', '2- Qualification', 'BMW Group', 'Land', 'Enterprise', 800000, 0.3, { major: true, expectedDeltaArr: 240000 }),
  opp('Q4-2026', '3- Proposal', 'Allianz', 'Expand', 'Enterprise', 200000, 0, { major: true, expectedDeltaArr: 150000 }),
  opp('Q4-2026', 'R2- Renewal Engagement', 'Zalando', 'Renewal', 'Renewal', -25000, 0, { expectedDeltaArr: -25000 }),
  opp('Q4-2026', '1- Discovery', 'Roche', 'Land', 'Enterprise', 600000, 0.1, { major: true, expectedDeltaArr: 60000 }),
  opp('Q1-2027', '2- Qualification', 'Deutsche Bank', 'Land', 'Enterprise', 950000, 0.4, { major: true, expectedDeltaArr: 380000 }),
  opp('Q1-2027', 'R1- Renewal Prep', 'Siemens', 'Renewal', 'Renewal', 60000, 0, { expectedDeltaArr: 60000 }),
]);

const ctx = vm.createContext(globals());
['Config.gs', 'Metrics.gs', 'Render.gs', 'RawData.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));

// Same shape as buildView() in Main.gs, with the ledger replaced by the seeded DACH values.
function sampleView(quarter = 'Q2-2026', opps = dach) {
  const goals = {
    'Q1-2026': { revenue: 2000000, logos: 0 }, 'Q2-2026': { revenue: 6700000, logos: 1 }, 'Q3-2026': { revenue: 7500000, logos: 2 },
    'Q4-2026': { revenue: 8000000, logos: 2 }, 'Q1-2027': { revenue: 8500000, logos: 3 },
  };
  const goalFor = q => goals[q] || { revenue: 0, logos: 0 };
  const accounts = [
    { id: 'Helaba', major: true, currentArr: 150000 }, { id: 'CompuGroup', major: false, currentArr: 900000 },
    { id: 'Serrala', major: false, currentArr: 120000 }, { id: 'BMW Group', major: true, currentArr: 0, hasOpenOpp: true },
    { id: 'Prospect', major: false, currentArr: 0, hasOpenOpp: true }, { id: 'Siemens', major: true, currentArr: 420000 },
    { id: 'Lufthansa Group', major: true, currentArr: 250000 }, { id: 'Allianz', major: true, currentArr: 610000 },
    { id: 'Zalando', major: false, currentArr: 96000 }, { id: 'SAP SE', major: true, currentArr: 0, hasOpenOpp: true },
  ];
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
  return { team: 'Europe - DACH', quarter, trend, current, future: [future1, future2], opps };
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
  const view = sampleView(quarter, previewOpps);
  const sheet = render(view);
  sheet.getRange('D1').setValue('SAMPLE DATA - illustrative numbers showing the layout Render.gs draws; the live tab is refreshed from Salesforce');
  const raw = new FakeSheet('Raw Data');
  ctx.SpreadsheetApp.getActive = () => ({ getSheetByName: () => raw, insertSheet: () => raw });
  ctx.writeRawData(view.team, view.opps);
  fs.writeFileSync(process.argv[process.argv.indexOf('--dump') + 1], JSON.stringify([sheet, raw], null, 1));
}
