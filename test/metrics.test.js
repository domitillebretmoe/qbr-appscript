// Runs the pure .gs files in a plain JS sandbox: `npm test`.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');

const ctx = vm.createContext({});
['Config.gs', 'Metrics.gs', 'Definitions.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));
const { quarterMetrics, forecastMetrics, accountMetrics, partnerMetrics, shiftQuarter, quarterOfDate, teamToken, quartersBetween,
  teamMatches, resolveTeam, assertSpecificTeam, definitionRows, quarterOptions } = ctx;

const { opp, dach } = require('./fixtures');

test('Q1-2026 DACH matches last quarter workbook', () => {
  const m = quarterMetrics(dach, 'Q1-2026', { revenue: 2000000, logos: 0 });
  assert.equal(m.netAddedArr, 127401.2);
  assert.equal(m.addedArr, 127401.2);
  assert.equal(m.churnArr, 0);
  assert.equal(m.attainment, 127401.2 / 2000000);
  assert.equal(m.lostPipelineCount, 11);
  assert.equal(m.lostPipelineArr, 2600401);
  assert.deepEqual(m.logosWon, []);
});

test('Q2-2026 DACH: -66,000 net added, one full churn, no downgrade', () => {
  const m = quarterMetrics(dach, 'Q2-2026', { revenue: 6700000, logos: 1 });
  assert.equal(m.netAddedArr, -66000);
  assert.equal(m.addedArr, 150000);
  assert.equal(m.fullChurnArr, -216000);
  assert.equal(m.fullChurnCount, 1);
  assert.equal(m.downgradeArr, 0);
  assert.equal(m.churnArr, -216000);
  assert.equal(m.churnCustomers, 1);
  assert.equal(m.renewals, 1);
  assert.equal(m.wonRenewals, 0);
  assert.equal(m.renewalRate, 0);
  assert.equal(m.logoAttainment, 0);
  assert.equal(m.lostPipelineCount, 1);
  assert.equal(m.lostPipelineArr, 2000);
  assert.deepEqual(m.topChurns, ['Deutsche Telekom (-$216K)']);
  assert.deepEqual(m.churnReasons, ['Budget']);
  assert.deepEqual(m.logosWon, ['Helaba']);
  assert.deepEqual(m.logosLost, ['Deutsche Telekom']);
  assert.equal(m.addedArr + m.downgradeArr + m.fullChurnArr, m.netAddedArr);
});

test('downgrade = Closed Won renewal with negative Delta ARR', () => {
  const rows = [opp('Q3-2026', 'Closed Won', 'Acme', 'Renewal', 'Renewal', -50000, 0), opp('Q3-2026', 'Closed Won', 'Fed', 'Renewal', 'Fed - Renewal', 10000, 0)];
  const m = quarterMetrics(rows, 'Q3-2026', {});
  assert.equal(m.downgradeArr, -50000);
  assert.equal(m.downgradeCount, 1);
  assert.equal(m.fullChurnCount, 0);
  assert.equal(m.netAddedArr, -40000);
  assert.equal(m.renewalRate, 1);
  assert.equal(m.churnCustomers, 0);
});

test('forecast quarter', () => {
  const f = forecastMetrics(dach, 'Q3-2026', 14040851.2, { revenue: 7500000, logos: 2 });
  assert.equal(f.forecastArr, 100000);
  assert.equal(f.forecastChurnArr, -30000);
  assert.equal(f.netForecastArr, 70000);
  assert.equal(f.pipelineArr, 400000);
  assert.equal(f.logoForecast, 0.5);
  assert.equal(f.forecastEndingArr, 14040851.2 + 70000);
});

test('account and partner metrics', () => {
  const accounts = [
    { id: 'Helaba', major: true, currentArr: 150000 }, { id: 'Serrala', major: false, currentArr: 120000 },
    { id: 'BMW Group', major: true, currentArr: 0 }, { id: 'Prospect', major: false, currentArr: 0, hasOpenOpp: true },
  ];
  const a = accountMetrics(accounts, dach, 'Q2-2026');
  assert.deepEqual(a, { activeCustomers: 2, majorCustomers: 1, enterpriseCustomers: 1, activatedProspects: 1, conversionRate: 0.5 });
  const p = partnerMetrics(dach, 'Q1-2026');
  assert.deepEqual(p, { partnerNetAddedArr: 0, partnerNewLogos: 0, partnerChurnArr: 0, partnerChurnCustomers: 0 });
  const partnerWin = opp('Q3-2026', 'Closed Won', 'Zalando', 'Land', 'Enterprise', 90000, 1, { oppGroup: 'Partnerships', oppTeam: 'Europe - DACH' });
  assert.deepEqual(partnerMetrics(dach.concat(partnerWin), 'Q3-2026'), { partnerNetAddedArr: 90000, partnerNewLogos: 1, partnerChurnArr: 0, partnerChurnCustomers: 0 });
});

test('forecast never double counts a renewal-record-type opp typed Expand', () => {
  const rows = [
    opp('Q4-2026', 'R2- Renewal Engagement', 'A', 'Expand', 'Renewal', 0, 0, { expectedDeltaArr: -20000 }),
    opp('Q4-2026', '1- Discovery', 'B', 'Expand', 'Enterprise', 0, 0, { expectedDeltaArr: 50000 }),
    opp('Q4-2026', 'R2- Renewal Engagement', 'C', 'Renewal', 'Fed - Renewal', 0, 0, { expectedDeltaArr: 5000 }),
  ];
  const f = forecastMetrics(rows, 'Q4-2026', 0, { revenue: 100000, logos: 0 });
  assert.equal(f.forecastArr, 55000);
  assert.equal(f.forecastChurnArr, -20000);
  assert.equal(f.netForecastArr, 35000);
});

test('team matching: token after the last " - ", leading segments must match too, never a bare region', () => {
  ['Europe - DACH', 'Europe Majors - DACH', 'Europe - DACH - Austria'].forEach(n => assert.ok(teamMatches(n, 'Europe - DACH'), n));
  assert.ok(!teamMatches('Europe - UKI', 'Europe - DACH'));
  assert.ok(!teamMatches('Europe', 'Europe - DACH'));
  assert.ok(teamMatches('Europe Majors - South', 'Europe - South'));
  assert.ok(!teamMatches('LATAM - South', 'Europe - South'));
  assert.ok(teamMatches('Japan Majors', 'Japan'));
  assert.ok(teamMatches('US Enterprise - East', 'US Enterprise'));
  assert.ok(teamMatches('US Majors - Media, Telco', 'US Majors - Media, Telco'));
  assert.ok(!teamMatches('', 'Japan'));

  assert.equal(resolveTeam('Europe Majors - DACH', 'Europe Majors - DACH - Switzerland', 'US Majors - Banking'), 'Europe Majors - DACH');
  assert.equal(resolveTeam('Europe', 'Europe - DACH - Germany', 'Europe'), 'Europe - DACH - Germany');
  assert.equal(resolveTeam('Europe', null, 'Europe - UKI'), 'Europe - UKI');
  assert.equal(resolveTeam('Europe', null, 'Other'), 'Europe');
  assert.equal(resolveTeam(null, null, null), '');

  assert.throws(() => assertSpecificTeam('Europe'), /region/);
  assert.doesNotThrow(() => assertSpecificTeam('Europe - DACH'));
  assert.doesNotThrow(() => assertSpecificTeam('Japan'));
});

test('quarter helpers', () => {
  assert.equal(shiftQuarter('Q4-2026', 1), 'Q1-2027');
  assert.equal(shiftQuarter('Q1-2026', -1), 'Q4-2025');
  assert.equal(quarterOfDate('2026-08-01'), 'Q3-2026');
  assert.equal(quarterOfDate('2027-01-15'), 'Q4-2026');
  assert.deepEqual(quartersBetween('Q1-2026', 'Q1-2027'), ['Q1-2026', 'Q2-2026', 'Q3-2026', 'Q4-2026', 'Q1-2027']);
  assert.equal(teamToken('Europe - DACH'), 'DACH');
  assert.equal(teamToken('US Majors - Media, Telco'), 'Media, Telco');
  assert.equal(teamToken('US Enterprise'), 'US Enterprise');
  assert.deepEqual(quarterOptions('2026-09-05'), ['Q1-2026', 'Q2-2026', 'Q3-2026', 'Q4-2026']);
  assert.equal(quarterOptions('2028-03-01').pop(), 'Q2-2028');
});

test('forecast churn count only counts renewals expected to shrink', () => {
  const rows = [
    opp('Q3-2026', 'R2- Renewal Engagement', 'Gone', 'Renewal', 'Renewal', 0, -1, { expectedDeltaArr: -50000 }),
    opp('Q3-2026', 'R2- Renewal Engagement', 'Flat', 'Renewal', 'Renewal', 0, -1, { expectedDeltaArr: 0 }),
    opp('Q3-2026', '1- Discovery', 'Land', 'Land', 'Enterprise', 10000, -1, { expectedDeltaArr: 5000 }),
  ];
  const f = forecastMetrics(rows, 'Q3-2026', 0, {});
  assert.equal(f.forecastChurnCount, 1);
  assert.equal(f.forecastChurnArr, -50000);
});

test('definitions tab covers every block of a team tab', () => {
  const rows = definitionRows();
  rows.forEach(row => assert.equal(row.length, 4, JSON.stringify(row)));
  const metrics = rows.map(r => r[1]).join(' | ');
  ['Net Added ARR', 'Downgrade', 'Full churn', 'Starting ARR', 'Ending ARR', 'Renewal rate', 'Logo attainment', 'Net Forecast',
    'Pipeline', 'Active customers', 'Partner'].forEach(name => assert.ok(metrics.indexOf(name) >= 0, name));
  assert.ok(rows.some(r => r[2].indexOf('Christian Lawless') >= 0));
  assert.ok(rows.some(r => r[2].indexOf('"Renewal" or "Fed - Renewal"') >= 0));
});
