// Runs the pure .gs files in a plain JS sandbox: `npm test`.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');

const ctx = vm.createContext({});
['Config.gs', 'Metrics.gs', 'Definitions.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));
const { quarterMetrics, forecastMetrics, accountMetrics, partnerMetrics, shiftQuarter, quarterOfDate, teamToken, quartersBetween,
  teamMatches, resolveTeam, assertSpecificTeam, definitionRows, quarterOptions, rollupMembers, withArr, withPace, quarterElapsed,
  quarterStart, ownerMetrics, dataQualityIssues } = ctx;
const ROLLUP_TEAMS = vm.runInContext('ROLLUP_TEAMS', ctx);
const TEAMS = vm.runInContext('TEAMS', ctx);

const { opp, dach, dachAccounts } = require('./fixtures');

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
  const { topMajors, topEnterprise, ...counts } = a;
  assert.deepEqual(counts, { activeCustomers: 2, majorCustomers: 1, enterpriseCustomers: 1, activatedProspects: 1, conversionRate: 0.5, activeArr: 270000 });
  assert.deepEqual(topMajors.map(x => x.id), ['Helaba']);
  assert.deepEqual(topEnterprise.map(x => x.id), ['Serrala']);
  const p = partnerMetrics(dach, 'Q1-2026');
  assert.deepEqual(p, { partnerNetAddedArr: 0, partnerNewLogos: 0, partnerChurnArr: 0, partnerChurnCustomers: 0 });
  const partnerWin = opp('Q3-2026', 'Closed Won', 'Zalando', 'Land', 'Enterprise', 90000, 1, { oppGroup: 'Partnerships', oppTeam: 'Europe - DACH' });
  assert.deepEqual(partnerMetrics(dach.concat(partnerWin), 'Q3-2026'), { partnerNetAddedArr: 90000, partnerNewLogos: 1, partnerChurnArr: 0, partnerChurnCustomers: 0 });
});

test('Q3-2026 lists: churned = Closed Lost renewals, downgrades = Closed Won renewals < 0, lost pipeline = Closed Lost non-renewal', () => {
  const m = quarterMetrics(dach, 'Q3-2026', { revenue: 7500000, logos: 2 });
  const names = rows => rows.map(o => o.account);
  assert.deepEqual(names(m.lists.logosWon), ['Zalando']);
  assert.deepEqual(names(m.lists.churned), ['Bolt']);
  assert.deepEqual(names(m.lists.renewalsLost), ['Bolt']);
  assert.deepEqual(names(m.lists.downgrades), ['Julius Baer']);
  assert.deepEqual(names(m.lists.renewalsWon), ['CompuGroup', 'Julius Baer']);
  assert.deepEqual(names(m.lists.lostPipeline), ['Siemens']);
  assert.equal(m.netAddedArr, 96000 + 12000 - 30000 - 84000);
  assert.equal(m.renewalRate, 2 / 3);
  m.lists.churned.concat(m.lists.logosWon).forEach(o => {
    assert.match(o.url, /\/lightning\/r\/Opportunity\/006\d{15}\/view$/);
    assert.match(o.accountUrl, /\/lightning\/r\/Account\/001[A-Za-z0-9]{15}\/view$/);
  });
  const f = forecastMetrics(dach, 'Q4-2026', 0, {});
  assert.deepEqual(f.topDeals.map(o => [o.account, o.deltaArr]), [['BMW Group', 540000], ['DKB', 120000], ['CompuGroup', 0]]);
  assert.ok(f.topDeals.every(o => !o.isClosed));
});

test('Europe is a roll-up of the five Europe sub-teams, never a team of its own', () => {
  assert.deepEqual(ROLLUP_TEAMS.Europe, ['Europe - Nordics', 'Europe - Benelux', 'Europe - UKI', 'Europe - DACH', 'Europe - South']);
  assert.deepEqual(rollupMembers('Europe'), ROLLUP_TEAMS.Europe);
  assert.equal(rollupMembers('Europe - DACH'), null);
  assert.ok(!TEAMS.includes('Europe'));
  assert.throws(() => assertSpecificTeam('Europe'), /region/);
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

test('GRR / NRR are dollar-based on Starting ARR: churn + downgrades for GRR, plus existing-customer expansion for NRR', () => {
  const m = quarterMetrics(dach, 'Q3-2026', { revenue: 7500000, logos: 2 });
  // Won: Zalando Land 96k, CompuGroup renewal +12k, Julius Baer renewal -30k (downgrade); Bolt renewal lost -84k.
  assert.equal(m.netAddedArr, -6000);
  assert.equal(m.churnArr, -114000);
  assert.equal(m.newLogoArr, 96000);
  assert.equal(m.expansionArr, 12000);
  const S = 13328250;
  withArr(m, { startingArr: S, endingArr: S - 6000 });
  assert.equal(m.grr, (S - 114000) / S);
  assert.equal(m.nrr, (S - 114000 + 12000) / S);
  assert.equal(withArr(quarterMetrics([], 'Q3-2026', {}), { startingArr: 0, endingArr: 0 }).grr, null);
});

test('current-quarter pipeline coverage = open pipeline / remaining goal, empty once the goal is met', () => {
  const m = quarterMetrics(dach, 'Q3-2026', { revenue: 7500000, logos: 2 });
  assert.equal(m.openPipelineArr, 400000); // Helaba 400k + Serrala 0
  assert.equal(m.remainingGoal, 7506000);
  assert.equal(m.pipelineCoverage, 400000 / 7506000);
  const met = quarterMetrics(dach, 'Q3-2026', { revenue: -10000, logos: 0 });
  assert.equal(met.remainingGoal, 0);
  assert.equal(met.pipelineCoverage, null);
});

test('quarter elapsed and pace: attainment relative to the share of the quarter gone', () => {
  assert.equal(quarterStart('Q3-2026'), '2026-08-01');
  assert.equal(quarterStart('Q4-2026'), '2026-11-01');
  assert.equal(quarterStart('Q1-2027'), '2027-02-01');
  assert.equal(quarterElapsed('Q3-2026', '2026-08-01'), 0);
  assert.equal(quarterElapsed('Q3-2026', '2026-09-15'), 45 / 92);
  assert.equal(quarterElapsed('Q2-2026', '2026-09-15'), 1);
  assert.equal(quarterElapsed('Q4-2026', '2026-09-15'), 0);
  const m = withPace({ quarter: 'Q3-2026', attainment: 0.25 }, '2026-09-15');
  assert.equal(m.quarterElapsedPct, 45 / 92);
  assert.equal(m.pace, 0.25 / (45 / 92));
  assert.equal(withPace({ quarter: 'Q3-2026', attainment: null }, '2026-09-15').pace, null);
});

test('future quarters list open renewals due with the account ARR at stake, each account counted once', () => {
  const q4 = forecastMetrics(dach, 'Q4-2026', 0, {}, dachAccounts);
  assert.equal(q4.renewalsDueCount, 1);
  assert.equal(q4.renewalArrDue, 410000);
  assert.deepEqual(q4.renewalsDue.map(o => [o.account, o.accountArr]), [['CompuGroup', 410000]]);
  const twice = dach.concat([opp('Q4-2026', 'R1- Renewal Planning', 'CompuGroup', 'Renewal', 'Fed - Renewal', 0, 0)]);
  assert.equal(forecastMetrics(twice, 'Q4-2026', 0, {}, dachAccounts).renewalsDueCount, 2);
  assert.equal(forecastMetrics(twice, 'Q4-2026', 0, {}, dachAccounts).renewalArrDue, 410000);
  assert.equal(forecastMetrics(dach, 'Q4-2026', 0, {}).renewalArrDue, 0);
});

test('owner table: net added, won count, churn and pipeline per opportunity owner, sorted by net added', () => {
  const rows = [
    opp('Q3-2026', 'Closed Won', 'A', 'Land', 'Enterprise', 100000, 1, { owner: 'Ben' }),
    opp('Q3-2026', 'Closed Won', 'B', 'Renewal', 'Renewal', -20000, 0, { owner: 'Ben' }),
    opp('Q3-2026', 'Closed Lost', 'C', 'Renewal', 'Renewal', -50000, -1, { owner: 'Anna' }),
    opp('Q3-2026', '3- Proposal', 'D', 'Expand', 'Enterprise', 70000, 0, { owner: 'Anna' }),
    opp('Q4-2026', '1- Discovery', 'E', 'Land', 'Enterprise', 30000, 1, { owner: 'Anna' }),
    opp('Q4-2026', '1- Discovery', 'F', 'Land', 'Enterprise', 5000, 1, { owner: 'Chris' }),
  ];
  const owners = ownerMetrics(rows, 'Q3-2026', 'Q4-2026');
  assert.deepEqual(owners.map(o => o.owner), ['Ben', 'Chris', 'Anna']);
  assert.deepEqual(owners[0], { owner: 'Ben', netAddedArr: 80000, wonCount: 2, churnArr: -20000, openPipelineArr: 0, nextPipelineArr: 0 });
  assert.deepEqual(owners[2], { owner: 'Anna', netAddedArr: -50000, wonCount: 0, churnArr: -50000, openPipelineArr: 70000, nextPipelineArr: 30000 });
  assert.equal(owners[1].nextPipelineArr, 5000);
});

test('data quality: stale open opps, $0 Closed Won, missing Expected ARR, region-only accounts', () => {
  const rows = [
    opp('Q3-2026', '3- Proposal', 'Stale', 'Land', 'Enterprise', 10000, 1, { closeDate: '2026-09-01' }),
    opp('Q3-2026', '3- Proposal', 'Fresh', 'Land', 'Enterprise', 10000, 1, { closeDate: '2026-10-01' }),
    opp('Q3-2026', 'Closed Won', 'Zero', 'Expand', 'Enterprise', 0, 0),
    opp('Q3-2026', 'Closed Won', 'Services', 'One Time', 'Enterprise', 0, 0),
    opp('Q3-2026', 'Closed Won', 'Flat renewal', 'Renewal', 'Renewal', 0, 0),
    opp('Q4-2026', '1- Discovery', 'NoExpected', 'Land', 'Enterprise', 40000, 1, { closeDate: '2026-12-01', expectedDeltaArrMissing: true }),
    opp('Q2-2026', '1- Discovery', 'Old quarter', 'Land', 'Enterprise', 40000, 1, { closeDate: '2026-05-01' }),
  ];
  const flutter = { id: '001F', name: 'Flutter Entertainment', url: 'u', team: 'Europe', currentArr: 50000, owner: 'Sam' };
  const issues = dataQualityIssues(rows, [flutter], ['Q3-2026', 'Q4-2026', 'Q1-2027'], '2026-09-15');
  assert.deepEqual(issues.map(i => [i.issue, i.opp ? i.opp.account : i.account.name]), [
    ['Open with close date in the past', 'Stale'],
    ['Closed Won with $0 Delta ARR', 'Zero'],
    ['Open without Expected Delta ARR', 'NoExpected'],
    ['Account team is a region only', 'Flutter Entertainment'],
  ]);
  assert.equal(issues[3].detail, 'Team = Europe, no sub-team');
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
