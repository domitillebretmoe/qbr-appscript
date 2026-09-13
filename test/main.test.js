// buildView() / defaultQuarter() from Main.gs with the Salesforce and ledger calls stubbed.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');
const { opp, account } = require('./fixtures');

function mainContext(stubs) {
  const ctx = vm.createContext(Object.assign({
    Utilities: { formatDate: d => d.toISOString().slice(0, 10) },
    Session: { getScriptTimeZone: () => 'UTC' },
    fetchRegionOnlyAccounts: () => [],
    fetchRepPerformance: () => [],
    fetchOwnerTeams: () => ({}),
  }, stubs));
  ['Config.gs', 'Metrics.gs', 'Main.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));
  return ctx;
}

const byTeam = {
  'Europe - DACH': {
    opps: [opp('Q3-2026', 'Closed Won', 'Zalando', 'Land', 'Enterprise', 96000, 1, { team: 'Europe - DACH' }),
      opp('Q3-2026', 'Closed Lost', 'Bolt', 'Renewal', 'Renewal', -84000, -1, { team: 'Europe - DACH' }),
      opp('Q4-2026', '3- Proposal', 'BMW Group', 'Land', 'Enterprise', 540000, 1, { team: 'Europe - DACH', expectedDeltaArr: 270000 })],
    accounts: [account('Zalando', 96000, false), account('Deutsche Telekom', 1200000, true)],
    goals: { 'Q3-2026': { revenue: 7500000, logos: 2 }, 'Q4-2026': { revenue: 8000000, logos: 2 } },
    ledger: { 'Q3-2026': { startingArr: 14040851.2, endingArr: 14052851.2 } },
  },
  'Europe - UKI': {
    opps: [opp('Q3-2026', 'Closed Won', 'Tesco', 'Expand', 'Enterprise', 200000, 0, { team: 'Europe - UKI' }),
      opp('Q4-2026', '1- Discovery', 'BP', 'Land', 'Enterprise', 900000, 1, { team: 'Europe - UKI', expectedDeltaArr: 90000 })],
    accounts: [account('Tesco', 700000, true, { team: 'Europe - UKI' })],
    goals: { 'Q3-2026': { revenue: 9000000, logos: 3 } },
    ledger: { 'Q3-2026': { startingArr: 31000000, endingArr: 31200000 } },
  },
};
const empty = { opps: [], accounts: [], goals: {}, ledger: { 'Q3-2026': { startingArr: 0, endingArr: 0 } } };
const data = team => byTeam[team] || empty;

test('Europe roll-up sums its member teams, keeps opportunities attributed to the sub-team, ledger = sum of member rows', () => {
  const fetched = [];
  const ledgerCalls = [];
  let ownerNames = null;
  const ctx = mainContext({
    fetchOwnerTeams: names => { ownerNames = names; return {}; },
    fetchOpportunities: team => { fetched.push(team); return data(team).opps; },
    fetchAccounts: team => data(team).accounts,
    fetchGoals: team => data(team).goals,
    rollLedger: (team, quarter, netAdded) => {
      ledgerCalls.push([team, quarter, netAdded]);
      const quarters = ctx.quartersBetween('Q1-2026', quarter);
      const out = {};
      quarters.forEach(q => { out[q] = data(team).ledger[q] || { startingArr: 0, endingArr: 0 }; });
      return out;
    },
  });
  const view = ctx.buildView('Europe', 'Q3-2026');

  assert.deepEqual(fetched, ['Europe - Nordics', 'Europe - Benelux', 'Europe - UKI', 'Europe - DACH', 'Europe - South']);
  assert.ok(!fetched.includes('Europe'));
  assert.deepEqual(view.members, fetched);
  assert.deepEqual(ownerNames, ['Anna Berger']);
  assert.deepEqual(view.opps.map(o => o.team).sort(), ['Europe - DACH', 'Europe - DACH', 'Europe - DACH', 'Europe - UKI', 'Europe - UKI']);

  const q3 = view.current;
  assert.equal(q3.netAddedArr, 96000 - 84000 + 200000);
  assert.equal(q3.revenueGoal, 7500000 + 9000000);
  assert.equal(q3.logoGoal, 5);
  assert.equal(q3.activeCustomers, 3);
  assert.deepEqual(q3.topMajors.map(a => a.name), ['Deutsche Telekom', 'Tesco']);
  assert.equal(q3.startingArr, 14040851.2 + 31000000);
  assert.equal(q3.endingArr, 14052851.2 + 31200000);
  // Each member's ledger is rolled forward with that member's own Net Added ARR.
  const dachCall = ledgerCalls.find(c => c[0] === 'Europe - DACH');
  assert.equal(dachCall[2]['Q3-2026'], 12000);
  assert.equal(ledgerCalls.find(c => c[0] === 'Europe - UKI')[2]['Q3-2026'], 200000);
  assert.ok(!ledgerCalls.some(c => c[0] === 'Europe'));

  const [q4] = view.future;
  assert.equal(q4.startingArr, q3.endingArr);
  assert.equal(q4.revenueGoal, 8000000);
  assert.deepEqual(q4.topDeals.map(o => [o.account, o.team]), [['BP', 'Europe - UKI'], ['BMW Group', 'Europe - DACH']]);
});

test('a plain team is fetched once and never goes through the roll-up path', () => {
  const fetched = [];
  const ctx = mainContext({
    fetchOpportunities: team => { fetched.push(team); return data(team).opps; },
    fetchAccounts: team => data(team).accounts,
    fetchGoals: team => data(team).goals,
    rollLedger: (team, quarter) => { const out = {}; ctx.quartersBetween('Q1-2026', quarter).forEach(q => { out[q] = data(team).ledger[q] || { startingArr: 0, endingArr: 0 }; }); return out; },
  });
  const view = ctx.buildView('Europe - DACH', 'Q3-2026');
  assert.deepEqual(fetched, ['Europe - DACH']);
  assert.deepEqual(view.members, ['Europe - DACH']);
  assert.equal(view.current.endingArr, 14052851.2);
});

test('region-only accounts (Team = Europe, no sub-team) surface as a data-quality issue on the region tabs', () => {
  const regions = [];
  const ctx = mainContext({
    fetchOpportunities: team => data(team).opps,
    fetchAccounts: team => data(team).accounts,
    fetchGoals: team => data(team).goals,
    fetchRegionOnlyAccounts: team => { regions.push(team); return [{ id: '001F', name: 'Flutter Entertainment', url: 'u', team: 'Europe', currentArr: 50000, owner: 'Sam' }]; },
    rollLedger: (team, quarter) => { const out = {}; ctx.quartersBetween('Q1-2026', quarter).forEach(q => { out[q] = data(team).ledger[q] || { startingArr: 0, endingArr: 0 }; }); return out; },
  });
  const view = ctx.buildView('Europe - DACH', 'Q3-2026');
  assert.deepEqual(regions, ['Europe - DACH']);
  const regionIssues = view.dataQuality.filter(i => i.account);
  assert.equal(regionIssues.length, 1);
  assert.equal(regionIssues[0].account.name, 'Flutter Entertainment');
  assert.equal(regionIssues[0].issue, vm.runInContext('DATA_QUALITY_CHECKS.regionOnly', ctx));
});

test('teamSheet keeps a valid quarter already in B2 and only fills in a missing / invalid one', () => {
  const sheets = {};
  const fakeSheet = (name, values) => {
    const cells = { 'A1:B2': values };
    return { name, getRange: a1 => ({ getValue: () => cells[a1] && cells[a1][0][1], getValues: () => cells['A1:B2'], setValues: v => { cells['A1:B2'] = v; } }), values: () => cells['A1:B2'] };
  };
  const ctx = mainContext({
    SpreadsheetApp: { getActive: () => ({ getSheetByName: name => sheets[name] || null, insertSheet: name => (sheets[name] = fakeSheet(name, [['', ''], ['', '']])) }) },
  });
  const today = ctx.defaultQuarter();
  sheets['Europe - DACH'] = fakeSheet('Europe - DACH', [['Team', 'Europe - DACH'], ['Quarter', 'Q1-2026']]);
  sheets['Japan'] = fakeSheet('Japan', [['', 'Japan'], ['', 'last quarter']]);
  assert.deepEqual(ctx.teamSheet('Europe - DACH').values(), [['Team', 'Europe - DACH'], ['Quarter', 'Q1-2026']]);
  assert.deepEqual(ctx.teamSheet('Japan').values(), [['Team', 'Japan'], ['Quarter', today]]);
  assert.deepEqual(ctx.teamSheet('Europe').values(), [['Team', 'Europe'], ['Quarter', today]]);
});

test('tabs default to the quarter we are in today', () => {
  const ctx = mainContext({});
  const today = new Date().toISOString().slice(0, 10);
  assert.equal(ctx.defaultQuarter(), ctx.quarterOfDate(today));
  assert.equal(vm.runInContext("quarterOfDate('2026-09-09')", ctx), 'Q3-2026');
});
