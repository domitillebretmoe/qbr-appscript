// Deck.gs against the fake Slides objects: token map, table filling, chart placement and the deck link on the tab.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');
const { FakeSheet, globals } = require('./fake-sheets');
const { FakePresentation, fakeChart } = require('./fake-slides');
const { opp, dach, dachAccounts } = require('./fixtures');

const env = globals();
const ctx = vm.createContext(env);
['Config.gs', 'Metrics.gs', 'Reps.gs', 'Ledger.gs', 'Formulas.gs', 'Render.gs', 'RawData.gs', 'Deck.gs']
  .forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));

const TODAY = '2026-09-15';
const opps = dach.concat([
  opp('Q3-2026', '3- Tech Validation', 'Allianz', 'Land', 'Enterprise', 180000, 1, { major: true, expectedDeltaArr: 90000, closeDate: '2025-10-20', pilotStatus: 'Active' }),
  opp('Q4-2026', '4- Proposal', 'SAP SE', 'Land', 'Enterprise', 800000, 1, { major: true, expectedDeltaArr: 560000, closeDate: '2025-12-19' }),
  opp('Q4-2026', '2- Qualification', 'BMW Group', 'Land', 'Enterprise', 300000, 1, { major: true, expectedDeltaArr: 60000, closeDate: '2026-01-10' }),
  opp('Q4-2026', 'R2- Renewal Engagement', 'Helaba', 'Renewal', 'Renewal', -25000, 0, { expectedDeltaArr: -25000, closeDate: '2025-11-28', accountArr: 150000 }),
]);
function sampleView(quarter = 'Q3-2026') {
  const goals = { 'Q1-2026': { revenue: 2000000, logos: 0 }, 'Q2-2026': { revenue: 6700000, logos: 1 }, 'Q3-2026': { revenue: 7500000, logos: 2 }, 'Q4-2026': { revenue: 8000000, logos: 2 }, 'Q1-2027': { revenue: 8500000, logos: 3 } };
  const seed = vm.runInContext("TEAM_SEEDS.filter(s => s[0] === 'Europe - DACH')[0]", ctx);
  const ledger = { 'Q1-2026': { startingArr: seed[1], endingArr: seed[2] }, 'Q2-2026': { startingArr: seed[2], endingArr: seed[3] } };
  ctx.quartersBetween('Q1-2026', quarter).forEach(q => {
    const prev = ledger[ctx.shiftQuarter(q, -1)];
    ledger[q] = ledger[q] || { startingArr: prev.endingArr, endingArr: prev.endingArr + ctx.quarterMetrics(opps, q, {}).netAddedArr };
  });
  const reps = [{ name: 'Kalle Harnos', url: 'https://sf/rep/1', monthsInSeat: 14.2, accountsOwned: 30, repGoal: 1200000, qWonArr: 96000, fyWonArr: 246000,
    attainmentPct: 0.205, meetings: 41, activities: 130, activityCoveragePct: 0.6, pipelineCreatedArr: 400000, stalledArr: 50000 }];
  return ctx.composeView({ team: 'Europe - DACH', members: ['Europe - DACH'], quarter, opps, accounts: dachAccounts, goals, ledgers: [ledger], unassignedAccounts: [], reps, today: TODAY });
}

// A template with one of each placeholder kind, like tools/build_qbr_deck.py emits.
function templateDeck() {
  const deck = new FakePresentation();
  const cover = deck.slide();
  cover.shape('{{team}}  ·  {{quarter}} QBR');
  cover.shape('{{status}}');
  cover.shape('{{refreshed}}');
  const score = deck.slide();
  score.shape('Net Added ARR {{netAddedArr}} of {{revenueGoal}} ({{attainment}}) {{qoq.netAddedArr}}');
  score.shape('{{rag.attainment}}');
  score.shape('Logos {{logosWonLand}} / {{logosWonMajors}} vs {{logoGoal}}; Q+1 {{q1}} forecast {{f1.netForecastArr}} ({{f1.netForecastPct}}), commit {{f1.commitArr}}, early {{f1.earlySharePct}}');
  score.table([['Account', 'Type', 'TCV', 'Delta ARR', 'Owner', 'Why we won'], ['{{rows:dealsWon}}', '', '', '', '', 'Why we won / what it unlocks']]);
  const bridge = deck.slide();
  bridge.shape('{{chart:ARR bridge}}', { left: 50, top: 60, width: 400, height: 250 });
  bridge.shape('{{chart:Renewals won vs lost}}');
  bridge.table([['Stage', '# Open', 'Delta ARR', '% of pipe'], ['{{rows:stagesQ1}}', '', '', ''], ['', '', '', ''], ['', '', '', '']]);
  bridge.table([['Account', 'Close', 'ARR', 'Expected', 'Risk driver', 'Mitigation'], ['{{rows:renewalsAtRisk}}', '', '', '', 'Risk driver', 'Mitigation + owner']]);
  bridge.table([['Account', 'Status', 'End', 'Delta ARR', 'Notes'], ['{{rows:pilots}}', '', '', '', 'Outcome + engineering notes'], ['', '', '', '', '']]);
  const commentary = deck.slide();
  commentary.shape('What worked and why - Sales leadership fills this in');
  return deck;
}

test('tokens carry the tab metrics, formatted, with forecast and QoQ variants', () => {
  const view = sampleView();
  const t = ctx.deckTokens(view);
  assert.equal(t.team, 'Europe - DACH');
  assert.equal(t.quarter, 'Q3-2026');
  assert.equal(t.q1, 'Q4-2026');
  assert.equal(t.q2, 'Q1-2027');
  assert.equal(t.status, view.current.status === 'ACTUALS' ? 'ACTUALS (quarter closed)' : ctx.statusText(view.current));
  assert.equal(t.netAddedArr, ctx.formatMoney(view.current.netAddedArr));
  assert.equal(t.revenueGoal, '$7.50M');
  assert.equal(t.logosWonLand, String(view.current.logosWonLand));
  assert.match(t.attainment, /^-?\d+%$/);
  assert.equal(t['f1.revenueGoal'], '$8.00M');
  assert.equal(t['f1.commitArr'], '$560K'); // SAP SE in stage 4
  const early = opps.filter(o => o.quarter === 'Q4-2026' && !o.isClosed && /^[0-2]/.test(o.stage)).reduce((s, o) => s + o.expectedDeltaArr, 0);
  assert.ok(early >= 60000, 'BMW in stage 2');
  assert.equal(t['f1.earlyArr'], ctx.formatMoney(early));
  assert.equal(t['f1.renewalsDueCount'], String(view.future[0].renewalsDueCount));
  assert.equal(t['rag.attainment'], ctx.ragLabel(view.current.attainment));
  assert.ok(Object.keys(t).every(k => t[k] !== undefined && t[k] !== 'undefined' && t[k] !== 'NaN'), 'no undefined tokens');
  assert.ok(t['qoq.netAddedArr'] !== undefined);
});

test('reps on pace compare FY attainment with the fiscal year elapsed, not the quarter', () => {
  assert.equal(Math.round(ctx.fiscalYearElapsed('Q3-2026', TODAY) * 1000) / 1000, Math.round((226 / 365) * 1000) / 1000); // 1 Feb -> 15 Sep
  assert.equal(ctx.fiscalYearElapsed('Q2-2025', TODAY), 1);
  const view = sampleView('Q2-2026'); // closed quarter: quarterElapsedPct = 1
  view.reps.push(Object.assign({}, view.reps[0], { name: 'Ahead Rep', attainmentPct: 0.7 }));
  assert.equal(ctx.deckTokens(view).repsOnPace, '1');
});

test('stage total shows no share when there is no positive open pipeline', () => {
  const withStages = stages => ctx.deckRows(Object.assign({}, sampleView(), { future: sampleView().future.map(f => Object.assign({}, f, { pipelineByStage: stages })) })).stagesQ1;
  assert.deepEqual(withStages([]), [['Total open', '0', '$0', '-']]);
  const negative = withStages([{ stage: 'R2- Renewal Engagement', count: 1, deltaArr: -25000, share: null }]);
  assert.equal(negative[negative.length - 1][3], '-');
  const positive = withStages([{ stage: '4- Proposal', count: 2, deltaArr: 500000, share: 1 }]);
  assert.equal(positive[positive.length - 1][3], '100%');
});

test('fillDeck replaces every token, resizes and fills tables, keeps manual columns and commentary', () => {
  const view = sampleView();
  const deck = templateDeck();
  const charts = [fakeChart('ARR bridge Q3-2026'), fakeChart('Attainment vs goal')];
  ctx.fillDeck(deck, view, charts);
  assert.ok(!/\{\{[^}]+\}\}/.test(deck.text()), `unreplaced tokens in:\n${deck.text()}`);
  assert.ok(deck.text().indexOf('Europe - DACH  ·  Q3-2026 QBR') >= 0);
  // Deals won: one row per Closed Won opp of the quarter (+ none hidden here), manual column kept.
  const dealsWon = deck.slides[1].tables[0].values();
  const shown = view.current.lists.dealsWon;
  assert.equal(dealsWon.length - 1, shown.length);
  assert.equal(dealsWon[1][0], shown[0].account);
  assert.equal(dealsWon[1][5], 'Why we won / what it unlocks');
  assert.equal(deck.slides[1].tables[0].rows[1][0].link, shown[0].accountUrl);
  // Stage table shrinks to the rows + total; at-risk table keeps its manual columns; pilots row filled.
  const stages = deck.slides[2].tables[0].values();
  assert.equal(stages[stages.length - 1][0], 'Total open');
  assert.equal(stages.length - 1, view.future[0].pipelineByStage.length + 1);
  const atRisk = deck.slides[2].tables[1].values();
  assert.ok(atRisk.slice(1).some(r => r[0] === 'Helaba'), JSON.stringify(atRisk));
  assert.equal(atRisk[1][4], 'Risk driver');
  const pilots = deck.slides[2].tables[2].values();
  assert.equal(pilots.length, 2);
  assert.equal(pilots[1][0], 'Allianz');
  assert.match(pilots[1][1], /^Active/);
  // Charts: the bridge frame is swapped for the linked chart at the same box; the missing one is cleared.
  assert.equal(deck.slides[2].charts.length, 1);
  assert.deepEqual([deck.slides[2].charts[0].left, deck.slides[2].charts[0].width], [50, 400]);
  assert.equal(deck.slides[2].shapes.length, 1);
  assert.equal(deck.slides[2].shapes[0].textRange.text, '');
  // RAG pill coloured, commentary untouched.
  const pill = deck.slides[1].shapes[1];
  assert.ok(['ON / ABOVE PLAN', 'WATCH', 'BEHIND'].includes(pill.textRange.text));
  assert.ok(pill.fill);
  assert.equal(deck.slides[3].shapes[0].textRange.text, 'What worked and why - Sales leadership fills this in');
});

test('a {{size:N}} marker restores the intended font size and is removed, RAG pill included', () => {
  const deck = new FakePresentation();
  const slide = deck.slide();
  const kpi = slide.shape('{{netAddedArr}}{{size:24}}');
  const pill = slide.shape('{{rag.attainment}}{{size:7}}');
  const plain = slide.shape('{{quarter}}');
  ctx.fillDeck(deck, sampleView(), []);
  assert.equal(kpi.textRange.fontSize, 24);
  assert.ok(!kpi.textRange.text.includes('{{'), kpi.textRange.text);
  assert.equal(pill.textRange.fontSize, 7);
  assert.ok(['ON / ABOVE PLAN', 'WATCH', 'BEHIND'].includes(pill.textRange.text));
  assert.ok(pill.fill);
  assert.equal(plain.textRange.fontSize, null);
});

test('an empty table shows a single "-" row', () => {
  const view = sampleView('Q1-2026'); // no Closed Won lands in Q1
  const deck = new FakePresentation();
  deck.slide().table([['Account', 'Type'], ['{{rows:logosWon}}', ''], ['', ''], ['', '']]);
  ctx.fillDeck(deck, view, []);
  assert.deepEqual(deck.slides[0].tables[0].values(), [['Account', 'Type'], ['-', '']]);
});

test('deck link is remembered per team / quarter and rewritten by renderTeamTab', () => {
  env.SpreadsheetApp.getActive = () => ({ getSheetByName: () => null, insertSheet: name => new FakeSheet(name) });
  const view = sampleView();
  ctx.rememberDeckLink('Europe - DACH', 'Q3-2026', 'https://docs.google.com/presentation/d/abc/edit', view.today);
  assert.equal(ctx.deckLinkFor('Europe - DACH', 'Q2-2026'), null);
  ctx.rememberDeckLink('Europe Majors - UKI', 'Q3-2026', 'https://docs.google.com/presentation/d/uki/edit', view.today);
  assert.equal(ctx.deckLinkFor('Europe - DACH', 'Q3-2026').url, 'https://docs.google.com/presentation/d/abc/edit');
  assert.equal(env.PropertiesService.getDocumentProperties().getProperty('QBR_DECK_LINKS'), null); // one property per team / quarter
  const sheet = new FakeSheet('Europe - DACH');
  ctx.writeDeckLink(sheet, 'Europe - DACH', 'Q3-2026');
  assert.equal(sheet.cell(1, 13).link, 'https://docs.google.com/presentation/d/abc/edit');
  assert.match(sheet.cell(1, 13).value, /^Open QBR deck \(built \d{4}-\d{2}-\d{2}\)$/);
});

test('slidesIdFrom accepts a URL or a bare ID', () => {
  assert.equal(ctx.slidesIdFrom('https://docs.google.com/presentation/d/1YMla6Dz5c2O5QyIWs3_l5ofhX1jCkTfX3dcBr4zCbxA/edit?slide=id.p1#slide=id.p1'), '1YMla6Dz5c2O5QyIWs3_l5ofhX1jCkTfX3dcBr4zCbxA');
  assert.equal(ctx.slidesIdFrom(' 1YMla6Dz5c2O5QyIWs3_l5ofhX1jCkTfX3dcBr4zCbxA '), '1YMla6Dz5c2O5QyIWs3_l5ofhX1jCkTfX3dcBr4zCbxA');
  assert.throws(() => ctx.slidesIdFrom('nope'));
});
