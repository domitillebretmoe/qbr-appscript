// "QBR > Build deck (this tab)": a Google Slides deck for the selected team / quarter, made from the template deck
// (tools/build_qbr_deck.py, imported once into Google Slides) and the tab's cockpit view. The template carries
// three kinds of placeholder, filled in this order:
//   {{rows:name}}   first data cell of a table: the table is resized to deckRows(view)[name] and filled row by row
//                   (a null cell keeps the template's text, i.e. the manual commentary columns);
//   {{chart:name}}  text of a placeholder frame: replaced by the cockpit chart of that name, linked to the sheet;
//   {{key}}         anywhere in text: replaced by deckTokens(view)[key] (formatted strings, '-' when unknown).
// Commentary slides have no placeholders and come through untouched. The deck is a new file in the user's Drive,
// its link is written to the tab (M1) and remembered per team / quarter so a refresh keeps it.
const DECK_TEMPLATE_KEY = 'QBR_DECK_TEMPLATE_ID';
const DECK_LINKS_KEY = 'QBR_DECK_LINKS';
const DEFAULT_DECK_TEMPLATE_ID = '1SNdm5jqL-6AtJOWgSmEb7Blk_knhv_uJXCNsT96R9XU';
const ROWS_TOKEN = /^\{\{rows:([A-Za-z0-9_]+)\}\}$/;
const CHART_TOKEN = /\{\{chart:([^}]+)\}\}/;
const FONT_SIZE_TAG = /\{\{size:(\d+(?:\.\d+)?)\}\}/;
const RAG_LABELS = {
  green: { text: 'ON / ABOVE PLAN', fill: '#D1FAE5', color: '#047857' },
  amber: { text: 'WATCH', fill: '#FEF3C7', color: '#D97706' },
  red: { text: 'BEHIND', fill: '#FEE2E2', color: '#DC2626' },
};
// Stage prefixes for the path-to-goal convention: commit = 4-5, best case = 3 (Tech Validation), early = 0-2.
const COMMIT_STAGES = [4, 5];
const BEST_CASE_STAGES = [3];

function buildDeckForActiveTab() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!isTeamTab(sheet)) throw new Error('Select a team tab first (A1 = Team, A2 = Quarter).');
  const templateId = deckTemplateId();
  // Refresh first so the deck and the tab show the same Salesforce pull.
  const view = withRefreshLock(() => refreshTab(sheet));
  const deck = buildDeck(view, templateId, sheet.getCharts());
  rememberDeckLink(view.team, view.quarter, deck.url, view.today);
  writeDeckLink(sheet, view.team, view.quarter);
  SpreadsheetApp.getActive().toast(`Deck ready: ${deck.name}`, 'QBR');
  SpreadsheetApp.getUi().alert('QBR deck built', `${deck.name}\n\n${deck.url}\n\nThe link is also in cell M1 of the tab.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function setDeckTemplate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Deck template', 'URL (or file ID) of the Google Slides template deck', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const id = slidesIdFrom(response.getResponseText());
  const template = SlidesApp.openById(id);
  PropertiesService.getDocumentProperties().setProperty(DECK_TEMPLATE_KEY, id);
  ui.alert(`Template set: "${template.getName()}" (${template.getSlides().length} slides)`);
}

function deckTemplateId() {
  return PropertiesService.getDocumentProperties().getProperty(DECK_TEMPLATE_KEY) || DEFAULT_DECK_TEMPLATE_ID;
}

// Accepts the presentation URL or the bare ID.
function slidesIdFrom(text) {
  const match = /\/presentation\/d\/([-\w]{20,})/.exec(text) || /^([-\w]{20,})$/.exec(String(text).trim());
  if (!match) throw new Error('That is not a Google Slides URL or file ID.');
  return match[1];
}

// Copies every template slide into a new presentation and fills it. `charts` are the tab's embedded charts.
function buildDeck(view, templateId, charts) {
  const template = SlidesApp.openById(templateId);
  const name = deckName(view);
  const deck = SlidesApp.create(name);
  template.getSlides().forEach(slide => deck.appendSlide(slide));
  deck.getSlides()[0].remove(); // the blank slide SlidesApp.create starts with
  fillDeck(deck, view, charts || []);
  const url = deck.getUrl();
  deck.saveAndClose();
  return { name, url, id: deck.getId() };
}

function deckName(view) {
  return `${view.team} ${view.quarter} QBR (${view.current.status === 'ACTUALS' ? 'actuals' : 'forecast'}, cockpit ${view.today})`;
}

function fillDeck(deck, view, charts) {
  const rows = deckRows(view);
  deck.getSlides().forEach(slide => {
    slide.getTables().forEach(table => fillTable(table, rows));
    slide.getShapes().forEach(shape => placeChart(slide, shape, charts));
  });
  const markers = new Set();
  deck.getSlides().forEach(slide => slide.getShapes().forEach(shape => restoreFontSize(shape, markers)));
  const tokens = deckTokens(view);
  Object.keys(tokens).forEach(key => deck.replaceAllText(`{{${key}}}`, tokens[key] == null ? '-' : String(tokens[key])));
  markers.forEach(marker => deck.replaceAllText(marker, ''));
  deck.getSlides().forEach(slide => slide.getShapes().forEach(paintRagPill));
}

// The template draws long {{tokens}} at a reduced size so they fit their card and ends the text with a 1pt "{{size:24}}"
// marker; the shape gets that size back here (the value replacing the token is short) and the marker is removed.
function restoreFontSize(shape, markers) {
  const match = FONT_SIZE_TAG.exec(shape.getText().asString());
  if (!match) return;
  shape.getText().getTextStyle().setFontSize(Number(match[1]));
  markers.add(match[0]);
}

// ---------------------------------------------------------------- tables
function fillTable(table, rows) {
  const marker = ROWS_TOKEN.exec(cellText(table.getCell(1, 0)).trim());
  if (!marker || !rows[marker[1]]) return;
  const data = rows[marker[1]];
  const columns = table.getNumColumns();
  const template = [];
  for (let c = 0; c < columns; c++) template.push(c === 0 ? '' : cellText(table.getCell(1, c)));
  const want = Math.max(1, data.length);
  while (table.getNumRows() - 1 < want) table.appendRow();
  while (table.getNumRows() - 1 > want) table.getRow(table.getNumRows() - 1).remove();
  if (!data.length) {
    for (let c = 0; c < columns; c++) setCell(table.getCell(1, c), c === 0 ? '-' : '');
    return;
  }
  data.forEach((row, r) => {
    for (let c = 0; c < columns; c++) setCell(table.getCell(r + 1, c), row[c] === null || row[c] === undefined ? template[c] : row[c]);
  });
}

function cellText(cell) {
  return cell.getText().asString().replace(/\n$/, '');
}

// A cell value is a string, a number (already formatted upstream) or a { text, url } link.
function setCell(cell, value) {
  const text = cell.getText();
  const isLink = value && typeof value === 'object';
  text.setText(String(isLink ? value.text : value));
  if (isLink && value.url) text.getTextStyle().setLinkUrl(value.url);
}

// ---------------------------------------------------------------- charts
// A shape whose text holds {{chart:name}} is swapped for the cockpit chart whose title starts with `name`, at the
// same position and size, linked to the spreadsheet so "Update all" in Slides refreshes it. Without a matching
// chart the token is cleared and the frame stays as a manual placeholder.
function placeChart(slide, shape, charts) {
  const match = CHART_TOKEN.exec(shape.getText().asString());
  if (!match) return;
  const wanted = match[1].trim().toLowerCase();
  const chart = charts.filter(c => String(chartTitle(c)).toLowerCase().indexOf(wanted) === 0)[0];
  if (!chart) {
    shape.getText().replaceAllText(match[0], '');
    return;
  }
  try {
    slide.insertSheetsChart(chart, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
  } catch (e) {
    shape.getText().replaceAllText(match[0], `Chart "${chartTitle(chart)}" could not be linked (${e.message}) - Insert > Chart > From Sheets.`);
    return;
  }
  shape.remove();
}

function chartTitle(chart) {
  const options = chart.getOptions();
  return options && options.get ? options.get('title') || '' : '';
}

// ---------------------------------------------------------------- RAG pills
// After token replacement a pill reads "ON / ABOVE PLAN" / "WATCH" / "BEHIND": colour it like the cockpit's RAG.
function paintRagPill(shape) {
  const text = shape.getText().asString().trim();
  const tone = Object.keys(RAG_LABELS).filter(k => RAG_LABELS[k].text === text)[0];
  if (!tone) return;
  shape.getFill().setSolidFill(RAG_LABELS[tone].fill);
  shape.getText().getTextStyle().setForegroundColor(RAG_LABELS[tone].color);
}

function ragLabel(value) {
  if (value == null) return '-';
  return RAG_LABELS[value >= RAG.green ? 'green' : value >= RAG.amber ? 'amber' : 'red'].text;
}

// Share of the selected quarter's fiscal year elapsed on `today` (rep goals and attainment are FY figures).
function fiscalYearElapsed(quarter, today) {
  const fy = parseQuarter(quarter).fy;
  const start = quarterStart(`Q1-${fy}`);
  return Math.min(1, Math.max(0, daysBetween(start, today) / daysBetween(start, quarterStart(`Q1-${fy + 1}`))));
}

// ---------------------------------------------------------------- values
const fmtMoney = v => (v == null ? '-' : formatMoney(v));
const fmtPct = (v, decimals) => (v == null ? '-' : `${(v * 100).toFixed(decimals || 0)}%`);
const fmtInt = v => (v == null ? '-' : String(Math.round(v)));
const fmtNum = v => (v == null ? '-' : String(Math.round(v * 10) / 10));
const fmtMult = v => (v == null ? '-' : `${v.toFixed(1)}x`);
const fmtDate = v => (v ? String(v).slice(0, 10) : '-');
const deckLink = (text, url) => ({ text: text || '-', url: url || '' });
const stageOrder = stage => { const m = /^(\d+)/.exec(stage || ''); return m ? Number(m[1]) : null; };
const inStages = (o, numbers) => numbers.indexOf(stageOrder(o.stage)) >= 0;

// Every {{key}} the template may use. Keys mirror the metric keys of the tab (Metrics.gs) so the appendix mapping
// slide stays true; f1. / f2. are the two forecast quarters, qoq. the tab's QoQ column, rag. the RAG pill text.
function deckTokens(view) {
  const m = view.current;
  const prev = view.previous;
  const [f1, f2] = view.future;
  const shown = m.lists.dealsWon;
  const shownArr = sum(shown, 'deltaArr');
  const pct = (part, whole) => (whole ? ratio(part, whole) : null);
  const t = {
    team: view.team,
    quarter: view.quarter,
    q1: f1.quarter,
    q2: f2.quarter,
    fy: `FY${parseQuarter(view.quarter).fy}`,
    status: statusText(m),
    refreshed: `cockpit refreshed ${view.today}`,
    today: view.today,
    qoqBasis: qoqBasis(prev),
    qoqHeader: prev && prev.samePoint ? `QoQ vs ${prev.quarter} at the same point (day ${prev.elapsedDays})` : `QoQ vs ${prev ? prev.quarter : 'n/a'}`,
    // Current quarter.
    revenueGoal: fmtMoney(m.revenueGoal),
    netAddedArr: fmtMoney(m.netAddedArr),
    attainment: fmtPct(m.attainment),
    quarterElapsedPct: fmtPct(m.quarterElapsedPct),
    pace: fmtPct(m.pace),
    openPipelineArr: fmtMoney(m.openPipelineArr),
    pipelineCoverage: fmtMult(m.pipelineCoverage),
    logoGoal: fmtInt(m.logoGoal),
    logoAttainment: fmtNum(m.logoAttainment),
    logosWonLand: fmtInt(m.logosWonLand),
    logosWonMajors: fmtInt(m.logosWonMajors),
    logoAttainmentPct: fmtPct(m.logoAttainmentPct),
    logoGap: fmtNum(m.logoGoal == null ? null : Math.max(0, m.logoGoal - m.logoAttainment)),
    renewals: fmtInt(m.renewals),
    wonRenewals: fmtInt(m.wonRenewals),
    lostRenewals: fmtInt(m.renewals - m.wonRenewals),
    renewalRate: fmtPct(m.renewalRate),
    churnArr: fmtMoney(m.churnArr),
    churnCustomers: fmtInt(m.churnCustomers),
    downgradeArr: fmtMoney(m.downgradeArr),
    downgradeCount: fmtInt(m.downgradeCount),
    fullChurnArr: fmtMoney(m.fullChurnArr),
    fullChurnCount: fmtInt(m.fullChurnCount),
    startingArr: fmtMoney(m.startingArr),
    addedArr: fmtMoney(m.addedArr),
    newLogoArr: fmtMoney(m.newLogoArr),
    expansionArr: fmtMoney(m.expansionArr),
    endingArr: fmtMoney(m.endingArr),
    endingArrQoq: fmtPct(pct(m.endingArr - m.startingArr, m.startingArr), 1),
    grr: fmtPct(m.grr, 1),
    nrr: fmtPct(m.nrr, 1),
    wonCount: fmtInt(m.wonCount),
    wonArr: fmtMoney(m.wonArr),
    dealsWonShown: fmtInt(shown.length),
    dealsWonOther: fmtInt(m.wonCount - shown.length),
    dealsWonOtherArr: fmtMoney(m.wonArr - shownArr),
    topDealShare: fmtPct(pct(shown.length ? shown[0].deltaArr : 0, m.wonArr)),
    lostPipelineCount: fmtInt(m.lostPipelineCount),
    lostPipelineArr: fmtMoney(m.lostPipelineArr),
    activeCustomers: fmtInt(m.activeCustomers),
    majorCustomers: fmtInt(m.majorCustomers),
    enterpriseCustomers: fmtInt(m.enterpriseCustomers),
    activatedProspects: fmtInt(m.activatedProspects),
    conversionRate: fmtPct(m.conversionRate),
    activePilots: fmtInt(m.activePilots),
    pilotsCompleted: fmtInt(m.pilotsCompleted),
    activeArr: fmtMoney(m.activeArr),
    topChurns: m.topChurns.length ? m.topChurns.join(', ') : '-',
    churnReasons: m.churnReasons.length ? m.churnReasons.join(', ') : '-',
    partnerNetAddedArr: fmtMoney(m.partnerNetAddedArr),
    partnerNewLogos: fmtInt(m.partnerNewLogos),
    // ARR bridge as % of Starting ARR.
    'bridge.addedPct': fmtPct(pct(m.addedArr, m.startingArr), 1),
    'bridge.downgradePct': fmtPct(pct(m.downgradeArr, m.startingArr), 1),
    'bridge.fullChurnPct': fmtPct(pct(m.fullChurnArr, m.startingArr), 1),
    'bridge.endingPct': fmtPct(pct(m.endingArr, m.startingArr), 1),
    // RAG pills.
    'rag.attainment': ragLabel(m.attainment),
    'rag.pace': ragLabel(m.pace),
    'rag.logoAttainmentPct': ragLabel(m.logoAttainmentPct),
    'rag.renewalRate': ragLabel(m.renewalRate),
    'rag.f1.netForecastPct': ragLabel(f1.netForecastPct),
    // Rep summary.
    repCount: fmtInt(view.reps.length),
    repsOnPace: fmtInt(view.reps.filter(r => r.attainmentPct != null && r.attainmentPct >= fiscalYearElapsed(view.quarter, view.today)).length),
  };
  // QoQ column, same wording as the tab (arrow + delta; '-' without a like-for-like baseline).
  [['netAddedArr', 'money'], ['attainment', 'pct'], ['pace', 'pct'], ['logosWonLand', 'int'], ['logosWonMajors', 'int'],
    ['logoAttainment', 'int'], ['renewals', 'int'], ['wonRenewals', 'int'], ['renewalRate', 'pct'], ['churnArr', 'money'],
    ['churnCustomers', 'int'], ['grr', 'pct'], ['nrr', 'pct'], ['endingArr', 'money'], ['startingArr', 'money'],
    ['openPipelineArr', 'money'], ['pipelineCoverage', 'mult'], ['pilotsCompleted', 'int'], ['wonCount', 'int'],
    ['lostPipelineArr', 'money']].forEach(([key, kind]) => {
    t[`qoq.${key}`] = prev ? qoqText(kind, m[key], prev[key]) : '-';
  });
  [f1, f2].forEach((f, i) => Object.assign(t, forecastTokens(`f${i + 1}`, f, view)));
  return t;
}

function forecastTokens(prefix, f, view) {
  const rows = inQuarter(view.opps, f.quarter);
  const open = rows.filter(o => !o.isClosed);
  const won = rows.filter(isWon);
  const wonArr = sum(won, 'deltaArr');
  const commitArr = sum(open.filter(o => inStages(o, COMMIT_STAGES)), 'expectedDeltaArr');
  const bestCaseArr = sum(open.filter(o => inStages(o, BEST_CASE_STAGES)), 'expectedDeltaArr');
  const earlyArr = sum(open.filter(o => { const n = stageOrder(o.stage); return n != null && n < BEST_CASE_STAGES[0]; }), 'expectedDeltaArr');
  const earlyPipeline = f.pipelineByStage.filter(s => { const n = stageOrder(s.stage); return n != null && n < BEST_CASE_STAGES[0]; });
  const openLogoMajors = open.filter(o => isNewLogo(o) && o.major);
  const t = {};
  t[`${prefix}.quarter`] = f.quarter;
  t[`${prefix}.revenueGoal`] = fmtMoney(f.revenueGoal);
  t[`${prefix}.netForecastArr`] = fmtMoney(f.netForecastArr);
  t[`${prefix}.netForecastPct`] = fmtPct(f.netForecastPct);
  t[`${prefix}.logoGoal`] = fmtInt(f.logoGoal);
  t[`${prefix}.logoForecast`] = fmtNum(f.logoForecast);
  t[`${prefix}.logoGap`] = fmtNum(f.logoGoal == null ? null : Math.max(0, f.logoGoal - f.logoForecast));
  t[`${prefix}.openLogoMajors`] = fmtInt(openLogoMajors.length);
  t[`${prefix}.pipelineArr`] = fmtMoney(f.pipelineArr);
  t[`${prefix}.pipelineCoverage`] = fmtMult(f.pipelineCoverage);
  t[`${prefix}.openCount`] = fmtInt(open.length);
  t[`${prefix}.renewalsDueCount`] = fmtInt(f.renewalsDueCount);
  t[`${prefix}.renewalArrDue`] = fmtMoney(f.renewalArrDue);
  t[`${prefix}.startingArr`] = fmtMoney(f.startingArr);
  t[`${prefix}.forecastArr`] = fmtMoney(f.forecastArr);
  t[`${prefix}.forecastChurnArr`] = fmtMoney(f.forecastChurnArr);
  t[`${prefix}.forecastChurnCount`] = fmtInt(f.predictedChurn.length);
  t[`${prefix}.forecastEndingArr`] = fmtMoney(f.forecastEndingArr);
  t[`${prefix}.wonArr`] = fmtMoney(wonArr);
  t[`${prefix}.wonCount`] = fmtInt(won.length);
  t[`${prefix}.commitArr`] = fmtMoney(commitArr);
  t[`${prefix}.bestCaseArr`] = fmtMoney(bestCaseArr);
  t[`${prefix}.earlyArr`] = fmtMoney(earlyArr);
  t[`${prefix}.gapArr`] = fmtMoney(Math.max(0, (f.revenueGoal || 0) - wonArr - commitArr));
  t[`${prefix}.gapToForecast`] = fmtMoney((f.revenueGoal || 0) - f.netForecastArr);
  t[`${prefix}.earlySharePct`] = fmtPct(f.pipelineArr > 0 ? sum(earlyPipeline, 'deltaArr') / f.pipelineArr : null);
  t[`${prefix}.topDealsArr`] = fmtMoney(sum(f.topDeals, 'deltaArr'));
  t[`${prefix}.topDealsSharePct`] = fmtPct(f.pipelineArr > 0 ? sum(f.topDeals, 'deltaArr') / f.pipelineArr : null);
  return t;
}

// Every {{rows:name}} table. Cells: string | { text, url } | null (keep the template's text: the manual columns).
function deckRows(view) {
  const m = view.current;
  const [f1, f2] = view.future;
  const shown = m.lists.dealsWon;
  const acct = o => deckLink(o.account, o.accountUrl);
  const dealsWon = shown.map(o => [acct(o), o.type || '-', fmtMoney(o.amount || 0), fmtMoney(o.deltaArr), o.owner || '-', null]);
  if (m.wonCount > shown.length) {
    dealsWon.push([`Other wins not shown (${m.wonCount - shown.length})`, '', '', fmtMoney(m.wonArr - sum(shown, 'deltaArr')), '', '']);
  }
  const quarterRows = inQuarter(view.opps, view.quarter);
  const openLogoOpps = byField(quarterRows.filter(o => !o.isClosed && isNewLogo(o) && o.major), 'expectedLogoImpact', true);
  const retention = [].concat(
    m.lists.churned.map(o => [acct(o), 'Churned', o.recordType || '-', fmtDate(o.closeDate), fmtMoney(o.deltaArr), o.lostReason || null]),
    m.lists.downgrades.map(o => [acct(o), 'Downgrade', o.recordType || '-', fmtDate(o.closeDate), fmtMoney(o.deltaArr), o.lostReason || null]),
    m.lists.renewalsWon.filter(o => o.deltaArr >= 0).map(o => [acct(o), 'Renewed', o.recordType || '-', fmtDate(o.closeDate), fmtMoney(o.deltaArr), null]),
  );
  const pilots = [].concat(
    m.completedPilotList.map(o => [acct(o), `Complete -> ${o.stage}`, fmtDate(o.pilotEndDate), fmtMoney(o.deltaArr), null]),
    m.activePilotList.map(o => [acct(o), `Active${o.pilotStatus ? ` (${o.pilotStatus})` : ''}`, fmtDate(o.pilotExpectedEnd || o.closeDate), fmtMoney(o.deltaArr), null]),
  );
  const stageRows = stages => stages.map(s => [s.stage, fmtInt(s.count), fmtMoney(s.deltaArr), fmtPct(s.share)])
    .concat([['Total open', fmtInt(sum(stages, 'count')), fmtMoney(sum(stages, 'deltaArr')), sum(stages, 'deltaArr') > 0 ? '100%' : '-']]);
  const topDeals = f => f.topDeals.map(o => [acct(o), o.stage || '-', fmtDate(o.closeDate), fmtMoney(o.deltaArr), null]);
  const atRisk = f => f.predictedChurn.map(o => [acct(o), fmtDate(o.closeDate), fmtMoney(o.accountArr), fmtMoney(o.expectedDeltaArr), null, null]);
  return {
    dealsWon,
    logosWon: m.lists.logosWon.map(o => [acct(o), o.type || '-', fmtMoney(o.amount || 0), fmtMoney(o.deltaArr), o.owner || '-']),
    openLogoOpps: openLogoOpps.map(o => [acct(o), o.stage || '-', fmtNum(o.expectedLogoImpact), fmtDate(o.closeDate)]),
    retention,
    pilots,
    stages: stageRows(m.pipelineByStage || []),
    stagesQ1: stageRows(f1.pipelineByStage),
    stagesQ2: stageRows(f2.pipelineByStage),
    topDealsQ1: topDeals(f1),
    topDealsQ2: topDeals(f2),
    renewalsAtRisk: atRisk(f1).concat(atRisk(f2)),
    reps: view.reps.map(r => [deckLink(r.name, r.url), fmtNum(r.monthsInSeat), fmtInt(r.accountsOwned), fmtMoney(r.repGoal), fmtMoney(r.qWonArr),
      fmtMoney(r.fyWonArr), fmtPct(r.attainmentPct), fmtInt(r.meetings), fmtInt(r.activities), fmtPct(r.activityCoveragePct),
      fmtMoney(r.pipelineCreatedArr), fmtMoney(r.stalledArr), null]),
    dataQuality: view.dataQuality.map(i => (i.opp
      ? [i.issue, acct(i.opp), i.detail, fmtDate(i.opp.closeDate), i.opp.owner || '-', null]
      : [i.issue, deckLink(i.account.name, i.account.url), i.detail, '-', i.account.owner || '-', null])),
    lostPipeline: m.lists.lostPipeline.map(o => [acct(o), o.type || '-', fmtDate(o.closeDate), fmtMoney(o.deltaArr), o.owner || '-', o.lostReason || null]),
    topCustomers: m.topMajors.map((a, i) => [fmtInt(i + 1), deckLink(a.name, a.url), fmtMoney(a.currentArr), fmtPct(ratio(a.currentArr, m.activeArr)), a.hasOpenOpp ? 'Yes' : 'No']),
  };
}

// ---------------------------------------------------------------- deck links on the tab
// One property per team / quarter, so concurrent builds never overwrite each other's link.
function rememberDeckLink(team, quarter, url, built) {
  PropertiesService.getDocumentProperties().setProperty(`${DECK_LINKS_KEY}:${team}|${quarter}`, JSON.stringify({ url, built }));
}

function deckLinkFor(team, quarter) {
  const raw = PropertiesService.getDocumentProperties().getProperty(`${DECK_LINKS_KEY}:${team}|${quarter}`);
  return raw ? JSON.parse(raw) : null;
}

// M1:N1, next to the PDF link in M2:N2 (resetSheet clears the tab, so renderTeamTab calls this on every refresh).
function writeDeckLink(sheet, team, quarter) {
  const entry = deckLinkFor(team, quarter);
  if (!entry) return;
  sheet.getRange(1, LAST_COL - 1, 1, 2).merge().setHorizontalAlignment('right').setFontSize(8)
    .setRichTextValue(SpreadsheetApp.newRichTextValue().setText(`Open QBR deck (built ${entry.built})`).setLinkUrl(entry.url).build());
}
