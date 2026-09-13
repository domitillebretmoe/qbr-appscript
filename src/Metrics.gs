// Pure metric calculations. Opportunities are the plain objects built in fetchOpportunities.
const RENEWAL_RECORD_TYPES = ['Renewal', 'Fed - Renewal'];
// Opportunity types that land a new logo when Closed Won (MSP = managed service provider deal on a new account).
const NEW_LOGO_TYPES = ['Land', 'MSP'];

const sum = (rows, field) => rows.reduce((total, row) => total + row[field], 0);
const unique = values => values.filter((v, i) => v && values.indexOf(v) === i);
const ratio = (num, den) => (den ? num / den : null);
const isRenewal = opp => RENEWAL_RECORD_TYPES.indexOf(opp.recordType) >= 0;
const isWon = opp => opp.stage === 'Closed Won';
const isNewLogo = opp => NEW_LOGO_TYPES.indexOf(opp.type) >= 0;
const isLost = opp => opp.stage === 'Closed Lost';
const inQuarter = (opps, quarter) => opps.filter(o => o.quarter === quarter);
const byField = (rows, field, descending) => rows.slice().sort((a, b) => (descending ? b[field] - a[field] : a[field] - b[field]));
const TOP_N = 10;

// Closed-quarter actuals. Net Added = Sum of Closed Won Delta ARR + Sum of Closed Lost renewal Delta ARR.
function quarterMetrics(opps, quarter, goal) {
  const rows = inQuarter(opps, quarter);
  const won = rows.filter(isWon);
  const lost = rows.filter(isLost);
  const fullChurn = lost.filter(isRenewal);
  const downgrades = won.filter(o => isRenewal(o) && o.deltaArr < 0);
  const wonRenewals = won.filter(isRenewal);
  const lostPipeline = lost.filter(o => !isRenewal(o));
  const churnOpps = downgrades.concat(fullChurn).sort((a, b) => a.deltaArr - b.deltaArr);

  const netAddedArr = sum(won, 'deltaArr') + sum(fullChurn, 'deltaArr');
  const downgradeArr = sum(downgrades, 'deltaArr');
  const fullChurnArr = sum(fullChurn, 'deltaArr');
  const churnArr = downgradeArr + fullChurnArr;
  const logoAttainment = sum(rows.filter(o => o.major), 'expectedLogoImpact');
  const logosWonLand = won.filter(o => o.type === 'Land' && o.major).length;
  const logosWonMajors = won.filter(o => isNewLogo(o) && o.major).length;
  const newLogoArr = sum(won.filter(isNewLogo), 'deltaArr');
  const open = rows.filter(o => !o.isClosed);
  const openPipelineArr = sum(open, 'deltaArr');
  const remainingGoal = Math.max(0, (goal.revenue || 0) - netAddedArr);

  return {
    quarter,
    revenueGoal: goal.revenue,
    netAddedArr,
    attainment: ratio(netAddedArr, goal.revenue),
    logoGoal: goal.logos,
    logoAttainment,
    logosWonLand,
    logosWonMajors,
    logoAttainmentPct: ratio(Math.max(0, logoAttainment), goal.logos),
    addedArr: netAddedArr - churnArr,
    wonCount: won.length,
    wonArr: sum(won, 'deltaArr'),
    downgradeArr,
    downgradeCount: downgrades.length,
    fullChurnArr,
    fullChurnCount: fullChurn.length,
    churnArr,
    newLogoArr,
    // Growth from existing customers: everything won that is not a new logo, net of downgrades.
    expansionArr: netAddedArr - churnArr - newLogoArr,
    openPipelineArr,
    remainingGoal,
    pipelineCoverage: remainingGoal ? openPipelineArr / remainingGoal : null,
    churnCustomers: fullChurn.length,
    renewals: wonRenewals.length + fullChurn.length,
    wonRenewals: wonRenewals.length,
    renewalRate: ratio(wonRenewals.length, wonRenewals.length + fullChurn.length),
    lostPipelineCount: lostPipeline.length,
    lostPipelineArr: sum(lostPipeline, 'deltaArr'),
    topChurns: churnOpps.slice(0, 3).map(o => `${o.account} (${formatMoney(o.deltaArr)})`),
    churnReasons: unique(churnOpps.map(o => o.lostReason)),
    logosWon: unique(won.filter(isNewLogo).map(o => o.account)),
    logosLost: unique(fullChurn.map(o => o.account)),
    newLogos: won.filter(isNewLogo).length,
    // Opportunity lists behind the linked tables on the tab.
    lists: {
      dealsWon: byField(won, 'deltaArr', true).slice(0, TOP_N),
      logosWon: byField(won.filter(isNewLogo), 'amount', true),
      lostPipeline: byField(lostPipeline, 'deltaArr', true),
      churned: byField(fullChurn, 'deltaArr', false),
      downgrades: byField(downgrades, 'deltaArr', false),
      renewalsWon: byField(wonRenewals, 'deltaArr', true),
      renewalsLost: byField(fullChurn, 'deltaArr', false),
    },
  };
}

// Forward-looking quarter. Net Forecast = Expected Delta ARR of Land / MSP + Expand + renewals expected to grow,
// plus forecast churn (renewals expected to shrink). Renewal = record type, as for actuals.
// `accounts` supplies the Current ARR that open renewals put up for renewal.
function forecastMetrics(opps, quarter, startingArr, goal, accounts) {
  const rows = inQuarter(opps, quarter);
  const arrOf = {};
  (accounts || []).forEach(a => { arrOf[a.id] = a.currentArr; });
  const withArr = list => list.map(o => Object.assign({}, o, { accountArr: arrOf[o.accountId] || 0 }));
  const renewalsDue = withArr(rows.filter(o => isRenewal(o) && !o.isClosed));
  const renewalArrDue = unique(renewalsDue.map(o => o.accountId)).reduce((total, id) => total + (arrOf[id] || 0), 0);
  const landExpand = rows.filter(o => (isNewLogo(o) || o.type === 'Expand') && !isRenewal(o));
  const renewalUp = rows.filter(o => isRenewal(o) && o.expectedDeltaArr > 0);
  const churn = withArr(rows.filter(o => isRenewal(o) && o.expectedDeltaArr < 0));
  const forecastArr = sum(landExpand, 'expectedDeltaArr') + sum(renewalUp, 'expectedDeltaArr');
  const forecastChurnArr = sum(churn, 'expectedDeltaArr');
  const netForecastArr = forecastArr + forecastChurnArr;
  const open = rows.filter(o => !o.isClosed);
  const pipelineArr = sum(open, 'deltaArr');

  return {
    quarter,
    revenueGoal: goal.revenue,
    netForecastArr,
    netForecastPct: ratio(netForecastArr, goal.revenue),
    logoGoal: goal.logos,
    logoForecast: sum(rows.filter(o => o.major), 'expectedLogoImpact'),
    pipelineArr,
    pipelineCoverage: ratio(pipelineArr, goal.revenue),
    renewalsDueCount: renewalsDue.length,
    renewalArrDue,
    renewalsDue: byField(renewalsDue, 'accountArr', true),
    startingArr,
    forecastArr,
    forecastChurnArr,
    forecastChurnCount: churn.filter(o => o.expectedLogoImpact < 0).length,
    predictedChurn: byField(churn, 'expectedDeltaArr', false),
    forecastEndingArr: startingArr + netForecastArr,
    topDeals: byField(open, 'deltaArr', true).slice(0, TOP_N),
  };
}

// Account counts are as of the refresh (Salesforce has no per-quarter account history).
function accountMetrics(accounts, opps, quarter) {
  const active = accounts.filter(a => a.currentArr > 0);
  const activated = accounts.filter(a => a.currentArr <= 0 && a.hasOpenOpp);
  const wonLands = inQuarter(opps, quarter).filter(o => isWon(o) && isNewLogo(o));
  return {
    activeCustomers: active.length,
    majorCustomers: active.filter(a => a.major).length,
    enterpriseCustomers: active.filter(a => !a.major).length,
    activatedProspects: activated.length,
    conversionRate: ratio(wonLands.length, wonLands.length + activated.length),
    activeArr: sum(active, 'currentArr'),
    topMajors: byField(active.filter(a => a.major), 'currentArr', true).slice(0, TOP_N),
    topEnterprise: byField(active.filter(a => !a.major), 'currentArr', true).slice(0, TOP_N),
  };
}

// Partner contribution = the team's opportunities in the Partnerships group (the "FY26 QBR COCKPIT - Partners"
// report filters Opportunity.Group__c = Partnerships).
function partnerMetrics(opps, quarter) {
  const m = quarterMetrics(opps.filter(o => o.oppGroup === PARTNER_GROUP), quarter, {});
  return { partnerNetAddedArr: m.netAddedArr, partnerNewLogos: m.newLogos, partnerChurnArr: m.churnArr, partnerChurnCustomers: m.churnCustomers };
}

// Starting/Ending ARR come from the ARR Ledger (seeded quarters keep last workbook's values). Dollar retention:
// GRR = (Starting ARR + downgrades + full churn) / Starting ARR, NRR = GRR + expansion from existing customers.
function withArr(metrics, ledgerEntry) {
  const { startingArr, endingArr } = ledgerEntry;
  return Object.assign(metrics, {
    startingArr,
    endingArr,
    grr: ratio(startingArr + metrics.churnArr, startingArr),
    nrr: ratio(startingArr + metrics.churnArr + metrics.expansionArr, startingArr),
  });
}

// Linearity: how far through the quarter we are vs how far through the goal. Past quarters are fully elapsed.
// A quarter's numbers are a forecast until its last day has passed, actuals afterwards; `phase` says whether the
// quarter has not started, is in progress or is closed.
function withPace(metrics, today) {
  const quarterElapsedPct = quarterElapsed(metrics.quarter, today);
  const phase = quarterElapsedPct >= 1 ? 'closed' : today < quarterStart(metrics.quarter) ? 'not started' : 'in progress';
  return Object.assign(metrics, {
    quarterElapsedPct,
    pace: metrics.attainment == null ? null : ratio(metrics.attainment, quarterElapsedPct),
    status: phase === 'closed' ? 'ACTUALS' : 'FORECAST',
    phase,
  });
}

function statusText(metrics) {
  if (metrics.status === 'ACTUALS') return 'ACTUALS (quarter closed)';
  if (metrics.phase === 'not started') return 'FORECAST (quarter not started)';
  return `FORECAST (quarter in progress, ${Math.round((metrics.quarterElapsedPct || 0) * 100)}% elapsed)`;
}

// One row per opportunity owner for the selected quarter, best Net Added ARR first. Owners with nothing to show
// (only Closed Lost pipeline) are skipped; `ownerTeams` (owner name -> Salesforce team) flags owners from another
// team than the tab's `members`.
function ownerMetrics(opps, quarter, nextQuarter, ownerTeams, members) {
  const rows = inQuarter(opps, quarter);
  const next = inQuarter(opps, nextQuarter).filter(o => !o.isClosed);
  const owners = unique(rows.concat(next).map(o => o.owner));
  const teams = ownerTeams || {};
  return owners.map(owner => {
    const mine = rows.filter(o => o.owner === owner);
    const won = mine.filter(isWon);
    const fullChurn = mine.filter(o => isLost(o) && isRenewal(o));
    const downgrades = won.filter(o => isRenewal(o) && o.deltaArr < 0);
    const team = teams[owner] || '';
    return {
      owner,
      team,
      own: !members || !team || members.some(m => teamMatches(team, m)),
      netAddedArr: sum(won, 'deltaArr') + sum(fullChurn, 'deltaArr'),
      wonCount: won.length,
      churnArr: sum(downgrades, 'deltaArr') + sum(fullChurn, 'deltaArr'),
      openPipelineArr: sum(mine.filter(o => !o.isClosed), 'deltaArr'),
      nextPipelineArr: sum(next.filter(o => o.owner === owner), 'deltaArr'),
    };
  }).filter(o => o.netAddedArr || o.wonCount || o.churnArr || o.openPipelineArr || o.nextPipelineArr)
    .sort((a, b) => b.netAddedArr - a.netAddedArr);
}

function ownerLabel(o) {
  return o.own ? o.owner || '-' : `${o.owner} (${o.team})`;
}

// Salesforce hygiene checks over the selected quarter and the two after it. Each issue: { issue, detail, opp | account }.
const DATA_QUALITY_CHECKS = {
  staleOpen: 'Open with close date in the past',
  wonZero: 'Closed Won with $0 Delta ARR',
  noExpected: 'Open without Expected Delta ARR',
  regionOnly: 'Account team is a region only',
};

function dataQualityIssues(opps, unassignedAccounts, quarters, today) {
  const rows = opps.filter(o => quarters.indexOf(o.quarter) >= 0);
  const issues = [];
  rows.filter(o => !o.isClosed && o.closeDate && o.closeDate < today)
    .forEach(o => issues.push({ issue: DATA_QUALITY_CHECKS.staleOpen, detail: o.stage, opp: o }));
  rows.filter(o => isWon(o) && !isRenewal(o) && o.type !== 'One Time' && o.deltaArr === 0)
    .forEach(o => issues.push({ issue: DATA_QUALITY_CHECKS.wonZero, detail: o.type || '', opp: o }));
  rows.filter(o => !o.isClosed && o.expectedDeltaArrMissing)
    .forEach(o => issues.push({ issue: DATA_QUALITY_CHECKS.noExpected, detail: `Delta ARR ${formatMoney(o.deltaArr)}`, opp: o }));
  (unassignedAccounts || []).forEach(a => issues.push({ issue: DATA_QUALITY_CHECKS.regionOnly, detail: `Team = ${a.team}, no sub-team`, account: a }));
  return issues;
}

// Everything a team tab shows, from already-fetched rows: `ledgers` holds one { quarter: { startingArr, endingArr } }
// map per member team (a roll-up sums them), `today` is yyyy-mm-dd.
function composeView({ team, members, quarter, opps, accounts, goals, ledgers, unassignedAccounts, reps, ownerTeams, today }) {
  const goalFor = q => goals[q] || { revenue: 0, logos: 0 };
  const next1 = shiftQuarter(quarter, 1);
  const next2 = shiftQuarter(quarter, 2);
  const quarters = quartersBetween(FIRST_QUARTER, quarter);
  const ledger = {};
  quarters.forEach(q => {
    ledger[q] = {
      startingArr: ledgers.reduce((total, l) => total + l[q].startingArr, 0),
      endingArr: ledgers.reduce((total, l) => total + l[q].endingArr, 0),
    };
  });
  const trend = quarters.map(q => withPace(withArr(
    Object.assign(quarterMetrics(opps, q, goalFor(q)), accountMetrics(accounts, opps, q), partnerMetrics(opps, q)), ledger[q]), today));
  const current = trend[trend.length - 1];
  const future1 = forecastMetrics(opps, next1, current.endingArr, goalFor(next1), accounts);
  const future2 = forecastMetrics(opps, next2, future1.forecastEndingArr, goalFor(next2), accounts);
  return {
    team, members, quarter, trend, current, future: [future1, future2], opps, goals, today,
    reps: reps || [],
    owners: ownerMetrics(opps, quarter, next1, ownerTeams, members),
    dataQuality: dataQualityIssues(opps, unassignedAccounts, [quarter, next1, next2], today),
  };
}

function formatMoney(value) {
  const abs = Math.abs(value);
  const text = abs >= 1e6 ? `$${(abs / 1e6).toFixed(2)}M` : abs >= 1e3 ? `$${(abs / 1e3).toFixed(0)}K` : `$${abs.toFixed(0)}`;
  return value < 0 ? `-${text}` : text;
}
