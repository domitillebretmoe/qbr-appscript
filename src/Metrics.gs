// Pure metric calculations. Opportunities are the plain objects built in fetchOpportunities.
const RENEWAL_RECORD_TYPES = ['Renewal', 'Fed - Renewal'];

const sum = (rows, field) => rows.reduce((total, row) => total + row[field], 0);
const unique = values => values.filter((v, i) => v && values.indexOf(v) === i);
const ratio = (num, den) => (den ? num / den : null);
const isRenewal = opp => RENEWAL_RECORD_TYPES.indexOf(opp.recordType) >= 0;
const isWon = opp => opp.stage === 'Closed Won';
const isLost = opp => opp.stage === 'Closed Lost';
const inQuarter = (opps, quarter) => opps.filter(o => o.quarter === quarter);

// Closed-quarter actuals. Net Added = Σ Closed Won Delta ARR + Σ Closed Lost renewal Delta ARR.
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

  return {
    quarter,
    revenueGoal: goal.revenue,
    netAddedArr,
    attainment: ratio(netAddedArr, goal.revenue),
    logoGoal: goal.logos,
    logoAttainment,
    logoAttainmentPct: ratio(Math.max(0, logoAttainment), goal.logos),
    addedArr: netAddedArr - churnArr,
    downgradeArr,
    downgradeCount: downgrades.length,
    fullChurnArr,
    fullChurnCount: fullChurn.length,
    churnArr,
    churnCustomers: fullChurn.length,
    renewals: wonRenewals.length + fullChurn.length,
    wonRenewals: wonRenewals.length,
    renewalRate: ratio(wonRenewals.length, wonRenewals.length + fullChurn.length),
    lostPipelineCount: lostPipeline.length,
    lostPipelineArr: sum(lostPipeline, 'deltaArr'),
    topChurns: churnOpps.slice(0, 3).map(o => `${o.account} (${formatMoney(o.deltaArr)})`),
    churnReasons: unique(churnOpps.map(o => o.lostReason)),
    logosWon: unique(won.filter(o => o.type === 'Land').map(o => o.account)),
    logosLost: unique(fullChurn.map(o => o.account)),
    newLogos: won.filter(o => o.type === 'Land').length,
  };
}

// Forward-looking quarter. Net Forecast = Expected Delta ARR of Land + Expand + renewals expected to grow,
// plus forecast churn (renewals expected to shrink). Renewal = record type, as for actuals.
function forecastMetrics(opps, quarter, startingArr, goal) {
  const rows = inQuarter(opps, quarter);
  const landExpand = rows.filter(o => (o.type === 'Land' || o.type === 'Expand') && !isRenewal(o));
  const renewalUp = rows.filter(o => isRenewal(o) && o.expectedDeltaArr > 0);
  const churn = rows.filter(o => isRenewal(o) && o.expectedDeltaArr < 0);
  const forecastArr = sum(landExpand, 'expectedDeltaArr') + sum(renewalUp, 'expectedDeltaArr');
  const forecastChurnArr = sum(churn, 'expectedDeltaArr');
  const netForecastArr = forecastArr + forecastChurnArr;
  const pipelineArr = sum(rows.filter(o => !o.isClosed), 'deltaArr');

  return {
    quarter,
    revenueGoal: goal.revenue,
    netForecastArr,
    netForecastPct: ratio(netForecastArr, goal.revenue),
    logoGoal: goal.logos,
    logoForecast: sum(rows.filter(o => o.major), 'expectedLogoImpact'),
    pipelineArr,
    pipelineCoverage: ratio(pipelineArr, goal.revenue),
    startingArr,
    forecastArr,
    forecastChurnArr,
    forecastChurnCount: churn.filter(o => o.expectedLogoImpact < 0).length,
    forecastEndingArr: startingArr + netForecastArr,
  };
}

// Account counts are as of the refresh (Salesforce has no per-quarter account history).
function accountMetrics(accounts, opps, quarter) {
  const active = accounts.filter(a => a.currentArr > 0);
  const activated = accounts.filter(a => a.currentArr <= 0 && a.hasOpenOpp);
  const wonLands = inQuarter(opps, quarter).filter(o => isWon(o) && o.type === 'Land');
  return {
    activeCustomers: active.length,
    majorCustomers: active.filter(a => a.major).length,
    enterpriseCustomers: active.filter(a => !a.major).length,
    activatedProspects: activated.length,
    conversionRate: ratio(wonLands.length, wonLands.length + activated.length),
  };
}

// Partner contribution = the team's opportunities in the Partnerships group (the "FY26 QBR COCKPIT - Partners"
// report filters Opportunity.Group__c = Partnerships).
function partnerMetrics(opps, quarter) {
  const m = quarterMetrics(opps.filter(o => o.oppGroup === PARTNER_GROUP), quarter, {});
  return { partnerNetAddedArr: m.netAddedArr, partnerNewLogos: m.newLogos, partnerChurnArr: m.churnArr, partnerChurnCustomers: m.churnCustomers };
}

// Starting/Ending ARR come from the ARR Ledger (seeded quarters keep last workbook's values).
function withArr(metrics, ledgerEntry) {
  return Object.assign(metrics, { startingArr: ledgerEntry.startingArr, endingArr: ledgerEntry.endingArr });
}

function formatMoney(value) {
  const abs = Math.abs(value);
  const text = abs >= 1e6 ? `$${(abs / 1e6).toFixed(2)}M` : abs >= 1e3 ? `$${(abs / 1e3).toFixed(0)}K` : `$${abs.toFixed(0)}`;
  return value < 0 ? `-${text}` : text;
}
