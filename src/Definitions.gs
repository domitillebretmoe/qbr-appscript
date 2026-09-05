// "Definitions" sheet: how every number on a team tab is calculated, and from which Salesforce fields.
// Rewritten on every "Set up workbook" so it always matches the code.
const DEFINITIONS_SHEET = 'Definitions';
const DEFINITIONS_HEADER = ['Section', 'Metric', 'Calculation', 'Salesforce source'];

function definitionRows() {
  const renewalTypes = RENEWAL_RECORD_TYPES.map(t => `"${t}"`).join(' or ');
  const owners = EXCLUDED_OWNERS.length ? EXCLUDED_OWNERS.join(', ') : 'none';
  return [
    ['Scope', 'Team (B1)',
      `Opportunities and accounts whose team contains the text after the last " - " of B1 (e.g. "DACH"); the leading words of B1 must appear too. Region names (${REGION_SEGMENTS.join(', ')}) are never accepted as a team.`,
      'Team = Account.Team__c, or Account.Subteam__c / Opportunity.Team__c when the account team is a region (e.g. "Europe") or blank'],
    ['Scope', 'Quarter (B2)', 'Fiscal quarter of the Close Date. Q1 = Feb-Apr, Q2 = May-Jul, Q3 = Aug-Oct, Q4 = Nov-Jan. "Q3-2026" = FY2026 Q3.', 'Opportunity.FiscalYear, Opportunity.FiscalQuarter'],
    ['Scope', 'Exclusions', `Accounts named "Test"; opportunities owned by ${owners} (as in the FY26 QBR COCKPIT report).`, 'Account.Name, Opportunity.Owner.Name'],
    ['Scope', 'Renewal', `Opportunity record type ${renewalTypes}.`, 'Opportunity.RecordType.Name'],
    ['Scope', 'Major account', 'Account with Major - Admin Tag = true.', 'Account.Major_Admin_Tag__c'],

    ['Quarter actuals', 'Revenue Goal', 'Sum of Goal Value of "Net ARR" goals whose Period Start falls in the quarter, over every Salesforce team matching B1.', 'Goal__c: Goal_Type__c = "Net ARR", Value__c, Period_Start__c, Team2__c'],
    ['Quarter actuals', 'Net Added ARR', 'Sum of Delta ARR of Closed Won opportunities + Sum of Delta ARR of Closed Lost renewals.', 'Opportunity.NACV__c, StageName'],
    ['Quarter actuals', 'Attainment (%)', 'Net Added ARR / Revenue Goal.', '-'],
    ['Quarter actuals', 'Logo Goal', 'Sum of Count Value of "New Logos" goals whose Period Start falls in the quarter.', 'Goal__c: Goal_Type__c = "New Logos", Count_Value__c'],
    ['Quarter actuals', 'Logo attainment', 'Sum of Expected Logo Impact of the quarter\'s opportunities on Major accounts (a churned logo counts -1).', 'Opportunity.Expected_Logo_Impact__c, Account.Major_Admin_Tag__c'],
    ['Quarter actuals', 'New logos', 'Number of Closed Won opportunities of Type "Land".', 'Opportunity.Type'],
    ['Quarter actuals', 'Downgrade ARR / #', 'Closed Won renewals with Delta ARR < 0: Sum of Delta ARR and count.', 'Opportunity.NACV__c, RecordType'],
    ['Quarter actuals', 'Full churn ARR / #', 'Closed Lost renewals: Sum of Delta ARR and count.', 'Opportunity.NACV__c, RecordType'],
    ['Quarter actuals', 'Churn ARR', 'Downgrade ARR + Full churn ARR.', '-'],
    ['Quarter actuals', 'Churn - Customer #', 'Number of Closed Lost renewals (full churns).', '-'],
    ['Quarter actuals', 'Added ARR', 'Net Added ARR - Churn ARR (gross new + expansion + renewal upsell).', '-'],
    ['Quarter actuals', 'Starting ARR', 'Previous quarter\'s Ending ARR from the ARR Ledger. Q1-2026 = Q4-2025 Ending ARR from the FY27 QBR Cockpit workbook.', 'ARR Ledger tab'],
    ['Quarter actuals', 'Ending ARR', 'Starting ARR + Net Added ARR. Q1-2026 and Q2-2026 keep the workbook values (Source != "Salesforce ..." in the ledger); later quarters are recomputed on refresh.', 'ARR Ledger tab'],
    ['Quarter actuals', 'QoQ', 'Metric this quarter - same metric the previous quarter.', '-'],
    ['Quarter actuals', 'Trend', `Same metric for every quarter from ${FIRST_QUARTER} to the selected quarter.`, '-'],
    ['Quarter actuals', '# Renewals / # Won renewals', 'Closed renewals (won + lost) / Closed Won renewals.', 'Opportunity.RecordType, StageName'],
    ['Quarter actuals', 'Renewal rate (%)', 'Won renewals / (won renewals + lost renewals).', '-'],
    ['Quarter actuals', 'Top churns', 'Three downgrades / full churns with the most negative Delta ARR.', 'Opportunity.NACV__c'],
    ['Quarter actuals', 'Churn reasons', 'Distinct Closed Lost reasons of the downgrades / full churns.', 'Opportunity.Closed_Lost_Reason_List__c, Closed_Lost_Reason__c'],
    ['Quarter actuals', 'Logos won / lost', 'Distinct accounts of Closed Won "Land" opportunities / of Closed Lost renewals.', 'Opportunity.Type, Account.Name'],
    ['Quarter actuals', 'Lost pipeline # / $', 'Closed Lost non-renewal opportunities: count and Sum of Delta ARR.', 'Opportunity.NACV__c'],

    ['Accounts (as of refresh)', 'Active customers', 'Accounts of the team with Current ARR > 0; Major / Enterprise split by the Major tag.', 'Account.Current_ARR__c, Major_Admin_Tag__c'],
    ['Accounts (as of refresh)', 'Activated prospects', 'Accounts of the team with Current ARR <= 0 and at least one open opportunity (any close date, excluded owners ignored).', 'Account.Current_ARR__c, Opportunity.IsClosed'],
    ['Accounts (as of refresh)', 'Conversion (%)', 'Closed Won "Land" opportunities in the quarter / (those + activated prospects).', '-'],

    ['Future quarters (Q+1, Q+2)', 'Net Forecast', 'Sum of Expected Delta ARR of Land + Expand (non-renewal) opportunities + renewals with Expected Delta ARR > 0 + Forecast churn.', 'Opportunity.Expected_NACV__c, Type, RecordType'],
    ['Future quarters (Q+1, Q+2)', 'Forecast churn ARR / #', 'Renewals with Expected Delta ARR < 0: Sum of Expected Delta ARR; count = those of them with Expected Logo Impact < 0 (expected full churns).', 'Opportunity.Expected_NACV__c, Expected_Logo_Impact__c'],
    ['Future quarters (Q+1, Q+2)', 'Net Forecast (%)', 'Net Forecast / Revenue Goal.', '-'],
    ['Future quarters (Q+1, Q+2)', 'Logo forecast', 'Sum of Expected Logo Impact of the quarter\'s opportunities on Major accounts.', 'Opportunity.Expected_Logo_Impact__c'],
    ['Future quarters (Q+1, Q+2)', 'Pipeline', 'Sum of Delta ARR of open (not closed) opportunities in the quarter.', 'Opportunity.NACV__c, IsClosed'],
    ['Future quarters (Q+1, Q+2)', 'Coverage', 'Pipeline / Revenue Goal.', '-'],
    ['Future quarters (Q+1, Q+2)', 'Forecast Ending ARR', 'Starting ARR (previous quarter\'s Ending / Forecast Ending ARR) + Net Forecast.', '-'],

    ['Partner contribution', 'Partner Net Added ARR / new logos / churn $ / churn #', `Same definitions as the quarter actuals, restricted to the team's opportunities in the "${PARTNER_GROUP}" group (as in the FY26 QBR COCKPIT - Partners report).`, 'Opportunity.Group__c'],

    ['ARR Ledger', 'Rows', 'One row per team and quarter: Starting ARR, Net Added ARR, Ending ARR, Source. Starting = previous quarter\'s Ending; Ending = Starting + Net Added.', 'ARR Ledger tab'],
    ['ARR Ledger', 'Source', `"${SEED_SOURCE}" rows (Q1-2026, Q2-2026) are fixed values from last quarter's workbook. Rows stamped "Salesforce <date>" are recomputed on every refresh. Any other text (e.g. "Locked") freezes the row as typed.`, '-'],
    ['Raw Data', 'Rows', 'Every Salesforce opportunity fetched for a team tab (all fiscal quarters from Q1-2026 to Q+2), one row each, with the Bucket it is counted in: Closed Won - Land/Expand, Won renewal, Won renewal - downgrade, Full churn, Lost pipeline, Open pipeline. Filter on Tab + Quarter to tie a metric out. Refreshing a tab replaces that tab\'s rows.', 'Raw Data tab'],
  ];
}

function writeDefinitionsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DEFINITIONS_SHEET) || ss.insertSheet(DEFINITIONS_SHEET);
  const rows = definitionRows();
  sheet.clear();
  sheet.getRange(1, 1, 1, DEFINITIONS_HEADER.length).setValues([DEFINITIONS_HEADER]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, DEFINITIONS_HEADER.length).setValues(rows).setVerticalAlignment('top').setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 620);
  sheet.setColumnWidth(4, 360);
  return sheet;
}
