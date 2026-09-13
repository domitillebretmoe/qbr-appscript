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
    ['Scope', 'Roll-up tab', `${Object.keys(ROLLUP_TEAMS).map(t => `"${t}" = ${ROLLUP_TEAMS[t].join(' + ')}`).join('; ')}. Opportunities stay attributed to their specific team; the roll-up sums the members' opportunities, accounts, goals and ARR Ledger rows.`, '-'],
    ['Scope', 'Links', 'Every account / opportunity name in the tables links to its Salesforce record (Lightning URL). Raw Data has the URLs as columns.', 'Opportunity.Id, AccountId'],

    ['Scope', 'Formula cells',
      'Dollar and count metrics in the metric blocks and the KPI cards are live formulas: SUMIFS / COUNTIFS over Raw Data keyed on B1 (Tab) and the quarter (B2, or the Q+1 / Q+2 label), Revenue / Logo Goal from the Goals tab, Starting / Ending ARR from the ARR Ledger, ratios from the cells next to them. Click a cell to see exactly which rows it adds up. Account counts, quarter elapsed, text rows, tables and charts are values written on refresh.',
      'Raw Data, Goals and ARR Ledger tabs'],
    ['Scope', 'Forecast vs actuals', 'The banner and the quarter block are labelled FORECAST while the selected quarter is in progress (its numbers are closed deals to date plus, where stated, open opportunities) and ACTUALS once the quarter has ended.', 'Refresh date vs quarter end'],

    ['Quarter actuals', 'Revenue Goal', 'Sum of Goal Value of "Net ARR" goals whose Period Start falls in the quarter, over every Salesforce team matching B1.', 'Goal__c: Goal_Type__c = "Net ARR", Value__c, Period_Start__c, Team2__c'],
    ['Quarter actuals', 'Net Added ARR', 'Sum of Delta ARR of Closed Won opportunities + Sum of Delta ARR of Closed Lost renewals. Delta ARR of a Closed Won "MSP" deal whose NACV is 0 = its ARR (Salesforce ARR field, else Amount), since MSP deals are booked without NACV.', 'Opportunity.NACV__c, ARR__c, Amount, Type, StageName'],
    ['Quarter actuals', 'Attainment (%)', 'Net Added ARR / Revenue Goal. Attainment-style percentages are coloured green >= 100%, amber 70-99%, red < 70%.', '-'],
    ['Quarter actuals', 'Logo Goal', 'Sum of Count Value of "New Logos" goals whose Period Start falls in the quarter.', 'Goal__c: Goal_Type__c = "New Logos", Count_Value__c'],
    ['Quarter actuals', 'Logo attainment (expected, incl. open opps)', 'Sum of Expected Logo Impact of all the quarter\'s opportunities on Major accounts, open ones included (a churned logo counts -1). Until the quarter closes this is a forecast, not landed logos.', 'Opportunity.Expected_Logo_Impact__c, Account.Major_Admin_Tag__c'],
    ['Quarter actuals', 'Logos Won (Land only, Majors)', 'Number of Closed Won opportunities of Type "Land" on Major accounts: logos actually landed in the quarter, MSP deals excluded.', 'Opportunity.Type, StageName, Account.Major_Admin_Tag__c'],
    ['Quarter actuals', 'Logos Won (Land + MSP, Majors)', 'Same, counting Closed Won "MSP" opportunities as well (an MSP deal on a new account lands the logo like a Land).', 'Opportunity.Type, StageName, Account.Major_Admin_Tag__c'],
    ['Quarter actuals', 'New logos', 'Number of Closed Won opportunities of Type "Land" or "MSP" (an MSP deal on a new account lands the logo like a Land).', 'Opportunity.Type'],
    ['Quarter actuals', 'Downgrade ARR / #', 'Closed Won renewals with Delta ARR < 0: Sum of Delta ARR and count.', 'Opportunity.NACV__c, RecordType'],
    ['Quarter actuals', 'Full churn ARR / #', 'Closed Lost renewals: Sum of Delta ARR and count.', 'Opportunity.NACV__c, RecordType'],
    ['Quarter actuals', 'Churn ARR', 'Downgrade ARR + Full churn ARR.', '-'],
    ['Quarter actuals', 'Churn - Customer #', 'Number of Closed Lost renewals (full churns).', '-'],
    ['Quarter actuals', 'Added ARR', 'Net Added ARR - Churn ARR (gross new + expansion + renewal upsell).', '-'],
    ['Quarter actuals', 'Starting ARR', 'Previous quarter\'s Ending ARR from the ARR Ledger. Q1-2026 = Q4-2025 Ending ARR from the FY27 QBR Cockpit workbook.', 'ARR Ledger tab'],
    ['Quarter actuals', 'Ending ARR', 'Starting ARR + Net Added ARR. Q1-2026 and Q2-2026 keep the workbook values (Source != "Salesforce ..." in the ledger); later quarters are recomputed on refresh.', 'ARR Ledger tab'],
    ['Quarter actuals', 'GRR (%)', 'Gross revenue retention = (Starting ARR + Downgrade $ + Full Churn $) / Starting ARR. Only existing-customer losses; new logos and expansion are ignored.', 'ARR Ledger tab, Opportunity.NACV__c'],
    ['Quarter actuals', 'NRR (%)', 'Net revenue retention = (Starting ARR + Downgrade $ + Full Churn $ + expansion) / Starting ARR, where expansion = Delta ARR of Closed Won opportunities that are not "Land" / "MSP" (Expand + renewal upsell).', 'ARR Ledger tab, Opportunity.NACV__c, Type'],
    ['Quarter actuals', 'Quarter elapsed (%)', 'Calendar days from the first day of the fiscal quarter to the refresh date / days in the quarter. 0% before the quarter starts, 100% once it is over.', 'Refresh date'],
    ['Quarter actuals', 'Pace (attainment / elapsed)', 'Attainment (%) / Quarter elapsed (%). 100% = on a straight-line path to the goal; above = ahead of linear pace, below = behind. Red / amber / green like attainment.', '-'],
    ['Quarter actuals', 'Open pipeline (this quarter)', 'Sum of Delta ARR of open (not closed) opportunities with a close date in the selected quarter.', 'Opportunity.NACV__c, IsClosed'],
    ['Quarter actuals', 'Pipeline coverage of remaining goal', 'Open pipeline / (Revenue Goal - Net Added ARR), shown as a multiple (e.g. 1.5x). Blank once the goal is already met.', '-'],
    ['Quarter actuals', 'QoQ', 'Metric this quarter - same metric the previous quarter.', '-'],
    ['Quarter actuals', 'Trend', `Same metric for every quarter from ${FIRST_QUARTER} to the selected quarter; the sparkline column appears once at least four quarters exist.`, '-'],
    ['Quarter actuals', '# Renewals / # Won renewals', 'Closed renewals (won + lost) / Closed Won renewals.', 'Opportunity.RecordType, StageName'],
    ['Quarter actuals', 'Renewal rate (%)', 'Won renewals / (won renewals + lost renewals).', '-'],
    ['Quarter actuals', 'Top churns', 'Three downgrades / full churns with the most negative Delta ARR.', 'Opportunity.NACV__c'],
    ['Quarter actuals', 'Churn reasons', 'Distinct Closed Lost reasons of the downgrades / full churns.', 'Opportunity.Closed_Lost_Reason_List__c, Closed_Lost_Reason__c'],
    ['Quarter actuals', 'Top 10 Deals Won (table)', 'The quarter\'s ten largest Closed Won opportunities (any type) by Delta ARR, with the booked Deal value (Salesforce Amount / TCV) next to Delta ARR - renewals carry a value but 0 Delta ARR; MSP deals show their ARR as Delta ARR. The title shows the Closed Won count and Sum of Delta ARR that feed Net Added ARR; with more than ten wins it reads "10 of N Closed Won shown: $shown of $total".', 'Opportunity.Amount, NACV__c, StageName'],
    ['Quarter actuals', 'Logos Won (Land + MSP) (table)', 'Closed Won "Land" / "MSP" opportunities of the quarter, largest Deal value (Salesforce Amount / TCV) first, with Delta ARR alongside.', 'Opportunity.Type, StageName, Amount, NACV__c'],
    ['Quarter actuals', 'Lost Pipeline (table) / Lost pipeline # / $', 'Closed Lost non-renewal opportunities (Land, Expand...): list, count and Sum of Delta ARR. This is lost pipeline, not churn.', 'Opportunity.NACV__c, RecordType'],
    ['Quarter actuals', 'Churned Customers (table)', 'Closed Lost renewals of the quarter (the full churns), most negative Delta ARR first. Same rows as Renewals Lost.', 'Opportunity.RecordType, StageName'],
    ['Quarter actuals', 'Downgrade Customers (table)', 'Closed Won renewals with Delta ARR < 0, most negative first.', 'Opportunity.RecordType, NACV__c'],
    ['Quarter actuals', 'Renewals Won / Renewals Lost (tables)', 'Closed Won renewals (incl. downgrades) / Closed Lost renewals of the quarter: the numerator and the rest of the denominator of the renewal rate.', 'Opportunity.RecordType, StageName'],

    ['Accounts (as of refresh)', 'Active customers', 'Accounts of the team with Current ARR > 0; Major / Enterprise split by the Major tag.', 'Account.Current_ARR__c, Major_Admin_Tag__c'],
    ['Accounts (as of refresh)', 'Activated prospects', 'Accounts of the team with Current ARR <= 0 and at least one open opportunity (any close date, excluded owners ignored).', 'Account.Current_ARR__c, Opportunity.IsClosed'],
    ['Accounts (as of refresh)', 'Conversion (%)', 'Closed Won "Land" / "MSP" opportunities in the quarter / (those + activated prospects).', '-'],
    ['Accounts (as of refresh)', 'Top 10 Major / Enterprise Customers (tables)', 'Active customers (Current ARR > 0) with / without the Major tag, ten largest by Current ARR; % = share of the team\'s total active Current ARR.', 'Account.Current_ARR__c, Major_Admin_Tag__c'],

    ['Future quarters (Q+1, Q+2)', 'Net Forecast', 'Sum of Expected Delta ARR of Land / MSP + Expand (non-renewal) opportunities + renewals with Expected Delta ARR > 0 + Forecast churn.', 'Opportunity.Expected_NACV__c, Type, RecordType'],
    ['Future quarters (Q+1, Q+2)', 'Forecast churn ARR / #', 'Renewals with Expected Delta ARR < 0: Sum of Expected Delta ARR; count = those of them with Expected Logo Impact < 0 (expected full churns).', 'Opportunity.Expected_NACV__c, Expected_Logo_Impact__c'],
    ['Future quarters (Q+1, Q+2)', 'Net Forecast (%)', 'Net Forecast / Revenue Goal.', '-'],
    ['Future quarters (Q+1, Q+2)', 'Logo forecast', 'Sum of Expected Logo Impact of the quarter\'s opportunities on Major accounts.', 'Opportunity.Expected_Logo_Impact__c'],
    ['Future quarters (Q+1, Q+2)', 'Pipeline', 'Sum of Delta ARR of open (not closed) opportunities in the quarter.', 'Opportunity.NACV__c, IsClosed'],
    ['Future quarters (Q+1, Q+2)', 'Pipeline coverage of goal', 'Pipeline / Revenue Goal, shown as a multiple (e.g. 2.0x).', '-'],
    ['Future quarters (Q+1, Q+2)', '# Renewals due / ARR up for renewal', 'Open renewal opportunities (record type Renewal or Fed - Renewal) closing in the quarter; ARR up for renewal = Current ARR of their accounts, each account counted once.', 'Opportunity.RecordType, IsClosed, CloseDate; Account.Current_ARR__c'],
    ['Future quarters (Q+1, Q+2)', 'Renewals due (tables)', 'The open renewals above, largest account Current ARR first, with the Expected Delta ARR entered on the opportunity.', 'Opportunity.Expected_NACV__c, Account.Current_ARR__c'],
    ['Future quarters (Q+1, Q+2)', 'Predicted Churn (tables)', 'Renewals closing in the quarter whose Expected Delta ARR is negative (open, or already Closed Lost), most negative first, with the account\'s Current ARR. Sum = Forecast Churn ARR; the count in the title includes expected downgrades, whereas Forecast Churn # counts expected full churns only. A renewal with no Expected Delta ARR entered is not predicted churn: see Renewals due.', 'Opportunity.RecordType, Expected_NACV__c, CloseDate; Account.Current_ARR__c'],
    ['Future quarters (Q+1, Q+2)', 'Forecast Ending ARR', 'Starting ARR (previous quarter\'s Ending / Forecast Ending ARR) + Net Forecast.', '-'],
    ['Future quarters (Q+1, Q+2)', 'Top 10 Deals (tables)', 'Open opportunities closing in the quarter, ten largest by Delta ARR, with stage, close date and owner.', 'Opportunity.IsClosed, NACV__c, StageName, CloseDate'],

    ['Partner contribution', 'Partner Net Added ARR / new logos / churn $ / churn #', `Same definitions as the quarter actuals, restricted to the team's opportunities in the "${PARTNER_GROUP}" group (as in the FY26 QBR COCKPIT - Partners report).`, 'Opportunity.Group__c'],

    ['Rep activity & performance', 'Reps (table)', 'Active users with a User Segment in the US Majors / Europe / Asia / LATAM groups whose Team is the tab\'s team (all member teams for a roll-up); Deployed Engineering, SDR and Pre-sales families excluded. Same measures as Salesforce > Global GTM Dashboard > Majors Rep Performance, one row per rep, best Won ARR first. Refreshed with the tab.', 'User_Segment__c (User__c, Team__c, Group__c), User.User_Family__c'],
    ['Rep activity & performance', 'Months in seat', 'Months since the Salesforce user was created.', 'User.CreatedDate'],
    ['Rep activity & performance', 'Accts owned', 'Accounts whose Owner is the rep.', 'Account.OwnerId'],
    ['Rep activity & performance', 'Rep goal (FY)', 'Sum of the Weighted Goal of the Account Goals assigned to the rep (Account Goal owner, else the account\'s owner).', 'Account_Goal__c.Weighted_Goal__c, Account_Owner__c'],
    ['Rep activity & performance', 'Won ARR (FY) / Attainment', 'Sum of Delta ARR (NACV) of Closed Won opportunities owned by the rep closing in the fiscal year of the selected quarter (Feb-Jan). Attainment = Won ARR / Rep goal.', 'Opportunity.OwnerId, IsWon, NACV__c, CloseDate'],
    ['Rep activity & performance', 'Coverage', 'Sum of Expected Delta ARR of the rep\'s open + Closed Won opportunities in the fiscal year / Rep goal.', 'Opportunity.Expected_NACV__c, IsClosed, IsWon'],
    ['Rep activity & performance', 'Meetings (30d)', 'Distinct Gong-synced meetings (Events) owned by the rep in the last 30 days, cancelled meetings excluded; an inactive duplicate user with the same name and email is credited to the rep.', 'Event.Gong__Gong_Activity_Id__c, Gong__Meeting_Prospect_Canceled__c, ActivityDate'],
    ['Rep activity & performance', 'Activities (30d)', 'Distinct Gong-synced Tasks (calls, emails) owned by the rep in the last 30 days.', 'Task.Gong__Gong_Activity_Id__c, ActivityDate'],
    ['Rep activity & performance', 'Acct coverage (30d)', 'Share of the rep\'s owned accounts with at least one Gong-synced meeting or activity by the rep in the last 30 days.', 'Event / Task.AccountId, Account.OwnerId'],
    ['Rep activity & performance', 'Pipeline created (Q)', 'Sum of Expected Delta ARR of opportunities owned by the rep created during the selected quarter (any stage).', 'Opportunity.CreatedDate, Expected_NACV__c'],
    ['Rep activity & performance', 'Stalled >60d', 'Sum of Expected Delta ARR of the rep\'s open opportunities whose stage has not changed for more than 60 days.', 'Opportunity.LastStageChangeDate, IsClosed'],
    ['Rep activity & performance', 'Renewal risk', 'ARR at risk / ARR up for renewal over the rep\'s open renewals in the fiscal year (Starting ARR > 0, not excluded from churn impact); ARR at risk = Starting ARR x (1 - Pwin).', 'Opportunity.Starting_ARR__c, Pwin__c, Exclude_From_Churn_Impact__c'],
    ['Rep activity & performance', 'Not reproduced from the dashboard', 'Slippage rate (needs the weekly pipeline snapshots) and meeting credit for reps who are participants but not the event owner.', '-'],

    ['Owners & data quality', 'Owner performance (table)', 'Per opportunity owner for the selected quarter: Net Added ARR, # Closed Won, Churn ARR (downgrades + full churn), open pipeline in the quarter and open pipeline in Q+1. Sorted by Net Added ARR.', 'Opportunity.OwnerId'],
    ['Owners & data quality', 'Open with close date in the past', 'Open opportunity in the selected quarter, Q+1 or Q+2 whose Close Date is before the refresh date (stale pipeline).', 'Opportunity.IsClosed, CloseDate'],
    ['Owners & data quality', 'Closed Won with $0 Delta ARR', 'Closed Won Land / Expand opportunity (not a renewal, not One Time, not an MSP deal with an ARR) with Delta ARR = 0: probably missing NACV.', 'Opportunity.NACV__c, Type, RecordType'],
    ['Owners & data quality', 'Open without Expected Delta ARR', 'Open opportunity whose Expected Delta ARR field is empty (it counts as 0 in the forecast).', 'Opportunity.Expected_NACV__c'],
    ['Owners & data quality', 'Account team is a region only', 'Account whose Team is the bare region (e.g. "Europe") with no Subteam, so it lands in no sub-team tab. Listed on every tab of that region when it has ARR or an open opportunity.', 'Account.Team__c, Subteam__c'],
    ['Owners & data quality', 'Hover notes', 'Every metric label on a team tab carries a note with its one-line definition (hover the cell).', '-'],
    ['Owners & data quality', 'Download this tab as PDF', 'Link in the tab header (M2) that exports the dashboard area of the current tab as a landscape PDF through Google Sheets\' own export, using your Google session; nothing is written to Drive.', '-'],

    ['ARR Ledger', 'Rows', 'One row per team and quarter: Starting ARR, Net Added ARR, Ending ARR, Source. Starting = previous quarter\'s Ending; Ending = Starting + Net Added.', 'ARR Ledger tab'],
    ['ARR Ledger', 'Source', `"${SEED_SOURCE}" rows (Q1-2026, Q2-2026) are fixed values from last quarter's workbook. Rows stamped "Salesforce <date>" are recomputed on every refresh. Any other text (e.g. "Locked") freezes the row as typed.`, '-'],
    ['ARR Ledger', 'Roll-up tabs', 'A roll-up tab (e.g. Europe) has no ledger rows: its Starting / Ending ARR is the sum of its member teams\' rows for the quarter.', '-'],
    ['Goals', 'Rows', 'One row per tab and fiscal quarter with the Net ARR Goal and New Logo Goal summed over the Salesforce teams matching the tab (a roll-up tab sums its members). Refreshing a tab replaces that tab\'s rows.', 'Goal__c'],
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
