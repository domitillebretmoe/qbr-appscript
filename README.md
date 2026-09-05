# qbr-appscript

Google Apps Script that fills the QBR workbook straight from Salesforce: one tab per team
(team in `B1`, quarter in `B2`), an `ARR Ledger` tab as the source of truth for Starting / Ending ARR,
a `Definitions` tab that spells out every calculation, and charts for the numbers that go on slides.

## Layout of a team tab

| Block | Contents |
| --- | --- |
| Previous quarter | Revenue goal, Net Added ARR, attainment, logo goal/attainment, Ending ARR, renewals, renewal rate, churn $ / #, top 3 churns, churn reasons; Starting → Added → Downgrade → Full churn → Ending ARR; active / major / enterprise customers, activated prospects, conversion, lost pipeline; logos won / lost |
| Future quarters (Q+1, Q+2) | Revenue goal, net forecast $ / %, logo goal / forecast, pipeline, coverage; forecast ARR / churn / ending ARR |
| Partner contribution | Net Added ARR, new logos, churn $ / # for the team's opportunities in the Partnerships group |
| Charts | ARR bridge (waterfall), attainment vs goal, renewal-rate donut, Q+1/Q+2 forecast vs goal, Net Added / Ending ARR trend since Q1-2026 |
| Column T onward | Per-quarter data behind the sparklines and charts |

Each metric row has a QoQ delta (vs the previous quarter) and a sparkline from Q1-2026 to the selected quarter.

## Definitions

Salesforce fields: Delta ARR = `NACV__c`, Expected Delta ARR = `Expected_NACV__c`, logo impact = `Expected_Logo_Impact__c`,
quarter = `FiscalYear` / `FiscalQuarter` (FY26 Q1 starts 1 Feb 2026), Major = `Account.Major_Admin_Tag__c`.

- **Team match**: the text after the last ` - ` in `B1` (whole text if there is none) must appear in the team name
  (and so must the words before it), so `Europe - DACH` covers `Europe - DACH` and `Europe Majors - DACH`; `Japan` covers
  `Japan` and `Japan Majors`; `Europe - South` does not pick up `LATAM - South`. The team of an opportunity is the Account's Team
  (`Account.Team__c`); when that is a region (`Europe`, `Asia`, `US Majors`) or blank, the Account Subteam and then the
  Opportunity Team are used instead, so a row is always attributed to the most detailed team and never to `Europe`.
  A region name in `B1` is rejected.
- **Exclusions**: accounts named `Test`; opportunities owned by Christian Lawless (`EXCLUDED_OWNERS`, as in the
  FY26 QBR COCKPIT report).
- **Renewal** = Opportunity record type `Renewal` or `Fed - Renewal`.
- **Net Added ARR** = Σ Delta ARR of Closed Won + Σ Delta ARR of Closed Lost renewals.
- **Downgrade** = Closed Won renewal with Delta ARR < 0. **Full churn** = Closed Lost renewal. **Churn ARR** = both.
  **Added ARR** = Net Added − Churn ARR, so Starting + Added + Downgrade + Full churn = Ending.
- **Starting ARR** = previous quarter's Ending ARR in the `ARR Ledger`; **Ending ARR** = Starting + Net Added.
  Q1-2026 Starting ARR (= Q4-2025 Ending) and the Q1/Q2-2026 Ending ARR are seeded from the FY27 QBR Cockpit workbook
  (`TEAM_SEEDS`) and never overwritten. Rows whose `Source` starts with `Salesforce` are recomputed on every refresh;
  type anything else in `Source` (e.g. `Locked`) to freeze a row.
- **Partner contribution** = the same metrics over the team's opportunities whose `Group__c` is `Partnerships`
  (the FY26 QBR COCKPIT - Partners report).
- **Renewal rate** = won renewals / (won + lost renewals). **Logo attainment** = Σ Expected Logo Impact on Major accounts.
- **Goals** come from `Goal__c` (`Net ARR` → revenue goal, `New Logos` → logo goal), summed over every Salesforce team matching the token.
- **Net Forecast** = Σ Expected Delta ARR of Land + Expand + positive Renewal, plus forecast churn (negative Expected Delta ARR on renewals).
  **Pipeline** = Σ Delta ARR of open opportunities; **coverage** = pipeline / revenue goal.
- **Active customers** = accounts with Current ARR > 0 (split by Major tag). **Activated prospects** = accounts with no ARR and an open opportunity.
  **Conversion** = Closed Won Lands / (Closed Won Lands + activated prospects). Account counts are as of the refresh.
- **Lost pipeline** = Closed Lost non-renewal opportunities in the quarter (count, Σ Delta ARR).

## Setup

1. **Salesforce Connected App** (Setup → App Manager → New Connected App): enable OAuth, scopes
   `Manage user data via APIs (api)`, tick *Enable Client Credentials Flow*, save, then under *Manage → Edit Policies*
   set the *Run As* user (a user who can read Opportunity, Account and Goal). Copy the consumer key and secret.
2. **Apps Script project**: in the Google Sheet, Extensions → Apps Script, then either paste the files from `src/`
   or use clasp:
   ```sh
   npm i -g @google/clasp && clasp login
   cp .clasp.json.example .clasp.json   # put your scriptId in it
   clasp push
   ```
3. Reload the sheet. Menu **QBR → Set Salesforce credentials…** and enter the My Domain URL
   (`https://<domain>.my.salesforce.com`), consumer key and secret. They live in Script Properties, not in the sheet.
4. **QBR → Set up workbook** creates the `Definitions` tab, the `ARR Ledger` (seeded) and one tab per team, then refreshes
   them all (a run that would exceed Apps Script's 6-minute limit continues by itself in the background).
   **QBR → Refresh on B1/B2 edit** installs the trigger so changing the team or quarter on a tab re-populates it.

To add a team later, use **QBR → Add team tab…** and add its `Q1-2026` Starting ARR row to the `ARR Ledger` first.

## Tests

```sh
npm test
```

Runs the pure calculation files (`Config.gs`, `Metrics.gs`, `Definitions.gs`) against last quarter's Europe - DACH opportunities.

## Reconciliation against the FY27 QBR Cockpit workbook

Replaying the exact queries in `Salesforce.gs` against the org reproduces the workbook's Q1-2026 and Q2-2026 Net Added ARR
to the dollar for 12 of the 19 teams. Every remaining difference has one of these causes (none is a calculation difference):

- accounts re-teamed in Salesforce since the workbook snapshot (KLA: Technology → US Enterprise; WPP: Partnerships → Europe
  Majors - UKI; ABLY: US Enterprise → APAC; Accelya: UKI → US Enterprise - East) and opportunities closed since (SAP NS2);
- the workbook counted the `Test` account (DACH Q1, −$15,000) and only the `Renewal` record type (Federal, `Fed - Renewal` −$159,000);
- tab formulas that disagree with each other: Technology / Partnerships / Federal add downgrades back, Europe - South Q2
  only counts Land + Expand, US Enterprise is typed in by hand.

Because the team of a row is its account's *current* team, historical quarters move with the account; the seeded
Starting / Ending ARR in the ledger do not.
