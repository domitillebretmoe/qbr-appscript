// Europe - DACH opportunities as returned by Salesforce (FY2026 Q1-Q3), shared by the tests and the preview.
'use strict';

const opp = (quarter, stage, account, type, recordType, deltaArr, logo, extra = {}) => Object.assign({
  quarter, stage, account, type, recordType, deltaArr, expectedLogoImpact: logo, isClosed: stage.startsWith('Closed'),
  expectedDeltaArr: 0, lostReason: '', accountId: account, major: false, oppTeam: 'Europe', oppGroup: 'Europe',
}, extra);
const dach = [
  opp('Q1-2026', 'Closed Won', 'CompuGroup', 'One Time', 'Enterprise', 0, 0),
  opp('Q1-2026', 'Closed Won', 'CompuGroup', 'Expand', 'Enterprise', 127401.2, 0),
  opp('Q1-2026', 'Closed Won', 'CompuGroup', 'One Time', 'Enterprise', 0, 0),
  opp('Q1-2026', 'Closed Lost', 'BMW Group', 'Land', 'Enterprise', 288000, 0, { major: true }),
  opp('Q1-2026', 'Closed Lost', 'SoftwareOne', 'Land', 'Enterprise', 144000, 0, { oppTeam: 'Partnerships', oppGroup: 'Partnerships' }),
  opp('Q1-2026', 'Closed Lost', 'Deutsche Telekom', 'Land', 'Enterprise', 876000, 0, { major: true }),
  opp('Q1-2026', 'Closed Lost', 'Roche', 'Land', 'Enterprise', 180000, 0, { major: true, oppTeam: 'Partnerships', oppGroup: 'Partnerships' }),
  opp('Q1-2026', 'Closed Lost', 'Helaba', 'Land', 'Enterprise', 252000, 0, { major: true }),
  opp('Q1-2026', 'Closed Lost', 'Siemens', 'Land', 'Enterprise', 1, 0),
  opp('Q1-2026', 'Closed Lost', 'DKB', 'Land', 'Enterprise', 144000, 0),
  opp('Q1-2026', 'Closed Lost', 'Julius Baer', 'Land', 'Enterprise', 125000, 0),
  opp('Q1-2026', 'Closed Lost', 'Bolt', 'Land', 'Enterprise', 360000, 0),
  opp('Q1-2026', 'Closed Lost', 'SAP SE', 'Land', 'Enterprise', 131400, 0, { major: true }),
  opp('Q1-2026', 'Closed Lost', 'SAP SE', 'Land', 'Enterprise', 100000, 0, { major: true }),
  opp('Q2-2026', 'Closed Won', 'CompuGroup', 'One Time', 'Enterprise', 0, 0),
  opp('Q2-2026', 'Closed Won', 'Helaba', 'Land', 'Enterprise', 150000, 1, { major: true }),
  opp('Q2-2026', 'Closed Lost', 'CompuGroup', 'Expand', 'Enterprise', 2000, 0),
  opp('Q2-2026', 'Closed Lost', 'Deutsche Telekom', 'Renewal', 'Renewal', -216000, -1, { major: true, lostReason: 'Budget' }),
  opp('Q3-2026', 'R2- Renewal Engagement', 'Serrala', 'Renewal', 'Renewal', 0, 0, { expectedDeltaArr: -30000 }),
  opp('Q3-2026', '1- Discovery', 'Helaba', 'Expand', 'Enterprise', 400000, 0.5, { major: true, expectedDeltaArr: 100000 }),
];

module.exports = { opp, dach };
