// Europe - DACH opportunities and accounts as returned by Salesforce (FY2026 Q1-Q3), shared by the tests and the preview.
'use strict';

const SF = 'https://codeium.lightning.force.com';
const slug = s => s.replace(/[^A-Za-z0-9]/g, '').padEnd(15, '0').slice(0, 15);
const accountId = account => `001${slug(account)}`;
const accountUrl = account => `${SF}/lightning/r/Account/${accountId(account)}/view`;
let oppSeq = 0;

const opp = (quarter, stage, account, type, recordType, deltaArr, logo, extra = {}) => {
  const id = `006${String(++oppSeq).padStart(15, '0')}`;
  const o = Object.assign({
    quarter, stage, account, type, recordType, deltaArr, expectedLogoImpact: logo, isClosed: stage.startsWith('Closed'),
    expectedDeltaArr: 0, lostReason: '', major: false, oppTeam: 'Europe', oppGroup: 'Europe',
    id, url: `${SF}/lightning/r/Opportunity/${id}/view`, accountId: accountId(account), accountUrl: accountUrl(account),
    name: `${account} - ${type}`, closeDate: '', owner: 'Anna Berger', team: 'Europe - DACH',
  }, extra);
  if (!('ownerId' in o)) o.ownerId = `005${slug(o.owner)}`;
  return o;
};
const dach = [
  opp('Q1-2026', 'Closed Won', 'CompuGroup', 'One Time', 'Enterprise', 0, 0, { closeDate: '2025-02-14' }),
  opp('Q1-2026', 'Closed Won', 'CompuGroup', 'Expand', 'Enterprise', 127401.2, 0, { closeDate: '2025-03-03' }),
  opp('Q1-2026', 'Closed Won', 'CompuGroup', 'One Time', 'Enterprise', 0, 0, { closeDate: '2025-04-11' }),
  opp('Q1-2026', 'Closed Lost', 'BMW Group', 'Land', 'Enterprise', 288000, 0, { major: true, closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'SoftwareOne', 'Land', 'Enterprise', 144000, 0, { oppTeam: 'Partnerships', oppGroup: 'Partnerships', closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'Deutsche Telekom', 'Land', 'Enterprise', 876000, 0, { major: true, closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'Roche', 'Land', 'Enterprise', 180000, 0, { major: true, oppTeam: 'Partnerships', oppGroup: 'Partnerships', closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'Helaba', 'Land', 'Enterprise', 252000, 0, { major: true, closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'Siemens', 'Land', 'Enterprise', 1, 0, { closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'DKB', 'Land', 'Enterprise', 144000, 0, { closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'Julius Baer', 'Land', 'Enterprise', 125000, 0, { closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'Bolt', 'Land', 'Enterprise', 360000, 0, { closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'SAP SE', 'Land', 'Enterprise', 131400, 0, { major: true, closeDate: '2025-04-30' }),
  opp('Q1-2026', 'Closed Lost', 'SAP SE', 'Land', 'Enterprise', 100000, 0, { major: true, closeDate: '2025-04-30' }),
  opp('Q2-2026', 'Closed Won', 'CompuGroup', 'One Time', 'Enterprise', 0, 0, { closeDate: '2025-05-20' }),
  opp('Q2-2026', 'Closed Won', 'Helaba', 'Land', 'Enterprise', 150000, 1, { major: true, closeDate: '2025-06-27' }),
  opp('Q2-2026', 'Closed Lost', 'CompuGroup', 'Expand', 'Enterprise', 2000, 0, { closeDate: '2025-07-31' }),
  opp('Q2-2026', 'Closed Lost', 'Deutsche Telekom', 'Renewal', 'Renewal', -216000, -1, { major: true, lostReason: 'Budget', closeDate: '2025-07-15' }),
  opp('Q3-2026', 'Closed Won', 'Zalando', 'Land', 'Enterprise', 96000, 1, { closeDate: '2025-08-20', amount: 288000 }),
  opp('Q3-2026', 'Closed Won', 'CompuGroup', 'Renewal', 'Renewal', 12000, 0, { closeDate: '2025-09-01' }),
  opp('Q3-2026', 'Closed Won', 'Julius Baer', 'Renewal', 'Renewal', -30000, 0, { closeDate: '2025-09-12', lostReason: 'Seat reduction' }),
  opp('Q3-2026', 'Closed Lost', 'Bolt', 'Renewal', 'Renewal', -84000, -1, { closeDate: '2025-09-30', lostReason: 'Competitor' }),
  opp('Q3-2026', 'Closed Lost', 'Siemens', 'Expand', 'Enterprise', 250000, 0, { closeDate: '2025-10-05' }),
  opp('Q3-2026', 'R2- Renewal Engagement', 'Serrala', 'Renewal', 'Renewal', 0, 0, { expectedDeltaArr: -30000, closeDate: '2025-10-20' }),
  opp('Q3-2026', '1- Discovery', 'Helaba', 'Expand', 'Enterprise', 400000, 0.5, { major: true, expectedDeltaArr: 100000, closeDate: '2025-10-28' }),
  opp('Q4-2026', 'R2- Renewal Engagement', 'CompuGroup', 'Renewal', 'Renewal', 0, 0, { expectedDeltaArr: -30000, closeDate: '2025-12-15' }),
  opp('Q4-2026', '3- Proposal', 'BMW Group', 'Land', 'Enterprise', 540000, 1, { major: true, expectedDeltaArr: 270000, closeDate: '2026-01-20' }),
  opp('Q4-2026', '2- Qualification', 'DKB', 'Land', 'Enterprise', 120000, 1, { expectedDeltaArr: 30000, closeDate: '2025-11-30' }),
  opp('Q1-2027', '1- Discovery', 'Roche', 'Land', 'Enterprise', 300000, 1, { major: true, expectedDeltaArr: 60000, closeDate: '2026-03-31' }),
  opp('Q1-2027', 'R1- Renewal Planning', 'Zalando', 'Renewal', 'Renewal', 20000, 0, { expectedDeltaArr: 10000, closeDate: '2026-04-15' }),
];

const account = (name, currentArr, major, extra = {}) => Object.assign({
  id: accountId(name), url: accountUrl(name), name, currentArr, major, team: 'Europe - DACH', hasOpenOpp: false, isActivated: false,
}, extra);
const dachAccounts = [
  account('Deutsche Telekom', 1200000, true),
  account('SAP SE', 850000, true),
  account('Helaba', 640000, true, { hasOpenOpp: true }),
  account('BMW Group', 0, true, { hasOpenOpp: true }),
  account('CompuGroup', 410000, false, { hasOpenOpp: true }),
  account('Zalando', 96000, false, { hasOpenOpp: true }),
  account('Serrala', 90000, false, { hasOpenOpp: true }),
  account('Julius Baer', 60000, false),
  account('DKB', 0, false, { hasOpenOpp: true, isActivated: true }),
  account('Roche', 0, true, { hasOpenOpp: true, isActivated: true }),
];

module.exports = { opp, dach, account, dachAccounts, SF, accountUrl };
