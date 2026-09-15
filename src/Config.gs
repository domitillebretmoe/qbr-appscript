// Fiscal quarters are labelled "Q3-2026". FY2026 Q1 starts 1 Feb 2026 (Salesforce fiscal calendar).
const FIRST_QUARTER = 'Q1-2026';
// B2 dropdown: FIRST_QUARTER through the quarter after the current one (the latest quarter a tab can be pointed at).
function quarterOptions(today) {
  return quartersBetween(FIRST_QUARTER, shiftQuarter(quarterOfDate(today), 1));
}

const LEDGER_SHEET = 'ARR Ledger';
const RAW_SHEET = 'Raw Data';
const GOALS_SHEET = 'Goals';
const SEED_SOURCE = 'Seeded from FY27 QBR Cockpit';

// Team names as they appear in Salesforce (Account > Team). The text after " - " is the contains-token
// used to match Account Team names, so "Europe - DACH" matches both "Europe - DACH" and "Europe Majors - DACH".
// [team, Q4-2025 Ending ARR (= Q1-2026 Starting), Q1-2026 Ending ARR, Q2-2026 Ending ARR], as in each tab of the
// FY27 QBR Cockpit workbook ("Raw - Q4 Ending ARR" for the starting point, the tab's Ending ARR rows as-is).
const TEAM_SEEDS = [
  ['US Majors - Media, Telco',        3300000,    8889104,    12996518],
  ['Partnerships',                    3300000,    4489407,    8390367],
  ['US Majors - Industrials',         1400000,    3233200,    8706000],
  ['US Majors - Retail',              5400000,    5771282,    12119282],
  ['US Majors - Healthcare',          6800000,    5675172,    9272772],
  ['Federal',                         4000000,    4925559,    4548831],
  ['US Majors - Banking',             30700000,   56070309,   81217859],
  ['US Majors - Insurance',           0,          1267600,    1987600],
  ['US Majors - Financial Services',  6000000,    9070239,    27455439],
  ['US Majors - Technology',          13000000,   19524630,   43390139],
  ['US Enterprise',                   23900000,   30826617,   44650370],
  ['Europe - Nordics',                3623750,    3161750,    897150],
  ['Europe - Benelux',                1895500,    3801700,    1812235],
  ['Europe - UKI',                    19847000,   37631000,   38731000],
  ['Europe - DACH',                   13324250,   14106851.2, 14040851.2],
  ['Europe - South',                  17059500,   17584699,   38522758],
  ['APAC',                            3200000,    4439000,    5153000],
  ['Japan',                           3900000,    9027005.6,  12413165.6],
  ['LATAM',                           25000000,   32889821,   33489821],
];

const TEAMS = TEAM_SEEDS.map(row => row[0]);

// Roll-up tabs: a view over several teams. Opportunities stay attributed to their specific team; the roll-up sums
// the members' opportunities, accounts and goals, and its Starting/Ending ARR is the sum of the members' ARR Ledger
// rows (a roll-up never gets ledger rows of its own).
const ROLLUP_TEAMS = { Europe: TEAMS.filter(team => team.indexOf('Europe - ') === 0) };

function rollupMembers(team) {
  return ROLLUP_TEAMS[String(team).trim()] || null;
}

// Region-level segments that must never be used as a team: opportunities are always attributed to the most
// specific sub-team (e.g. Europe - DACH, Europe Majors - UKI), never to "Europe".
const REGION_SEGMENTS = ['Europe', 'Asia', 'US Majors'];

// Opportunity owners excluded by the "FY26 QBR COCKPIT" Salesforce report.
const EXCLUDED_OWNERS = ['Christian Lawless'];

const PARTNER_GROUP = 'Partnerships';

// Pilots: an open opportunity in the Tech Validation stage is an active pilot; a pilot is completed once Pilot Status is
// Complete, in the quarter of its Pilot Actual End Date.
const PILOT_STAGE = '3- Tech Validation';
const PILOT_COMPLETE_STATUS = 'Complete';

function teamSegments(team) {
  return String(team).split(' - ').map(s => s.trim()).filter(Boolean);
}

function teamToken(team) {
  const parts = teamSegments(team);
  return parts[parts.length - 1];
}

// Contains match on the token after the last " - " of B1 (e.g. DACH), so "Europe - DACH" covers
// "Europe - DACH", "Europe Majors - DACH" and "Europe - DACH - Austria". The leading segments of B1
// must also appear, so "Europe - South" never picks up "LATAM - South".
function teamMatches(segmentName, team) {
  const name = String(segmentName || '').toLowerCase();
  return !!name && teamSegments(team).every(part => name.indexOf(part.toLowerCase()) >= 0);
}

function isRegionSegment(name) {
  return REGION_SEGMENTS.indexOf(String(name || '').trim()) >= 0;
}

function assertSpecificTeam(team) {
  if (isRegionSegment(team)) {
    throw new Error(`"${team}" is a region, not a team. Use the most detailed team, e.g. ${team} - DACH.`);
  }
}

// The most detailed team of an opportunity: the account's team unless it is only a region ("Europe"),
// in which case fall back to the account's sub-team, then the opportunity's team.
function resolveTeam(accountTeam, accountSubteam, oppTeam) {
  const candidates = [accountTeam, accountSubteam, oppTeam].map(n => String(n || '').trim()).filter(Boolean);
  const specific = candidates.filter(n => !isRegionSegment(n) && n !== 'Other');
  return specific[0] || candidates[0] || '';
}

function parseQuarter(label) {
  const m = /^Q([1-4])-(\d{4})$/.exec(String(label).trim());
  if (!m) throw new Error(`Quarter must look like Q3-2026, got "${label}"`);
  return { q: Number(m[1]), fy: Number(m[2]) };
}

function quarterLabel(q, fy) {
  return `Q${q}-${fy}`;
}

function shiftQuarter(label, n) {
  const { q, fy } = parseQuarter(label);
  const index = fy * 4 + (q - 1) + n;
  return quarterLabel((index % 4) + 1, Math.floor(index / 4));
}

// Q1 starts in February, so Feb-Apr = Q1, May-Jul = Q2, Aug-Oct = Q3, Nov-Jan = Q4.
function quarterOfDate(isoDate) {
  const [year, month] = isoDate.split('-').map(Number);
  if (month === 1) return quarterLabel(4, year - 1);
  return quarterLabel(Math.floor((month - 2) / 3) + 1, year);
}

// First day of the quarter as yyyy-mm-dd (Q1 = 1 Feb of the fiscal year, Q4 = 1 Nov).
function quarterStart(label) {
  const { q, fy } = parseQuarter(label);
  const month = 2 + (q - 1) * 3;
  return `${fy}-${String(month).padStart(2, '0')}-01`;
}

// Share of the quarter's days already elapsed on `today` (yyyy-mm-dd): 0 before it starts, 1 once it is over.
function quarterElapsed(label, today) {
  const start = quarterStart(label);
  return Math.min(1, Math.max(0, daysBetween(start, today) / daysBetween(start, quarterStart(shiftQuarter(label, 1)))));
}

function daysBetween(fromIso, toIso) {
  const day = iso => Date.UTC(...iso.slice(0, 10).split('-').map((n, i) => Number(n) - (i === 1 ? 1 : 0)));
  return Math.round((day(toIso) - day(fromIso)) / 86400000);
}

function shiftDate(iso, days) {
  const date = new Date(`${iso.slice(0, 10)}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

// The day of `previousQuarter` that is as far in as `today` is in `quarter` (e.g. day 41 of both), for like-for-like
// quarter-on-quarter comparisons while `quarter` is in progress.
function samePointInQuarter(quarter, previousQuarter, today) {
  return shiftDate(quarterStart(previousQuarter), daysBetween(quarterStart(quarter), today));
}

function quartersBetween(first, last) {
  const out = [];
  for (let q = first; ; q = shiftQuarter(q, 1)) {
    out.push(q);
    if (q === last) return out;
    if (out.length > 40) throw new Error(`${last} is before ${first}`);
  }
}
