// Salesforce access via a Connected App using the OAuth 2.0 client-credentials flow.
// Script properties: SF_LOGIN_URL (e.g. https://codeium.my.salesforce.com), SF_CLIENT_ID, SF_CLIENT_SECRET.
const SF_API = '/services/data/v60.0';

let sfSession = null;

function scriptProperty(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) throw new Error(`Script property ${name} is missing. Use QBR > Set Salesforce credentials.`);
  return value;
}

function sfConnect() {
  if (sfSession) return sfSession;
  const response = UrlFetchApp.fetch(`${expectSalesforceUrl(scriptProperty('SF_LOGIN_URL'))}/services/oauth2/token`, {
    method: 'post',
    muteHttpExceptions: true,
    payload: {
      grant_type: 'client_credentials',
      client_id: scriptProperty('SF_CLIENT_ID'),
      client_secret: scriptProperty('SF_CLIENT_SECRET'),
    },
  });
  if (response.getResponseCode() !== 200) throw new Error(`Salesforce login failed: ${response.getContentText()}`);
  const body = JSON.parse(response.getContentText());
  sfSession = { token: body.access_token, instanceUrl: body.instance_url };
  return sfSession;
}

// UrlFetchApp rejects URLs longer than 2 KB; long queries (e.g. an IN list of 140 user Ids) go through the Composite
// API as a POST body instead. Subsequent pages use the short nextRecordsUrl.
const SF_MAX_URL = 1900;

function soql(query) {
  const session = sfConnect();
  const queryPath = `${SF_API}/query?q=${encodeURIComponent(query)}`;
  const records = [];
  let page = queryPath.length + session.instanceUrl.length > SF_MAX_URL
    ? soqlComposite(session, queryPath, query)
    : soqlGet(session, session.instanceUrl + queryPath, query);
  while (page) {
    records.push(...page.records);
    page = page.nextRecordsUrl ? soqlGet(session, session.instanceUrl + page.nextRecordsUrl, query) : null;
  }
  return records;
}

function soqlGet(session, url, query) {
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${session.token}` },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) throw new Error(`SOQL failed: ${response.getContentText()}\n${query}`);
  return JSON.parse(response.getContentText());
}

function soqlComposite(session, queryPath, query) {
  const response = UrlFetchApp.fetch(`${session.instanceUrl}${SF_API}/composite`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${session.token}` },
    payload: JSON.stringify({ allOrNone: true, compositeRequest: [{ method: 'GET', url: queryPath, referenceId: 'q' }] }),
    muteHttpExceptions: true,
  });
  const text = response.getContentText();
  const sub = response.getResponseCode() === 200 ? JSON.parse(text).compositeResponse[0] : null;
  if (!sub || sub.httpStatusCode !== 200) throw new Error(`SOQL failed: ${sub ? JSON.stringify(sub.body) : text}\n${query}`);
  return sub.body;
}

function soqlLiteral(text) {
  return `'${text.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`;
}

// Broad SOQL pre-filter on the token; the exact rule (teamMatches on the resolved team) is applied in JS.
// With no team (bulk mode) the clause is a no-op and every team's rows come back.
function teamClause(fields, team) {
  if (!team) return 'Id != null';
  const like = soqlLiteral(`%${teamToken(team)}%`);
  return `(${fields.map(field => `${field} LIKE ${like}`).join(' OR ')})`;
}

// Bulk mode (refresh of several tabs): each object is queried once for all teams and cached for the execution, then
// filtered per team in JS. The per-team LIKE is only a pre-filter for the same teamMatches rule, so results are identical.
let sfBulk = null;

function withBulkFetch(fn) {
  sfBulk = {};
  try {
    return fn();
  } finally {
    sfBulk = null;
  }
}

function bulkOrTeam(key, team, load) {
  return sfBulk ? bulkCached(key, () => load(null)) : load(team);
}

function bulkCached(key, load) {
  if (!sfBulk) return load();
  if (!(key in sfBulk)) sfBulk[key] = load();
  return sfBulk[key];
}

function ownerClause() {
  return EXCLUDED_OWNERS.length ? `AND Owner.Name NOT IN (${EXCLUDED_OWNERS.map(soqlLiteral).join(', ')})` : '';
}

// Lightning record link shown next to every account / opportunity name on the tabs.
function recordUrl(objectType, id) {
  return id ? `${sfConnect().instanceUrl}/lightning/r/${objectType}/${id}/view` : '';
}

function expectSalesforceUrl(url) {
  const match = /^https:\/\/([a-z0-9.-]+)\/?$/i.exec(String(url).trim());
  if (!match || !/\.(my\.salesforce\.com|salesforce\.com|force\.com)$/i.test(match[1])) {
    throw new Error(`SF_LOGIN_URL must be an https Salesforce My Domain such as https://codeium.my.salesforce.com, got "${url}"`);
  }
  return `https://${match[1]}`;
}

function fetchOpportunities(team, lastFiscalYear) {
  return bulkOrTeam(`opps:${lastFiscalYear}`, team, t => queryOpportunities(t, lastFiscalYear)).filter(o => teamMatches(o.team, team));
}

function queryOpportunities(team, lastFiscalYear) {
  const query = `
    SELECT Id, Name, StageName, IsClosed, IsWon, Type, RecordType.Name, CloseDate, FiscalYear, FiscalQuarter,
           Amount, NACV__c, Expected_NACV__c, Expected_Logo_Impact__c, Closed_Lost_Reason_List__c, Closed_Lost_Reason__c,
           AccountId, Account.Name, Account.Team__r.Name, Account.Subteam__r.Name, Account.Major_Admin_Tag__c,
           Team__r.Name, Group__r.Name, Owner.Name
    FROM Opportunity
    WHERE ${teamClause(['Account.Team__r.Name', 'Account.Subteam__r.Name', 'Team__r.Name'], team)}
      AND Account.Name != 'Test'
      ${ownerClause()}
      AND FiscalYear >= ${parseQuarter(FIRST_QUARTER).fy} AND FiscalYear <= ${lastFiscalYear}`;
  return soql(query).map(r => ({
    id: r.Id,
    url: recordUrl('Opportunity', r.Id),
    name: r.Name,
    stage: r.StageName,
    isClosed: r.IsClosed,
    isWon: r.IsWon,
    type: r.Type,
    recordType: r.RecordType ? r.RecordType.Name : '',
    closeDate: r.CloseDate,
    quarter: quarterLabel(r.FiscalQuarter, r.FiscalYear),
    deltaArr: r.NACV__c || 0,
    amount: r.Amount || 0,
    expectedDeltaArr: r.Expected_NACV__c || 0,
    expectedDeltaArrMissing: r.Expected_NACV__c == null,
    expectedLogoImpact: r.Expected_Logo_Impact__c || 0,
    lostReason: r.Closed_Lost_Reason_List__c || r.Closed_Lost_Reason__c || '',
    accountId: r.AccountId,
    accountUrl: recordUrl('Account', r.AccountId),
    account: r.Account.Name,
    accountTeam: r.Account.Team__r ? r.Account.Team__r.Name : '',
    team: resolveTeam(r.Account.Team__r && r.Account.Team__r.Name, r.Account.Subteam__r && r.Account.Subteam__r.Name, r.Team__r && r.Team__r.Name),
    major: r.Account.Major_Admin_Tag__c === true,
    oppTeam: r.Team__r ? r.Team__r.Name : '',
    oppGroup: r.Group__r ? r.Group__r.Name : '',
    owner: r.Owner ? r.Owner.Name : '',
  }));
}

// Only accounts that can count as an active customer or an activated prospect (any open opportunity, whatever its
// close date). SOQL does not allow a semi-join inside OR, so the two populations are fetched separately and merged.
function fetchAccounts(team) {
  return bulkOrTeam('accounts', team, queryAccounts).filter(a => teamMatches(a.team, team));
}

function queryAccounts(team) {
  const base = `
    SELECT Id, Name, Major_Admin_Tag__c, Current_ARR__c, Team__r.Name, Subteam__r.Name
    FROM Account
    WHERE ${teamClause(['Team__r.Name', 'Subteam__r.Name'], team)}
      AND Name != 'Test'`;
  const customers = soql(`${base} AND Current_ARR__c > 0`);
  const prospects = soql(`${base} AND Id IN (SELECT AccountId FROM Opportunity WHERE IsClosed = false ${ownerClause()})`);
  const hasOpenOpp = {};
  prospects.forEach(r => { hasOpenOpp[r.Id] = true; });
  const seen = {};
  return customers.concat(prospects).filter(r => !seen[r.Id] && (seen[r.Id] = true)).map(r => ({
    id: r.Id,
    url: recordUrl('Account', r.Id),
    name: r.Name,
    major: r.Major_Admin_Tag__c === true,
    currentArr: r.Current_ARR__c || 0,
    hasOpenOpp: hasOpenOpp[r.Id] === true,
    team: resolveTeam(r.Team__r && r.Team__r.Name, r.Subteam__r && r.Subteam__r.Name, ''),
  }));
}

// Accounts left on the region alone (Team = "Europe", no sub-team): they match no team tab, so they are surfaced
// as a data-quality issue on every tab of that region. Only customers and accounts with open pipeline.
function fetchRegionOnlyAccounts(team) {
  const segments = teamSegments(team);
  const region = isRegionSegment(segments[0]) ? segments[0] : null;
  if (!region) return [];
  return bulkCached(`region:${region}`, () => {
    const base = `
      SELECT Id, Name, Team__r.Name, Current_ARR__c, Owner.Name
      FROM Account
      WHERE Team__r.Name = ${soqlLiteral(region)}
        AND Subteam__c = null
        AND Name != 'Test'`;
    const customers = soql(`${base} AND Current_ARR__c > 0`);
    const prospects = soql(`${base} AND Id IN (SELECT AccountId FROM Opportunity WHERE IsClosed = false ${ownerClause()})`);
    const seen = {};
    return customers.concat(prospects).filter(a => !seen[a.Id] && (seen[a.Id] = true)).map(a => ({
      id: a.Id,
      url: recordUrl('Account', a.Id),
      name: a.Name,
      team: a.Team__r ? a.Team__r.Name : region,
      currentArr: a.Current_ARR__c || 0,
      owner: a.Owner ? a.Owner.Name : '',
    }));
  });
}

// Returns { 'Q3-2026': { revenue, logos }, ... } summed over every Salesforce team matching the token
// (the "Goals By Quarter - Net ARR" report: Goal__c grouped by Team2__c and Period_Start__c).
function fetchGoals(team) {
  const rows = bulkOrTeam('goals', team, queryGoals).filter(r => teamMatches(r.Team2__r && r.Team2__r.Name, team));
  const goals = {};
  rows.forEach(r => {
    const quarter = quarterOfDate(r.Period_Start__c);
    const goal = goals[quarter] || (goals[quarter] = { revenue: 0, logos: 0 });
    if (r.Goal_Type__c === 'Net ARR') goal.revenue += r.Value__c || 0;
    if (r.Goal_Type__c === 'New Logos') goal.logos += r.Count_Value__c || 0;
  });
  return goals;
}

function queryGoals(team) {
  const query = `
    SELECT Goal_Type__c, Period_Start__c, Value__c, Count_Value__c, Team2__r.Name
    FROM Goal__c
    WHERE ${teamClause(['Team2__r.Name'], team)}
      AND Goal_Type__c IN ('Net ARR', 'New Logos')
      AND Period_Start__c != null`;
  return soql(query);
}
