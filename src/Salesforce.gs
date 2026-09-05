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

function soql(query) {
  const session = sfConnect();
  let url = `${session.instanceUrl}${SF_API}/query?q=${encodeURIComponent(query)}`;
  const records = [];
  while (url) {
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${session.token}` },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) throw new Error(`SOQL failed: ${response.getContentText()}\n${query}`);
    const page = JSON.parse(response.getContentText());
    records.push(...page.records);
    url = page.nextRecordsUrl ? session.instanceUrl + page.nextRecordsUrl : null;
  }
  return records;
}

function soqlLiteral(text) {
  return `'${text.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`;
}

// Broad SOQL pre-filter on the token; the exact rule (teamMatches on the resolved team) is applied in JS.
function teamClause(fields, team) {
  const like = soqlLiteral(`%${teamToken(team)}%`);
  return `(${fields.map(field => `${field} LIKE ${like}`).join(' OR ')})`;
}

function ownerClause() {
  return EXCLUDED_OWNERS.length ? `AND Owner.Name NOT IN (${EXCLUDED_OWNERS.map(soqlLiteral).join(', ')})` : '';
}

function expectSalesforceUrl(url) {
  const match = /^https:\/\/([a-z0-9.-]+)\/?$/i.exec(String(url).trim());
  if (!match || !/\.(my\.salesforce\.com|salesforce\.com|force\.com)$/i.test(match[1])) {
    throw new Error(`SF_LOGIN_URL must be an https Salesforce My Domain such as https://codeium.my.salesforce.com, got "${url}"`);
  }
  return `https://${match[1]}`;
}

function fetchOpportunities(team, lastFiscalYear) {
  const query = `
    SELECT Id, Name, StageName, IsClosed, IsWon, Type, RecordType.Name, CloseDate, FiscalYear, FiscalQuarter,
           NACV__c, Expected_NACV__c, Expected_Logo_Impact__c, Closed_Lost_Reason_List__c, Closed_Lost_Reason__c,
           AccountId, Account.Name, Account.Team__r.Name, Account.Subteam__r.Name, Account.Major_Admin_Tag__c,
           Team__r.Name, Group__r.Name, Owner.Name
    FROM Opportunity
    WHERE ${teamClause(['Account.Team__r.Name', 'Account.Subteam__r.Name', 'Team__r.Name'], team)}
      AND Account.Name != 'Test'
      ${ownerClause()}
      AND FiscalYear >= ${parseQuarter(FIRST_QUARTER).fy} AND FiscalYear <= ${lastFiscalYear}`;
  return soql(query).map(r => ({
    id: r.Id,
    name: r.Name,
    stage: r.StageName,
    isClosed: r.IsClosed,
    isWon: r.IsWon,
    type: r.Type,
    recordType: r.RecordType ? r.RecordType.Name : '',
    closeDate: r.CloseDate,
    quarter: quarterLabel(r.FiscalQuarter, r.FiscalYear),
    deltaArr: r.NACV__c || 0,
    expectedDeltaArr: r.Expected_NACV__c || 0,
    expectedLogoImpact: r.Expected_Logo_Impact__c || 0,
    lostReason: r.Closed_Lost_Reason_List__c || r.Closed_Lost_Reason__c || '',
    accountId: r.AccountId,
    account: r.Account.Name,
    accountTeam: r.Account.Team__r ? r.Account.Team__r.Name : '',
    team: resolveTeam(r.Account.Team__r && r.Account.Team__r.Name, r.Account.Subteam__r && r.Account.Subteam__r.Name, r.Team__r && r.Team__r.Name),
    major: r.Account.Major_Admin_Tag__c === true,
    oppTeam: r.Team__r ? r.Team__r.Name : '',
    oppGroup: r.Group__r ? r.Group__r.Name : '',
    owner: r.Owner ? r.Owner.Name : '',
  })).filter(o => teamMatches(o.team, team));
}

// Only accounts that can count as an active customer or an activated prospect (any open opportunity, whatever its
// close date). SOQL does not allow a semi-join inside OR, so the two populations are fetched separately and merged.
function fetchAccounts(team) {
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
    name: r.Name,
    major: r.Major_Admin_Tag__c === true,
    currentArr: r.Current_ARR__c || 0,
    hasOpenOpp: hasOpenOpp[r.Id] === true,
    team: resolveTeam(r.Team__r && r.Team__r.Name, r.Subteam__r && r.Subteam__r.Name, ''),
  })).filter(a => teamMatches(a.team, team));
}

// Returns { 'Q3-2026': { revenue, logos }, ... } summed over every Salesforce team matching the token
// (the "Goals By Quarter - Net ARR" report: Goal__c grouped by Team2__c and Period_Start__c).
function fetchGoals(team) {
  const query = `
    SELECT Goal_Type__c, Period_Start__c, Value__c, Count_Value__c, Team2__r.Name
    FROM Goal__c
    WHERE ${teamClause(['Team2__r.Name'], team)}
      AND Goal_Type__c IN ('Net ARR', 'New Logos')
      AND Period_Start__c != null`;
  const goals = {};
  soql(query).filter(r => teamMatches(r.Team2__r && r.Team2__r.Name, team)).forEach(r => {
    const quarter = quarterOfDate(r.Period_Start__c);
    const goal = goals[quarter] || (goals[quarter] = { revenue: 0, logos: 0 });
    if (r.Goal_Type__c === 'Net ARR') goal.revenue += r.Value__c || 0;
    if (r.Goal_Type__c === 'New Logos') goal.logos += r.Count_Value__c || 0;
  });
  return goals;
}
