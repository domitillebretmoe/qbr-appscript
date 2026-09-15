// Rep activity & performance: the measures of the Salesforce "Global GTM Dashboard > Majors Rep Performance" tab
// (Apex GTMDashRepPerformanceController), recomputed for the reps of the tab's team with plain SOQL.
// Reps = active users with a User_Segment__c in a sales-carrying group whose Team matches the tab.
// Attribution follows the dashboard: accounts owned, goals and activity coverage by Account.OwnerId; won ARR,
// coverage, pipeline created, stalled pipeline and renewal risk by Opportunity.OwnerId; Gong-backed activity only.
// Scoped to the tab's quarter: Won ARR, meetings, activities, account coverage and pipeline created cover the
// selected quarter up to today (the whole quarter once it is closed); rep goals only exist per fiscal year in
// Salesforce, so goal, FY Won ARR, attainment and coverage stay FY-level.
// Simplifications vs the dashboard: meetings are credited to the event owner (no participant credit), deduped by
// Gong activity id; no slippage rate (needs the weekly snapshots).
const REP_GROUPS = ['US Majors', 'Europe', 'Asia', 'LATAM', 'Federal', 'US Enterprise', 'Partnerships'];
const REP_EXCLUDED_FAMILIES = ['Deployed Engineering', 'SDR', 'Pre-sales'];
const REP_STALLED_DAYS = 60;
const REP_RENEWAL_RECORD_TYPES = ['Support_Renewal', 'Fed_Renewal'];

// One row per rep of `team` (all member teams for a roll-up), best quarter Won ARR first. Reps are always taken
// from every team's segments so that a user's current (most recent) segment decides the team, whichever tab refreshes.
function fetchRepPerformance(team, quarter, today) {
  const members = rollupMembers(team) || [team];
  const universe = bulkCached('reps', () => queryReps(null));
  const reps = universe.filter(r => members.some(m => teamMatches(r.team, m)));
  if (!reps.length) return [];
  const stats = bulkCached(`repStats:${quarter}`, () => queryRepStats(sfBulk ? universe : reps, quarter, today));
  return reps.map(rep => Object.assign({}, rep, repMetrics(rep, stats[rep.userId] || {}, today)))
    .sort((a, b) => (b.qWonArr || 0) - (a.qWonArr || 0) || (b.fyWonArr || 0) - (a.fyWonArr || 0) || a.name.localeCompare(b.name));
}

function repMetrics(rep, s, today) {
  const seat = rep.createdDate ? Math.max(0, daysBetween(rep.createdDate, today)) / 30.4375 : null;
  const accountsOwned = s.accountsOwned || 0;
  const repGoal = s.repGoal || 0;
  const renewalArr = s.renewalArr || 0;
  return {
    monthsInSeat: seat == null ? null : Math.round(seat * 10) / 10,
    accountsOwned,
    repGoal,
    qWonArr: s.qWonArr || 0,
    fyWonArr: s.fyWonArr || 0,
    attainmentPct: ratio(s.fyWonArr || 0, repGoal),
    coverageArr: s.coverageArr || 0,
    coveragePct: ratio(s.coverageArr || 0, repGoal),
    meetings: s.meetings || 0,
    activities: s.activities || 0,
    coveredAccounts: s.coveredAccounts || 0,
    activityCoveragePct: ratio(s.coveredAccounts || 0, accountsOwned),
    pipelineCreatedArr: s.pipelineCreatedArr || 0,
    pipelineCreatedCount: s.pipelineCreatedCount || 0,
    stalledArr: s.stalledArr || 0,
    stalledCount: s.stalledCount || 0,
    renewalArr,
    renewalAtRiskArr: s.renewalAtRiskArr || 0,
    renewalRiskPct: ratio(s.renewalAtRiskArr || 0, renewalArr),
  };
}

function queryReps(team) {
  const query = `
    SELECT User__c, User__r.Name, User__r.Email, User__r.CreatedDate, User__r.User_Family__c, Group__r.Name, Team__r.Name
    FROM User_Segment__c
    WHERE ${teamClause(['Team__r.Name'], team)}
      AND User__r.IsActive = true
      AND User__c != null
      AND Group__r.Name IN (${REP_GROUPS.map(soqlLiteral).join(', ')})
    ORDER BY LastModifiedDate DESC`;
  const seen = {};
  return soql(query).filter(r => {
    if (seen[r.User__c] || REP_EXCLUDED_FAMILIES.indexOf(r.User__r.User_Family__c) >= 0) return false;
    seen[r.User__c] = true;
    return true;
  }).map(r => ({
    userId: r.User__c,
    name: r.User__r.Name,
    url: recordUrl('User', r.User__c),
    email: r.User__r.Email || '',
    createdDate: r.User__r.CreatedDate ? String(r.User__r.CreatedDate).slice(0, 10) : null,
    group: r.Group__r ? r.Group__r.Name : '',
    team: r.Team__r ? r.Team__r.Name : '',
  }));
}

// Salesforce team (User Segment, most recently modified first) keyed by user Id for the given opportunity owner Ids;
// `ids` null = everyone.
function fetchOwnerTeams(ids) {
  if (ids && !ids.length) return {};
  return sfBulk ? bulkCached('ownerTeams', () => queryOwnerTeams(null)) : queryOwnerTeams(ids);
}

function queryOwnerTeams(ids) {
  const where = ids ? `WHERE User__c IN (${unique(ids).map(soqlLiteral).join(', ')})` : 'WHERE User__c != null';
  const teams = {};
  soql(`SELECT User__c, Team__r.Name FROM User_Segment__c ${where} ORDER BY LastModifiedDate DESC`).forEach(r => {
    if (r.User__c && !(r.User__c in teams)) teams[r.User__c] = r.Team__r ? r.Team__r.Name : '';
  });
  return teams;
}

// Activity Gong still syncs onto a rep's retired duplicate user (inactive, same name, same email local-part) is
// credited to the active rep.
function queryRepAliases(reps) {
  if (!reps.length) return {};
  const byLocalPart = {};
  reps.forEach(r => { if (r.email) byLocalPart[r.email.split('@')[0].toLowerCase()] = r.userId; });
  const query = `
    SELECT Id, Email FROM User
    WHERE IsActive = false AND Name IN (${unique(reps.map(r => r.name)).map(soqlLiteral).join(', ')})`;
  const alias = {};
  soql(query).forEach(u => {
    const rep = byLocalPart[String(u.Email || '').split('@')[0].toLowerCase()];
    if (rep && rep !== u.Id) alias[u.Id] = rep;
  });
  return alias;
}

// Per-rep measures: goal / coverage / renewal risk over the fiscal year of `quarter`, Won ARR for both the fiscal year
// and the quarter, activity and pipeline created within the quarter up to `today`.
function queryRepStats(reps, quarter, today) {
  const stats = {};
  const stat = id => stats[id] || (stats[id] = {});
  const ids = reps.map(r => r.userId);
  const idList = `(${ids.map(soqlLiteral).join(', ')})`;
  const fy = parseQuarter(quarter).fy;
  const fyStart = `${fy}-02-01`;
  const fyEnd = `${fy + 1}-01-31`;
  const qStart = quarterStart(quarter);
  const qEnd = quarterStart(shiftQuarter(quarter, 1));
  const aliases = queryRepAliases(reps);
  const repOf = ownerId => aliases[ownerId] || ownerId;
  const isRep = {};
  ids.forEach(id => { isRep[id] = true; });

  soql(`SELECT OwnerId o, COUNT(Id) c FROM Account WHERE OwnerId IN ${idList} GROUP BY OwnerId`)
    .forEach(r => { stat(r.o).accountsOwned = r.c; });

  soql(`SELECT Weighted_Goal__c, Account_Owner__c, Account__r.OwnerId
        FROM Account_Goal__c
        WHERE Weighted_Goal__c != null AND (Account_Owner__c IN ${idList} OR Account__r.OwnerId IN ${idList})`)
    .forEach(g => {
      const owner = g.Account_Owner__c || (g.Account__r && g.Account__r.OwnerId);
      if (isRep[owner]) stat(owner).repGoal = (stat(owner).repGoal || 0) + (g.Weighted_Goal__c || 0);
    });

  soql(`SELECT OwnerId, Type, IsWon, RecordType.Name, NACV__c, ARR__c, Amount, CloseDate FROM Opportunity
        WHERE OwnerId IN ${idList} AND IsWon = true AND CloseDate >= ${fyStart} AND CloseDate <= ${fyEnd}`)
    .forEach(o => {
      const s = stat(o.OwnerId);
      const v = deltaArrOf(o);
      s.fyWonArr = (s.fyWonArr || 0) + v;
      if (o.CloseDate >= qStart && o.CloseDate < qEnd) s.qWonArr = (s.qWonArr || 0) + v;
    });

  soql(`SELECT OwnerId o, SUM(Expected_NACV__c) v FROM Opportunity
        WHERE OwnerId IN ${idList} AND (IsClosed = false OR IsWon = true) AND CloseDate >= ${fyStart} AND CloseDate <= ${fyEnd}
        GROUP BY OwnerId`)
    .forEach(r => { stat(r.o).coverageArr = r.v || 0; });

  soql(`SELECT OwnerId o, SUM(Expected_NACV__c) v, COUNT(Id) c FROM Opportunity
        WHERE OwnerId IN ${idList} AND CreatedDate >= ${qStart}T00:00:00Z AND CreatedDate < ${qEnd}T00:00:00Z GROUP BY OwnerId`)
    .forEach(r => { stat(r.o).pipelineCreatedArr = r.v || 0; stat(r.o).pipelineCreatedCount = r.c; });

  const stalledCut = Date.parse(`${today}T00:00:00Z`) - REP_STALLED_DAYS * 86400000;
  soql(`SELECT OwnerId, Expected_NACV__c, LastStageChangeDate, CreatedDate FROM Opportunity
        WHERE OwnerId IN ${idList} AND IsClosed = false`)
    .forEach(o => {
      if (Date.parse(o.LastStageChangeDate || o.CreatedDate) >= stalledCut) return;
      const s = stat(o.OwnerId);
      s.stalledArr = (s.stalledArr || 0) + (o.Expected_NACV__c || 0);
      s.stalledCount = (s.stalledCount || 0) + 1;
    });

  soql(`SELECT OwnerId, Starting_ARR__c, Pwin__c FROM Opportunity
        WHERE OwnerId IN ${idList} AND RecordType.DeveloperName IN (${REP_RENEWAL_RECORD_TYPES.map(soqlLiteral).join(', ')})
          AND IsClosed = false AND Exclude_From_Churn_Impact__c = false
          AND CloseDate >= ${fyStart} AND CloseDate <= ${fyEnd} AND Starting_ARR__c > 0`)
    .forEach(o => {
      if (o.Pwin__c == null) return;
      const s = stat(o.OwnerId);
      s.renewalArr = (s.renewalArr || 0) + o.Starting_ARR__c;
      s.renewalAtRiskArr = (s.renewalAtRiskArr || 0) + o.Starting_ARR__c * (1 - o.Pwin__c / 100);
    });

  const ownerList = `(${ids.concat(Object.keys(aliases)).map(soqlLiteral).join(', ')})`;
  const window = `ActivityDate >= ${qStart} AND ActivityDate < ${qEnd} AND ActivityDate <= ${today}`;
  const covered = {};
  const touch = (ownerId, accountOwner, accountId) => {
    const rep = repOf(ownerId);
    if (!isRep[rep] || !accountId || accountOwner !== rep) return;
    (covered[rep] || (covered[rep] = {}))[accountId] = true;
  };
  soql(`SELECT OwnerId o, COUNT_DISTINCT(Gong__Gong_Activity_Id__c) c FROM Event
        WHERE OwnerId IN ${ownerList} AND Gong__Gong_Activity_Id__c != null AND ${window}
          AND Gong__Meeting_Prospect_Canceled__c = false AND (NOT Subject LIKE '%Meeting Canceled') GROUP BY OwnerId`)
    .forEach(r => { const s = stat(repOf(r.o)); s.meetings = (s.meetings || 0) + r.c; });
  soql(`SELECT OwnerId o, COUNT_DISTINCT(Gong__Gong_Activity_Id__c) c FROM Task
        WHERE OwnerId IN ${ownerList} AND Gong__Gong_Activity_Id__c != null AND ${window} GROUP BY OwnerId`)
    .forEach(r => { const s = stat(repOf(r.o)); s.activities = (s.activities || 0) + r.c; });
  ['Event', 'Task'].forEach(object => {
    soql(`SELECT OwnerId o, AccountId a, Account.OwnerId FROM ${object}
          WHERE OwnerId IN ${ownerList} AND Gong__Gong_Activity_Id__c != null AND ${window} AND AccountId != null
          GROUP BY OwnerId, AccountId, Account.OwnerId`)
      .forEach(r => touch(r.o, r.Account ? r.Account.OwnerId : r.OwnerId, r.a)); // grouped relationship field comes back flat as OwnerId
  });
  Object.keys(covered).forEach(rep => { stat(rep).coveredAccounts = Object.keys(covered[rep]).length; });
  return stats;
}
