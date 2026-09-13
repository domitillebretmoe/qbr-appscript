// soql() from Salesforce.gs with UrlFetchApp stubbed: short queries go as GET, long ones (Apps Script rejects URLs
// over 2 KB) through the Composite API POST, and pagination follows nextRecordsUrl either way.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const vm = require('node:vm');

const INSTANCE = 'https://codeium.my.salesforce.com';

function sfContext(handler) {
  const calls = [];
  const ctx = vm.createContext({
    UrlFetchApp: {
      fetch(url, opts) {
        assert.ok(url.length <= 2048, `URL too long for UrlFetchApp: ${url.length}`);
        calls.push({ url, opts });
        const body = handler(url, opts);
        return { getResponseCode: () => body.status || 200, getContentText: () => JSON.stringify(body.json) };
      },
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'x' }) },
  });
  ['Config.gs', 'Metrics.gs', 'Salesforce.gs', 'Reps.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));
  vm.runInContext(`sfSession = { token: 't', instanceUrl: ${JSON.stringify(INSTANCE)} }`, ctx);
  return { ctx, calls };
}

const page = (records, next) => ({ json: Object.assign({ done: !next, records }, next ? { nextRecordsUrl: next } : {}) });

test('short query: GET with the query in the URL, follows nextRecordsUrl', () => {
  const { ctx, calls } = sfContext(url => url.includes('/query/01g')
    ? page([{ Id: 'b' }])
    : page([{ Id: 'a' }], '/services/data/v60.0/query/01g-2000'));
  const records = ctx.soql('SELECT Id FROM Account');
  assert.strictEqual(records.map(r => r.Id).join(), 'a,b');
  assert.strictEqual(calls.length, 2);
  assert.strictEqual(calls[0].url, `${INSTANCE}/services/data/v60.0/query?q=SELECT%20Id%20FROM%20Account`);
  assert.strictEqual(calls[0].opts.method, undefined);
  assert.strictEqual(calls[0].opts.headers.Authorization, 'Bearer t');
  assert.strictEqual(calls[1].url, `${INSTANCE}/services/data/v60.0/query/01g-2000`);
});

test('long query (IN list of 140 Ids): POST to /composite with the query as a subrequest, pages by GET', () => {
  const ids = Array.from({ length: 140 }, (_, i) => `'005${String(i).padStart(15, '0')}'`);
  const query = `SELECT OwnerId o, COUNT(Id) c FROM Account WHERE OwnerId IN (${ids.join(', ')}) GROUP BY OwnerId`;
  const { ctx, calls } = sfContext((url, opts) => {
    if (url.endsWith('/composite')) {
      const body = JSON.parse(opts.payload);
      assert.strictEqual(opts.method, 'post');
      assert.strictEqual(body.compositeRequest.length, 1);
      assert.strictEqual(body.compositeRequest[0].method, 'GET');
      assert.strictEqual(body.compositeRequest[0].url, `/services/data/v60.0/query?q=${encodeURIComponent(query)}`);
      return { json: { compositeResponse: [{ httpStatusCode: 200, referenceId: 'q',
        body: page([{ o: 'x', c: 1 }], '/services/data/v60.0/query/01g-2000').json }] } };
    }
    return page([{ o: 'y', c: 2 }]);
  });
  const records = ctx.soql(query);
  assert.strictEqual(records.map(r => r.o).join(), 'x,y');
  assert.strictEqual(calls.length, 2);
  assert.strictEqual(calls[0].url, `${INSTANCE}/services/data/v60.0/composite`);
  assert.strictEqual(calls[0].opts.headers.Authorization, 'Bearer t');
  assert.strictEqual(calls[1].url, `${INSTANCE}/services/data/v60.0/query/01g-2000`);
});

test('composite subrequest error surfaces the Salesforce message and the query', () => {
  const query = `SELECT Id FROM Account WHERE Id IN (${Array(120).fill("'001000000000000AAA'").join(', ')})`;
  const { ctx } = sfContext(() => ({ json: { compositeResponse: [{ httpStatusCode: 400, referenceId: 'q',
    body: [{ errorCode: 'MALFORMED_QUERY', message: 'unexpected token' }] }] } }));
  assert.throws(() => ctx.soql(query), /MALFORMED_QUERY[\s\S]*SELECT Id FROM Account/);
});

test('queryOpportunities: a Closed Won MSP deal with NACV 0 takes its ARR (else Amount) as Delta ARR', () => {
  const acct = { Name: 'SIEMENS AG', Team__r: { Name: 'Europe Majors - DACH' }, Subteam__r: null, Major_Admin_Tag__c: true };
  const base = { StageName: 'Closed Won', IsClosed: true, IsWon: true, RecordType: { Name: 'Enterprise' }, CloseDate: '2026-09-03',
    FiscalYear: 2027, FiscalQuarter: 3, AccountId: '001A', Account: acct, Team__r: null, Group__r: null, Owner: null };
  const records = [
    Object.assign({ Id: '1', Name: 'Siemens MSP', Type: 'MSP', Amount: 95520, ARR__c: 95520, NACV__c: 0 }, base),
    Object.assign({ Id: '2', Name: 'Ford MSP', Type: 'MSP', Amount: 273600, ARR__c: 547200, NACV__c: 0 }, base),
    Object.assign({ Id: '3', Name: 'MSP no ARR', Type: 'MSP', Amount: 16000, ARR__c: null, NACV__c: 0 }, base),
    Object.assign({ Id: '4', Name: 'MSP with NACV', Type: 'MSP', Amount: 206400, ARR__c: 206400, NACV__c: 189120 }, base),
    Object.assign({ Id: '5', Name: 'Renewal', Type: 'Renewal', Amount: 4980800, ARR__c: 4980800, NACV__c: 0 }, base),
    Object.assign({ Id: '6', Name: 'Lost MSP', Type: 'MSP', Amount: 100000, ARR__c: null, NACV__c: 0 }, base, { StageName: 'Closed Lost', IsWon: false }),
    Object.assign({ Id: '7', Name: 'Open MSP', Type: 'MSP', Amount: 500000, ARR__c: null, NACV__c: 0 }, base, { StageName: '1- Discovery', IsClosed: false, IsWon: false }),
    Object.assign({ Id: '8', Name: 'MSP blank NACV', Type: 'MSP', Amount: 75000, ARR__c: 50000, NACV__c: null }, base),
    Object.assign({ Id: '9', Name: 'MSP renewal', Type: 'MSP', Amount: 200000, ARR__c: 200000, NACV__c: 0 }, base, { RecordType: { Name: 'Renewal' } }),
    Object.assign({ Id: '10', Name: 'MSP fed renewal', Type: 'MSP', Amount: 200000, ARR__c: 200000, NACV__c: 0 }, base, { RecordType: { Name: 'Fed - Renewal' } }),
  ];
  const { ctx, calls } = sfContext(() => page(records));
  const opps = ctx.queryOpportunities('Europe - DACH', 2027);
  assert.strictEqual(opps.map(o => [o.name, o.deltaArr, o.amount].join(':')).join('\n'), [
    'Siemens MSP:95520:95520', 'Ford MSP:547200:273600', 'MSP no ARR:16000:16000', 'MSP with NACV:189120:206400',
    'Renewal:0:4980800', 'Lost MSP:0:100000', 'Open MSP:0:500000', 'MSP blank NACV:0:75000', 'MSP renewal:0:200000', 'MSP fed renewal:0:200000',
  ].join('\n'));
  assert.match(decodeURIComponent(calls[0].url), /Amount, ARR__c, NACV__c/);
});

test('queryRepStats: Won ARR split FY vs selected quarter (MSP fallback applied), activity windows are the quarter to date', () => {
  const queries = [];
  const { ctx } = sfContext(url => {
    const q = decodeURIComponent(url.split('?q=')[1] || '').replace(/\s+/g, ' ');
    queries.push(q);
    if (q.startsWith('SELECT OwnerId, Type, IsWon, RecordType.Name')) {
      return page([
        { OwnerId: 'u1', Type: 'Land', IsWon: true, RecordType: { Name: 'Enterprise' }, NACV__c: 1000000, ARR__c: 1000000, Amount: 1000000, CloseDate: '2026-03-15' },
        { OwnerId: 'u1', Type: 'MSP', IsWon: true, RecordType: { Name: 'Enterprise' }, NACV__c: 0, ARR__c: 95520, Amount: 95520, CloseDate: '2026-09-03' },
        { OwnerId: 'u1', Type: 'Expand', IsWon: true, RecordType: { Name: 'Enterprise' }, NACV__c: 200000, ARR__c: null, Amount: 200000, CloseDate: '2026-08-01' },
        { OwnerId: 'u1', Type: 'Expand', IsWon: true, RecordType: { Name: 'Enterprise' }, NACV__c: 50000, ARR__c: null, Amount: 50000, CloseDate: '2026-11-01' },
        { OwnerId: 'u2', Type: 'Renewal', IsWon: true, RecordType: { Name: 'Renewal' }, NACV__c: 0, ARR__c: 300000, Amount: 300000, CloseDate: '2026-08-20' },
      ]);
    }
    if (q.includes('FROM Event')) return page([{ o: 'u1', c: 7 }]);
    if (q.includes('FROM Task')) return page([{ o: 'u1', c: 40 }]);
    return page([]);
  });
  const reps = [{ userId: 'u1', name: 'A', email: 'a@x.com' }, { userId: 'u2', name: 'B', email: 'b@x.com' }];
  const stats = ctx.queryRepStats(reps, 'Q3-2026', '2026-09-12');
  assert.strictEqual(stats.u1.fyWonArr, 1345520);
  assert.strictEqual(stats.u1.qWonArr, 295520);
  assert.strictEqual(stats.u2.fyWonArr, 0);
  assert.strictEqual(stats.u2.qWonArr || 0, 0);
  assert.strictEqual(stats.u1.meetings, 7);
  assert.strictEqual(stats.u1.activities, 40);
  const activity = queries.filter(q => q.includes('FROM Event') || q.includes('FROM Task'));
  assert.ok(activity.length >= 2);
  activity.forEach(q => assert.match(q, /ActivityDate >= 2026-08-01 AND ActivityDate < 2026-11-01 AND ActivityDate <= 2026-09-12/));
  assert.ok(queries.every(q => !q.includes('LAST_N_DAYS')));
  assert.ok(queries.some(q => q.includes('CreatedDate >= 2026-08-01T00:00:00Z AND CreatedDate < 2026-11-01T00:00:00Z')));
});
