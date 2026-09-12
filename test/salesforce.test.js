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
  ['Config.gs', 'Metrics.gs', 'Salesforce.gs'].forEach(f => vm.runInContext(fs.readFileSync(`${__dirname}/../src/${f}`, 'utf8'), ctx));
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
