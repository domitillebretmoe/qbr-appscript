// Menu, triggers and the refresh pipeline. A team tab is any sheet with "Team" in A1 and "Quarter" in A2.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('QBR')
    .addItem('Refresh this tab', 'refreshActiveTab')
    .addItem('Refresh all team tabs', 'refreshAllTabs')
    .addSeparator()
    .addItem('Set up workbook (all teams + ARR Ledger + Definitions)', 'setupWorkbook')
    .addItem('Add team tab...', 'addTeamTab')
    .addItem('Refresh on B1/B2 edit (install trigger)', 'installEditTrigger')
    .addSeparator()
    .addItem('Set Salesforce credentials...', 'setSalesforceCredentials')
    .addToUi();
}

function isTeamTab(sheet) {
  const [[a1], [a2]] = sheet.getRange('A1:A2').getValues();
  return a1 === 'Team' && a2 === 'Quarter';
}

function teamTabs() {
  return SpreadsheetApp.getActive().getSheets().filter(isTeamTab);
}

function refreshActiveTab() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!isTeamTab(sheet)) throw new Error('Select a team tab first (A1 = Team, A2 = Quarter).');
  withRefreshLock(() => refreshTab(sheet));
}

// One refresh at a time per workbook, so ARR Ledger rows and the pending-tabs queue are never written concurrently.
function withRefreshLock(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('Another QBR refresh is running on this workbook. Try again in a few minutes.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// Apps Script stops a run after 6 minutes, so a long refresh queues the remaining tabs for a one-off trigger.
const REFRESH_BUDGET_MS = 4.5 * 60 * 1000;
const PENDING_TABS_KEY = 'QBR_PENDING_TABS';
const CONTINUE_HANDLER = 'continueRefresh';

function refreshAllTabs() {
  refreshTabsNamed(teamTabs().map(sheet => sheet.getName()));
}

function refreshTabsNamed(names) {
  withRefreshLock(() => runRefreshQueue(names));
}

function runRefreshQueue(names) {
  const ss = SpreadsheetApp.getActive();
  const started = Date.now();
  const pending = names.slice();
  while (pending.length && Date.now() - started < REFRESH_BUDGET_MS) {
    const sheet = ss.getSheetByName(pending.shift());
    if (sheet && isTeamTab(sheet)) refreshTab(sheet);
  }
  if (pending.length) {
    const queued = takePendingTabs().filter(name => pending.indexOf(name) < 0);
    PropertiesService.getDocumentProperties().setProperty(PENDING_TABS_KEY, JSON.stringify(pending.concat(queued)));
    if (!continueTriggers().length) ScriptApp.newTrigger(CONTINUE_HANDLER).timeBased().after(1000).create();
    ss.toast(`Refreshed ${names.length - pending.length} tabs, ${pending.length + queued.length} more continue in the background`, 'QBR');
  } else {
    ss.toast(`Refreshed ${names.length} team tabs`, 'QBR');
  }
}

function continueTriggers() {
  return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === CONTINUE_HANDLER);
}

function takePendingTabs() {
  const props = PropertiesService.getDocumentProperties();
  const pending = JSON.parse(props.getProperty(PENDING_TABS_KEY) || '[]');
  props.deleteProperty(PENDING_TABS_KEY);
  return pending;
}

function continueRefresh(e) {
  withRefreshLock(() => {
    continueTriggers().filter(t => !e || !e.triggerUid || t.getUniqueId() === e.triggerUid).forEach(t => ScriptApp.deleteTrigger(t));
    const pending = takePendingTabs();
    if (pending.length) runRefreshQueue(pending);
  });
}

function refreshTab(sheet) {
  const team = String(sheet.getRange('B1').getValue()).trim();
  const quarter = String(sheet.getRange('B2').getValue()).trim();
  if (!team) throw new Error(`${sheet.getName()}: B1 must hold the team name.`);
  assertSpecificTeam(team);
  parseQuarter(quarter);
  renderTeamTab(sheet, buildView(team, quarter));
}

// Installable trigger target: re-renders a team tab when B1 or B2 changes (including a paste over B1:B2).
function onTabEdit(e) {
  const sheet = e.range.getSheet();
  const touchesSelector = e.range.getColumn() <= 2 && e.range.getLastColumn() >= 2 && e.range.getRow() <= 2;
  if (!isTeamTab(sheet) || !touchesSelector) return;
  withRefreshLock(() => refreshTab(sheet));
}

function installEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  const installed = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onTabEdit');
  if (!installed) ScriptApp.newTrigger('onTabEdit').forSpreadsheet(ss).onEdit().create();
  ss.toast('Editing B1 or B2 on a team tab now refreshes it', 'QBR');
}

function setupWorkbook() {
  writeDefinitionsSheet();
  ledgerSheet();
  TEAMS.forEach(team => teamSheet(team));
  refreshAllTabs();
}

function addTeamTab() {
  const response = SpreadsheetApp.getUi().prompt('Team name as in Salesforce (e.g. Europe - DACH)');
  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
  const sheet = teamSheet(response.getResponseText().trim());
  withRefreshLock(() => refreshTab(sheet));
}

function teamSheet(team) {
  const ss = SpreadsheetApp.getActive();
  const existing = ss.getSheetByName(team);
  if (existing) return existing;
  const sheet = ss.insertSheet(team);
  sheet.getRange('A1:B2').setValues([['Team', team], ['Quarter', defaultQuarter()]]);
  return sheet;
}

// The most recent quarter that has fully closed.
function defaultQuarter() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return shiftQuarter(quarterOfDate(today), -1);
}

function setSalesforceCredentials() {
  const ui = SpreadsheetApp.getUi();
  const ask = (name, hint, normalize) => {
    const response = ui.prompt(`${name}`, hint, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) throw new Error('Cancelled');
    const value = response.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty(name, normalize ? normalize(value) : value);
  };
  ask('SF_LOGIN_URL', 'My Domain URL, e.g. https://codeium.my.salesforce.com', expectSalesforceUrl);
  ask('SF_CLIENT_ID', 'Connected App consumer key');
  ask('SF_CLIENT_SECRET', 'Connected App consumer secret');
  sfSession = null;
  sfConnect();
  ui.alert('Salesforce connection OK');
}

// Everything a team tab needs, computed from Salesforce + the ARR Ledger.
function buildView(team, quarter) {
  const next1 = shiftQuarter(quarter, 1);
  const next2 = shiftQuarter(quarter, 2);
  const opps = fetchOpportunities(team, parseQuarter(next2).fy);
  const accounts = fetchAccounts(team);
  const goals = fetchGoals(team);
  const goalFor = q => goals[q] || { revenue: 0, logos: 0 };

  const quarters = quartersBetween(FIRST_QUARTER, quarter);
  const actuals = quarters.map(q => Object.assign(quarterMetrics(opps, q, goalFor(q)), accountMetrics(accounts, opps, q), partnerMetrics(opps, q)));
  const netAddedByQuarter = {};
  actuals.forEach(m => { netAddedByQuarter[m.quarter] = m.netAddedArr; });
  const ledger = rollLedger(team, quarter, netAddedByQuarter);
  const trend = actuals.map(m => withArr(m, ledger[m.quarter]));
  const current = trend[trend.length - 1];

  const future1 = forecastMetrics(opps, next1, current.endingArr, goalFor(next1));
  const future2 = forecastMetrics(opps, next2, future1.forecastEndingArr, goalFor(next2));
  return { team, quarter, trend, current, future: [future1, future2] };
}
