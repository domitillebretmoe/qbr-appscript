// "Raw Data" sheet: every Salesforce opportunity behind the team tabs, one row per opportunity, so teams can
// filter to their tab and tie the numbers out. Refreshing a tab replaces that tab's rows only.
const RAW_HEADER = ['Tab', 'Resolved Team', 'Quarter', 'Account', 'Opportunity', 'Stage', 'Type', 'Record Type', 'Bucket',
  'Close Date', 'Delta ARR', 'Deal Value (TCV)', 'Expected Delta ARR', 'Expected Logo Impact', 'Major Account', 'Group', 'Opp Team', 'Owner',
  'Lost Reason', 'Pilot Status', 'Pilot End Date', 'Pilot End Quarter', 'Salesforce Id', 'Opportunity URL', 'Account URL'];
const RAW_MONEY_COLS = [11, 12, 13];

// How the opportunity is counted in the metrics (see Definitions).
function rawBucket(opp) {
  if (isWon(opp)) return isRenewal(opp) ? (opp.deltaArr < 0 ? 'Won renewal - downgrade' : 'Won renewal') : `Closed Won - ${opp.type || 'other'}`;
  if (isLost(opp)) return isRenewal(opp) ? 'Full churn' : 'Lost pipeline';
  return 'Open pipeline';
}

function rawRow(tab, opp) {
  return [tab, opp.team, opp.quarter, opp.account, opp.name, opp.stage, opp.type, opp.recordType, rawBucket(opp),
    opp.closeDate, opp.deltaArr, opp.amount || 0, opp.expectedDeltaArr, opp.expectedLogoImpact, opp.major ? 'Yes' : 'No', opp.oppGroup, opp.oppTeam,
    opp.owner, opp.lostReason, opp.pilotStatus || '', opp.pilotEndDate || '', opp.pilotEndDate ? quarterOfDate(opp.pilotEndDate) : '',
    opp.id, opp.url || '', opp.accountUrl || ''];
}

function rawSheet() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(RAW_SHEET) || ss.insertSheet(RAW_SHEET);
}

// Other tabs' rows, remapped by header name onto RAW_HEADER so rows written by an older column layout keep
// their values under the right headers (columns the old layout lacked are left blank).
function keptRawRows(sheet, tab) {
  const width = Math.max(sheet.getLastColumn(), RAW_HEADER.length);
  const header = sheet.getRange(1, 1, 1, width).getValues()[0].map(String);
  const from = RAW_HEADER.map(name => header.indexOf(name));
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues()
    .filter(row => row[0] !== tab)
    .map(row => from.map(i => (i < 0 ? '' : row[i])));
}

// Replaces the rows of `tab` with `opps`, keeps every other tab's rows, and rewrites the sheet sorted.
function writeRawData(tab, opps) {
  const sheet = rawSheet();
  const kept = sheet.getLastRow() > 1 ? keptRawRows(sheet, tab) : [];
  const rows = kept.concat(opps.map(opp => rawRow(tab, opp)))
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])) || String(a[2]).localeCompare(String(b[2])) || String(a[3]).localeCompare(String(b[3])));

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();
  sheet.setHiddenGridlines(false);
  sheet.getRange(1, 1, 1, RAW_HEADER.length).setValues([RAW_HEADER]).setFontWeight('bold')
    .setBackground(COLORS.ink).setFontColor('#ffffff');
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, RAW_HEADER.length).setValues(rows);
    RAW_MONEY_COLS.forEach(c => sheet.getRange(2, c, rows.length, 1).setNumberFormat(FORMATS.money));
    sheet.getRange(2, 14, rows.length, 1).setNumberFormat('0.##');
  }
  sheet.getRange(1, 1, Math.max(2, rows.length + 1), RAW_HEADER.length).setFontFamily(FONT).setFontSize(9);
  sheet.getRange(1, 1, rows.length + 1, RAW_HEADER.length).createFilter();
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidths(1, RAW_HEADER.length, 110);
  [4, 5].forEach(c => sheet.setColumnWidth(c, 220));
  sheet.setColumnWidth(9, 170);
  sheet.setColumnWidth(19, 200);
  sheet.setColumnWidths(21, 2, 260);
  return sheet;
}
