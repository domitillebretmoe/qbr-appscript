// "ARR Ledger" sheet: Team | Quarter | Starting ARR | Net Added ARR | Ending ARR | Source.
// Rows stamped "Salesforce <date>" are recomputed on every refresh. Any other Source (the seeded rows from last
// quarter's workbook, or a value typed by hand such as "Locked") is kept exactly as written.
const LEDGER_HEADER = ['Team', 'Quarter', 'Starting ARR', 'Net Added ARR', 'Ending ARR', 'Source'];
const LIVE_SOURCE_PREFIX = 'Salesforce ';

function isFixedLedgerRow(row) {
  return !!row && String(row.source).indexOf(LIVE_SOURCE_PREFIX) !== 0;
}

function ledgerSheet() {
  const ss = SpreadsheetApp.getActive();
  const existing = ss.getSheetByName(LEDGER_SHEET);
  if (existing) return existing;
  const sheet = ss.insertSheet(LEDGER_SHEET);
  const rows = [];
  TEAM_SEEDS.forEach(([team, q1Start, q1End, q2End]) => {
    rows.push([team, 'Q1-2026', q1Start, q1End - q1Start, q1End, SEED_SOURCE]);
    rows.push([team, 'Q2-2026', q1End, q2End - q1End, q2End, SEED_SOURCE]);
  });
  sheet.getRange(1, 1, 1, LEDGER_HEADER.length).setValues([LEDGER_HEADER]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, LEDGER_HEADER.length).setValues(rows);
  sheet.getRange(2, 3, rows.length, 3).setNumberFormat('$#,##0.00');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, LEDGER_HEADER.length);
  return sheet;
}

function ledgerRows(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, LEDGER_HEADER.length).getValues()
    .map((row, i) => ({ rowNumber: i + 2, team: row[0], quarter: row[1], startingArr: row[2], netAddedArr: row[3], endingArr: row[4], source: row[5] }));
}

function ledgerEntry(rows, team, quarter) {
  return rows.filter(r => r.team === team && r.quarter === quarter)[0] || null;
}

// Computes Starting/Ending ARR for every quarter from FIRST_QUARTER to lastQuarter, writing non-seeded rows back.
// netAddedByQuarter maps quarter label -> Net Added ARR from Salesforce.
function rollLedger(team, lastQuarter, netAddedByQuarter) {
  const sheet = ledgerSheet();
  const rows = ledgerRows(sheet);
  const first = ledgerEntry(rows, team, FIRST_QUARTER);
  if (!first) throw new Error(`ARR Ledger has no ${FIRST_QUARTER} row for "${team}". Add its Starting ARR first.`);

  const source = `${LIVE_SOURCE_PREFIX}${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`;
  const result = {};
  let startingArr = Number(first.startingArr);
  quartersBetween(FIRST_QUARTER, lastQuarter).forEach(quarter => {
    const existing = ledgerEntry(rows, team, quarter);
    if (isFixedLedgerRow(existing)) {
      result[quarter] = { startingArr: Number(existing.startingArr), netAddedArr: Number(existing.netAddedArr), endingArr: Number(existing.endingArr) };
    } else {
      if (netAddedByQuarter[quarter] == null) throw new Error(`No Net Added ARR for ${team} ${quarter}`);
      const netAddedArr = netAddedByQuarter[quarter];
      const endingArr = startingArr + netAddedArr;
      result[quarter] = { startingArr, netAddedArr, endingArr };
      writeLedgerRow(sheet, existing, [team, quarter, startingArr, netAddedArr, endingArr, source]);
    }
    startingArr = result[quarter].endingArr;
  });
  return result;
}

function writeLedgerRow(sheet, existing, values) {
  const rowNumber = existing ? existing.rowNumber : sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  sheet.getRange(rowNumber, 3, 1, 3).setNumberFormat('$#,##0.00');
}
