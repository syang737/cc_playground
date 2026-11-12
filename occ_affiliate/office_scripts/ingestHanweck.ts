/**
 * Office Script: ingestHanweck
 * Appends the apex-STANSATTR-BULL_*.csv rows to the `data` tab and refreshes the pivots.
 * Idempotent: if the CSV's DATE is already in `data`, nothing is appended.
 *
 * Power Automate input:  csv  = base64ToString(<attachment ContentBytes>)
 * Manual testing:        leave `csv` blank and paste the CSV file's text (from Notepad) into
 *                        cell A1 of a sheet named `csv_input`; one line per row in column A.
 * Returns:               { date: "YYYY-MM-DD", rowsAdded, skipped }
 */

const DATA_COLUMNS = ["DATE", "PID", "GROUP", "ACCTID", "PRISK", "PNAV", "MREQT", "MRISK", "MNAV",
  "RREQT", "RRISK", "RNAV", "BREQT", "BRISK", "BNAV", "NRFACT", "NPOS"];

interface IngestResult {
  date: string;
  rowsAdded: number;
  skipped: boolean;
}

function main(workbook: ExcelScript.Workbook, csv?: string): IngestResult {
  if (!csv || csv.trim() === "") csv = readCsvSheet(workbook);
  const lines = csv.replace(/^\uFEFF/, "").split(/\r?\n/).filter(l => l.trim() !== "");
  const header = lines[0].split(",").map(h => h.trim().toUpperCase());
  if (header.slice(0, DATA_COLUMNS.length).join(",") !== DATA_COLUMNS.join(",")) {
    throw new Error(`Unexpected STANSATTR header: ${header.join(",")}`);
  }

  const rows: (string | number)[][] = [];
  let serial: number | null = null;
  for (const line of lines.slice(1)) {
    const c = line.split(",").map(x => x.trim());
    const s = toSerial(c[0]);
    if (serial !== null && s !== serial) throw new Error("STANSATTR file contains more than one DATE");
    serial = s;
    const nums = c.slice(4, DATA_COLUMNS.length).map(x => (x === "" ? 0 : Number(x)));
    if (nums.some(n => isNaN(n))) throw new Error(`Non-numeric value in row: ${line}`);
    rows.push([s, c[1], c[2], c[3], ...nums]);
  }
  if (serial === null) throw new Error("STANSATTR file has no data rows");
  const iso = serialToIso(serial);

  const data = sheet(workbook, "data");
  // Column A only: column T holds unrelated notes, so the sheet's used range is misleading.
  const colA = data.getRange("A:A").getUsedRange(true);
  const lastRow = colA.getRowIndex() + colA.getRowCount(); // 1-based last used row
  const existing = colA.getValues().map(r => r[0]);
  if (existing.some(v => typeof v === "number" && Math.floor(v) === serial)) {
    return { date: iso, rowsAdded: 0, skipped: true };
  }

  const target = data.getRangeByIndexes(lastRow, 0, rows.length, DATA_COLUMNS.length);
  target.setValues(rows);
  target.getColumn(0).setNumberFormatLocal("mm-dd-yy");

  workbook.refreshAllPivotTables();
  return { date: iso, rowsAdded: rows.length, skipped: false };
}

/** Rebuilds CSV text from the `csv_input` sheet (works whether Excel split the columns or not). */
function readCsvSheet(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet("csv_input");
  if (!sheet) throw new Error("csv is blank and there is no `csv_input` sheet to read from");
  const used = sheet.getUsedRange(true);
  if (!used) throw new Error("`csv_input` sheet is empty");
  return used.getValues()
    .map(r => r.filter(c => c !== "").map(c => String(c)).join(","))
    .filter(l => l !== "")
    .join("\n");
}

/** Accepts YYYYMMDD, YYYY-MM-DD, M/D/YYYY. Returns the Excel date serial. */
function toSerial(s: string): number {
  let y: number, m: number, d: number;
  let r: RegExpMatchArray | null;
  if ((r = s.match(/^(\d{4})(\d{2})(\d{2})$/)) || (r = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) {
    [y, m, d] = [Number(r[1]), Number(r[2]), Number(r[3])];
  } else if ((r = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) {
    [y, m, d] = [Number(r[3]), Number(r[1]), Number(r[2])];
  } else {
    throw new Error(`Unrecognised DATE value: ${s}`);
  }
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

function serialToIso(serial: number): string {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 10);
}

function sheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const ws = workbook.getWorksheet(name);
  if (!ws) throw new Error(`Worksheet \`${name}\` not found`);
  return ws;
}
