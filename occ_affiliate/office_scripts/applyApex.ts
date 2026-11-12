/**
 * Office Script: applyApex
 * Parses Apex's "OCC 157 Req" email, writes the Apex figures into the workbook, and
 * returns the Hanweck-vs-Apex check plus the `email` tab rendered as HTML.
 * Run AFTER ingestHanweck (so the pivot has the new day).
 *
 * Power Automate inputs: emailBody   = Body of the Apex email (HTML or text)
 *                        hanweckDate = `date` output of ingestHanweck ("YYYY-MM-DD")
 */

const ALERT_THRESHOLD = 0.03;
const PIVOT_FIRST_ROW = 24; // first data row under the pivot headers (row 23)

interface ApexResult {
  date: string;
  hanweck: number;   // pivot J (BULL)
  apex: number;      // pivot L (negated Fluctuated Daily OCC Requirement)
  diffPct: number;   // pivot N
  alert: boolean;
  omni: number;
  html: string;
}

function main(workbook: ExcelScript.Workbook, emailBody: string, hanweckDate: string): ApexResult {
  const apex = parseApex(emailBody);
  if (apex.date !== hanweckDate) {
    throw new Error(`Date mismatch: Hanweck ${hanweckDate} vs Apex ${apex.date}`);
  }
  const serial = isoToSerial(apex.date);

  // ---- pivot tab: Apex Req (L) + Diff/DoD formulas (M..S) ----
  const pivot = workbook.getWorksheet("pivot");
  const colA = pivot.getRange(`A${PIVOT_FIRST_ROW}:A2000`).getValues().map(r => r[0]);
  const idx = colA.findIndex(v => typeof v === "number" && Math.floor(v) === serial);
  if (idx < 0) throw new Error(`${apex.date} not found in pivot!A — did ingestHanweck run?`);
  const r = PIVOT_FIRST_ROW + idx;

  pivot.getRange(`L${r}`).setValue(-apex.fluctuated);
  // Fill formulas for this row and any earlier rows left blank (historically O..S lag a day).
  for (let row = Math.max(PIVOT_FIRST_ROW + 1, r - 5); row <= r; row++) {
    if (pivot.getRange(`L${row}`).getValue() === "") continue;
    pivot.getRange(`M${row}:N${row}`).setFormulas([[`=J${row}-L${row}`, `=M${row}/J${row}`]]);
    pivot.getRange(`O${row}:S${row}`).setFormulas([[
      `=J${row - 1}-J${row}`, `=L${row - 1}-L${row}`,
      `=SUM(B${row}:I${row})`, `=Q${row}-L${row}`, `=R${row}/L${row}`,
    ]]);
  }

  // ---- email tab ----
  const email = workbook.getWorksheet("email");
  email.getRange("C7").setValue(apex.omni);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const emailDate = email.getRange("B2").getValue() as number;
  if (Math.floor(emailDate) !== serial) {
    throw new Error(`email!B2 shows ${serialToIso(emailDate)}, expected ${apex.date}`);
  }

  const hanweck = pivot.getRange(`J${r}`).getValue() as number;
  const diffPct = pivot.getRange(`N${r}`).getValue() as number;
  return {
    date: apex.date,
    hanweck,
    apex: -apex.fluctuated,
    diffPct,
    alert: Math.abs(diffPct) > ALERT_THRESHOLD,
    omni: apex.omni,
    html: renderHtml(email.getRange("B1:K11").getValues(), apex.date),
  };
}

interface Apex { date: string; fluctuated: number; omni: number; }

function parseApex(body: string): Apex {
  const text = body.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&#36;/g, "$");
  const amt = String.raw`\$\s*(-?[\d,]+(?:\.\d+)?)`;
  const grab = (label: string, re: RegExp): string => {
    const m = text.match(re);
    if (!m) throw new Error(`Apex email: could not find ${label}`);
    return m[1];
  };
  const money = (s: string) => Number(s.replace(/,/g, ""));
  return {
    date: grab("date", /requirement amounts for\s+(\d{4}-\d{2}-\d{2})/i),
    fluctuated: money(grab("Fluctuated Daily OCC Requirement",
      new RegExp(String.raw`Fluctuated Daily OCC Requirement[^:]*:\s*` + amt, "i"))),
    omni: money(grab("Apex Omni Requirement", new RegExp(String.raw`Apex Omni Requirement[^:]*:\s*` + amt, "i"))),
  };
}

/** Renders email!B1:K11 (values) like the manually pasted table. */
function renderHtml(v: (string | number | boolean)[][], isoDate: string): string {
  const n = (x: string | number | boolean) => (typeof x === "number" ? x : 0);
  const acct = (x: number) => (x < 0 ? `(${fmt(-x)})` : fmt(x));
  const dollars = (x: number) => (x === 0 ? "$ -" : `$ ${fmt(x)}`);
  const td = 'style="border:1px solid #999;padding:2px 8px;text-align:right"';
  const th = 'style="border:1px solid #999;padding:2px 8px;background:#DDEBF7"';
  const lab = 'style="border:1px solid #999;padding:2px 8px;text-align:left"';
  const tbl = 'style="border-collapse:collapse;font-family:Calibri,Arial;font-size:11pt"';
  const cols = [1, 2, 3, 4, 5, 6, 7, 8]; // C..J relative to B
  const [y, m, d] = isoDate.split("-");

  const head = cols.map(c => `<th ${th}>${v[2][c]}</th>`).join("") + `<th ${th}>Total</th>`;
  const sim = cols.map(c => `<td ${td}>${acct(n(v[3][c]))}</td>`).join("") + `<td ${td}>${acct(n(v[3][9]))}</td>`;
  const prop = cols.map(c => `<td ${td}>${(n(v[4][c]) * 100).toFixed(2)}%</td>`).join("") + `<td ${td}></td>`;
  const adjHead = cols.map(c => `<th ${th}>${v[8][c]}</th>`).join("") + `<th ${th}>Total</th>`;
  const adj = cols.map(c => `<td ${td}>${dollars(n(v[9][c]))}</td>`).join("") + `<td ${td}>${dollars(n(v[9][9]))}</td>`;
  const apexTotal = Math.round(n(v[6][1])).toLocaleString("en-US");

  return `<p>Affiliate OCC Requirement below:</p>
<table ${tbl}>
<tr><th ${th}>Date</th><th ${th} colspan="9">Simulated OCC Requirement</th></tr>
<tr><td ${lab}>${m}/${d}/${y}</td>${head}</tr>
<tr><td ${lab}>Simulated Requirement</td>${sim}</tr>
<tr><td ${lab}>Proportion</td>${prop}</tr>
</table><br>
<table ${tbl}><tr><td ${lab}><b>Apex Total Requirement</b></td><td ${td}><b>$ ${apexTotal}</b></td></tr></table><br>
<table ${tbl}>
<tr><th ${th}></th>${adjHead}</tr>
<tr><td ${lab}><b>Total Adj. Requirement</b></td>${adj}</tr>
</table>
<p><i>${v[10][0]}</i></p>`;
}

function fmt(x: number): string {
  return x.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function isoToSerial(iso: string): number {
  const [y, m, d] = iso.split("-").map(Number);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

function serialToIso(serial: number): string {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 10);
}
