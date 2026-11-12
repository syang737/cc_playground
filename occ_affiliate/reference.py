"""Reference implementation of the daily Affiliate OCC 157 Requirement calculation.

Mirrors the Office Scripts (office_scripts/*.ts) so the logic can be tested in
Python, and serves as the fallback runner if Power Automate is unavailable.

Usage:
    python reference.py --csv apex-STANSATTR-BULL_20261112_060000.csv --apex-body apex_email.txt [--html out.html]
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime

# Order/labels used by the `email` tab. Header row 3 uses pivot ACCTIDs; row 9 uses affiliate names.
AFFILIATES = ["AU", "EU", "HK", "JP", "MY", "SG", "TH", "US"]
ADJ_LABELS = {"MY": "ML", "US": "US_OMNI"}
DATA_COLUMNS = ["DATE", "PID", "GROUP", "ACCTID", "PRISK", "PNAV", "MREQT", "MRISK", "MNAV",
                "RREQT", "RRISK", "RNAV", "BREQT", "BRISK", "BNAV", "NRFACT", "NPOS"]
ALERT_THRESHOLD = 0.03


@dataclass(frozen=True)
class ApexReq:
    date: date
    daily_occ: float
    monthly_addon: float
    fluctuated: float  # "Fluctuated Daily OCC Requirement" -> pivot!L (stored negated)
    omni: float        # "Apex Omni Requirement" -> email!C7


def _money(s: str) -> float:
    return float(s.replace(",", "").replace("$", ""))


_AMT = r"\$\s*(-?[\d,]+(?:\.\d+)?)"
_APEX_PATTERNS = {
    "date": r"requirement amounts for\s+(\d{4}-\d{2}-\d{2})",
    "daily_occ": r"Daily OCC 157 Requirement:\s*" + _AMT,
    "monthly_addon": r"Monthly Intraday Exposure Add-On[^:]*:\s*" + _AMT,
    "fluctuated": r"Fluctuated Daily OCC Requirement[^:]*:\s*" + _AMT,
    "omni": r"Apex Omni Requirement[^:]*:\s*" + _AMT,
}


def parse_apex_email(body: str) -> ApexReq:
    """Extract the requirement figures from Apex's daily 'OCC 157 Req' email (plain text or HTML)."""
    text = re.sub(r"<[^>]+>", " ", body).replace("&nbsp;", " ").replace("&#36;", "$")
    found = {}
    for key, pat in _APEX_PATTERNS.items():
        m = re.search(pat, text, re.IGNORECASE)
        if not m:
            raise ValueError(f"Apex email: could not find {key!r}")
        found[key] = m.group(1)
    return ApexReq(
        date=datetime.strptime(found["date"], "%Y-%m-%d").date(),
        **{k: _money(v) for k, v in found.items() if k != "date"},
    )


def _parse_date(s: str) -> date:
    s = s.strip()
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unrecognised DATE value: {s!r}")


def parse_stansattr(csv_text: str) -> list[dict]:
    """Parse apex-STANSATTR-BULL_*.csv into rows matching the workbook's `data` tab."""
    reader = csv.reader(io.StringIO(csv_text.lstrip("﻿")))
    rows = [r for r in reader if any(c.strip() for c in r)]
    header = [h.strip().upper() for h in rows[0]]
    if header[: len(DATA_COLUMNS)] != DATA_COLUMNS:
        raise ValueError(f"Unexpected STANSATTR header: {header}")
    out = []
    for r in rows[1:]:
        rec = dict(zip(DATA_COLUMNS, (c.strip() for c in r)))
        rec["DATE"] = _parse_date(rec["DATE"])
        for k in DATA_COLUMNS[4:]:
            rec[k] = float(rec[k]) if rec[k] != "" else 0.0
        out.append(rec)
    if len({r["DATE"] for r in out}) != 1:
        raise ValueError("STANSATTR file contains more than one DATE")
    return out


@dataclass(frozen=True)
class DailyResult:
    date: date
    simulated: dict[str, float]   # email row 4 (sub-account MREQT)
    proportion: dict[str, float]  # email row 5
    adjusted: dict[str, float]    # email row 10
    apex_total: float             # email C7
    hanweck_bull: float           # pivot J (BULL MREQT)
    apex_req: float               # pivot L (negated fluctuated)
    diff_pct: float               # pivot N = (J - L) / J

    @property
    def simulated_total(self) -> float:
        return sum(self.simulated.values())

    @property
    def adjusted_total(self) -> float:
        return sum(self.adjusted.values())

    @property
    def alert(self) -> bool:
        return abs(self.diff_pct) > ALERT_THRESHOLD


def compute(rows: list[dict], apex: ApexReq) -> DailyResult:
    """Replicates the pivot + `email` tab formulas for a single day."""
    day = rows[0]["DATE"]
    if apex.date != day:
        raise ValueError(f"Date mismatch: Hanweck {day} vs Apex {apex.date}")

    mreqt: dict[str, float] = {}
    for r in rows:
        mreqt[r["ACCTID"]] = mreqt.get(r["ACCTID"], 0.0) + r["MREQT"]  # pivot uses SUM
    if "BULL" not in mreqt:
        raise ValueError("STANSATTR file has no BULL (ACCT) row")

    simulated = {a: mreqt.get(a, 0.0) for a in AFFILIATES}
    total = sum(simulated.values())
    proportion = {a: v / total for a, v in simulated.items()}
    # Matches email!C10:J10 = MAX(prop * ApexTotal, 0). NB: not renormalised, so the
    # sum can exceed apex_total when some affiliates have negative share.
    adjusted = {a: max(p * apex.omni, 0.0) for a, p in proportion.items()}

    j = mreqt["BULL"]
    l = -apex.fluctuated
    return DailyResult(day, simulated, proportion, adjusted, apex.omni, j, l, (j - l) / j)


def _fmt_acct(v: float) -> str:
    return f"({abs(v):,.2f})" if v < 0 else f"{v:,.2f}"


def render_html(res: DailyResult) -> str:
    """HTML table equivalent to the `email` tab, for pasting/drafting."""
    td = 'style="border:1px solid #999;padding:2px 8px;text-align:right"'
    th = 'style="border:1px solid #999;padding:2px 8px;background:#DDEBF7"'
    lab = 'style="border:1px solid #999;padding:2px 8px;text-align:left"'
    head = "".join(f"<th {th}>{a}</th>" for a in AFFILIATES) + f"<th {th}>Total</th>"
    sim = "".join(f"<td {td}>{_fmt_acct(res.simulated[a])}</td>" for a in AFFILIATES)
    prop = "".join(f"<td {td}>{res.proportion[a]:.2%}</td>" for a in AFFILIATES)
    adj_head = "".join(f"<th {th}>{ADJ_LABELS.get(a, a)}</th>" for a in AFFILIATES) + f"<th {th}>Total</th>"
    adj = "".join(
        f"<td {td}>{'$ -' if res.adjusted[a] == 0 else f'$ {res.adjusted[a]:,.2f}'}</td>" for a in AFFILIATES
    )
    return f"""<p>Affiliate OCC Requirement below:</p>
<table style="border-collapse:collapse;font-family:Calibri,Arial;font-size:11pt">
<tr><th {th}>Date</th><th {th} colspan="9">Simulated OCC Requirement</th></tr>
<tr><td {lab}>{res.date:%m/%d/%Y}</td>{head}</tr>
<tr><td {lab}>Simulated Requirement</td>{sim}<td {td}>{_fmt_acct(res.simulated_total)}</td></tr>
<tr><td {lab}>Proportion</td>{prop}<td {td}></td></tr>
</table><br>
<table style="border-collapse:collapse;font-family:Calibri,Arial;font-size:11pt">
<tr><td {lab}><b>Apex Total Requirement</b></td><td {td} style="background:#FCE4D6"><b>$ {res.apex_total:,.0f}</b></td></tr>
</table><br>
<table style="border-collapse:collapse;font-family:Calibri,Arial;font-size:11pt">
<tr><th {th}></th>{adj_head}</tr>
<tr><td {lab}><b>Total Adj. Requirement</b></td>{adj}<td {td}>$ {res.adjusted_total:,.2f}</td></tr>
</table>
<p><i>Total Adj. Requirement = Fluctuated Daily OCC Requirement * 1.4 + Monthly Add-On Charge</i></p>"""


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", required=True, help="apex-STANSATTR-BULL_*.csv")
    ap.add_argument("--apex-body", required=True, help="text/HTML body of Apex's OCC 157 Req email")
    ap.add_argument("--html", help="write the email HTML here")
    args = ap.parse_args(argv)

    with open(args.csv, encoding="utf-8-sig") as f:
        rows = parse_stansattr(f.read())
    with open(args.apex_body, encoding="utf-8") as f:
        apex = parse_apex_email(f.read())
    res = compute(rows, apex)

    print(f"Date {res.date}  Hanweck {res.hanweck_bull:,.2f}  Apex {res.apex_req:,.2f}  Diff {res.diff_pct:.2%}")
    for a in AFFILIATES:
        print(f"  {ADJ_LABELS.get(a, a):8s} {res.adjusted[a]:>18,.2f}")
    print(f"  {'Total':8s} {res.adjusted_total:>18,.2f}  (Apex total {res.apex_total:,.2f})")
    if res.alert:
        print(f"ALERT: Hanweck vs Apex diff {res.diff_pct:.2%} exceeds {ALERT_THRESHOLD:.0%}", file=sys.stderr)
    if args.html:
        with open(args.html, "w", encoding="utf-8") as f:
            f.write(render_html(res))
    return 2 if res.alert else 0


if __name__ == "__main__":
    raise SystemExit(main())
