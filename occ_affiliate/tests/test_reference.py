"""Tests for reference.py.

Unit tests use synthetic numbers. Workbook regression tests run only when
OCC_WORKBOOK points at a copy of Daily_Hanweck_Logs.xlsx (not committed: it holds
internal risk data). They compare against the values Excel cached on save.
"""

import os
import sys
from datetime import date, datetime
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import reference as ref  # noqa: E402

APEX_BODY = """Hello Webull Team,

Please see the OCC and Apex 157 requirement amounts for {d}:

    Daily OCC 157 Requirement: ${daily:,.2f}
    Monthly Intraday Exposure Add-On (Effective 2026-01-01): ${addon:,.2f}
    Fluctuated Daily OCC Requirement (after subtracting Monthly Add-On): ${fluct:,.2f}
    Apex Omni Requirement (1.4x Daily OCC + Monthly Add-On): ${omni:,.2f}

    Account Equities:
    5XX00000: $-1,234.56
Best regards,
Apex
"""


def apex_body(d, fluct, omni, addon=1_000_000.0):
    return APEX_BODY.format(d=d, daily=fluct + addon, addon=addon, fluct=fluct, omni=omni)


def csv_text(day, mreqt_by_acct, date_fmt="%Y%m%d"):
    lines = [",".join(ref.DATA_COLUMNS)]
    for acct, v in mreqt_by_acct.items():
        group = "ACCT" if acct == "BULL" else "SUBACCT"
        nums = [0.0] * 13
        nums[2] = v  # MREQT
        lines.append(",".join([day.strftime(date_fmt), "Apex", group, acct] + [str(x) for x in nums]))
    return "\n".join(lines) + "\n"


SYNTH = {"BULL": -100.0, "AU": -10.0, "EU": 5.0, "HK": -20.0, "JP": 0.0,
         "MY": -5.0, "SG": -10.0, "TH": 0.0, "US": -60.0}


# ---------- unit tests ----------

def test_parse_apex_email():
    a = ref.parse_apex_email(apex_body("2026-01-02", 184_613_677.48, 317_115_733.99))
    assert a.date == date(2026, 1, 2)
    assert a.fluctuated == pytest.approx(184_613_677.48)
    assert a.omni == pytest.approx(317_115_733.99)
    assert a.daily_occ == pytest.approx(185_613_677.48)


def test_parse_apex_email_html():
    html = "<p>" + apex_body("2026-01-02", 1.5, 2.5).replace("\n", "<br>") + "</p>"
    a = ref.parse_apex_email(html)
    assert (a.fluctuated, a.omni) == (1.5, 2.5)


def test_parse_apex_email_missing_field():
    with pytest.raises(ValueError, match="omni"):
        ref.parse_apex_email(apex_body("2026-01-02", 1, 2).replace("Apex Omni", "Other"))


@pytest.mark.parametrize("fmt", ["%Y%m%d", "%Y-%m-%d", "%m/%d/%Y"])
def test_parse_stansattr_dates(fmt):
    rows = ref.parse_stansattr("﻿" + csv_text(date(2026, 1, 2), SYNTH, fmt))
    assert {r["DATE"] for r in rows} == {date(2026, 1, 2)}
    assert len(rows) == len(SYNTH)


def test_parse_stansattr_bad_header():
    with pytest.raises(ValueError, match="header"):
        ref.parse_stansattr("FOO,BAR\n1,2\n")


def test_compute_synthetic():
    d = date(2026, 1, 2)
    res = ref.compute(ref.parse_stansattr(csv_text(d, SYNTH)), ref.parse_apex_email(apex_body(d, 98.0, 200.0)))
    assert res.simulated_total == -100.0
    assert res.proportion["US"] == pytest.approx(0.6)
    assert res.adjusted["EU"] == 0.0            # negative share floored at 0
    assert res.adjusted["US"] == pytest.approx(120.0)
    assert res.adjusted_total == pytest.approx(210.0)  # > apex total: MAX(.,0) not renormalised
    assert res.apex_req == -98.0
    assert res.diff_pct == pytest.approx((-100 + 98) / -100)
    assert not res.alert
    html = ref.render_html(res)
    assert "US_OMNI" in html and "$ 120.00" in html and "01/02/2026" in html


def test_compute_alert_and_date_mismatch():
    d = date(2026, 1, 2)
    rows = ref.parse_stansattr(csv_text(d, SYNTH))
    assert ref.compute(rows, ref.parse_apex_email(apex_body(d, 90.0, 1.0))).alert
    with pytest.raises(ValueError, match="mismatch"):
        ref.compute(rows, ref.parse_apex_email(apex_body("2026-01-05", 90.0, 1.0)))


# ---------- workbook regression (opt-in) ----------

WB = os.environ.get("OCC_WORKBOOK")
needs_wb = pytest.mark.skipif(not WB, reason="set OCC_WORKBOOK to run workbook regression tests")


@pytest.fixture(scope="module")
def workbook():
    import openpyxl
    return openpyxl.load_workbook(WB, data_only=True)


def _rows_by_date(wb):
    out = {}
    for r in wb["data"].iter_rows(min_row=2, max_col=17, values_only=True):
        if not isinstance(r[0], datetime):
            continue
        rec = dict(zip(ref.DATA_COLUMNS, r))
        rec["DATE"] = r[0].date()
        rec["MREQT"] = float(rec["MREQT"] or 0)
        out.setdefault(rec["DATE"], []).append(rec)
    return out


def _pivot_rows(wb):
    for r in wb["pivot"].iter_rows(min_row=24, max_col=19, values_only=True):
        if isinstance(r[0], datetime):
            yield r


@needs_wb
def test_backtest_against_pivot_history(workbook):
    rows = _rows_by_date(workbook)
    checked = 0
    for p in _pivot_rows(workbook):
        d, apex_l, cached_n = p[0].date(), p[11], p[13]
        if apex_l is None:
            continue
        res = ref.compute(rows[d], ref.parse_apex_email(apex_body(d.isoformat(), -apex_l, 1.0)))
        assert res.hanweck_bull == pytest.approx(p[9], abs=0.01), d
        for col, acct in zip(p[1:9], ["US", "AU", "EU", "HK", "JP", "MY", "SG", "TH"]):
            assert res.simulated[acct] == pytest.approx(col or 0.0, abs=0.01), (d, acct)
        if isinstance(cached_n, float):
            assert res.diff_pct == pytest.approx(cached_n, abs=1e-9), d
        checked += 1
    assert checked > 100


@needs_wb
def test_latest_email_table_matches_workbook(workbook):
    e = workbook["email"]
    d = e["B2"].value.date()
    omni = e["C7"].value
    apex_l = next(p[11] for p in _pivot_rows(workbook) if p[0].date() == d)
    res = ref.compute(_rows_by_date(workbook)[d], ref.parse_apex_email(apex_body(d.isoformat(), -apex_l, omni)))
    for col, acct in zip("CDEFGHIJ", ref.AFFILIATES):
        assert res.simulated[acct] == pytest.approx(e[f"{col}4"].value, abs=0.01), acct
        assert res.adjusted[acct] == pytest.approx(e[f"{col}10"].value, abs=0.01), acct
    assert res.adjusted_total == pytest.approx(e["K10"].value, abs=0.01)
