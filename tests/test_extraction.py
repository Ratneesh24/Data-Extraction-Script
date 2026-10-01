"""Regression tests on the synthetic multi-layout ZIP.  Run:  python -m pytest -q"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import extract_production as ep  # noqa: E402
import make_sample_zip  # noqa: E402


@pytest.fixture(scope="module")
def out(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("run")
    z = tmp / "in" / "sample.zip"
    make_sample_zip.main(z)
    ep.main(["--input", str(z.parent), "--output", str(tmp / "out")])
    return tmp / "out"


def read(out, name, sheet=None):
    p = out / name
    return pd.read_csv(p) if p.suffix == ".csv" else pd.read_excel(p, sheet_name=sheet)


def test_outputs_exist(out):
    for f in ["production_master.csv", "production_master_long.csv", "unmapped_data.xlsx",
              "extraction_report.xlsx", "inspection_report.xlsx"]:
        assert (out / f).exists(), f


def test_nothing_silently_lost(out):
    acct = read(out, "extraction_report.xlsx", "Cell_Accounting")
    assert acct["Balanced"].all()
    assert (acct["Unaccounted"] == 0).all()


def test_sources_unchanged(out):
    files = read(out, "extraction_report.xlsx", "Files")
    assert files["UnchangedAfterRun"].all()


def test_failed_and_other_files_reported(out):
    files = read(out, "extraction_report.xlsx", "Files")
    failed = files[files["Status"] != "processed"]
    assert list(failed["SourcePath"]) == ["Production/2024/broken_Oct_2024.xlsx"]
    other = read(out, "extraction_report.xlsx", "Other_Files")
    assert other["Path"].str.endswith(".xls").any()


def test_only_first_sheet(out):
    long = read(out, "production_master_long.csv")
    assert not (long["Value"] == 999).any()          # value lives on 2nd sheet "Delays"
    assert set(long["SourceSheet"]) >= {"Apr23", "Prod", "Daily", "Shiftwise", "Summary"}


def test_vertical_merged_values_and_trace(out):
    w = read(out, "production_master.csv")
    r = w[(w.SourceFile == "Production_April_2023.xlsx") & (w.Date == "2023-04-01") & (w["Mill/Area"] == "CRM04")]
    assert len(r) == 1
    assert r.iloc[0]["Output"] == 973.0 and r.iloc[0]["SourceRow"] == "5"
    assert "Output=C5" in r.iloc[0]["SourceCells"]


def test_totals_not_counted_as_daily(out):
    long = read(out, "production_master_long.csv")
    apr = long[long.SourceFile == "Production_April_2023.xlsx"]
    totals = apr[apr.RecordType == "Total"]
    assert len(totals) > 0 and not totals["CountsAsProduction"].any()
    rec = read(out, "extraction_report.xlsx", "Reconciliation")
    b35 = rec[(rec.SourceFile == "Production/2023/Production_April_2023.xlsx") & (rec.Cell == "B35")]
    assert b35["Status"].iloc[0] != "MATCH"           # deliberate +7 error is detected


def test_horizontal_layout(out):
    w = read(out, "production_master.csv")
    r = w[(w.SourceFile == "June-24 Production.xlsx") & (w.Date == "2024-06-01") & (w["Mill/Area"] == "CRM06")]
    assert r.iloc[0]["Input"] == 799.0 and r.iloc[0]["SourceRow"] == "8;9;10"


def test_stacked_sections_day_numbers(out):
    w = read(out, "production_master.csv")
    crm = w[(w.SourceFile == "Jul_2024_stacked.xlsx") & (w["Mill/Area"] == "CRM04") & (w.RecordType == "Daily")]
    assert len(crm) == 31 and crm["Date"].min() == "2024-07-01"
    pick = w[(w.SourceFile == "Jul_2024_stacked.xlsx") & (w["Mill/Area"] == "PICKLING LINE")]
    assert len(pick) == 32  # 31 days + Total


def test_unlabelled_rows_go_to_unmapped(out):
    u = read(out, "unmapped_data.xlsx", "Unmapped")
    aug = u[u.SourceFile.str.contains("Aug 2024")]
    assert set(aug["Cell"]) == {"D33", "E33"}
    assert (u["ReasonCategory"] == "Placeholder").any()


def test_shift_long_format(out):
    w = read(out, "production_master.csv")
    aug = w[w.SourceFile == "Aug 2024 shiftwise.xlsx"]
    assert len(aug) == 30 and set(aug["Shift"]) == {"A", "B", "C"} and set(aug["Mill/Area"]) == {"CRM04"}


def test_duplicate_file_flagged(out):
    w = read(out, "production_master.csv")
    copy = w[w.SourcePath.str.contains("copy/")]
    assert copy["DuplicateFlag"].fillna("").str.contains("DUPLICATE-EXACT").all()
