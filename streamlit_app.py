"""Streamlit front-end for extract_production.py.

Run locally:   streamlit run streamlit_app.py
Deploy:        Streamlit Community Cloud -> New app -> this repo, main file streamlit_app.py

Uploaded files are processed in a temporary folder that is deleted after the run;
nothing is written back to the uploaded workbooks and nothing is stored server-side.
"""
from __future__ import annotations

import contextlib
import hashlib
import io
import json
import logging
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

import extract_production as ep

OUTPUT_FILES = [
    ("production_master.csv", "text/csv"),
    ("production_master_long.csv", "text/csv"),
    ("unmapped_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ("extraction_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ("inspection_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ("extraction.log", "text/plain"),
]
# validated categorical palette, fixed order (slot 1..8); colour follows the mill, not its rank
SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
ACCEPTED = ["zip", "xlsx", "xlsm"]


# --------------------------------------------------------------------------- processing
def run_extraction(uploads: list[tuple[str, bytes]], config_bytes: bytes | None) -> dict:
    """Run the extractor on uploaded files in a temp folder and return everything the UI needs
    (all in memory, so the temp folder can be deleted immediately)."""
    with tempfile.TemporaryDirectory(prefix="prodx_") as tmp:
        tmp = Path(tmp)
        in_dir, out_dir = tmp / "input", tmp / "output"
        in_dir.mkdir()
        out_dir.mkdir()
        for name, data in uploads:
            (in_dir / Path(name).name).write_bytes(data)
        cfg_path = None
        if config_bytes:
            cfg_path = tmp / "config.json"
            cfg_path.write_bytes(config_bytes)

        log_stream = io.StringIO()
        handler = logging.StreamHandler(log_stream)
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        ep.log.addHandler(handler)
        ep.log.setLevel(logging.INFO)
        console = io.StringIO()
        try:
            ep.load_config(cfg_path)
            with contextlib.redirect_stdout(console):
                exit_code = ep.run(in_dir, out_dir)
        finally:
            ep.log.removeHandler(handler)
        (out_dir / "extraction.log").write_text(log_stream.getvalue() + "\n" + console.getvalue(),
                                                encoding="utf-8")

        files = {name: (out_dir / name).read_bytes() for name, _ in OUTPUT_FILES if (out_dir / name).exists()}
        summary = json.loads((out_dir / "run_summary.json").read_text(encoding="utf-8"))
        master = pd.read_csv(out_dir / "production_master.csv", low_memory=False) \
            if (out_dir / "production_master.csv").stat().st_size > 3 else pd.DataFrame()
        long = pd.read_csv(out_dir / "production_master_long.csv", low_memory=False) \
            if (out_dir / "production_master_long.csv").stat().st_size > 3 else pd.DataFrame()
        report = pd.read_excel(out_dir / "extraction_report.xlsx", sheet_name=None)
        unmapped = pd.read_excel(out_dir / "unmapped_data.xlsx", sheet_name="Unmapped")
    return {"exit_code": exit_code, "summary": summary, "files": files, "master": master, "long": long,
            "report": report, "unmapped": unmapped, "console": console.getvalue()}


def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Mixed-type text columns (e.g. Value in unmapped cells) are shown as text so the table renders."""
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].map(lambda v: v if v is None or isinstance(v, str) or pd.isna(v) else str(v))
    return out


def zip_outputs(files: dict[str, bytes]) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return bio.getvalue()


def monthly_mill_output(long: pd.DataFrame) -> pd.DataFrame:
    """Base production Output per month and mill (no totals, no product breakdown)."""
    if long.empty:
        return pd.DataFrame()
    base = long[(long["CountsAsProduction"] == True) & (long["Measure"] == "Output")  # noqa: E712
                & long["Year"].notna() & long["Month"].notna()].copy()
    if "Dimension" in base:
        base = base[base["Dimension"] == "Mill/Line"]
    if base.empty:
        return pd.DataFrame()
    base["Month"] = pd.to_datetime(dict(year=base["Year"].astype(int), month=base["Month"].astype(int), day=1))
    base["Area"] = base["Area"].fillna("(no area)")
    table = base.pivot_table(index="Month", columns="Area", values="Value", aggfunc="sum").sort_index()
    order = table.sum().sort_values(ascending=False).index.tolist()
    if len(order) > len(SERIES_COLORS):  # never invent a 9th hue - fold the tail into "Other"
        keep = order[:len(SERIES_COLORS) - 1]
        table["Other"] = table[order[len(keep):]].sum(axis=1)
        table = table[keep + ["Other"]]
    else:
        table = table[sorted(order)]
    return table.round(3)


# --------------------------------------------------------------------------- UI
def metric_row(summary: dict) -> None:
    def g(key):
        return summary.get(key, "-")

    cols = st.columns(4)
    cols[0].metric("Files found", g("Files found (Excel .xlsx/.xlsm)"))
    cols[1].metric("Processed", g("Files successfully processed"))
    cols[2].metric("Failed", g("Files failed"))
    cols[3].metric("Rows inspected", g("Total rows inspected (first sheets, used range)"))
    cols = st.columns(4)
    cols[0].metric("Records extracted", g("Total records extracted (production_master.csv rows)"))
    cols[1].metric("Unmapped cells", g("Total unmapped (cells in unmapped_data.xlsx)"))
    cols[2].metric("Duplicates", g("Total duplicates (extra occurrences)"))
    cols[3].metric("Errors", g("Total errors"))


def checks(summary: dict) -> None:
    ok_acct = str(summary.get("Cell accounting balanced for all files")) == "True"
    ok_src = str(summary.get("Source files unchanged (SHA-256)")) == "True"
    mism = int(summary.get("Reconciliation mismatches", 0) or 0)
    (st.success if ok_acct else st.error)(
        "✔ Every non-empty cell is accounted for (extracted, header/label, or unmapped)." if ok_acct
        else "✖ Cell accounting does not balance - see the Cell_Accounting sheet.")
    (st.success if ok_src else st.error)(
        "✔ Uploaded files were not modified (SHA-256 before = after)." if ok_src
        else "✖ A source file changed during processing.")
    if mism:
        st.warning(f"⚠ {mism} reconciliation mismatch(es) found in the data - see the Reconciliation tab.")
    else:
        st.success("✔ All reconciliation checks match.")


def filter_master(master: pd.DataFrame) -> pd.DataFrame:
    c1, c2, c3, c4 = st.columns([1, 1, 2, 1.4])
    years = sorted(master["Year"].dropna().astype(int).unique()) if "Year" in master else []
    year = c1.selectbox("Year", ["All"] + years)
    months = sorted(master["Month"].dropna().astype(int).unique()) if "Month" in master else []
    month = c2.selectbox("Month", ["All"] + months)
    areas = sorted(master["Mill/Area"].dropna().astype(str).unique()) if "Mill/Area" in master else []
    area = c3.multiselect("Mill/Area", areas)
    only_base = c4.toggle("Daily production only", value=True,
                          help="Hide totals, calculated columns and product-segment breakdown "
                               "(CountsAsProduction = True).")
    df = master
    if year != "All":
        df = df[df["Year"] == year]
    if month != "All":
        df = df[df["Month"] == month]
    if area:
        df = df[df["Mill/Area"].astype(str).isin(area)]
    if only_base and "CountsAsProduction" in df:
        df = df[df["CountsAsProduction"] == True]  # noqa: E712
    return df.dropna(axis=1, how="all")


def show_results(res: dict) -> None:
    summary = res["summary"]
    st.subheader("Validation summary")
    metric_row(summary)
    checks(summary)

    st.subheader("Downloads")
    st.download_button("⬇ Download all outputs (.zip)", zip_outputs(res["files"]),
                       file_name="production_extraction_outputs.zip", mime="application/zip",
                       type="primary", width="stretch")
    dcols = st.columns(3)
    for i, (name, mime) in enumerate(OUTPUT_FILES):
        if name in res["files"]:
            dcols[i % 3].download_button(name, res["files"][name], file_name=name, mime=mime,
                                         width="stretch", key=f"dl_{name}")

    report = res["report"]
    tabs = st.tabs(["Production master", "Monthly output", "Unmapped", "Reconciliation", "Files & structure",
                    "Special handling", "Run log"])
    with tabs[0]:
        master = res["master"]
        if master.empty:
            st.info("No records extracted.")
        else:
            df = filter_master(master)
            st.caption(f"{len(df):,} of {len(master):,} rows. Every row keeps SourceFile, SourceRow and "
                       "SourceCells for tracing back to Excel.")
            st.dataframe(arrow_safe(df), width="stretch", height=480)
    with tabs[1]:
        table = monthly_mill_output(res["long"])
        if table.empty:
            st.info("No dated daily production values to chart.")
        else:
            st.markdown("**Monthly output by mill (MT)** - sum of daily base production; totals, calculated "
                        "columns and product-segment breakdown excluded.")
            st.line_chart(table, color=SERIES_COLORS[:len(table.columns)], height=380)
            st.markdown("Table view")
            show = table.copy()
            show.index = show.index.strftime("%Y-%m")
            st.dataframe(arrow_safe(show), width="stretch")
    with tabs[2]:
        unm = res["unmapped"]
        if unm.empty:
            st.success("Nothing unmapped.")
        else:
            st.caption("Cells that could not be mapped confidently. Nothing is discarded - review them here or in "
                       "unmapped_data.xlsx.")
            cats = unm["ReasonCategory"].value_counts()
            st.dataframe(cats.rename("Cells"), width="content")
            pick = st.multiselect("Reason category", list(cats.index))
            st.dataframe(arrow_safe(unm[unm["ReasonCategory"].isin(pick)] if pick else unm), width="stretch",
                         height=420)
    with tabs[3]:
        rec = report.get("Reconciliation", pd.DataFrame())
        if rec.empty:
            st.info("No reconciliation checks.")
        else:
            st.dataframe(arrow_safe(rec.groupby(["Check", "Status"]).size().rename("Checks").reset_index()),
                         width="stretch", hide_index=True)
            only_bad = st.toggle("Show mismatches only", value=True)
            view = rec[rec["Status"].astype(str).str.startswith("MISMATCH")] if only_bad else rec
            st.dataframe(arrow_safe(view), width="stretch", height=420)
    with tabs[4]:
        st.markdown("**Files**")
        st.dataframe(arrow_safe(report.get("Files", pd.DataFrame())), width="stretch", height=300)
        st.markdown("**Layout groups** (files with the same first-sheet structure)")
        st.dataframe(arrow_safe(report.get("Layout_Groups", pd.DataFrame())), width="stretch")
        other = report.get("Other_Files", pd.DataFrame())
        if not other.empty and "Info" not in other.columns:
            st.markdown("**Other files in the upload (not processed)**")
            st.dataframe(arrow_safe(other), width="stretch")
    with tabs[5]:
        st.dataframe(arrow_safe(report.get("Special_Handling", pd.DataFrame())), width="stretch", height=480)
    with tabs[6]:
        st.code(res["files"].get("extraction.log", b"").decode("utf-8", "replace")[-20000:] or "(empty)")


def main() -> None:
    st.set_page_config(page_title="Production Data Extraction", page_icon="🏭", layout="wide")
    st.title("Production Data Extraction")
    st.markdown(
        "Upload the ZIP of monthly production workbooks. Only the **first worksheet** of every `.xlsx` is read; "
        "every value is traced to its source file, row and cell, and anything that cannot be mapped confidently "
        "is listed for review instead of being dropped.")
    with st.sidebar:
        st.header("About")
        st.markdown(
            "- Files are processed in a temporary folder and deleted after the run.\n"
            "- Original workbooks are never modified (verified by SHA-256).\n"
            "- Totals, calculated columns and product-segment breakdown are kept but excluded from "
            "`CountsAsProduction`, so nothing is double counted.")
        st.caption(f"Extractor version {ep.SCRIPT_VERSION}")

    uploads = st.file_uploader("ZIP file(s) or workbooks", type=ACCEPTED, accept_multiple_files=True)
    with st.expander("Advanced: keyword configuration (optional)"):
        st.markdown("Upload a JSON file to extend mill names, measure synonyms or product segments "
                    "(same format as `--config` on the command line).")
        cfg = st.file_uploader("config.json", type=["json"], key="cfg")

    if uploads:
        payload = [(u.name, u.getvalue()) for u in uploads]
        cfg_bytes = cfg.getvalue() if cfg else None
        digest = hashlib.sha256(b"".join(hashlib.sha256(d).digest() for _, d in payload)
                                + (cfg_bytes or b"")).hexdigest()
        if st.button("Run extraction", type="primary"):
            with st.spinner(f"Extracting {len(payload)} upload(s) - large ZIPs can take a minute or two..."):
                try:
                    st.session_state["result"] = run_extraction(payload, cfg_bytes)
                    st.session_state["digest"] = digest
                except Exception as exc:  # noqa: BLE001 - show the error instead of a blank page
                    st.session_state.pop("result", None)
                    st.error(f"Extraction failed: {type(exc).__name__}: {exc}")
                    st.exception(exc)
        if st.session_state.get("result") and st.session_state.get("digest") != digest:
            st.info("The uploaded files changed - press **Run extraction** to process them.")
    if st.session_state.get("result") and uploads and st.session_state.get("digest") == digest:
        show_results(st.session_state["result"])
    elif not uploads:
        st.session_state.pop("result", None)


if __name__ == "__main__":
    main()
