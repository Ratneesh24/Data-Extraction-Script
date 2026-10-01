#!/usr/bin/env python3
"""
extract_production.py
=====================

Consolidate monthly production workbooks (delivered inside one or more ZIP
files) into a single traceable dataset.

Guarantees
----------
* Only the FIRST worksheet of every workbook is read; all other sheets are
  listed in the report as "ignored" but never parsed.
* Original files are never modified: ZIPs are extracted into a separate work
  folder, workbooks are opened from read-only binary handles and never saved,
  and a SHA-256 hash of every source file is compared before/after the run.
* Every non-empty cell of every first sheet receives exactly one disposition
  (title / section / header / date / label / attribute / value / unmapped).
  A cell-accounting check proves that nothing was silently dropped.
* Anything that cannot be mapped confidently goes to unmapped_data.xlsx with a
  reason; anything mapped with an assumption is flagged (ReviewFlag) for review.
* Totals / subtotals / averages / cumulative (MTD) values are kept but marked
  (IsSummary / CountsAsProduction) so they are never double counted.

Layouts handled (detected per file, never assumed)
--------------------------------------------------
* vertical   : dates down a column, areas/processes/measures across (multi-row
               and merged headers, stacked sections, side-by-side tables).
* horizontal : dates across a row, areas/processes/measures down label columns.
* matrix     : no date axis (monthly summary tables) - rows/columns are labels.

Usage
-----
    python extract_production.py                      # ./input -> ./output
    python extract_production.py --input my.zip --output out
    python extract_production.py --inspect-only       # structure report only
    python extract_production.py --config keywords.json   # extend keyword lists

Outputs (in the output folder)
------------------------------
    production_master.csv        one row per source row x area x process,
                                 measures (Input/Output/Target/...) as columns
    production_master_long.csv   one row per extracted value cell (full detail)
    unmapped_data.xlsx           Unmapped cells + records needing review
    extraction_report.xlsx       files, structure, errors, duplicates,
                                 reconciliation, cell accounting, samples
    inspection_report.xlsx       structure inspection of every first sheet
    extraction.log
"""
from __future__ import annotations

import argparse
import calendar
import datetime as dt
import hashlib
import json
import logging
import math
import re
import shutil
import sys
import traceback
import warnings
import zipfile
from collections import Counter, OrderedDict, defaultdict
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

SCRIPT_VERSION = "1.0.0"
log = logging.getLogger("extract_production")

EXCEL_EXTS = {".xlsx", ".xlsm"}
OTHER_SHEET_EXTS = {".xls", ".xlsb", ".ods", ".csv", ".xltx", ".xltm"}
NUMERIC_KINDS = {"NUM", "NUMTEXT", "TIME"}

# --------------------------------------------------------------------------
# Keyword configuration. Everything here can be extended with --config.
# Regexes are case-insensitive.
# --------------------------------------------------------------------------
DEFAULT_CONFIG: dict[str, Any] = {
    # canonical measure name -> regex. "secondary" measures yield to primary.
    "measures": [
        ["Target", r"\b(?:target|tgt|plan(?:ned)?|budget|abp|aop)\b"],
        ["Input", r"\b(?:input|i/p|inp|feed|charged?|consumption)\b"],
        ["Output", r"\b(?:output|o/p|production|prodn|prod|actual|actuals|act|achieved)\b"],
        ["Yield", r"\byield\b"],
        ["Achievement", r"\b(?:achievement|achv|ach)\b"],
        ["Variance", r"\b(?:variance|var|deviation|diff(?:erence)?|shortfall|gap)\b"],
        ["Despatch", r"\b(?:despatch(?:ed)?|dispatch(?:ed)?)\b"],
        ["Rejection", r"\b(?:rejection|rejected|reject|scrap|downgraded?|diversion|defects?)\b"],
        ["Hold", r"\b(?:hold|on\s*hold)\b"],
        ["Stock", r"\b(?:stock|inventory|wip|opening|closing|backlog)\b"],
        ["Delay", r"\b(?:delays?|breakdown|b/d|downtime|stoppage|shutdown)\b"],
        ["Hours", r"\b(?:running\s*h(?:ou)?rs|run\s*hrs|working\s*hrs|availability|utili[sz]ation)\b"],
        ["Speed", r"\b(?:speed|mpm)\b"],
        ["Productivity", r"\b(?:tph|productivity)\b"],
        ["Power", r"\b(?:power|kwh|energy)\b"],
        ["Thickness", r"\b(?:thickness|thk|gauge)\b"],
        ["Width", r"\bwidth\b"],
        ["Count", r"\b(?:no\.?\s*of\s*coils|coils?|nos)\b"],
    ],
    "secondary_measures": ["Count"],
    # measures whose unit is appended to the column name if it is not tonnage
    "quantity_measures": ["Input", "Output", "Target", "Despatch", "Rejection", "Hold", "Stock", "Variance"],
    "tonnage_units": ["MT", "T", "Kg"],
    "summary": [
        ["Grand Total", r"\bgrand\s*total\b|\bg\.\s*total\b"],
        ["Subtotal", r"\bsub[\s\-]*total\b"],
        ["Total", r"\btotal\b|\btot\b|\bsum\b"],
        ["Average", r"\bavg\b|\baverage\b|\bmean\b"],
        ["Maximum", r"\bmax(?:imum)?\b|\bhighest\b"],
        ["Minimum", r"\bmin(?:imum)?\b|\blowest\b"],
    ],
    "cumulative": r"\b(?:cum\w*|mtd|ytd|ftm|progressive|prog|till\s*date|to\s*date|todate|upto\s*date|month\s*to\s*date|year\s*to\s*date)\b",
    "units": [
        ["MT", r"\(\s*mt\s*\)|\bmts?\b|\bm\.t\.?|\btonnes?\b|\btons?\b|\btonnage\b"],
        ["T", r"\(\s*t\s*\)"],
        ["%", r"%|\bpercent(?:age)?\b|\bpct\b"],
        ["Hrs", r"\bhrs?\b|\bhours?\b"],
        ["Min", r"\bmins?\b|\bminutes?\b"],
        ["Nos", r"\bnos\.?\b|\(\s*no\.?s?\s*\)"],
        ["kWh", r"\bkwh\b"],
        ["Kg", r"\bkgs?\b"],
        ["mm", r"\(\s*mm\s*\)"],
        ["MPM", r"\bmpm\b"],
        ["TPH", r"\btph\b"],
    ],
    "date_header": r"^\s*(?:date|dates|day|days|dt\.?|dated|date\s*/\s*day|day\s*/\s*date|date\s*&\s*day)\s*$",
    "serial_header": r"^\s*(?:s\.?\s*no\.?|sl\.?\s*no\.?|sr\.?\s*no\.?|serial(?:\s*no\.?)?|#)\s*$",
    "shift_header": r"\bshifts?\b",
    "remarks_header": r"\b(?:remarks?|comments?|reasons?|notes?|observations?|cause|explanation)\b",
    "weekday_header": r"^\s*(?:weekday|day\s*name|week\s*day|day\s*of\s*week)\s*$",
    "row_area_header": r"\b(?:mill|area|line|unit|plant|shop|section|dept|department|equipment|machine|stand|facility|location)\b",
    "row_process_header": r"\b(?:process|operation|activity|product|grade|item|description|particulars?|parameter|category|type|customer|route)\b",
    "title_markers": r"\b(?:report|statement|summary|for\s+the\s+month|month\s+of|ltd|limited|pvt|company|corporation|dated|period)\b",
    "note_markers": r"^\s*(?:note|nb|n\.b\.|remarks?|prepared|checked|approved|verified|sign|signature|submitted|source|ref|\*)",
    "placeholders": r"^\s*(?:-+|–|—|_+|\.+|nil|na|n/?a|none|x+|\*+|sd|s/d|shut|shutdown|off|idle|bd|b/d|holiday|nr|not\s*reported|no\s*production|no\s*prod\.?|---)\s*$",
    "known_areas": r"\b(?:CRM\s*-?\s*\d+|PLTCM|CPL|PL\s*-?\s*\d+|BAF|HPM|SPM|CRS|CTL|ECL|ETL|HDGL|CGL|CAL|ARP|RCL|SKIN\s*PASS|TEMPER\s*MILL|PICKLING(?:\s*LINE)?|SLITTING(?:\s*LINE)?|SLITTER|REWINDING|RECOILING|PACKING|PACKAGING|FINISHING|ANNEALING|COLD\s*MILL|TANDEM\s*MILL)\b",
    "noise_words": r"\b(?:qty|quantity|wt|weight|of|the|for|data|details|value|values|figures?|in|daily|ftd|for\s+the\s+day|today|on\s*date|day)\b",
    "weekday_names": r"^\s*(?:mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun)(?:day|nesday|sday|urday|rsday)?\.?\s*$",
}

CFG: dict[str, Any] = {}
RX: dict[str, Any] = {}


def load_config(path: Path | None) -> None:
    """Merge user config (JSON) into defaults and compile regexes."""
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if path:
        user = json.loads(Path(path).read_text(encoding="utf-8"))
        for k, v in user.items():
            if k in ("measures", "summary", "units") and isinstance(v, list):
                cfg[k] = v + cfg[k]  # user entries take priority
            else:
                cfg[k] = v
    CFG.clear()
    CFG.update(cfg)
    RX.clear()
    RX["measures"] = [(n, re.compile(p, re.I)) for n, p in cfg["measures"]]
    RX["summary"] = [(n, re.compile(p, re.I)) for n, p in cfg["summary"]]
    RX["units"] = [(n, re.compile(p, re.I)) for n, p in cfg["units"]]
    for k in ("cumulative", "date_header", "serial_header", "shift_header", "remarks_header",
              "weekday_header", "row_area_header", "row_process_header", "title_markers",
              "note_markers", "placeholders", "known_areas", "noise_words", "weekday_names"):
        RX[k] = re.compile(cfg[k], re.I)
    analyse_label.cache_clear()


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------
ERROR_RE = re.compile(r"^#(?:REF!|DIV/0!|N/A|VALUE!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|FIELD!|BLOCKED!|UNKNOWN!)")
NUMTEXT_RE = re.compile(r"^[+-]?(?:\d+|\d{1,3}(?:,\d{2,3})+)(?:\.\d+)?$")
MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
          "sep": 9, "oct": 10, "nov": 11, "dec": 12}
MONTH_WORD = r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
PERIOD_RXS = [
    ("my", re.compile(r"\b" + MONTH_WORD + r"[\s\-./',’`]*((?:19|20)\d{2}|\d{2})\b", re.I)),
    ("ym", re.compile(r"\b((?:19|20)\d{2})[\s\-./]*" + MONTH_WORD + r"\b", re.I)),
    ("yn", re.compile(r"\b((?:19|20)\d{2})[\-./](0?[1-9]|1[0-2])\b")),
    ("ny", re.compile(r"\b(0?[1-9]|1[0-2])[\-./]((?:19|20)\d{2})\b")),
]
MONTH_ONLY_RX = re.compile(r"\b" + MONTH_WORD + r"\b", re.I)
YEAR_ONLY_RX = re.compile(r"\b((?:19|20)\d{2})\b")
WEEKDAY_STRIP = re.compile(r"\(?\b(?:mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun)(?:day|nesday|sday|urday|rsday)?\b\.?\)?", re.I)


def month_num(word: str) -> int:
    return MONTHS[word.lower()[:3]]


def year_full(y: str) -> int:
    y = int(y)
    return 2000 + y if y < 100 else y


def find_period(text: str, loose: bool = False) -> tuple[int, int] | None:
    """Return (year, month) mentioned in text, or None."""
    if not text:
        return None
    t = str(text).replace("_", " ")
    for kind, rx in PERIOD_RXS:
        m = rx.search(t)
        if not m:
            continue
        try:
            if kind == "my":
                return year_full(m.group(2)), month_num(m.group(1))
            if kind == "ym":
                return int(m.group(1)), month_num(m.group(2))
            if kind == "yn":
                return int(m.group(1)), int(m.group(2))
            if kind == "ny":
                return int(m.group(2)), int(m.group(1))
        except (KeyError, ValueError):
            continue
    if loose:  # month word and a 4-digit year anywhere (file paths)
        mm, yy = MONTH_ONLY_RX.findall(t), YEAR_ONLY_RX.findall(t)
        if mm and yy and len({month_num(x) for x in mm}) == 1 and len(set(yy)) == 1:
            return int(yy[0]), month_num(mm[0])
    return None


def safe_date(y: int, m: int, d: int) -> dt.date | None:
    try:
        return dt.date(y, m, d)
    except ValueError:
        return None


def parse_date_text(s: str, period: tuple[int, int] | None) -> tuple[dt.date | None, str | None]:
    """Parse a date written as text. Returns (date, status) where status is
    'ok', 'ambiguous' (looks like a date but cannot be resolved) or None."""
    s0 = s.strip()
    if not re.search(r"\d", s0) or len(s0) > 30:
        return None, None
    s1 = WEEKDAY_STRIP.sub(" ", s0).strip(" ,-/()")
    m = re.match(r"^(\d{1,2})[./\-\s](\d{1,2})[./\-\s](\d{2,4})$", s1)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), year_full(m.group(3))
        cands = {d for d in (safe_date(y, b, a), safe_date(y, a, b)) if d}
        if len(cands) == 1:
            return cands.pop(), "ok"
        if len(cands) == 2 and period:
            hits = [d for d in cands if (d.year, d.month) == period]
            if len(hits) == 1:
                return hits[0], "ok"
        return (None, "ambiguous") if cands else (None, None)
    m = re.match(r"^(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})$", s1)
    if m:
        d = safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return (d, "ok") if d else (None, None)
    m = re.match(r"^(\d{1,2})(?:st|nd|rd|th)?[\s\-./,']*" + MONTH_WORD + r"\.?[\s\-./,'’]*(\d{2,4})?$", s1, re.I)
    if m:
        day, mon = int(m.group(1)), month_num(m.group(2))
        if m.group(3):
            d = safe_date(year_full(m.group(3)), mon, day)
            return (d, "ok") if d else (None, None)
        if period and period[1] == mon:
            d = safe_date(period[0], mon, day)
            return (d, "ok") if d else (None, None)
        return None, "ambiguous"
    m = re.match(r"^" + MONTH_WORD + r"\.?[\s\-./]*(\d{1,2})(?:st|nd|rd|th)?[\s,\-./'’]*(\d{2,4})?$", s1, re.I)
    if m:
        mon, day = month_num(m.group(1)), int(m.group(2))
        if m.group(3):
            d = safe_date(year_full(m.group(3)), mon, day)
            return (d, "ok") if d else (None, None)
        if period and period[1] == mon:
            d = safe_date(period[0], mon, day)
            return (d, "ok") if d else (None, None)
        return None, "ambiguous"
    return None, None


def fmt_num(x: float) -> str:
    if x is None:
        return ""
    if float(x).is_integer():
        return str(int(x))
    return repr(round(x, 10))


def norm_key(s: str | None) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def clean_text(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v)).strip()


def sheet_safe(v: Any) -> Any:
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
        if v.startswith("="):  # keep formula text as text, never re-evaluate it
            v = "'" + v
    return v


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------
# Label analysis (headers, row labels, section titles)
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class LabelInfo:
    text: str
    measures: tuple = ()
    unit: str | None = None
    summary: str | None = None
    cumulative: bool = False
    period: tuple | None = None
    descriptor: str | None = None
    is_date_header: bool = False


@lru_cache(maxsize=50000)
def analyse_label(text: str) -> LabelInfo:
    t = clean_text(text)
    if not t:
        return LabelInfo(text="")
    if RX["date_header"].match(t):
        return LabelInfo(text=t, is_date_header=True)
    measures = []
    for name, rx in RX["measures"]:
        if rx.search(t) and name not in measures:
            measures.append(name)
    unit = None
    for name, rx in RX["units"]:
        if rx.search(t):
            unit = name
            break
    primary = [m for m in measures if m not in CFG["secondary_measures"]]
    if primary and len(primary) < len(measures):
        if "Count" in measures and not unit:
            unit = "Nos"
        measures = primary
    summary = None
    for name, rx in RX["summary"]:
        if rx.search(t):
            summary = name
            break
    cumulative = bool(RX["cumulative"].search(t))
    period = find_period(t)
    rem = t
    for _, rx in RX["measures"]:
        rem = rx.sub(" ", rem)
    for _, rx in RX["units"]:
        rem = rx.sub(" ", rem)
    for _, rx in RX["summary"]:
        rem = rx.sub(" ", rem)
    rem = RX["cumulative"].sub(" ", rem)
    for _, rx in PERIOD_RXS:
        rem = rx.sub(" ", rem)
    rem = RX["noise_words"].sub(" ", rem)
    rem = re.sub(r"[()\[\]{}:;,|\\_=]+", " ", rem)
    rem = re.sub(r"(?:^|\s)[-/&.+*]+(?=\s|$)", " ", rem)
    rem = re.sub(r"\s+", " ", rem).strip(" -/&.+*'\"")
    descriptor = rem or None
    if descriptor and (RX["date_header"].match(descriptor) or RX["serial_header"].match(descriptor)):
        descriptor = None
    return LabelInfo(text=t, measures=tuple(measures), unit=unit, summary=summary,
                     cumulative=cumulative, period=period, descriptor=descriptor)


def is_title_like(text: str) -> bool:
    return bool(RX["title_markers"].search(text)) or find_period(text) is not None


def is_note_like(text: str) -> bool:
    return bool(RX["note_markers"].search(text)) or len(text.split()) > 10 or len(text) > 80


# --------------------------------------------------------------------------
# Data structures
# --------------------------------------------------------------------------
@dataclass
class Cell:
    r: int
    c: int
    raw: Any
    kind: str                      # NUM NUMTEXT TIME DATE TEXT ERROR BLANK
    num: float | None = None
    date: dt.date | None = None
    date_status: str | None = None  # ok / ambiguous for TEXT dates
    text: str | None = None
    formula: str | None = None
    span: tuple | None = None       # merged range (r1, c1, r2, c2) if anchor
    disposition: str | None = None
    detail: str = ""

    @property
    def ref(self) -> str:
        return f"{get_column_letter(self.c)}{self.r}"


@dataclass
class FileEntry:
    idx: int
    path: Path              # file actually read (extracted copy or original)
    display: str            # path relative to the ZIP / input folder
    zip_name: str           # ZIP the file came from ('' if loose file)
    ext: str
    status: str = "pending"
    reason: str = ""
    size: int = 0
    sha_before: str = ""
    sha_after: str = ""


@dataclass
class SheetCtx:
    fe: FileEntry
    sheet_name: str = ""
    all_sheets: list = field(default_factory=list)
    cells: dict = field(default_factory=dict)
    merged_parent: dict = field(default_factory=dict)
    hidden_rows: set = field(default_factory=set)
    hidden_cols: set = field(default_factory=set)
    formula_nocache: list = field(default_factory=list)
    max_row: int = 0
    max_col: int = 0
    min_row: int = 0
    period: tuple | None = None
    period_source: str = ""
    period_candidates: dict = field(default_factory=dict)
    mode: str = ""
    axis_scores: str = ""
    titles: list = field(default_factory=list)
    section_rows: dict = field(default_factory=dict)   # orig row -> text
    section_cells: dict = field(default_factory=dict)  # orig row -> Cell
    records: list = field(default_factory=list)
    unmapped: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    special: list = field(default_factory=list)
    blocks: list = field(default_factory=list)
    columns: list = field(default_factory=list)
    recon: list = field(default_factory=list)
    load_warnings: list = field(default_factory=list)

    # -- helpers -----------------------------------------------------------
    def note(self, category: str, detail: str, location: str = "") -> None:
        self.special.append({"SourceFile": self.fe.display, "Sheet": self.sheet_name,
                             "Category": category, "Location": location, "Detail": detail})

    def error(self, category: str, detail: str, location: str = "") -> None:
        self.errors.append({"SourceFile": self.fe.display, "Sheet": self.sheet_name,
                            "Category": category, "Location": location, "Detail": detail})

    def unmap(self, cell: Cell, reason: str, category: str, context: str = "") -> None:
        if cell.disposition is not None:
            return
        cell.disposition = "unmapped"
        cell.detail = reason
        self.unmapped.append({
            "SourceFile": self.fe.display, "ZipFile": self.fe.zip_name, "Sheet": self.sheet_name,
            "Row": cell.r, "Column": get_column_letter(cell.c), "Cell": cell.ref,
            "Value": cell.raw if not isinstance(cell.raw, (dt.date, dt.time, dt.timedelta)) else str(cell.raw),
            "ValueType": cell.kind, "Formula": cell.formula or "", "ReasonCategory": category,
            "Reason": reason, "Context": context,
            "HiddenRow": cell.r in self.hidden_rows, "HiddenColumn": cell.c in self.hidden_cols,
        })


def classify_value(v: Any) -> tuple[str, dict]:
    if isinstance(v, bool):
        return "TEXT", {"text": str(v)}
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return "TEXT", {"text": str(v)}
        return "NUM", {"num": float(v)}
    if isinstance(v, dt.datetime):
        return "DATE", {"date": v.date()}
    if isinstance(v, dt.date):
        return "DATE", {"date": v}
    if isinstance(v, dt.time):
        return "TIME", {"num": v.hour + v.minute / 60 + v.second / 3600}
    if isinstance(v, dt.timedelta):
        return "TIME", {"num": v.total_seconds() / 3600}
    s = clean_text(v)
    if not s:
        return "BLANK", {"text": ""}
    if ERROR_RE.match(s):
        return "ERROR", {"text": s}
    if NUMTEXT_RE.match(s):
        return "NUMTEXT", {"num": float(s.replace(",", "")), "text": s}
    return "TEXT", {"text": s}


# --------------------------------------------------------------------------
# Workbook loading
# --------------------------------------------------------------------------
def load_first_sheet(fe: FileEntry) -> SheetCtx:
    S = SheetCtx(fe=fe)
    with warnings.catch_warnings(record=True) as wlist:
        warnings.simplefilter("always")
        with open(fe.path, "rb") as fh:
            wb_v = openpyxl.load_workbook(fh, data_only=True)
        with open(fe.path, "rb") as fh:
            wb_f = openpyxl.load_workbook(fh, data_only=False)
    S.load_warnings = sorted({str(w.message)[:300] for w in wlist})
    S.all_sheets = list(wb_v.sheetnames)
    if not wb_v.worksheets:
        raise ValueError("workbook contains no worksheets (only chartsheets)")
    ws = wb_v.worksheets[0]
    wsf = wb_f[ws.title]
    S.sheet_name = ws.title
    if wb_v.sheetnames[0] != ws.title:
        S.note("First tab is not a worksheet",
               f"Tab '{wb_v.sheetnames[0]}' is a chartsheet; first worksheet '{ws.title}' used")
    if ws.sheet_state != "visible":
        S.note("Hidden first sheet", f"First worksheet '{ws.title}' is {ws.sheet_state}; processed anyway")

    for rng in ws.merged_cells.ranges:
        a = (rng.min_row, rng.min_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != a:
                    S.merged_parent[(r, c)] = a
    spans = {(rng.min_row, rng.min_col): (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
             for rng in ws.merged_cells.ranges}

    fcells = wsf._cells
    for (r, c), cell in ws._cells.items():
        v = cell.value
        fcell = fcells.get((r, c))
        formula = None
        if fcell is not None and fcell.data_type == "f":
            fv = fcell.value
            formula = fv if isinstance(fv, str) else getattr(fv, "text", None) or str(fv)
        if v is None:
            if formula and (r, c) not in S.merged_parent:
                S.formula_nocache.append((r, c, formula))
            continue
        if (r, c) in S.merged_parent:
            continue
        kind, extra = classify_value(v)
        S.cells[(r, c)] = Cell(r=r, c=c, raw=v, kind=kind, formula=formula, span=spans.get((r, c)), **extra)

    for (r, c, f) in S.formula_nocache:  # keep them visible in the grid as unresolved
        S.cells[(r, c)] = Cell(r=r, c=c, raw=f, kind="ERROR", formula=f, text="formula without cached value")

    for r, dim in ws.row_dimensions.items():
        if dim.hidden:
            S.hidden_rows.add(r)
    for _, dim in ws.column_dimensions.items():
        if dim.hidden and dim.min:
            S.hidden_cols.update(range(dim.min, (dim.max or dim.min) + 1))

    if S.cells:
        S.max_row = max(r for r, _ in S.cells)
        S.min_row = min(r for r, _ in S.cells)
        S.max_col = max(c for _, c in S.cells)
    wb_v.close()
    wb_f.close()
    return S


# --------------------------------------------------------------------------
# Period (month/year of the file)
# --------------------------------------------------------------------------
def resolve_period(S: SheetCtx) -> None:
    cands: dict[str, tuple] = OrderedDict()
    dates = [c.date for c in S.cells.values() if c.kind == "DATE"]
    if dates:
        cnt = Counter((d.year, d.month) for d in dates)
        (ym, n) = cnt.most_common(1)[0]
        if n >= max(3, 0.5 * len(dates)):
            cands["dates in sheet"] = ym
    rows = defaultdict(list)
    for c in S.cells.values():
        if c.kind == "TEXT":
            rows[c.r].append(c)
    top = sorted(rows)[:12]
    for r in top:
        for c in rows[r]:
            p = find_period(c.text)
            if p and "title text" not in cands:
                cands["title text"] = p
    p = find_period(S.sheet_name, loose=True)
    if p:
        cands["sheet name"] = p
    p = find_period(Path(S.fe.display).stem, loose=True)
    if p:
        cands["file name"] = p
    p = find_period(str(Path(S.fe.display).parent), loose=True)
    if p:
        cands["folder name"] = p
    S.period_candidates = dict(cands)
    if cands:
        src, per = next(iter(cands.items()))
        S.period, S.period_source = per, src
        distinct = set(cands.values())
        if len(distinct) > 1:
            S.note("Period conflict",
                   "Different month/year found: " + "; ".join(f"{k}={v[1]:02d}/{v[0]}" for k, v in cands.items())
                   + f" -> using {src}", "")
    else:
        S.note("Period unknown", "No month/year found in dates, titles, sheet name, file name or folder")

    for c in S.cells.values():
        if c.kind == "TEXT":
            d, st = parse_date_text(c.text, S.period)
            if st:
                c.date, c.date_status = d, st


# --------------------------------------------------------------------------
# Grid view (optionally transposed so that one algorithm handles both
# "dates down" and "dates across" layouts)
# --------------------------------------------------------------------------
class View:
    def __init__(self, S: SheetCtx, transposed: bool):
        self.S, self.t = S, transposed
        self.by_row: dict[int, dict[int, Cell]] = defaultdict(dict)
        self.by_col: dict[int, dict[int, Cell]] = defaultdict(dict)
        for (r, c), cell in S.cells.items():
            vr, vc = (c, r) if transposed else (r, c)
            self.by_row[vr][vc] = cell
            self.by_col[vc][vr] = cell
        self.max_row = max(self.by_row) if self.by_row else 0

    def orig(self, vr: int, vc: int) -> tuple[int, int]:
        return (vc, vr) if self.t else (vr, vc)

    def get(self, vr: int, vc: int) -> Cell | None:
        return self.by_row.get(vr, {}).get(vc)

    def filled(self, vr: int, vc: int) -> Cell | None:
        cell = self.get(vr, vc)
        if cell is not None:
            return cell
        p = self.S.merged_parent.get(self.orig(vr, vc))
        return self.S.cells.get(p) if p else None

    def span_width(self, cell: Cell) -> int:
        if not cell.span:
            return 1
        r1, c1, r2, c2 = cell.span
        return (r2 - r1 + 1) if self.t else (c2 - c1 + 1)

    def colref(self, vc: int) -> str:
        return f"row {vc}" if self.t else get_column_letter(vc)


@dataclass
class Axis:
    vc: int
    rows: dict                    # vr -> ("date", date) | ("day", n) | ("serial", date) | ("ambig", None)
    header_hint: bool
    n_real: int
    has_days: bool
    attr_cols: list = field(default_factory=list)


def longest_day_run(colcells: dict[int, Cell]) -> list[int]:
    best, cur, prev = [], [], None
    for vr in sorted(colcells):
        cell = colcells[vr]
        ok = cell.kind == "NUM" and float(cell.num).is_integer() and 1 <= cell.num <= 31
        if not ok:
            if cell.kind != "BLANK":
                best = cur if len(cur) > len(best) else best
                cur, prev = [], None
            continue
        n = int(cell.num)
        if prev is None or n == prev + 1 or n == prev or (prev >= 28 and n == 1):
            cur.append(vr)
        else:
            best = cur if len(cur) > len(best) else best
            cur = [vr]
        prev = n
    return cur if len(cur) > len(best) else best


def find_date_axes(view: View) -> list[Axis]:
    out = []
    for vc, colcells in view.by_col.items():
        rows: dict[int, tuple] = {}
        hint = any(c.kind == "TEXT" and RX["date_header"].match(c.text) for c in colcells.values())
        n_real = 0
        for vr, cell in colcells.items():
            if cell.kind == "DATE":
                rows[vr] = ("date", cell.date)
                n_real += 1
            elif cell.kind == "TEXT" and cell.date_status == "ok":
                rows[vr] = ("date", cell.date)
                n_real += 1
            elif cell.kind == "TEXT" and cell.date_status == "ambiguous":
                rows[vr] = ("ambig", None)
            elif hint and cell.kind == "NUM" and float(cell.num).is_integer() and 30000 <= cell.num <= 80000:
                rows[vr] = ("serial", from_excel(cell.num).date())
        run = longest_day_run(colcells)
        has_days = False
        if len(run) >= 5 or (hint and len(run) >= 3):
            for vr in run:
                rows.setdefault(vr, ("day", int(colcells[vr].num)))
            has_days = True
        if len(rows) >= 5 or (hint and len(rows) >= 3):
            out.append(Axis(vc=vc, rows=rows, header_hint=hint, n_real=n_real, has_days=has_days))
    return out


def choose_anchors(view: View, cands: list[Axis]) -> list[Axis]:
    anchors: list[Axis] = []
    for cand in sorted(cands, key=lambda a: a.vc):
        if anchors:
            prev = anchors[-1]
            overlap = len(set(prev.rows) & set(cand.rows))
            between_num = any(
                (cell := view.get(vr, vc)) is not None and cell.kind in NUMERIC_KINDS
                for vc in range(prev.vc + 1, cand.vc) for vr in cand.rows)
            if not between_num and overlap >= 0.5 * min(len(prev.rows), len(cand.rows)):
                better, other = (cand, prev) if cand.n_real > prev.n_real else (prev, cand)
                better.attr_cols = prev.attr_cols + cand.attr_cols + [other.vc]
                anchors[-1] = better
                continue
        anchors.append(cand)
    return anchors


# --------------------------------------------------------------------------
# Titles and section titles (original orientation)
# --------------------------------------------------------------------------
def detect_titles_sections(S: SheetCtx) -> None:
    rows = defaultdict(list)
    for c in S.cells.values():
        if c.kind != "BLANK":
            rows[c.r].append(c)
    data_rows = sorted(r for r, cs in rows.items() if any(c.kind in NUMERIC_KINDS or c.kind == "DATE" for c in cs))
    first_data = data_rows[0] if data_rows else S.max_row + 1
    srows = sorted(rows)
    for i, r in enumerate(srows):
        cs = rows[r]
        if len(cs) != 1 or cs[0].kind != "TEXT":
            continue
        cell = cs[0]
        txt = cell.text
        li = analyse_label(txt)
        # topmost text naming only a measure ("Shift wise production") is a title, not an area
        generic_title = r < first_data and (
            (i == 0 and bool(li.measures) and not RX["known_areas"].search(txt))
            or bool(re.search(r"\bwise\b", txt, re.I)))
        if (is_title_like(txt) and (r < first_data or cell.span)) or generic_title:
            S.titles.append(txt)
            cell.disposition, cell.detail = "title", "sheet/report title"
            continue
        if is_note_like(txt):
            if r < first_data:
                S.titles.append(txt)
                cell.disposition, cell.detail = "title", "text above first table"
            continue
        upcoming = [x for x in srows[i + 1:i + 7] if x <= r + 6]
        table_follows = any(len(rows[x]) >= 2 or any(c.kind in NUMERIC_KINDS or c.kind == "DATE" for c in rows[x])
                            for x in upcoming)
        if table_follows:
            S.section_rows[r] = txt
            S.section_cells[r] = cell
        elif r < first_data:
            S.titles.append(txt)
            cell.disposition, cell.detail = "title", "text above first table"


def section_for_row(S: SheetCtx, r: int, lower_bound: int = 0) -> tuple[str | None, int | None]:
    """Nearest section title above original row r (and below lower_bound)."""
    best = None
    for sr in S.section_rows:
        if lower_bound < sr <= r and (best is None or sr > best):
            best = sr
    return (S.section_rows[best], best) if best else (None, None)


# --------------------------------------------------------------------------
# Table / block analysis
# --------------------------------------------------------------------------
@dataclass
class RowStat:
    n_num: int = 0
    n_text: int = 0
    n_other: int = 0
    empty: bool = True
    single_text: bool = False
    merged_multi: bool = False
    header_like: bool = False
    summary_label: bool = False


def row_stat(view: View, vr: int, colset: set, avc: int | None) -> RowStat:
    st = RowStat()
    cells = [(vc, c) for vc, c in view.by_row.get(vr, {}).items() if vc in colset and c.kind != "BLANK"]
    for vc, c in cells:
        if c.kind in NUMERIC_KINDS and vc != avc:
            st.n_num += 1
        elif c.kind == "TEXT" and not RX["placeholders"].match(c.text):
            st.n_text += 1
            if any(rx.search(c.text) for _, rx in RX["summary"]) or RX["cumulative"].search(c.text):
                st.summary_label = True
            if view.span_width(c) > 1:
                st.merged_multi = True
        else:
            st.n_other += 1
    st.empty = not cells
    st.single_text = len(cells) == 1 and st.n_text == 1
    st.header_like = (st.n_text >= 2 and st.n_num == 0) or (st.n_text >= 3 and st.n_num * 3 <= st.n_text)
    return st


def order_key(item: tuple | None):
    if not item:
        return None
    kind, val = item
    if kind in ("date", "serial") and val:
        return ("d", val.toordinal())
    if kind == "day":
        return ("n", val)
    return None


def process_sheet(S: SheetCtx) -> None:
    if not S.cells:
        S.mode = "empty"
        S.note("Empty sheet", "First worksheet has no values")
        return
    resolve_period(S)
    detect_titles_sections(S)
    vview, hview = View(S, False), View(S, True)
    vc_cands, hc_cands = find_date_axes(vview), find_date_axes(hview)
    bv = max((len(a.rows) for a in vc_cands), default=0)
    bh = max((len(a.rows) for a in hc_cands), default=0)
    S.axis_scores = f"vertical={bv}, horizontal={bh}"
    if bv == 0 and bh == 0:
        S.mode, view, anchors = "matrix", vview, [None]
        S.note("No date axis", "No column/row of dates found - parsed as a period (monthly) summary table")
    elif bv >= bh:
        S.mode, view, anchors = "vertical", vview, choose_anchors(vview, vc_cands)
    else:
        S.mode, view, anchors = "horizontal", hview, choose_anchors(hview, hc_cands)
    if bv and bh:
        S.note("Both date axes present",
               f"Dates found down columns ({bv}) and across rows ({bh}); using {S.mode}. "
               "Cells of the other table go to unmapped if not covered.")

    all_cols = sorted(view.by_col)
    if anchors == [None]:
        tables = [(None, all_cols)]
    else:
        assign = defaultdict(list)
        avcs = [a.vc for a in anchors]
        for vc in all_cols:
            left = [a for a in avcs if a <= vc]
            assign[left[-1] if left else avcs[0]].append(vc)
        tables = [(a, assign[a.vc]) for a in anchors]
        if len(anchors) > 1:
            S.note("Multiple tables", f"{len(anchors)} date axes found ({', '.join(view.colref(a.vc) for a in anchors)})")
    for ti, (anchor, cols) in enumerate(tables, 1):
        process_table(S, view, anchor, cols, ti)
    finalise_leftovers(S)


def process_table(S: SheetCtx, view: View, anchor: Axis | None, cols: list[int], ti: int) -> None:
    colset = set(cols)
    avc = anchor.vc if anchor else None
    rs_cache: dict[int, RowStat] = {}

    def rs(vr: int) -> RowStat:
        if vr not in rs_cache:
            rs_cache[vr] = row_stat(view, vr, colset, avc)
        return rs_cache[vr]

    if anchor:
        key_rows = sorted(anchor.rows)
    else:
        key_rows = sorted(vr for vr in view.by_row if rs(vr).n_num > 0)
    if not key_rows:
        return

    # ---- split key rows into blocks --------------------------------------
    groups = [[key_rows[0]]]
    for p, q in zip(key_rows, key_rows[1:]):
        between = range(p + 1, q)
        hdr = any(rs(r).header_like for r in between)
        single = any(rs(r).single_text and not rs(r).summary_label for r in between)
        dec = False
        if anchor:
            kp, kq = order_key(anchor.rows.get(p)), order_key(anchor.rows.get(q))
            dec = kp is not None and kq is not None and kp[0] == kq[0] and kq[1] < kp[1]
            gap_split = (q - p - 1) > 6
            split = hdr or (single and dec) or gap_split
        else:
            split = hdr or single or (q - p - 1) > 2
        if split:
            groups.append([q])
        else:
            groups[-1].append(q)

    # ---- header bands ----------------------------------------------------
    blocks = []
    prev_last = 0
    for g in groups:
        pre, hdrs, blanks = [], [], 0
        r = g[0] - 1
        while r > prev_last and len(hdrs) < 10:
            st = rs(r)
            if st.empty:
                if hdrs:
                    break
                blanks += 1
                if blanks > 3:
                    break
            elif st.single_text and not st.merged_multi:
                break
            elif st.header_like or (st.n_num == 0 and st.n_text > 0):
                hdrs.append(r)
            elif st.n_num > 0 and not hdrs and anchor is not None:
                pre.append(r)
            else:
                break
            r -= 1
        hdrs.reverse()
        pre.reverse()
        blocks.append({"key": g, "hdrs": hdrs, "pre": pre})
        prev_last = g[-1]

    # ---- post zones (summary rows after the data) -------------------------
    for i, b in enumerate(blocks):
        if i + 1 < len(blocks):
            nb = blocks[i + 1]
            limit = min(nb["hdrs"] + nb["pre"] + [nb["key"][0]])
        else:
            limit = view.max_row + 1
        post, blanks, r = [], 0, b["key"][-1] + 1
        while r < limit:
            st = rs(r)
            if st.empty:
                blanks += 1
                if blanks >= 3:
                    break
            elif st.header_like and not st.summary_label:
                break
            elif st.single_text and not st.merged_multi and not st.summary_label:
                break
            elif st.n_num > 0 or st.summary_label or st.n_other > 0:
                post.append(r)
                blanks = 0
            else:
                break
            r += 1
        keyset = set(b["key"])
        intra = [r for r in range(b["key"][0], b["key"][-1] + 1) if r not in keyset and not rs(r).empty]
        b["post"], b["intra"] = post, intra

    prev_hdrs: list[int] = []
    prev_block_last = 0
    for bi, b in enumerate(blocks, 1):
        inherited = False
        if not b["hdrs"] and prev_hdrs:
            b["hdrs"], inherited = prev_hdrs, True
        process_block(S, view, anchor, cols, ti, bi, b, inherited, prev_block_last)
        prev_hdrs = b["hdrs"]
        prev_block_last = (b["post"] or b["key"])[-1]


def header_paths(view: View, hdrs: list[int], cols: list[int], avc: int | None) -> dict:
    """Return {vc: [(text, cell, forward_filled), ...]} with forward-fill of
    unmerged group headers (only into columns that have lower-level headers)."""
    out: dict[int, list] = defaultdict(list)
    scols = sorted(cols)
    for i, h in enumerate(hdrs):
        last = None
        for vc in scols:
            cell = view.filled(h, vc)
            if cell is not None and cell.kind != "BLANK":
                txt = label_of(cell)
                is_anchor_hdr = vc == avc or (cell.kind == "TEXT" and RX["date_header"].match(txt))
                last = None if is_anchor_hdr else (txt, cell)
                out[vc].append((txt, cell, False))
            else:
                lower = any((x := view.filled(h2, vc)) is not None and x.kind != "BLANK" for h2 in hdrs[i + 1:])
                if i < len(hdrs) - 1 and last and lower:
                    out[vc].append((last[0], last[1], True))
                else:
                    last = None
    for vc in list(out):
        dedup = []
        for item in out[vc]:
            if not dedup or norm_key(dedup[-1][0]) != norm_key(item[0]):
                dedup.append(item)
        out[vc] = dedup
    return out


def label_of(cell: Cell) -> str:
    if cell.kind == "DATE" and cell.date:
        return cell.date.isoformat()
    if cell.kind in ("NUM", "NUMTEXT", "TIME"):
        return fmt_num(cell.num) if cell.kind != "NUMTEXT" else cell.text
    return cell.text or str(cell.raw)


@dataclass
class Combined:
    measure: str | None = None
    measure_conflict: str | None = None
    unit: str | None = None
    summary: str | None = None
    cumulative: bool = False
    period: tuple | None = None
    descriptors: list = field(default_factory=list)
    date_header: bool = False


def combine_labels(texts: list[str]) -> Combined:
    out = Combined()
    for t in texts:  # top -> bottom; bottom-most measure wins
        li = analyse_label(t)
        if li.is_date_header:
            out.date_header = True
            continue
        if li.measures:
            if len(li.measures) > 1:
                out.measure = "/".join(li.measures)
                out.measure_conflict = f"multiple measure keywords in '{t}'"
            else:
                out.measure, out.measure_conflict = li.measures[0], None
        out.unit = li.unit or out.unit
        out.summary = li.summary or out.summary
        out.cumulative = out.cumulative or li.cumulative
        out.period = li.period or out.period
        if li.descriptor:
            out.descriptors.append(li.descriptor)
    return out


def classify_column(view: View, vc: int, hp: list, block_rows: list[int], anchor: Axis | None,
                    mode: str) -> tuple[str, Combined]:
    texts = [t for t, _, _ in hp]
    comb = combine_labels(texts)
    joined = " ".join(texts)
    cells = [c for vr in block_rows if (c := view.get(vr, vc)) is not None and c.kind != "BLANK"]
    n_num = sum(1 for c in cells if c.kind in NUMERIC_KINDS)
    n_txt = sum(1 for c in cells if c.kind == "TEXT" and not RX["placeholders"].match(c.text)
                and not c.date_status)
    n_ph = sum(1 for c in cells if c.kind == "TEXT" and RX["placeholders"].match(c.text))
    n_err = sum(1 for c in cells if c.kind == "ERROR")
    n_date = sum(1 for c in cells if c.kind == "DATE" or c.date_status)
    nonempty = len(cells)
    if anchor and vc in anchor.attr_cols:
        return "DATEATTR", comb
    if comb.date_header or (n_date and n_date >= 0.8 * nonempty):
        return "DATEATTR", comb
    if any(RX["serial_header"].match(t) for t in texts):
        return "SERIAL", comb
    if anchor and vc < anchor.vc and n_num and not texts:
        vals = [c.num for c in cells if c.kind == "NUM"]
        if len(vals) >= 3 and vals == list(range(1, len(vals) + 1)):
            return "SERIAL?", comb
    if RX["shift_header"].search(joined):
        return "SHIFT", comb
    if RX["remarks_header"].search(joined) and n_num <= n_txt:
        return "REMARKS", comb
    if RX["weekday_header"].match(joined) or (n_txt and all(
            c.kind != "TEXT" or RX["weekday_names"].match(c.text) for c in cells)):
        return "WEEKDAY", comb
    if n_num > 0 or n_err > 0 or (n_ph and not n_txt) or (comb.measure and not n_txt and nonempty):
        return "DATA", comb
    if n_txt > 0:
        if RX["row_area_header"].search(joined):
            return "ROW_AREA", comb
        if RX["row_process_header"].search(joined):
            return "ROW_PROCESS", comb
        if mode == "matrix" or not texts:
            return "ROW_LABEL", comb
        return "ATTR", comb
    return "EMPTY", comb


def resolve_row_date(S: SheetCtx, acell: Cell | None, anchor: Axis | None):
    """Return (date, day, granular_note, flag) for the anchor cell of a row."""
    if anchor is None or acell is None:
        return None, None, None
    if acell.kind == "DATE":
        return acell.date, acell.date.day, None
    if acell.kind == "TEXT" and acell.date_status == "ok":
        return acell.date, acell.date.day, None
    if acell.kind == "TEXT" and acell.date_status == "ambiguous":
        return None, None, f"ambiguous date text '{acell.text}'"
    if acell.kind == "NUM" and float(acell.num).is_integer():
        n = int(acell.num)
        if anchor.has_days and 1 <= n <= 31:
            if S.period:
                d = safe_date(S.period[0], S.period[1], n)
                if d:
                    return d, n, "date built from day number + file month/year"
                return None, n, f"day {n} does not exist in {S.period[1]:02d}/{S.period[0]}"
            return None, n, "day number but file month/year unknown"
        if anchor.header_hint and 30000 <= n <= 80000:
            d = from_excel(n).date()
            return d, d.day, "Excel serial number converted to date"
    return None, None, None


def process_block(S: SheetCtx, view: View, anchor: Axis | None, cols: list[int], ti: int, bi: int,
                  b: dict, inherited: bool, prev_block_last: int) -> None:
    avc = anchor.vc if anchor else None
    block_id = f"T{ti}B{bi}"
    rows = sorted(set(b["pre"] + b["key"] + b["intra"] + b["post"]))
    hdrs = b["hdrs"]
    hp_all = header_paths(view, hdrs, cols, avc)

    # dispositions for header cells
    if not inherited:
        for h in hdrs:
            for vc in cols:
                cell = view.get(h, vc)
                if cell is not None and cell.disposition is None:
                    cell.disposition, cell.detail = "header", block_id
    if inherited:
        S.note("Header inherited", f"Block {block_id} has no header rows; reused headers of previous block",
               f"view row {b['key'][0]}")

    # section (vertical: nearest single-text row above the block band)
    band_start = min(hdrs + b["pre"] + [b["key"][0]]) if not inherited else min(b["pre"] + [b["key"][0]])
    block_section = None
    if S.mode != "horizontal":
        block_section, srow = section_for_row(S, band_start, prev_block_last)
        if srow:
            cell = S.section_cells.get(srow)
            if cell and cell.disposition is None:
                cell.disposition, cell.detail = "section", block_id

    # column roles
    roles: dict[int, tuple[str, Combined]] = {}
    for vc in cols:
        if vc == avc:
            continue
        hp = hp_all.get(vc, [])
        roles[vc] = classify_column(view, vc, hp, rows, anchor, S.mode)
        role, comb = roles[vc]
        if role == "SERIAL?":
            S.note("Serial number column inferred", f"Unlabelled column {view.colref(vc)} holds 1..N; kept as attribute",
                   block_id)
        S.columns.append({
            "SourceFile": S.fe.display, "Sheet": S.sheet_name, "Block": block_id,
            "Column" if not view.t else "SourceRowOfLabel": view.colref(vc),
            "HeaderPath": " > ".join(t for t, _, _ in hp), "Role": role,
            "Measure": comb.measure or "", "Unit": comb.unit or "", "Summary": comb.summary or "",
            "Cumulative": comb.cumulative, "Descriptors": " | ".join(comb.descriptors),
            "HeaderForwardFilled": any(ff for _, _, ff in hp),
        })
    if any(ff for vc in hp_all for _, _, ff in hp_all[vc]):
        S.note("Unmerged group header", "Group header text forward-filled across following columns "
               "(e.g. area name written once above Input/Output/Target)", block_id)

    title_text = " / ".join(S.titles)
    title_comb = combine_labels(S.titles) if S.titles else Combined()
    fallback_area, fallback_src = None, None
    for src, txt in (("sheet title", title_text), ("sheet name", S.sheet_name), ("file name", Path(S.fe.display).stem)):
        m = RX["known_areas"].search(txt or "")
        if m:
            fallback_area, fallback_src = re.sub(r"\s+", " ", m.group(0)).strip().upper(), src
            break

    key_info = {}
    granularity = "Daily"
    if anchor:
        real = [anchor.rows[r][1] for r in b["key"] if anchor.rows[r][0] in ("date", "serial")]
        if real and all(d.day == 1 for d in real) and len({(d.year, d.month) for d in real}) >= 2:
            granularity = "Monthly"
            S.note("Monthly axis", "Date axis holds 1st-of-month dates spanning several months -> treated as monthly",
                   block_id)

    # day-number wrap detection (e.g. 26..31, 1..25)
    wrapped_after = None
    if anchor and anchor.has_days and not anchor.n_real:
        prev = None
        for r in b["key"]:
            k = anchor.rows[r]
            if k[0] == "day":
                if prev is not None and k[1] < prev - 20:
                    wrapped_after = r
                    S.note("Day-number wrap", "Day numbers restart inside a block (cross-month reporting period); "
                           "dates after the wrap left blank and flagged", block_id)
                    break
                prev = k[1]

    unlabelled_rows: list[int] = []
    block_vals: dict[int, dict[int, float]] = defaultdict(dict)   # vc -> vr -> value (dated rows)
    summary_cells: list[tuple] = []
    recs_this_block: list[dict] = []
    dated_rows: list[int] = []

    for vr in rows:
        acell = view.filled(vr, avc) if anchor else None
        date, day, dnote = resolve_row_date(S, acell, anchor)
        if wrapped_after is not None and vr >= wrapped_after and day is not None and date is not None:
            date, dnote = None, "day number after month wrap - month ambiguous"
        is_key = date is not None or day is not None or (dnote and "ambiguous" in dnote)

        # anchor cell disposition
        real_acell = view.get(vr, avc) if anchor else None
        if real_acell is not None and real_acell.disposition is None:
            if is_key and real_acell.kind in ("DATE", "TEXT", "NUM", "NUMTEXT"):
                real_acell.disposition, real_acell.detail = "date", block_id
            elif real_acell.kind == "TEXT":
                real_acell.disposition, real_acell.detail = "label", block_id
            elif real_acell.kind in NUMERIC_KINDS:
                S.unmap(real_acell, "Numeric value in the date column that is not a date/day", "Unrecognised date",
                        f"block {block_id}")
            elif real_acell.kind == "ERROR":
                S.unmap(real_acell, f"Excel error / unresolved formula in date column: {real_acell.text}",
                        "Excel error", f"block {block_id}")
            elif real_acell.kind == "BLANK":
                real_acell.disposition = "blank"

        # row labels and attributes
        label_texts, attrs, row_area, row_proc, row_generic = [], OrderedDict(), [], [], []
        shift = remarks = None
        if acell is not None and acell.kind == "TEXT" and not is_key:
            label_texts.append(acell.text)
        for vc in cols:
            if vc == avc:
                continue
            role, comb = roles[vc]
            if role in ("DATA", "EMPTY"):
                continue
            fcell = view.filled(vr, vc)
            rcell = view.get(vr, vc)
            if fcell is None or fcell.kind == "BLANK":
                if rcell is not None and rcell.disposition is None:
                    rcell.disposition = "blank"
                continue
            val = label_of(fcell)
            hname = " ".join(t for t, _, _ in hp_all.get(vc, [])) or view.colref(vc)
            if role == "SHIFT":
                shift = val
            elif role == "REMARKS":
                remarks = val if remarks is None else f"{remarks}; {val}"
            elif role == "ROW_AREA":
                row_area.append(val)
            elif role == "ROW_PROCESS":
                row_proc.append(val)
            elif role == "ROW_LABEL":
                row_generic.append(val)
            else:
                attrs[hname] = val
            if role in ("ROW_AREA", "ROW_PROCESS", "ROW_LABEL", "ATTR") and fcell.kind == "TEXT":
                label_texts.append(fcell.text)
            if rcell is not None and rcell.disposition is None:
                rcell.disposition, rcell.detail = "attribute", f"{block_id}:{role}"
        # text written in DATA columns on non-dated rows (e.g. 'TOTAL' in first data column)
        if not is_key:
            for vc in cols:
                if vc != avc and roles[vc][0] == "DATA":
                    c = view.get(vr, vc)
                    if c is not None and c.kind == "TEXT" and (
                            any(rx.search(c.text) for _, rx in RX["summary"]) or RX["cumulative"].search(c.text)):
                        label_texts.append(c.text)
                        c.disposition, c.detail = "label", block_id

        row_comb = combine_labels([t for t in label_texts if not RX["date_header"].match(t)])
        row_label = " | ".join(label_texts)
        row_type, is_summary = None, False
        if row_comb.summary:
            row_type, is_summary = row_comb.summary, True
        elif row_comb.cumulative and not is_key:
            row_type, is_summary = "Cumulative", True
        elif is_key:
            row_type = granularity
        elif row_comb.measure and not is_key:
            row_type = f"{row_comb.measure} row"
        elif S.mode == "matrix" and (row_generic or row_area or row_proc or row_comb.descriptors):
            row_type = "Period"
        elif label_texts:
            row_type = "Labelled row"
        else:
            row_type = None   # unlabelled
        if is_key:
            dated_rows.append(vr)

        # data cells
        data_cells = []
        for vc in cols:
            if vc == avc or roles[vc][0] != "DATA":
                continue
            cell = view.get(vr, vc)
            if cell is None or cell.disposition is not None:
                continue
            if cell.kind == "BLANK":
                cell.disposition = "blank"
                continue
            ctx = f"block {block_id}; header '{' > '.join(t for t, _, _ in hp_all.get(vc, []))}'; " \
                  f"row label '{row_label}'; date {date or ''}"
            if cell.kind in NUMERIC_KINDS:
                data_cells.append((vc, cell))
            elif cell.kind == "TEXT" and RX["placeholders"].match(cell.text):
                S.unmap(cell, f"Placeholder / non-numeric marker '{cell.text}' in a data column (not converted to 0)",
                        "Placeholder", ctx)
            elif cell.kind == "TEXT":
                S.unmap(cell, "Text in a numeric data column", "Text in data column", ctx)
            elif cell.kind == "ERROR":
                S.unmap(cell, f"Excel error or formula without cached value: {cell.text or cell.raw}",
                        "Excel error", ctx)
            elif cell.kind == "DATE":
                S.unmap(cell, "Date value in a numeric data column", "Date in data column", ctx)
        if not data_cells:
            continue
        if row_type is None:
            unlabelled_rows.append(vr)
            continue

        for vc, cell in data_cells:
            role, ccomb = roles[vc]
            rec = build_record(S, view, cell, vc, vr, block_id, ccomb, hp_all.get(vc, []), row_comb,
                               row_type, is_summary, is_key, date, day, dnote, shift, remarks, attrs,
                               row_area, row_proc, row_generic, row_label, block_section, title_comb,
                               fallback_area, fallback_src, granularity, inherited)
            recs_this_block.append(rec)
            if is_key and not rec["IsSummary"]:
                block_vals[vc][vr] = rec["Value"]
            if is_summary or rec["IsSummary"]:
                summary_cells.append((vc, vr, rec))

    # unlabelled numeric rows: accept as totals only if reconciliation proves it
    for vr in unlabelled_rows:
        cells = [(vc, c) for vc in cols if vc != avc and roles[vc][0] == "DATA"
                 and (c := view.get(vr, vc)) is not None and c.kind in NUMERIC_KINDS and c.disposition is None]
        matches = 0
        for vc, c in cells:
            vals = block_vals.get(vc, {})
            if vals and abs(sum(vals.values()) - c.num) <= max(0.01, 1e-6 * abs(c.num)):
                matches += 1
        if cells and vr > max(b["key"]) and matches >= max(1, math.ceil(0.8 * len(cells))):
            S.note("Unlabelled total row", f"Row without label equals the column sums in {matches}/{len(cells)} "
                   "columns -> kept as 'Total (unlabelled, verified)' and flagged", f"{block_id} view row {vr}")
            for vc, c in cells:
                role, ccomb = roles[vc]
                rec = build_record(S, view, c, vc, vr, block_id, ccomb, hp_all.get(vc, []), Combined(),
                                   "Total (unlabelled, verified)", True, False, None, None,
                                   "unlabelled row identified as total by reconciliation", None, None,
                                   OrderedDict(), [], [], [], "", block_section, title_comb,
                                   fallback_area, fallback_src, granularity, inherited)
                recs_this_block.append(rec)
                summary_cells.append((vc, vr, rec))
        else:
            prev_dates = [r for r in dated_rows if r < vr]
            for vc, c in cells:
                S.unmap(c, "Numeric value on a row with no date and no label (e.g. shift/continuation row "
                           "where the date is not repeated, or an unlabelled total that does not reconcile)",
                        "Row without date/label",
                        f"block {block_id}; header '{' > '.join(t for t, _, _ in hp_all.get(vc, []))}'; "
                        f"previous dated row (view) {prev_dates[-1] if prev_dates else '-'}")

    S.records.extend(recs_this_block)
    reconcile_block(S, view, block_id, roles, hp_all, block_vals, summary_cells, recs_this_block)

    dates = [r["Date"] for r in recs_this_block if isinstance(r.get("Date"), dt.date)]
    S.blocks.append({
        "SourceFile": S.fe.display, "Sheet": S.sheet_name, "Layout": S.mode, "Block": block_id,
        "DateAxis": view.colref(avc) if anchor else "(none)",
        "DateAxisType": ("dates" if anchor and anchor.n_real else "day numbers" if anchor and anchor.has_days
                         else "-") if anchor else "-",
        "HeaderRows": ", ".join(str(view.orig(h, avc or cols[0])[0 if not view.t else 1]) for h in hdrs),
        "HeaderInherited": inherited,
        "Section": block_section or "",
        "DataRows(dated)": len(b["key"]),
        "PreRows": len(b["pre"]), "IntraRows": len(b["intra"]), "PostRows": len(b["post"]),
        "DataColumns": sum(1 for v in roles.values() if v[0] == "DATA"),
        "AttributeColumns": ", ".join(f"{view.colref(vc)}:{r[0]}" for vc, r in roles.items()
                                      if r[0] not in ("DATA", "EMPTY")),
        "FirstDate": min(dates) if dates else None, "LastDate": max(dates) if dates else None,
        "Records": len(recs_this_block),
        "SummaryRecords": sum(1 for r in recs_this_block if r["IsSummary"]),
        "HeaderSignature": " || ".join(sorted({r["HeaderPath"] for r in recs_this_block})),
    })


def build_record(S, view, cell, vc, vr, block_id, ccomb: Combined, hp, row_comb: Combined, row_type,
                 row_is_summary, is_key, date, day, dnote, shift, remarks, attrs, row_area, row_proc,
                 row_generic, row_label, block_section, title_comb: Combined, fallback_area, fallback_src,
                 granularity, inherited) -> dict:
    review: list[str] = []
    r0, c0 = cell.r, cell.c
    # section for horizontal layouts follows the original row of the label
    section = block_section
    if S.mode == "horizontal":
        section, srow = section_for_row(S, r0)
        if srow:
            sc = S.section_cells.get(srow)
            if sc and sc.disposition is None:
                sc.disposition, sc.detail = "section", block_id
    sec_comb = combine_labels([section]) if section else Combined()

    # --- measure ---------------------------------------------------------
    measure, msrc = None, None
    if row_comb.measure and not is_key:
        measure, msrc = row_comb.measure, "row label"
        if ccomb.measure and ccomb.measure != row_comb.measure:
            review.append(f"row label measure '{row_comb.measure}' combined with column measure '{ccomb.measure}'")
            measure = f"{row_comb.measure} ({ccomb.measure})"
    elif ccomb.measure:
        measure, msrc = ccomb.measure, "header"
    else:
        for comb, src in ((combine_labels(row_area + row_proc + row_generic), "row label column"),
                          (row_comb, "row label"), (sec_comb, "section title")):
            if comb.measure:
                measure, msrc = comb.measure, src
                break
    if measure is None and title_comb.measure:
        measure, msrc = title_comb.measure, "sheet title"
        review.append(f"measure '{measure}' inferred from sheet title (no measure in headers)")
    if measure is None:
        measure = "Unspecified"
        review.append("no measure (Input/Output/Target/...) identified in headers or labels")
    if ccomb.measure_conflict:
        review.append(ccomb.measure_conflict)
    unit = ccomb.unit or row_comb.unit or sec_comb.unit or title_comb.unit
    cumulative = ccomb.cumulative or row_comb.cumulative

    # --- area / process ---------------------------------------------------
    hierarchy, seen = [], set()
    area_src = None
    parts = []
    if section and sec_comb.descriptors:
        parts.append(("section", sec_comb.descriptors))
    parts.append(("row label column", [d for v in row_area if (d := analyse_label(v).descriptor)]))
    parts.append(("header", ccomb.descriptors))
    parts.append(("row label column", [d for v in row_proc if (d := analyse_label(v).descriptor)]))
    if S.mode == "matrix" or (not is_key and row_type == "Labelled row"):
        parts.append(("row label", [d for v in row_generic if (d := analyse_label(v).descriptor)]))
        if not row_generic and row_type == "Labelled row":
            parts.append(("row label", row_comb.descriptors))
    elif S.mode == "horizontal":
        parts.append(("row label", [d for v in row_generic if (d := analyse_label(v).descriptor)]))
    for src, items in parts:
        for d in items:
            k = norm_key(d)
            if k and k not in seen and k != norm_key(measure):
                seen.add(k)
                hierarchy.append(d)
                area_src = area_src or src
    col_summary = ccomb.summary
    if not hierarchy and col_summary:
        hierarchy, area_src = ["ALL (column total)"], "header"
    if not hierarchy and row_is_summary:
        hierarchy, area_src = [f"ALL ({row_type.lower()} row)"], "row label"
    if not hierarchy and fallback_area:
        hierarchy, area_src = [fallback_area], fallback_src
    if not hierarchy:
        review.append("no mill/area identified")
    area = hierarchy[0] if hierarchy else ""
    process = " | ".join(hierarchy[1:])

    # --- record type ----------------------------------------------------
    is_summary = bool(row_is_summary)
    rtype = row_type
    if col_summary and not row_is_summary:
        rtype = f"{col_summary} (across columns)" if is_key else f"{row_type} / {col_summary} column"
        is_summary = True
    counts = (not is_summary) and (not cumulative) and rtype in ("Daily", "Monthly", "Period")

    # --- value -----------------------------------------------------------
    value = float(cell.num)
    if cell.kind == "NUMTEXT":
        review.append("number stored as text in Excel (converted)")
    if cell.kind == "TIME":
        unit = unit or "Hrs"
        review.append("time/duration value converted to decimal hours")
    if dnote:
        if "converted" in (dnote or "") or "built from day" in (dnote or ""):
            pass
        else:
            review.append(dnote)
    if is_key and date is None and day is None:
        review.append("date could not be resolved")
    if date and S.period and (date.year, date.month) != S.period and granularity == "Daily":
        review.append(f"date {date.isoformat()} outside file period {S.period[1]:02d}/{S.period[0]}")
    if inherited:
        review.append("headers inherited from previous block")
    if cell.formula:
        pass

    # --- period ------------------------------------------------------------
    if isinstance(date, dt.date):
        year, month = date.year, date.month
        period_src = "date"
    elif ccomb.period:
        (year, month), period_src = ccomb.period, "header"
    elif S.period:
        (year, month), period_src = S.period, f"file ({S.period_source})"
    else:
        year = month = None
        period_src = "unknown"

    measure_name = measure
    if unit and measure in CFG["quantity_measures"] and unit not in CFG["tonnage_units"]:
        measure_name = f"{measure} [{unit}]"
    if cumulative:
        measure_name = f"{measure_name} (Cum)"

    cell.disposition, cell.detail = "value", block_id
    rec = OrderedDict(
        RecordID=f"F{S.fe.idx:03d}-{cell.ref}",
        Date=date, Day=day, Month=month, Year=year,
        MonthName=calendar.month_abbr[month] if month else "",
        PeriodSource=period_src, Granularity=granularity if is_key else ("Period" if rtype else ""),
        Area=area, Process=process, AreaSource=area_src or "", Measure=measure, MeasureName=measure_name,
        MeasureSource=msrc or "", Basis="Cumulative" if cumulative else "Point",
        Value=value, RawValue=cell.raw if not isinstance(cell.raw, (dt.time, dt.timedelta)) else str(cell.raw),
        Unit=unit or "", Shift=shift or "", RecordType=rtype, IsSummary=is_summary,
        CountsAsProduction=counts,
        Section=section or "", SheetTitle=" / ".join(S.titles), HeaderPath=" > ".join(t for t, _, _ in hp),
        RowLabel=row_label, Remarks=remarks or "",
        RowAttributes="; ".join(f"{k}={v}" for k, v in attrs.items()),
        SourceFile=Path(S.fe.display).name, SourcePath=S.fe.display, ZipFile=S.fe.zip_name,
        SourceSheet=S.sheet_name, SourceRow=r0, SourceColumn=get_column_letter(c0), SourceCell=cell.ref,
        Formula=cell.formula or "", HiddenRow=r0 in S.hidden_rows, HiddenColumn=c0 in S.hidden_cols,
        Layout=S.mode, BlockID=block_id,
        ReviewFlag=bool(review), ReviewReason="; ".join(review),
        DuplicateFlag="", DuplicateGroup="",
        _vrow=vr, _vcol=vc, _file=S.fe.idx,
    )
    return rec


def reconcile_block(S, view, block_id, roles, hp_all, block_vals, summary_cells, recs) -> None:
    def status(expected, actual):
        diff = actual - expected
        if abs(diff) <= max(0.01, 1e-6 * abs(expected)):
            return "MATCH", diff
        if expected and abs(diff) / abs(expected) <= 0.005:
            return "CLOSE (<=0.5%)", diff
        return "MISMATCH", diff

    for vc, vr, rec in summary_cells:
        vals = list(block_vals.get(vc, {}).values())
        if not vals or rec["Basis"] == "Cumulative" or "across columns" in rec["RecordType"]:
            continue
        rt = rec["RecordType"]
        if rt.startswith(("Total", "Grand Total", "Subtotal", "Cumulative")):
            exp, check = sum(vals), "sum of dated rows"
            if rt.startswith("Subtotal"):
                before = [v for r, v in block_vals[vc].items() if r < vr]
                exp, check = sum(before), "sum of dated rows above"
        elif rt.startswith("Average"):
            exp, check = sum(vals) / len(vals), "mean of dated rows"
        elif rt.startswith("Maximum"):
            exp, check = max(vals), "max of dated rows"
        elif rt.startswith("Minimum"):
            exp, check = min(vals), "min of dated rows"
        else:
            continue
        st, diff = status(exp, rec["Value"])
        if st == "MISMATCH" and rt.startswith("Average"):
            nz = [v for v in vals if v]
            if nz:
                st2, diff2 = status(sum(nz) / len(nz), rec["Value"])
                if st2 != "MISMATCH":
                    st, diff, check = st2, diff2, "mean of non-zero dated rows"
        S.recon.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Block": block_id,
                        "Check": f"{rt} vs {check}", "Cell": rec["SourceCell"], "HeaderPath": rec["HeaderPath"],
                        "Area": rec["Area"], "MeasureName": rec["MeasureName"], "Expected": exp,
                        "Reported": rec["Value"], "Difference": diff, "Status": st, "ValuesUsed": len(vals)})

    # column totals: compare with sum of sibling columns of same measure in same row
    by_row = defaultdict(list)
    for rec in recs:
        by_row[rec["_vrow"]].append(rec)
    for vr, rlist in by_row.items():
        for rec in rlist:
            if "across columns" not in rec["RecordType"] and "column" not in rec["RecordType"]:
                continue
            sib = [x["Value"] for x in rlist if x is not rec and x["MeasureName"] == rec["MeasureName"]
                   and "column" not in x["RecordType"]]
            if len(sib) < 2:
                continue
            st, diff = status(sum(sib), rec["Value"])
            S.recon.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Block": block_id,
                            "Check": "Total column vs sum of other columns (same measure, same row)",
                            "Cell": rec["SourceCell"], "HeaderPath": rec["HeaderPath"], "Area": rec["Area"],
                            "MeasureName": rec["MeasureName"], "Expected": sum(sib), "Reported": rec["Value"],
                            "Difference": diff, "Status": st, "ValuesUsed": len(sib)})

    # cumulative columns must be non-decreasing
    cum = defaultdict(list)
    for rec in recs:
        if rec["Basis"] == "Cumulative" and rec["RecordType"] in ("Daily", "Monthly"):
            cum[rec["_vcol"]].append(rec)
    for vc, rl in cum.items():
        rl.sort(key=lambda x: x["_vrow"])
        bad = [x["SourceCell"] for a, x in zip(rl, rl[1:]) if x["Value"] < a["Value"] - 0.01]
        S.recon.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Block": block_id,
                        "Check": "Cumulative column is non-decreasing", "Cell": rl[0]["SourceCell"],
                        "HeaderPath": rl[0]["HeaderPath"], "Area": rl[0]["Area"],
                        "MeasureName": rl[0]["MeasureName"], "Expected": None, "Reported": None,
                        "Difference": None, "Status": "MATCH" if not bad else "MISMATCH (decreases at "
                        + ", ".join(bad[:5]) + ")", "ValuesUsed": len(rl)})


def finalise_leftovers(S: SheetCtx) -> None:
    """Every cell that no parser consumed goes to unmapped with a reason."""
    rows = defaultdict(dict)
    for c in S.cells.values():
        rows[c.r][c.c] = c

    def context(cell: Cell) -> str:
        left = [x for cc, x in sorted(rows[cell.r].items()) if cc < cell.c and x.kind == "TEXT"]
        above = [S.cells[(r, cell.c)] for r in range(cell.r - 1, max(0, cell.r - 30), -1)
                 if (r, cell.c) in S.cells and S.cells[(r, cell.c)].kind == "TEXT"]
        parts = []
        if left:
            parts.append(f"label left: '{left[-1].text}'")
        if above:
            parts.append(f"text above: '{above[0].text}'")
        return "; ".join(parts)

    for cell in sorted(S.cells.values(), key=lambda c: (c.r, c.c)):
        if cell.disposition is not None:
            continue
        if cell.kind == "BLANK":
            cell.disposition = "blank"
        elif cell.kind == "TEXT" and cell.r in S.section_rows:
            S.unmap(cell, "Section-like title not followed by a recognised table", "Unused section title",
                    context(cell))
        elif cell.kind == "TEXT":
            cat = "Note / free text" if is_note_like(cell.text) else "Text outside tables"
            S.unmap(cell, "Text not used as title, header or label", cat, context(cell))
        elif cell.kind in NUMERIC_KINDS:
            S.unmap(cell, "Numeric value outside any detected data table/block", "Numeric outside table",
                    context(cell))
        elif cell.kind == "DATE":
            S.unmap(cell, "Date value outside any detected date axis", "Date outside table", context(cell))
        elif cell.kind == "ERROR":
            S.unmap(cell, f"Excel error / formula without cached value: {cell.text or cell.raw}",
                    "Excel error", context(cell))
        else:
            S.unmap(cell, "Cell not consumed by parser", "Other", context(cell))
    for (r, c, f) in S.formula_nocache:
        S.error("Formula without cached value",
                f"{get_column_letter(c)}{r}: {f} (workbook not recalculated/saved by Excel - value unknown)",
                f"{get_column_letter(c)}{r}")


# --------------------------------------------------------------------------
# File discovery / ZIP extraction
# --------------------------------------------------------------------------
def safe_extract(zpath: Path, dest: Path, prefix: str, entries: list, ignored: list, depth: int = 0) -> None:
    with zipfile.ZipFile(zpath) as zf:
        for info in zf.infolist():
            name = info.filename
            if info.is_dir():
                continue
            target = (dest / name).resolve()
            if not str(target).startswith(str(dest.resolve())):
                ignored.append({"Path": f"{prefix}{name}", "Reason": "unsafe path in ZIP (skipped)"})
                continue
            base = Path(name).name
            if name.startswith("__MACOSX/") or base.startswith("._") or base in ("Thumbs.db", ".DS_Store"):
                ignored.append({"Path": f"{prefix}{name}", "Reason": "OS metadata file"})
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
            ext = Path(name).suffix.lower()
            if ext == ".zip" and depth < 5:
                sub = target.with_suffix("")
                sub = sub.parent / (sub.name + "__unzipped")
                sub.mkdir(parents=True, exist_ok=True)
                try:
                    safe_extract(target, sub, f"{prefix}{name}/", entries, ignored, depth + 1)
                except zipfile.BadZipFile as e:
                    ignored.append({"Path": f"{prefix}{name}", "Reason": f"nested ZIP unreadable: {e}"})
            else:
                entries.append((target, f"{prefix}{name}"))


def discover(input_path: Path, work_dir: Path) -> tuple[list[FileEntry], list[dict], list[dict]]:
    """Return (excel_files, other_files, zip_info)."""
    raw: list[tuple[Path, str, str]] = []
    ignored: list[dict] = []
    zips_info: list[dict] = []
    zips: list[Path] = []
    loose: list[Path] = []
    if input_path.is_file():
        if input_path.suffix.lower() == ".zip":
            zips = [input_path]
        else:
            loose = [input_path]
    else:
        for p in sorted(input_path.rglob("*")):
            if not p.is_file():
                continue
            try:
                p.resolve().relative_to(work_dir.resolve())
                continue
            except ValueError:
                pass
            if p.suffix.lower() == ".zip":
                zips.append(p)
            else:
                loose.append(p)
    if work_dir.exists():
        shutil.rmtree(work_dir)
    for z in zips:
        dest = work_dir / f"{z.stem}__{sha256(z)[:8]}"
        dest.mkdir(parents=True, exist_ok=True)
        entries: list = []
        info = {"ZipFile": z.name, "ZipPath": str(z), "SHA256": sha256(z), "Status": "extracted",
                "Members": 0, "Error": ""}
        try:
            safe_extract(z, dest, "", entries, ignored)
            info["Members"] = len(entries)
        except zipfile.BadZipFile as e:
            info["Status"], info["Error"] = "FAILED", f"bad ZIP: {e}"
        zips_info.append(info)
        raw += [(p, disp, z.name) for p, disp in entries]
    base = input_path if input_path.is_dir() else input_path.parent
    for p in loose:
        try:
            disp = str(p.relative_to(base))
        except ValueError:
            disp = p.name
        raw.append((p, disp, ""))

    excel, others = [], []
    for p, disp, zname in sorted(raw, key=lambda x: (x[2], x[1].lower())):
        ext = p.suffix.lower()
        if Path(disp).name.startswith("~$"):
            others.append({"Path": disp, "ZipFile": zname, "Reason": "Excel lock/temporary file (ignored)"})
        elif ext in EXCEL_EXTS:
            excel.append(FileEntry(idx=len(excel) + 1, path=p, display=disp, zip_name=zname, ext=ext,
                                   size=p.stat().st_size))
        elif ext in OTHER_SHEET_EXTS:
            others.append({"Path": disp, "ZipFile": zname,
                           "Reason": f"spreadsheet format {ext} not supported - convert to .xlsx (NOT processed)"})
        elif input_path.is_file() or zname:
            others.append({"Path": disp, "ZipFile": zname, "Reason": "not an Excel workbook (ignored)"})
    others += [{"Path": i["Path"], "ZipFile": "", "Reason": i["Reason"]} for i in ignored]
    return excel, others, zips_info


# --------------------------------------------------------------------------
# Post processing: wide master, duplicates
# --------------------------------------------------------------------------
BASE_MEASURE_ORDER = ["Input", "Output", "Target"]


def mark_duplicates(records: list[dict]) -> list[dict]:
    groups = defaultdict(list)
    for rec in records:
        if isinstance(rec["Date"], dt.date):
            when = rec["Date"].isoformat()
        else:
            when = f"{rec['Year']}-{rec['Month']}"
        key = (when, rec["Granularity"], norm_key(rec["Area"]), norm_key(rec["Process"]), rec["MeasureName"],
               norm_key(rec["Shift"]), rec["RecordType"], norm_key(rec["RowLabel"]) if not rec["Date"] else "")
        groups[key].append(rec)
    out = []
    gid = 0
    for key, recs in groups.items():
        if len(recs) < 2:
            continue
        gid += 1
        vals = {round(r["Value"], 6) for r in recs}
        kind = "EXACT" if len(vals) == 1 else "CONFLICT"
        recs.sort(key=lambda r: (r["_file"], r["SourceRow"], r["SourceColumn"]))
        for i, r in enumerate(recs):
            r["DuplicateGroup"] = f"D{gid:05d}"
            r["DuplicateFlag"] = f"DUPLICATE-{kind}" + (" (first occurrence)" if i == 0 else "")
            out.append({"DuplicateGroup": f"D{gid:05d}", "Type": kind, "Occurrence": i + 1,
                        "Date": r["Date"], "Year": r["Year"], "Month": r["Month"], "Area": r["Area"],
                        "Process": r["Process"], "MeasureName": r["MeasureName"], "RecordType": r["RecordType"],
                        "Value": r["Value"], "SourcePath": r["SourcePath"], "SourceCell": r["SourceCell"],
                        "RecordID": r["RecordID"]})
    return out


def build_wide(records: list[dict]) -> pd.DataFrame:
    groups: dict[tuple, list] = OrderedDict()
    for rec in records:
        key = (rec["_file"], rec["BlockID"], rec["_vrow"], norm_key(rec["Area"]), norm_key(rec["Process"]),
               rec["RecordType"], str(rec["Date"]), rec["Day"], rec["Shift"], rec["Section"])
        groups.setdefault(key, []).append(rec)
    rows, measures = [], set()
    for key, recs in groups.items():
        first = recs[0]
        vals: dict[str, float] = OrderedDict()
        cells_by_measure = []
        for r in recs:
            name = r["MeasureName"]
            n = 2
            while name in vals:
                name = f"{r['MeasureName']} #{n}"
                n += 1
            vals[name] = r["Value"]
            measures.add(name)
            cells_by_measure.append(f"{name}={r['SourceCell']}")
        units = OrderedDict((r["MeasureName"], r["Unit"]) for r in recs if r["Unit"])
        reviews = [r["ReviewReason"] for r in recs if r["ReviewReason"]]
        srows = sorted({r["SourceRow"] for r in recs})
        row = OrderedDict(
            RecordID="W-" + first["RecordID"],
            Date=first["Date"], Day=first["Day"], Month=first["Month"], Year=first["Year"],
            MonthName=first["MonthName"], **{"Mill/Area": first["Area"]}, Process=first["Process"],
        )
        row["_vals"] = vals
        row.update(OrderedDict(
            Units="; ".join(f"{k}={v}" for k, v in units.items()),
            Shift=first["Shift"], RecordType=first["RecordType"],
            IsSummary=any(r["IsSummary"] for r in recs),
            CountsAsProduction=any(r["CountsAsProduction"] for r in recs),
            Granularity=first["Granularity"], PeriodSource=first["PeriodSource"],
            Section=first["Section"], SheetTitle=first["SheetTitle"], RowLabel=first["RowLabel"],
            Remarks=first["Remarks"], RowAttributes=first["RowAttributes"],
            HeaderPaths=" || ".join(dict.fromkeys(r["HeaderPath"] for r in recs)),
            SourceFile=first["SourceFile"], SourceSheet=first["SourceSheet"],
            SourceRow=srows[0] if len(srows) == 1 else ";".join(map(str, srows)),
            SourceCells="; ".join(cells_by_measure), SourcePath=first["SourcePath"], ZipFile=first["ZipFile"],
            Layout=first["Layout"], BlockID=first["BlockID"],
            HiddenSource=any(r["HiddenRow"] or r["HiddenColumn"] for r in recs),
            ReviewFlag=any(r["ReviewFlag"] for r in recs),
            ReviewReason=" | ".join(dict.fromkeys(reviews)),
            DuplicateFlag="; ".join(dict.fromkeys(r["DuplicateFlag"] for r in recs if r["DuplicateFlag"])),
            LongRecordIDs=";".join(r["RecordID"] for r in recs),
        ))
        rows.append(row)

    def morder(m: str):
        base = m.split(" ")[0].split("[")[0]
        return (BASE_MEASURE_ORDER.index(base) if base in BASE_MEASURE_ORDER else 99, m)

    mcols = [m for m in BASE_MEASURE_ORDER] + sorted((m for m in measures if m not in BASE_MEASURE_ORDER),
                                                     key=morder)
    final = []
    for row in rows:
        vals = row.pop("_vals")
        out = OrderedDict()
        for k, v in row.items():
            out[k] = v
            if k == "Process":
                for m in mcols:
                    out[m] = vals.get(m)
        final.append(out)
    return pd.DataFrame(final)


# --------------------------------------------------------------------------
# Run
# --------------------------------------------------------------------------
LONG_COLS = ["RecordID", "Date", "Day", "Month", "Year", "MonthName", "PeriodSource", "Granularity", "Area",
             "Process", "AreaSource", "Measure", "MeasureName", "MeasureSource", "Basis", "Value", "RawValue",
             "Unit", "Shift", "RecordType", "IsSummary", "CountsAsProduction", "Section", "SheetTitle",
             "HeaderPath", "RowLabel", "Remarks", "RowAttributes", "SourceFile", "SourcePath", "ZipFile",
             "SourceSheet", "SourceRow", "SourceColumn", "SourceCell", "Formula", "HiddenRow", "HiddenColumn",
             "Layout", "BlockID", "ReviewFlag", "ReviewReason", "DuplicateFlag", "DuplicateGroup"]


def process_file(fe: FileEntry) -> SheetCtx:
    S = load_first_sheet(fe)
    process_sheet(S)
    return S


def accounting(S: SheetCtx) -> dict:
    disp = Counter(c.disposition for c in S.cells.values())
    total = len(S.cells)
    accounted = sum(disp.values()) - disp.get(None, 0)
    nonempty_rows = len({c.r for c in S.cells.values()})
    return {
        "SourceFile": S.fe.display, "Sheet": S.sheet_name,
        "UsedRange": f"A1:{get_column_letter(S.max_col or 1)}{S.max_row}",
        "RowsInspected": S.max_row, "NonEmptyRows": nonempty_rows, "NonEmptyCells": total,
        "Title": disp.get("title", 0), "Section": disp.get("section", 0), "Header": disp.get("header", 0),
        "Date": disp.get("date", 0), "Label": disp.get("label", 0), "Attribute": disp.get("attribute", 0),
        "ValueExtracted": disp.get("value", 0), "Unmapped": disp.get("unmapped", 0),
        "BlankText": disp.get("blank", 0), "Unaccounted": disp.get(None, 0),
        "Balanced": accounted == total and disp.get("value", 0) == len(S.records)
        and disp.get("unmapped", 0) == len(S.unmapped),
    }


def write_excel(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df = df.copy()
            if df.empty and not len(df.columns):
                df = pd.DataFrame({"Info": ["(none)"]})
            for col in df.columns:
                if df[col].dtype == object or pd.api.types.is_string_dtype(df[col]):
                    df[col] = df[col].map(sheet_safe).astype(object)
            max_rows = 1_000_000
            if len(df) <= max_rows:
                df.to_excel(xw, sheet_name=name[:31], index=False)
            else:
                for i in range(0, len(df), max_rows):
                    df.iloc[i:i + max_rows].to_excel(xw, sheet_name=f"{name[:26]}_{i // max_rows + 1}", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)):
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 8), 60)


def run(input_path: Path, out_dir: Path, inspect_only: bool = False) -> int:
    out_dir.mkdir(parents=True, exist_ok=True)
    work_dir = out_dir / "_extracted"
    started = dt.datetime.now()
    excel, others, zips_info = discover(input_path, work_dir)
    log.info("Found %d Excel workbook(s), %d other file(s)", len(excel), len(others))

    contexts: list[SheetCtx] = []
    files_rows, errors = [], []
    for fe in excel:
        fe.sha_before = sha256(fe.path)
        log.info("[%d/%d] %s", fe.idx, len(excel), fe.display)
        try:
            S = process_file(fe)
            fe.status = "processed"
            contexts.append(S)
        except Exception as e:  # noqa: BLE001 - report every failure, continue with others
            fe.status, fe.reason = "FAILED", f"{type(e).__name__}: {e}"
            errors.append({"SourceFile": fe.display, "Sheet": "", "Category": "File failed",
                           "Location": "", "Detail": fe.reason + "\n" + traceback.format_exc(limit=3)})
            log.error("  FAILED: %s", fe.reason)
            S = None
        fe.sha_after = sha256(fe.path)
        if fe.sha_after != fe.sha_before:
            errors.append({"SourceFile": fe.display, "Sheet": "", "Category": "SOURCE MODIFIED",
                           "Location": "", "Detail": "SHA-256 changed during processing"})
    zip_after = {}
    for zi in zips_info:
        zip_after[zi["ZipFile"]] = sha256(Path(zi["ZipPath"]))
        zi["UnchangedAfterRun"] = zip_after[zi["ZipFile"]] == zi["SHA256"]

    records = [r for S in contexts for r in S.records]
    unmapped = [u for S in contexts for u in S.unmapped]
    for S in contexts:
        errors.extend(S.errors)
    dups = mark_duplicates(records)
    special = [s for S in contexts for s in S.special]
    acct = [accounting(S) for S in contexts]
    for a in acct:
        if not a["Balanced"]:
            errors.append({"SourceFile": a["SourceFile"], "Sheet": a["Sheet"], "Category": "Cell accounting",
                           "Location": "", "Detail": "cell accounting does not balance"})
    for S in contexts:
        for w in S.load_warnings:
            special.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Category": "openpyxl warning",
                            "Location": "", "Detail": w})
        if S.hidden_rows or S.hidden_cols:
            special.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Category": "Hidden rows/columns",
                            "Location": "", "Detail": f"{len(S.hidden_rows)} hidden row(s), {len(S.hidden_cols)} "
                            "hidden column(s) - included in extraction and flagged (HiddenRow/HiddenColumn)"})
        if len(S.all_sheets) > 1:
            special.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Category": "Other sheets ignored",
                            "Location": "", "Detail": ", ".join(S.all_sheets[1:])})
        if S.formula_nocache:
            special.append({"SourceFile": S.fe.display, "Sheet": S.sheet_name, "Category": "Formulas without values",
                            "Location": "", "Detail": f"{len(S.formula_nocache)} formula cell(s) have no cached value; "
                            "open and save the file in Excel to resolve"})

    # ---- per-file summary ------------------------------------------------
    ctx_by_idx = {S.fe.idx: S for S in contexts}
    first_by_hash: dict[str, str] = {}
    dup_of: dict[int, str] = {}
    for fe in excel:
        if fe.sha_before in first_by_hash:
            dup_of[fe.idx] = first_by_hash[fe.sha_before]
            special.append({"SourceFile": fe.display, "Sheet": "", "Category": "Duplicate file",
                            "Location": "", "Detail": f"byte-identical to {first_by_hash[fe.sha_before]} - processed, "
                            "its records are flagged in DuplicateFlag"})
        else:
            first_by_hash[fe.sha_before] = fe.display
    for fe in excel:
        S = ctx_by_idx.get(fe.idx)
        recs = S.records if S else []
        dated = sorted(r["Date"] for r in recs if isinstance(r["Date"], dt.date))
        missing = ""
        if S and S.period and S.mode != "matrix" and dated:
            y, m = S.period
            present = {d for d in dated if (d.year, d.month) == (y, m)}
            days = calendar.monthrange(y, m)[1]
            miss = [d for d in range(1, days + 1) if dt.date(y, m, d) not in present]
            missing = ", ".join(map(str, miss)) if miss else "none"
        files_rows.append({
            "FileNo": fe.idx, "SourcePath": fe.display, "ZipFile": fe.zip_name, "Extension": fe.ext,
            "SizeBytes": fe.size, "Status": fe.status, "FailureReason": fe.reason,
            "FirstSheet": S.sheet_name if S else "", "AllSheets": ", ".join(S.all_sheets) if S else "",
            "SheetsIgnored": max(len(S.all_sheets) - 1, 0) if S else "",
            "Layout": S.mode if S else "", "AxisScores": S.axis_scores if S else "",
            "Blocks": sum(1 for b in S.blocks) if S else 0,
            "Period": f"{S.period[0]}-{S.period[1]:02d}" if S and S.period else "",
            "PeriodSource": S.period_source if S else "",
            "RowsInspected": S.max_row if S else 0, "NonEmptyCells": len(S.cells) if S else 0,
            "ValueRecords": len(recs), "DailyRecords": sum(1 for r in recs if r["CountsAsProduction"]),
            "SummaryRecords": sum(1 for r in recs if r["IsSummary"]),
            "ReviewFlagged": sum(1 for r in recs if r["ReviewFlag"]),
            "Unmapped": len(S.unmapped) if S else 0, "Errors": len(S.errors) if S else 0,
            "FirstDate": dated[0] if dated else None, "LastDate": dated[-1] if dated else None,
            "MissingDaysInPeriod": missing,
            "IdenticalToFile": dup_of.get(fe.idx, ""),
            "SHA256": fe.sha_before, "UnchangedAfterRun": fe.sha_after == fe.sha_before,
        })

    # ---- structure signature groups ---------------------------------------
    sig_groups = defaultdict(list)
    for S in contexts:
        sig = S.mode + " :: " + " ## ".join(sorted({b["HeaderSignature"] for b in S.blocks}))
        sig_groups[sig].append(S.fe.display)
    layout_rows = []
    base_cols = None
    for gi, (sig, files) in enumerate(sorted(sig_groups.items(), key=lambda x: -len(x[1])), 1):
        cols_set = set(sig.split(" :: ", 1)[-1].replace(" ## ", " || ").split(" || "))
        if base_cols is None:
            base_cols = cols_set
        layout_rows.append({"LayoutGroup": f"L{gi}", "Files": len(files), "FileList": "\n".join(files),
                            "Layout": sig.split(" :: ")[0],
                            "HeaderPathsOnlyInThisGroup": "\n".join(sorted(cols_set - base_cols))[:30000],
                            "HeaderPathsMissingVsGroupL1": "\n".join(sorted(base_cols - cols_set))[:30000],
                            "Signature": sig[:30000]})
    file_group = {f: row["LayoutGroup"] for row in layout_rows for f in row["FileList"].split("\n")}
    for fr in files_rows:
        fr["LayoutGroup"] = file_group.get(fr["SourcePath"], "")

    long_df = pd.DataFrame([{k: r.get(k) for k in LONG_COLS} for r in records], columns=LONG_COLS)
    wide_df = build_wide(records) if records else pd.DataFrame()
    unm_df = pd.DataFrame(unmapped, columns=["SourceFile", "ZipFile", "Sheet", "Row", "Column", "Cell", "Value",
                                             "ValueType", "Formula", "ReasonCategory", "Reason", "Context",
                                             "HiddenRow", "HiddenColumn"])
    review_df = long_df[long_df["ReviewFlag"] == True] if not long_df.empty else long_df  # noqa: E712

    # sample records: first 5 daily records of each layout group
    samples = []
    for row in layout_rows:
        f0 = row["FileList"].split("\n")[0]
        sub = [r for r in records if r["SourcePath"] == f0]
        sub = [r for r in sub if r["CountsAsProduction"]][:6] + [r for r in sub if r["IsSummary"]][:2]
        for r in sub:
            samples.append({"LayoutGroup": row["LayoutGroup"], **{k: r.get(k) for k in LONG_COLS[:27]},
                            "SourceCell": r["SourceCell"]})
    sample_df = pd.DataFrame(samples)

    n_found = len(excel)
    n_ok = sum(1 for fe in excel if fe.status == "processed")
    n_fail = n_found - n_ok
    rows_inspected = sum(a["RowsInspected"] for a in acct)
    n_wide = len(wide_df)
    n_dupe = sum(1 for d in dups if d["Occurrence"] > 1)
    recon_df = pd.DataFrame(S_recon := [x for S in contexts for x in S.recon])
    n_mismatch = int((recon_df["Status"].astype(str).str.startswith("MISMATCH")).sum()) if not recon_df.empty else 0
    n_close = int((recon_df["Status"].astype(str).str.startswith("CLOSE")).sum()) if not recon_df.empty else 0
    if not long_df.empty and n_wide:
        mcols = [c for c in wide_df.columns if c in set(long_df["MeasureName"]) or "#" in c]
        wide_sum = float(pd.to_numeric(wide_df[mcols].stack(), errors="coerce").sum())
        long_sum = float(long_df["Value"].sum())
        recon_df = pd.concat([recon_df, pd.DataFrame([{
            "SourceFile": "(all)", "Check": "Sum of values: production_master.csv vs production_master_long.csv",
            "Expected": long_sum, "Reported": wide_sum, "Difference": wide_sum - long_sum,
            "Status": "MATCH" if abs(wide_sum - long_sum) < 1e-6 * max(1, abs(long_sum)) else "MISMATCH"},
            {"SourceFile": "(all)", "Check": "Value cells: extracted + unmapped + consumed == non-empty cells",
             "Expected": sum(a["NonEmptyCells"] for a in acct),
             "Reported": sum(a["NonEmptyCells"] - a["Unaccounted"] for a in acct),
             "Difference": -sum(a["Unaccounted"] for a in acct),
             "Status": "MATCH" if all(a["Balanced"] for a in acct) else "MISMATCH"}])], ignore_index=True)
    summary = OrderedDict([
        ("Run started", started.strftime("%Y-%m-%d %H:%M:%S")),
        ("Script version", SCRIPT_VERSION),
        ("Input", str(input_path)),
        ("ZIP files", len(zips_info)),
        ("Files found (Excel .xlsx/.xlsm)", n_found),
        ("Files successfully processed", n_ok),
        ("Files failed", n_fail),
        ("Other files found (not processed)", len(others)),
        ("Total rows inspected (first sheets, used range)", rows_inspected),
        ("Non-empty rows inspected", sum(a["NonEmptyRows"] for a in acct)),
        ("Non-empty cells inspected", sum(a["NonEmptyCells"] for a in acct)),
        ("Total records extracted (production_master.csv rows)", n_wide),
        ("Value cells extracted (production_master_long.csv rows)", len(long_df)),
        ("  of which daily/base production values", int(long_df["CountsAsProduction"].sum()) if len(long_df) else 0),
        ("  of which totals/subtotals/averages/cumulative", int(long_df["IsSummary"].sum()) if len(long_df) else 0),
        ("Records flagged for review", len(review_df)),
        ("Total unmapped (cells in unmapped_data.xlsx)", len(unm_df)),
        ("Total duplicates (extra occurrences)", n_dupe),
        ("Duplicate groups", len({d['DuplicateGroup'] for d in dups})),
        ("Total errors", len(errors)),
        ("Reconciliation checks", len(recon_df)),
        ("Reconciliation mismatches", n_mismatch),
        ("Reconciliation close (<=0.5%, likely rounding)", n_close),
        ("Byte-identical duplicate files", len(dup_of)),
        ("Cell accounting balanced for all files", all(a["Balanced"] for a in acct)),
        ("Source files unchanged (SHA-256)", all(fe.sha_after == fe.sha_before for fe in excel)
         and all(z.get("UnchangedAfterRun", True) for z in zips_info)),
    ])

    # ---- write outputs -----------------------------------------------------
    structure_df = pd.DataFrame([b for S in contexts for b in S.blocks])
    columns_df = pd.DataFrame([c for S in contexts for c in S.columns])
    inspection = OrderedDict([
        ("Files", pd.DataFrame(files_rows)),
        ("Layout_Groups", pd.DataFrame(layout_rows)),
        ("Blocks", structure_df),
        ("Columns", columns_df),
        ("Special_Handling", pd.DataFrame(special)),
        ("Sample_Records", sample_df),
    ])
    write_excel(out_dir / "inspection_report.xlsx", inspection)
    print_inspection(excel, others, contexts, layout_rows, special, sample_df)
    if inspect_only:
        log.info("Inspection only - wrote %s", out_dir / "inspection_report.xlsx")
        return 0

    long_df.to_csv(out_dir / "production_master_long.csv", index=False, encoding="utf-8-sig")
    wide_df.to_csv(out_dir / "production_master.csv", index=False, encoding="utf-8-sig")
    write_excel(out_dir / "unmapped_data.xlsx", OrderedDict([
        ("Unmapped", unm_df),
        ("Review_Required", review_df),
        ("Reason_Summary", unm_df.groupby(["ReasonCategory"]).size().reset_index(name="Cells")
         if not unm_df.empty else pd.DataFrame()),
    ]))
    write_excel(out_dir / "extraction_report.xlsx", OrderedDict([
        ("Summary", pd.DataFrame({"Metric": list(summary), "Value": [str(v) for v in summary.values()]})),
        ("Files", pd.DataFrame(files_rows)),
        ("Files_Failed", pd.DataFrame([f for f in files_rows if f["Status"] != "processed"])),
        ("Other_Files", pd.DataFrame(others)),
        ("ZIPs", pd.DataFrame(zips_info)),
        ("Layout_Groups", pd.DataFrame(layout_rows)),
        ("Structure_Blocks", structure_df),
        ("Structure_Columns", columns_df),
        ("Errors", pd.DataFrame(errors)),
        ("Duplicates", pd.DataFrame(dups)),
        ("Reconciliation", recon_df),
        ("Cell_Accounting", pd.DataFrame(acct)),
        ("Special_Handling", pd.DataFrame(special)),
        ("Sample_Records", sample_df),
    ]))
    print_summary(summary, out_dir)
    return 0 if n_fail == 0 else 2


def print_inspection(excel, others, contexts, layout_rows, special, sample_df) -> None:
    ctx = {S.fe.idx: S for S in contexts}
    p = print
    p("\n" + "=" * 78)
    p("INSPECTION")
    p("=" * 78)
    p(f"Excel files found: {len(excel)}")
    for fe in excel:
        S = ctx.get(fe.idx)
        if S:
            per = f"{S.period[1]:02d}/{S.period[0]} ({S.period_source})" if S.period else "UNKNOWN"
            p(f"  [{fe.idx:3d}] {fe.display}\n        first sheet: '{S.sheet_name}' (of {len(S.all_sheets)})"
              f"  layout: {S.mode}  blocks: {len(S.blocks)}  period: {per}  "
              f"records: {len(S.records)}  unmapped: {len(S.unmapped)}")
        else:
            p(f"  [{fe.idx:3d}] {fe.display}\n        FAILED: {fe.reason}")
    if others:
        p(f"Other files (not processed): {len(others)}")
        for o in others[:30]:
            p(f"   - {o['Path']}: {o['Reason']}")
    p(f"\nStructure groups: {len(layout_rows)}")
    for row in layout_rows:
        p(f"  {row['LayoutGroup']}: {row['Files']} file(s), layout={row['Layout']}")
        if row["HeaderPathsOnlyInThisGroup"]:
            extra = row["HeaderPathsOnlyInThisGroup"].split("\n")
            p(f"      header paths not in L1 ({len(extra)}): " + "; ".join(extra[:6]) + (" ..." if len(extra) > 6 else ""))
        if row["HeaderPathsMissingVsGroupL1"]:
            miss = row["HeaderPathsMissingVsGroupL1"].split("\n")
            p(f"      header paths of L1 missing ({len(miss)}): " + "; ".join(miss[:6]) + (" ..." if len(miss) > 6 else ""))
    if special:
        p(f"\nSpecial handling notes: {len(special)} (see Special_Handling sheet)")
        for s in special[:25]:
            p(f"   - {s['SourceFile']}: [{s['Category']}] {s['Detail'][:140]}")
        if len(special) > 25:
            p(f"   ... {len(special) - 25} more")
    if not sample_df.empty:
        p("\nSample records:")
        cols = [c for c in ["LayoutGroup", "Date", "Area", "Process", "MeasureName", "Value", "RecordType",
                            "SourceFile", "SourceCell"] if c in sample_df.columns]
        with pd.option_context("display.width", 200, "display.max_columns", 20, "display.max_colwidth", 28):
            p(sample_df[cols].head(20).to_string(index=False))


def print_summary(summary: dict, out_dir: Path) -> None:
    keys = [("Files found", "Files found (Excel .xlsx/.xlsm)"),
            ("Files successfully processed", "Files successfully processed"),
            ("Files failed", "Files failed"),
            ("Total rows inspected", "Total rows inspected (first sheets, used range)"),
            ("Total records extracted", "Total records extracted (production_master.csv rows)"),
            ("  (value cells)", "Value cells extracted (production_master_long.csv rows)"),
            ("Total unmapped", "Total unmapped (cells in unmapped_data.xlsx)"),
            ("Total review-flagged", "Records flagged for review"),
            ("Total duplicates", "Total duplicates (extra occurrences)"),
            ("Total errors", "Total errors"),
            ("Reconciliation mismatches", "Reconciliation mismatches"),
            ("Cell accounting balanced", "Cell accounting balanced for all files"),
            ("Source files unchanged", "Source files unchanged (SHA-256)")]
    print("\n" + "=" * 78)
    print("VALIDATION SUMMARY")
    print("=" * 78)
    for label, k in keys:
        print(f"{label + ':':34s} {summary[k]}")
    print(f"\nOutputs written to: {out_dir.resolve()}")


def main(argv: list[str] | None = None) -> int:
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description="Extract production data from the first sheet of every workbook.")
    ap.add_argument("--input", type=Path, default=here / "input",
                    help="ZIP file, or folder containing ZIP(s)/workbooks (default: ./input)")
    ap.add_argument("--output", type=Path, default=here / "output", help="output folder (default: ./output)")
    ap.add_argument("--config", type=Path, default=None, help="JSON file extending keyword lists")
    ap.add_argument("--inspect-only", action="store_true", help="only inspect structure, do not extract")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    args.output.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(sys.stdout),
                                  logging.FileHandler(args.output / "extraction.log", mode="w", encoding="utf-8")])
    load_config(args.config)
    if not args.input.exists():
        log.error("Input not found: %s", args.input)
        return 1
    try:
        args.input.resolve().relative_to(args.output.resolve())
        log.error("Input must not be inside the output folder")
        return 1
    except ValueError:
        pass
    return run(args.input, args.output, args.inspect_only)


if __name__ == "__main__":
    sys.exit(main())
