"""Report Generator: upload an Excel/CSV export (e.g. from Metabase), get a PowerPoint deck.

How it is organised (one file, three layers of settings):

  1. Data rules    what is allowed in: blocked columns, row filters, small group hiding
  2. Report        title, currency, date column and range, time grain, footer
  3. Charts        a list of charts, each with its own X, Y, calculation, type, top N...

All the logic lives in build_slides(). It returns plain "slide data" that both the
on-screen preview and the PowerPoint writer use, so their numbers always match.
Everything above the line marked "STREAMLIT UI" works without Streamlit installed,
so you can test it with plain Python.

Run with:  streamlit run app.py
"""

import copy
import hashlib
import io
import json
import re
import uuid
import warnings
from datetime import date, timedelta
from types import SimpleNamespace

import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Application integrations

import streamlit as st

# ---------------------------------------------------------------------------

# Saved starting points for reports you run often. A preset is used when a file
# has every column listed in "detect". Charts only need the fields that differ
# from the defaults (see new_spec). Charts that use a missing column are skipped.
PRESETS = {
    "Loan requests": {
        "detect": ["Loan ID", "Amount Requested", "Status"],
        "settings": {
            "title": "Monthly Loan Report",
            "date_col": "Date Created",
            "blocked": ["Loan ID", "Customer Petra ID", "Customer Name"],
        },
        "charts": [
            {"chart_type": "kpi", "agg": "count", "title": "Loan requests"},
            {"chart_type": "kpi", "agg": "sum", "y": "Amount Requested", "title": "Total requested"},
            {"chart_type": "kpi", "agg": "sum", "y": "Disbursed Amount", "title": "Total disbursed"},
            {"chart_type": "column", "x": "Date Created", "agg": "sum", "y": "Amount Requested"},
            {"x": "Status", "agg": "sum", "y": "Amount Requested"},
            {"x": "Product", "agg": "sum", "y": "Amount Requested"},
            {"x": "Provider", "agg": "sum", "y": "Amount Requested"},
            {"x": "Source", "agg": "sum", "y": "Amount Requested"},
            {"x": "Loan Type", "agg": "sum", "y": "Amount Requested"},
        ],
    },
}

# Column name hints (lower case, matched anywhere in the name)
MEASURE_HINTS = ("amount", "balance", "principal", "disbursed", "repaid",
                 "outstanding", "value", "total", "fee", "interest")
SENSITIVE_HINTS = ("name", "phone", "mobile", "email", "address", "passport",
                   "ssn", "bvn", "nin", "birth")

# Type checks when a column is used: share of values that can be converted
OK_RATE = 0.95      # at or above: fine
BLOCK_RATE = 0.50   # below: the column cannot be used that way (between: warning)

# Identifier guard: a column used for grouping is blocked when almost every row differs
ID_RATIO = 0.5
ID_MIN_VALUES = 20

PIE_MAX = 6          # most slices a pie or doughnut may have
PIE_SHOW_MAX = 8     # pie is only offered when the column has this few values
MAX_SPLIT = 6        # most series when a chart is split by a second column
TABLE_MAX_ROWS = 15  # most rows on a table slide
MAX_KPI_CARDS = 8
MAX_TREND_POINTS = 120
FEW_VALUES = 12      # a number column with this few values can be used as categories

# Theme taken from the Loans report screenshots
NAVY = RGBColor(0x0E, 0x7C, 0x61)    # main green (titles, accents, big numbers)
MINT = RGBColor(0x86, 0xD0, 0xB8)    # light green (second series)
INK = RGBColor(0x1A, 0x2B, 0x25)     # near-black text
GREY = RGBColor(0x4A, 0x55, 0x50)
LIGHT = RGBColor(0xFF, 0xFF, 0xFF)   # card fill
PAGE_BG = RGBColor(0xF3, 0xF6, 0xF4) # slide background
BORDER = RGBColor(0xDD, 0xE5, 0xE1)
RED = RGBColor(0xC0, 0x39, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SERIES_COLORS = [RGBColor(0x0E, 0x7C, 0x61), RGBColor(0x86, 0xD0, 0xB8), RGBColor(0x2F, 0xA8, 0x8A),
                 RGBColor(0x5B, 0x8C, 0x7E), RGBColor(0xB5, 0xE3, 0xD4), RGBColor(0x3D, 0x5A, 0x50)]
HEX_COLORS = ["#0E7C61", "#86D0B8", "#2FA88A", "#5B8C7E", "#B5E3D4", "#3D5A50"]

# ---------------------------------------------------------------------------
# Option lists
# ---------------------------------------------------------------------------

AUTO = "bar/column (auto)"
OTHER = "Other"
BLANK = "(blank)"

AGGS = ["sum", "average", "count", "distinct count", "median", "min", "max"]
ADDITIVE = ("sum", "count", "distinct count")   # empty groups can safely be zero
MONEY_AGGS = ("sum", "average", "median", "min", "max")
SORTS = ["auto", "value (high to low)", "label (A to Z)", "natural order"]
NUMBER_FORMATS = ["full", "thousands (K)", "millions (M)", "percent of total"]
EXCEL_FORMATS = {"full": "#,##0", "thousands (K)": '#,##0.0,"K"',
                 "millions (M)": '#,##0.0,,"M"', "percent of total": "0.0%"}
GRAINS = ["day", "week", "month", "quarter", "year"]
DATE_PRESETS = ["All data", "Last 7 days", "Last 30 days", "Last 90 days", "This month", "Last month",
                "Quarter to date", "Year to date", "Custom"]
DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAY_CODES = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]
MONTH_CODES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
FILTER_MODES = ["include only", "exclude", "not blank"]
SUPPRESS_MODES = ["merge into Other", "hide"]
UNDATED_MODES = ["exclude", "include in non-trend charts"]

STACKED = ("stacked column", "stacked bar")
SPLIT_TYPES = ("line",) + STACKED
KEEP_ZERO_TYPES = ("column", "bar", "area") + STACKED

PPTX_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "stacked column": XL_CHART_TYPE.COLUMN_STACKED,
    "stacked bar": XL_CHART_TYPE.BAR_STACKED,
}

DEFAULT_SETTINGS = {
    "title": "Monthly Loan Report", "currency": "", "footer": "",
    "date_col": None, "date_preset": "All data", "start": None, "end": None,
    "grain": "month", "week_start": "Monday", "fiscal_start": 1,
    "undated": "exclude", "fill_empty": True, "compare_prev": False,
    "quality_slide": False, "summary_headline": "", "summary_notes": "",
    "min_group": 0, "suppress_mode": "merge into Other",
    "blocked": [], "filters": [],
}


def new_id():
    return uuid.uuid4().hex[:8]


def new_spec(**changes):
    """One chart. x=None with chart_type 'kpi' makes a summary card."""
    spec = {
        "id": new_id(), "include": True, "title": "",
        "x": None, "y": None, "agg": "count", "chart_type": AUTO,
        "top_n": 10, "other_bucket": False, "sort": "auto",
        "time_grain": "report default", "split": None, "filters": [],
        "number_format": "full", "show_labels": True, "gridlines": False,
        "show_blanks": True, "rename": "", "notes": "",
        "section": "", "headline": "",
    }
    spec.update(changes)
    return spec


def new_filter(**changes):
    f = {"id": new_id(), "col": None, "mode": "include only", "values": []}
    f.update(changes)
    return f


# ---------------------------------------------------------------------------
# Loading and converting columns
# ---------------------------------------------------------------------------


def load_file(source, name, sheet=None):
    """Read a CSV, or one sheet of an Excel file, with tidy column names."""
    if name.lower().endswith(".csv"):
        df = pd.read_csv(source)
    else:
        df = pd.read_excel(source, sheet_name=sheet or 0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    for col in df.columns:  # empty strings count as blank
        if not pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].map(lambda v: np.nan if isinstance(v, str) and not v.strip() else v)
    return df.reset_index(drop=True)


def is_id_like(name):
    return "id" in re.split(r"[\W_]+", name.lower())


def is_sensitive(name):
    return any(h in name.lower() for h in SENSITIVE_HINTS)


def to_numeric(s):
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)
    cleaned = s.astype(str).str.replace(r"[,\s$£€₦%]", "", regex=True)
    return pd.to_numeric(cleaned.where(s.notna()), errors="coerce")


def to_datetime(s):
    if pd.api.types.is_numeric_dtype(s):  # numbers are not dates
        return pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            parsed = pd.to_datetime(s, errors="coerce", utc=True, format="mixed")
        except (TypeError, ValueError):
            parsed = pd.to_datetime(s, errors="coerce", utc=True)
    return parsed.dt.tz_localize(None)


def category_key(series):
    """Clean text labels for grouping, with blanks made visible."""
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
        return series.map(lambda v: BLANK if pd.isna(v) else f"{v:.10g}")
    return series.astype(object).where(series.notna(), BLANK).astype(str)


def convert_column(series, kind):
    """Try to read a column as 'date', 'number' or 'category'.

    Returns (converted, status, message). Status is ok, warn or block.
    """
    if kind == "category":
        return category_key(series), "ok", ""
    converted = to_datetime(series) if kind == "date" else to_numeric(series)
    non_null = int(series.notna().sum())
    if non_null == 0:
        return converted, "block", "This column is empty."
    good = int(converted.notna().sum())
    rate = good / non_null
    lost = non_null - good
    if rate >= OK_RATE:
        return converted, "ok", ""
    if rate >= BLOCK_RATE:
        return converted, "warn", (f"{lost:,} of {non_null:,} values ({lost / non_null:.0%}) "
                                   f"could not be read as {kind}s and are left out.")
    return converted, "block", f"Only {rate:.0%} of values can be read as {kind}s, so it cannot be used that way."


class Converter:
    """Converts each column at most once per report build."""

    def __init__(self, df):
        self.df = df
        self.cache = {}

    def get(self, col, kind):
        if (col, kind) not in self.cache:
            self.cache[(col, kind)] = convert_column(self.df[col], kind)
        return self.cache[(col, kind)]


def profile_columns(df):
    """Guess each column's type. Only a label: the real check happens when a column is used."""
    info = {}
    for col in df.columns:
        s = df[col]
        non_null = int(s.notna().sum())
        kind = "category"
        if pd.api.types.is_bool_dtype(s):
            pass
        elif pd.api.types.is_datetime64_any_dtype(s):
            kind = "date"
        elif pd.api.types.is_numeric_dtype(s):
            kind = "number"
        elif non_null:
            sample = s.dropna().astype(str).head(200)
            if to_numeric(sample).notna().mean() >= 0.8:
                kind = "number"
            elif sample.str.contains(r"\d").mean() >= 0.8 and to_datetime(sample).notna().mean() >= 0.8:
                kind = "date"
        info[col] = {"kind": kind, "non_null": non_null, "rows": len(df),
                     "nunique": int(s.nunique()),
                     "id_like": is_id_like(col), "sensitive": is_sensitive(col)}
    return info


def looks_like_identifier(meta):
    if meta["id_like"]:
        return True
    return (meta["nunique"] >= ID_MIN_VALUES and meta["non_null"] > 0
            and meta["nunique"] / meta["non_null"] >= ID_RATIO)


def column_problem(col, info, blocked):
    """Why a column cannot be used to group a chart (None if it can)."""
    if not col:
        return "Pick an X column."
    if col not in info:
        return f"'{col}' is not in this file."
    if col in blocked:
        return f"'{col}' is blocked by your data rules."
    if info[col]["kind"] != "date" and looks_like_identifier(info[col]):
        return (f"'{col}' looks like an identifier (almost every row is different), "
                "so a chart grouped by it would expose individuals.")
    return None


# ---------------------------------------------------------------------------
# Dates
# ---------------------------------------------------------------------------


def to_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def add_months(d, n):
    years, month0 = divmod(d.month - 1 + n, 12)
    return date(d.year + years, month0 + 1, 1)


def resolve_range(preset, custom_start, custom_end, data_min, data_max, today, fiscal_start):
    """Turn a preset like 'Last month' into a (start, end) pair of dates."""
    if preset in ("Last 7 days", "Last 30 days", "Last 90 days"):
        days = int(preset.split()[1])
        return today - timedelta(days=days - 1), today
    if preset == "This month":
        return today.replace(day=1), today
    if preset == "Last month":
        first = today.replace(day=1)
        return add_months(first, -1), first - timedelta(days=1)
    if preset in ("Quarter to date", "Year to date"):
        year = today.year if today.month >= fiscal_start else today.year - 1
        year_start = date(year, fiscal_start, 1)
        if preset == "Year to date":
            return year_start, today
        months_in = (today.month - fiscal_start) % 12
        return add_months(year_start, (months_in // 3) * 3), today
    if preset == "Custom" and to_date(custom_start) and to_date(custom_end):
        return to_date(custom_start), to_date(custom_end)
    return data_min, data_max


def to_periods(dates, grain, week_start, fiscal_start):
    """Bucket dates into day/week/month/quarter/year periods. Returns (periods, freq)."""
    year_end = MONTH_CODES[(fiscal_start - 2) % 12]
    if grain == "day":
        freq = "D"
    elif grain == "week":
        freq = "W-" + DAY_CODES[(DAY_NAMES.index(week_start) - 1) % 7]
    elif grain == "month":
        freq = "M"
    elif grain == "quarter":
        freq = "Q-" + year_end
    else:
        freq = "Y-" + year_end
    try:
        periods = dates.dt.to_period(freq)
    except (ValueError, KeyError):
        freq = freq.replace("Y-", "A-")  # older pandas names
        periods = dates.dt.to_period(freq)
    return periods, periods.dtype.freq


def period_label(p, grain):
    if grain == "day":
        return p.strftime("%d %b %Y")
    if grain == "week":
        return "Week of " + p.start_time.strftime("%d %b %y")
    if grain == "month":
        return p.strftime("%b %Y")
    if grain == "quarter":
        return f"{p.year} Q{p.quarter}"
    return str(p.year)


# ---------------------------------------------------------------------------
# Small helpers: filters, labels, numbers
# ---------------------------------------------------------------------------


def apply_filters(df, filters, warns):
    for f in filters:
        col = f.get("col")
        if not col:
            continue
        if col not in df.columns:
            warns.append(f"Filter column '{col}' is not in this file and was ignored.")
            continue
        if f.get("mode") == "not blank":
            df = df[df[col].notna()]
        elif f.get("values"):
            mask = category_key(df[col]).isin(f["values"])
            df = df[~mask] if f.get("mode") == "exclude" else df[mask]
    return df


def parse_rename(text):
    """'Y=Yes, N=No' becomes {'Y': 'Yes', 'N': 'No'}."""
    out = {}
    for part in re.split(r"[,\n]", text or ""):
        if "=" in part:
            old, new = (p.strip() for p in part.split("=", 1))
            if old:
                out[old] = new
    return out


def labels_for(series, rename):
    key = category_key(series)
    return key.map(lambda v: rename.get(v, v)) if rename else key


def format_value(v, fmt="full", currency=""):
    if v is None or pd.isna(v):
        return "n/a"
    if fmt.startswith("thousands"):
        body = f"{v / 1e3:,.1f}K"
    elif fmt.startswith("millions"):
        body = f"{v / 1e6:,.1f}M"
    elif float(v).is_integer() or abs(v) >= 1000:
        body = f"{v:,.0f}"
    else:
        body = f"{v:,.2f}"
    return f"{currency} {body}".strip() if currency else body


def excel_format(fmt, values):
    if fmt == "full":
        arr = np.asarray(values, dtype=float)
        arr = arr[~np.isnan(arr)]
        if len(arr) and np.abs(arr).max() < 100 and not np.all(arr == np.round(arr)):
            return "#,##0.00"
    return EXCEL_FORMATS.get(fmt, "#,##0")


def value_label(agg, y):
    if agg == "count":
        return "Records"
    names = {"sum": "Total", "average": "Average", "median": "Median",
             "min": "Minimum", "max": "Maximum", "distinct count": "Distinct"}
    return f"{names[agg]} {y}"


def chart_types_for(xkind, nunique=0):
    """Chart types that make sense for the type of the X column."""
    if xkind == "date":
        return ["column", "line", "area", "stacked column", "table"]
    types = [AUTO, "column", "bar"]
    if nunique <= PIE_SHOW_MAX:
        types += ["pie", "doughnut"]
    return types + ["stacked column", "stacked bar", "table"]


def aggregate(work, keys, agg):
    """Group work (columns x, y, s) by keys and calculate agg."""
    if agg == "count":
        out = work.groupby(keys).size()
    else:
        g = work.groupby(keys)["y"]
        out = {"sum": g.sum, "average": g.mean, "median": g.median,
               "min": g.min, "max": g.max, "distinct count": g.nunique}[agg]()
    return out.dropna()


def scalar(sub, y, agg):
    """One number for a summary card."""
    if agg == "count":
        return float(len(sub))
    vals = y.loc[sub.index]
    if agg == "distinct count":
        return float(vals.nunique())
    vals = vals.dropna()
    if vals.empty:
        return float("nan")
    return float({"sum": vals.sum, "average": vals.mean, "median": vals.median,
                  "min": vals.min, "max": vals.max}[agg]())


def get_measure(spec, ctx, errors, warns):
    """The Y column ready to calculate with (None for count)."""
    agg, y = spec["agg"], spec["y"]
    if agg not in AGGS:
        errors.append(f"Unknown calculation '{agg}'.")
        return None
    if agg == "count":
        return None
    if not y or y not in ctx.df.columns:
        errors.append("Pick the column to calculate.")
        return None
    if y in ctx.blocked and agg != "distinct count":
        errors.append(f"'{y}' is blocked by your data rules (only distinct count is allowed).")
        return None
    if agg == "distinct count":
        return ctx.df[y]
    series, status, msg = ctx.conv.get(y, "number")
    if status == "block":
        errors.append(f"'{y}': {msg}")
    elif status == "warn":
        warns.append(f"'{y}': {msg}")
    return series


def auto_title(spec, grain=None):
    label = value_label(spec["agg"], spec["y"])
    if spec["chart_type"] == "kpi" or not spec["x"]:
        return label
    by = f"{grain} ({spec['x']})" if grain else spec["x"]
    return f"{label} by {by}"


# ---------------------------------------------------------------------------
# Building slide data (no Streamlit in here)
# ---------------------------------------------------------------------------


def skipped(spec, errors, warns):
    return {"kind": "skipped", "id": spec["id"], "title": spec["title"] or auto_title(spec),
            "errors": errors, "warnings": warns}


def hide_other_if_small(work, col, k, warns):
    if k > 0 and col in work and (work[col] == OTHER).any() and (work[col] == OTHER).sum() < k:
        warns.append(f"Some groups were hidden because they hold fewer than {k} records in total.")
        return work[work[col] != OTHER]
    return work


def suppress_small(work, col, k, mode, warns, foot, name):
    """Groups with fewer than k records are merged into Other or hidden."""
    counts = work[col].value_counts()
    small = counts[counts < k].index
    if len(small) == 0:
        return work
    work = work.copy()
    if mode == "merge into Other":
        work.loc[work[col].isin(small), col] = OTHER
        text = f"Groups with fewer than {k} records are combined into Other."
    else:
        work = work[~work[col].isin(small)]
        text = f"Groups with fewer than {k} records are not shown."
    warns.append(f"{len(small)} small group(s) in '{name}' affected by the small group rule.")
    if text not in foot:
        foot.append(text)
    return work


def render_chart(spec, ctx):
    s = ctx.settings
    x, agg, ctype = spec["x"], spec["agg"], spec["chart_type"]
    split = spec["split"] if ctype in SPLIT_TYPES else None
    errors, warns, foot = [], [], []

    # 1. Check the choices before touching any data
    problem = column_problem(x, ctx.info, ctx.blocked)
    if problem:
        errors.append(problem)
    xkind = "date" if x in ctx.info and ctx.info[x]["kind"] == "date" else "category"
    if ctype not in chart_types_for(xkind):
        errors.append(f"'{ctype}' is not available when X is a {xkind} column.")
    if ctype in ("pie", "doughnut") + STACKED and agg not in ("sum", "count"):
        errors.append(f"'{ctype}' charts only work with sum or count.")
    if ctype in STACKED and not split:
        errors.append("Stacked charts need a 'Split by' column.")
    if split:
        p = column_problem(split, ctx.info, ctx.blocked)
        if p:
            errors.append(p)
        elif ctx.info[split]["kind"] == "date":
            errors.append("'Split by' must be a category column, not a date.")
        elif split == x:
            errors.append("'Split by' must differ from X.")
    elif spec["split"]:
        warns.append("'Split by' is ignored for this chart type.")
    yseries = get_measure(spec, ctx, errors, warns)
    if errors:
        return skipped(spec, errors, warns)

    # 2. Rows for this chart
    sub = apply_filters(ctx.cur, spec["filters"], warns)
    if sub.empty:
        return skipped(spec, ["No rows left after the filters."], warns)
    rename = parse_rename(spec["rename"])
    work = pd.DataFrame(index=sub.index)
    grain = freq = None
    if xkind == "date":
        dser, status, msg = ctx.conv.get(x, "date")
        if status == "block":
            return skipped(spec, [f"'{x}': {msg}"], warns)
        if status == "warn":
            warns.append(f"'{x}': {msg}")
        grain = spec["time_grain"] if spec["time_grain"] in GRAINS else s["grain"]
        work["x"], freq = to_periods(dser.loc[sub.index], grain, s["week_start"], s["fiscal_start"])
        missing = int(work["x"].isna().sum())
        if missing:
            warns.append(f"{missing:,} rows have no {x} and are not in this chart.")
        work = work[work["x"].notna()]
    else:
        work["x"] = labels_for(sub[x], rename)
    if yseries is not None:
        work["y"] = yseries.loc[sub.index]
    if split:
        work["s"] = labels_for(sub[split], rename)
    if not spec["show_blanks"]:
        work = work[work["x"] != BLANK]
        if split:
            work = work[work["s"] != BLANK]
    if work.empty:
        return skipped(spec, ["No rows left for this chart."], warns)

    # 3. Small groups, top N and Other (category X only)
    k = int(s["min_group"])
    trimmed, top_n = False, int(spec["top_n"])
    if k > 0 and xkind == "category":
        work = suppress_small(work, "x", k, s["suppress_mode"], warns, foot, x)
    if k > 0 and split:
        work = suppress_small(work, "s", k, "merge into Other", warns, foot, split)
    if xkind == "category":
        if ctype in ("pie", "doughnut"):
            top_n = min(top_n, PIE_MAX - (1 if spec["other_bucket"] else 0))
        if ctype == "table":
            top_n = min(top_n, TABLE_MAX_ROWS)
        if work.empty:
            return skipped(spec, ["No rows left after hiding small groups."], warns)
        ranked = aggregate(work, ["x"], agg).drop(OTHER, errors="ignore").sort_values(ascending=False)
        keep = list(ranked.index[:top_n])
        outside = ~work["x"].isin(keep) & (work["x"] != OTHER)
        if outside.any():
            trimmed = True
            work = work.copy()
            if spec["other_bucket"]:
                work.loc[outside, "x"] = OTHER
            else:
                work = work[~outside]
        work = hide_other_if_small(work, "x", k, warns)
    if split:
        ranked = aggregate(work, ["s"], agg).drop(OTHER, errors="ignore").sort_values(ascending=False)
        keep = list(ranked.index[:MAX_SPLIT])
        outside = ~work["s"].isin(keep) & (work["s"] != OTHER)
        if outside.any():
            work = work.copy()
            work.loc[outside, "s"] = OTHER
            warns.append(f"Only the {MAX_SPLIT} largest 'Split by' values are shown, the rest are in Other.")
        work = hide_other_if_small(work, "s", k, warns)
        if k > 0 and (work.groupby(["x", "s"]).size() < k).any():
            warns.append(f"Some combinations of X and Split by have fewer than {k} records.")
    if work.empty:
        return skipped(spec, ["No rows left for this chart."], warns)

    # 4. Calculate
    name = value_label(agg, spec["y"])
    if split:
        res = aggregate(work, ["x", "s"], agg).unstack("s")
    else:
        res = aggregate(work, ["x"], agg).to_frame(name=name)
    if agg in ADDITIVE:
        res = res.fillna(0)
    if xkind == "date" and s["fill_empty"] and agg in ADDITIVE and len(res):
        res = res.reindex(pd.period_range(res.index.min(), res.index.max(), freq=freq), fill_value=0)
    res.columns = [str(c) for c in res.columns]
    if res.empty:
        return skipped(spec, ["Nothing to show."], warns)
    if ctype in ("pie", "doughnut"):
        if (res < 0).any().any():
            return skipped(spec, ["Pie charts cannot show negative values."], warns)
        if res.sum().sum() == 0:
            return skipped(spec, ["All values are zero."], warns)
    if xkind == "date" and len(res) > MAX_TREND_POINTS:
        warns.append(f"This chart has {len(res)} points. A coarser time grain would be easier to read.")

    # 5. Sort
    sort = spec["sort"]
    if sort == "auto":
        sort = "natural order" if xkind == "date" else "value (high to low)"
    if sort == "value (high to low)":
        res = res.loc[res.sum(axis=1).sort_values(ascending=False).index]
    elif sort == "label (A to Z)":
        res = res.loc[sorted(res.index, key=str)]
    elif xkind == "category":
        order = [v for v in pd.unique(work["x"]) if v in res.index]
        try:
            order = sorted(order, key=float)
        except ValueError:
            pass
        res = res.loc[order]
    if OTHER in res.index:
        res = res.loc[[i for i in res.index if i != OTHER] + [OTHER]]

    # 6. Number format and labels
    fmt = spec["number_format"]
    if fmt == "percent of total":
        total = res.to_numpy().sum()
        if agg in ADDITIVE and total > 0:
            res = res / total
        else:
            warns.append("Percent of total only works with sum or count, so full numbers are shown.")
            fmt = "full"
    labels = [period_label(i, grain) for i in res.index] if xkind == "date" else [str(i) for i in res.index]
    res.index = labels
    res.index.name = x

    # 7. Title and chart type
    title = spec["title"]
    if not title:
        title = auto_title(spec, grain)
        if trimmed:
            title += f" (top {top_n}{' plus Other' if spec['other_bucket'] else ''})"
    resolved = ctype
    if ctype == AUTO:
        longest = max(len(str(i)) for i in labels)
        resolved = "bar" if len(labels) > 6 or longest > 12 else "column"
    money = agg in MONEY_AGGS and spec["y"] and fmt != "percent of total"
    currency = s["currency"] if money else ""
    slide = {
        "kind": "table" if ctype == "table" else "chart", "id": spec["id"], "title": title,
        "chart_type": resolved, "categories": labels, "frame": res,
        "series": [(c, list(res[c])) for c in res.columns],
        "number_format": excel_format(fmt, res.to_numpy()), "fmt_name": fmt,
        "show_labels": spec["show_labels"] and len(labels) <= 24, "gridlines": spec["gridlines"],
        "notes": spec["notes"], "footnotes": foot, "warnings": warns,
        "section": spec.get("section", ""), "headline": spec.get("headline", ""),
    }
    if ctype == "table":
        slide["header"] = [x] + list(res.columns)
        slide["rows"] = [[lab] + [format_value(v, "full" if fmt == "percent of total" else fmt, currency)
                                  if fmt != "percent of total" else f"{v:.1%}" for v in row]
                         for lab, row in zip(labels, res.to_numpy())]
    return slide


def render_kpi(spec, ctx):
    errors, warns = [], []
    yseries = get_measure(spec, ctx, errors, warns)
    if errors:
        return None, errors, warns
    agg, s = spec["agg"], ctx.settings
    sub = apply_filters(ctx.cur, spec["filters"], warns)
    value = scalar(sub, yseries, agg)
    fmt = "full" if spec["number_format"] == "percent of total" else spec["number_format"]
    money = agg in MONEY_AGGS and spec["y"]
    card = {"label": spec["title"] or value_label(agg, spec["y"]),
            "value": format_value(value, fmt, s["currency"] if money else ""),
            "small": "", "delta": None}
    if ctx.prev is not None:
        prev = scalar(apply_filters(ctx.prev, spec["filters"], []), yseries, agg)
        if prev and not np.isnan(prev) and not np.isnan(value):
            change = (value - prev) / abs(prev)
            card["delta"] = f"{change:+.1%}"
            card["small"] = f"{change:+.1%} vs previous period"
        else:
            card["small"] = "No previous period data"
    return card, errors, warns


def build_slides(df, info, settings, specs, today=None):
    """Turn the data, report settings and chart list into finished slide data."""
    today = today or date.today()
    s = {**DEFAULT_SETTINGS, **settings}
    conv = Converter(df)
    warns = []
    stats = [("Rows in the file", f"{len(df):,}")]
    blocked = set(s["blocked"])

    base = apply_filters(df, s["filters"], warns)
    stats.append(("Rows after row filters", f"{len(base):,}"))

    cur, prev, date_col = base, None, s["date_col"]
    period_text = None
    if date_col:
        if date_col not in df.columns:
            warns.append(f"Date column '{date_col}' is not in this file. Dates are ignored.")
            date_col = None
        else:
            dser, status, msg = conv.get(date_col, "date")
            d = dser.loc[base.index]
            if status == "block" or d.dropna().empty:
                warns.append(f"Date column '{date_col}': {msg or 'no valid dates'}. Dates are ignored.")
                date_col = None
            else:
                if status == "warn":
                    warns.append(f"Date column '{date_col}': {msg}")
                valid = d.dropna()
                start, end = resolve_range(s["date_preset"], s["start"], s["end"],
                                           valid.min().date(), valid.max().date(),
                                           today, s["fiscal_start"])
                in_range = (d >= pd.Timestamp(start)) & (d < pd.Timestamp(end) + pd.Timedelta(days=1))
                keep = in_range | d.isna() if s["undated"] != "exclude" else in_range
                cur = base[keep]
                stats.append((f"Rows from {start:%d %b %Y} to {end:%d %b %Y}", f"{int(in_range.sum()):,}"))
                undated = int(d.isna().sum())
                if undated:
                    stats.append((f"Rows with no {date_col}",
                                  f"{undated:,} ({'kept in non-trend charts' if s['undated'] != 'exclude' else 'left out'})"))
                if s["compare_prev"]:
                    if s["date_preset"] == "All data":
                        warns.append("Compare to previous period needs a date range other than 'All data'.")
                    else:
                        length = (end - start).days + 1
                        prev_end = start - timedelta(days=1)
                        prev_start = prev_end - timedelta(days=length - 1)
                        prev = base[(d >= pd.Timestamp(prev_start)) &
                                    (d < pd.Timestamp(prev_end) + pd.Timedelta(days=1))]
                inside = d[in_range]
                if not inside.empty:
                    period_text = f"{inside.min():%d %b %Y} to {inside.max():%d %b %Y}"
    footer = (s["footer"] or "").replace("{date}", f"{today:%d %b %Y}")
    result = {"slides": [], "warnings": warns, "stats": stats, "footer": footer, "title": s["title"]}
    if cur.empty:
        warns.append("No rows are left after the filters and date range.")
        return result

    ctx = SimpleNamespace(df=df, info=info, settings=s, conv=conv, cur=cur, prev=prev, blocked=blocked)
    result["slides"].append({"kind": "title", "title": s["title"],
                             "subtitle": period_text or f"{len(cur):,} records", "warnings": []})

    summary, used_cols = None, set()
    for spec in specs:
        if not spec.get("include", True):
            continue
        spec = {**new_spec(), **spec}
        for c in (spec["x"], spec["y"], spec["split"]):
            if c:
                used_cols.add(c)
        if spec["chart_type"] == "kpi":
            card, errors, kwarns = render_kpi(spec, ctx)
            if summary is None:
                summary = {"kind": "kpi", "title": "Where we stand", "cards": [], "warnings": [],
                           "section": "At a glance", "headline": s["summary_headline"], "notes": s["summary_notes"]}
                result["slides"].append(summary)
            summary["warnings"] += [f"Card '{spec['title'] or auto_title(spec)}': {w}" for w in kwarns]
            if errors:
                summary["warnings"] += [f"Card '{spec['title'] or auto_title(spec)}' skipped: {e}" for e in errors]
            elif len(summary["cards"]) >= MAX_KPI_CARDS:
                summary["warnings"].append(f"Only {MAX_KPI_CARDS} summary cards fit on a slide.")
            else:
                summary["cards"].append(card)
        else:
            result["slides"].append(render_chart(spec, ctx))

    if s["quality_slide"]:
        rows = [list(r) for r in stats]
        for col in sorted(used_cols):
            if col in cur.columns:
                n = int(cur[col].isna().sum())
                rows.append([f"Blank values in {col}", f"{n:,} ({n / len(cur):.1%})"])
        result["slides"].append({"kind": "quality", "title": "Data quality", "header": ["Check", "Value"],
                                 "rows": rows[:TABLE_MAX_ROWS + 3], "warnings": []})
    return result


# ---------------------------------------------------------------------------
# Suggested charts and saved settings
# ---------------------------------------------------------------------------


def suggest_specs(df, info, settings):
    blocked = set(settings["blocked"])
    specs = [new_spec(chart_type="kpi", agg="count", title="Records")]
    first_small = True
    measures = [c for c, m in info.items()
                if m["kind"] == "number" and any(h in c.lower() for h in MEASURE_HINTS)
                and c not in blocked and not m["id_like"]]
    for m in measures[:3]:
        specs.append(new_spec(chart_type="kpi", agg="sum", y=m, title=f"Total {m}"))
    agg, y = ("sum", measures[0]) if measures else ("count", None)
    if settings["date_col"]:
        specs.append(new_spec(chart_type="column", x=settings["date_col"], agg=agg, y=y, section="Trend"))
    cats = [c for c, m in info.items()
            if c not in blocked and c not in measures and 2 <= m["nunique"] <= 15
            and (m["kind"] == "category" or (m["kind"] == "number" and m["nunique"] <= FEW_VALUES))
            and not looks_like_identifier(m)]
    for c in cats[:6]:
        small = info[c]["nunique"] <= PIE_MAX and first_small
        first_small = first_small and not small
        specs.append(new_spec(x=c, agg=agg, y=y, section="Breakdown",
                              chart_type="doughnut" if small else AUTO))
    return specs


def find_preset(columns):
    for name, preset in PRESETS.items():
        if all(c in columns for c in preset["detect"]):
            return name, preset
    return None, None


def preset_specs(preset, columns):
    specs = []
    for chart in preset["charts"]:
        needed = [chart.get(k) for k in ("x", "y") if chart.get(k)]
        if all(c in columns for c in needed):
            specs.append(new_spec(**chart))
    return specs


def starting_config(df, info, filename):
    settings = copy.deepcopy(DEFAULT_SETTINGS)
    settings["blocked"] = [c for c in df.columns if info[c]["id_like"] or info[c]["sensitive"]]
    dates = [c for c in df.columns if info[c]["kind"] == "date"]
    settings["date_col"] = dates[0] if dates else None
    name, preset = find_preset(df.columns)
    if preset:
        extra = preset["settings"]
        settings.update({k: v for k, v in extra.items() if k != "blocked"})
        if settings["date_col"] not in df.columns:
            settings["date_col"] = dates[0] if dates else None
        settings["blocked"] = sorted(set(settings["blocked"]) | {c for c in extra.get("blocked", []) if c in df.columns})
        return settings, preset_specs(preset, df.columns), name
    return settings, suggest_specs(df, info, settings), None


def config_to_json(settings, specs):
    return json.dumps({"version": 1, "settings": settings, "charts": specs}, indent=2, default=str)


def config_from_json(text):
    data = json.loads(text)
    settings = {**copy.deepcopy(DEFAULT_SETTINGS), **data.get("settings", {})}
    settings["filters"] = [new_filter(**{k: v for k, v in f.items() if k != "id"}) for f in settings["filters"]]
    known = new_spec().keys()
    specs = []
    for chart in data.get("charts", []):
        spec = new_spec(**{k: v for k, v in chart.items() if k in known and k != "id"})
        spec["filters"] = [new_filter(**{k: v for k, v in f.items() if k != "id"}) for f in spec["filters"]]
        specs.append(spec)
    return settings, specs


# ---------------------------------------------------------------------------
# PowerPoint writer
# ---------------------------------------------------------------------------


def add_text(slide, text, left, top, width, height, size=18, bold=False, color=GREY, align=None):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.word_wrap = True
    para = box.text_frame.paragraphs[0]
    if align:
        para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def card_shape(slide, left, top, width, height, accent=True):
    """White card with a thin border and a green bar on the left, like the dashboard."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT
    shape.line.color.rgb = BORDER
    shape.shadow.inherit = False
    if accent:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(0.06), Inches(height))
        bar.fill.solid()
        bar.fill.fore_color.rgb = NAVY
        bar.line.fill.background()
        bar.shadow.inherit = False
    return shape


def blank_slide(prs, title, footer="", eyebrow="", headline="", number=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = PAGE_BG
    if eyebrow:
        add_text(slide, eyebrow.upper(), 0.6, 0.35, 12, 0.4, size=12, bold=True, color=NAVY)
    add_text(slide, title, 0.6, 0.65, 12, 0.8, size=30, bold=True, color=INK)
    if headline:
        add_text(slide, headline, 0.6, 1.3, 12.1, 0.5, size=15, color=GREY)
    if footer:
        add_text(slide, footer, 0.6, 7.05, 11, 0.35, size=10)
    if number:
        add_text(slide, str(number), 12.2, 7.05, 0.6, 0.35, size=11)
    return slide


def commentary_box(slide, text, top, height=1.2):
    """Space for the written explanation under a chart or the cards."""
    card_shape(slide, 0.6, top, 12.1, height)
    add_text(slide, "WHAT THIS SHOWS", 0.8, top + 0.08, 4, 0.3, size=10, bold=True, color=NAVY)
    add_text(slide, text, 0.8, top + 0.38, 11.7, height - 0.45, size=14, color=INK)


def finish(slide, data, notes_top=None):
    if data.get("notes") and notes_top:
        commentary_box(slide, data["notes"], notes_top)
    if data.get("footnotes"):
        add_text(slide, " ".join(data["footnotes"]), 0.6, 6.7, 12, 0.35, size=10)


def eyebrow_for(data, number):
    return f"{number:02d} \u00b7 {data['section']}" if data.get("section") else f"{number:02d}"


def add_title_slide(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY
    add_text(slide, data["title"], 0.8, 2.6, 11.5, 1.2, size=44, bold=True, color=WHITE)
    add_text(slide, data["subtitle"], 0.8, 3.9, 11.5, 0.8, size=20, color=RGBColor(0xD8, 0xF0, 0xE8))


def add_kpi_slide(prs, data, footer, number):
    slide = blank_slide(prs, data["title"], footer, eyebrow_for(data, number), data.get("headline", ""), number)
    cards = data["cards"]
    two_rows = len(cards) > 4
    tops = [1.95, 3.75] if two_rows else [2.0]
    for i, card in enumerate(cards):
        left = 0.6 + (i % 4) * 3.05
        top = tops[i // 4]
        card_shape(slide, left, top, 2.85, 1.6)
        add_text(slide, card["label"], left + 0.2, top + 0.1, 2.55, 0.4, size=12, bold=True, color=GREY)
        add_text(slide, card["value"], left + 0.2, top + 0.45, 2.55, 0.7, size=28, bold=True, color=INK)
        if card["small"]:
            down = (card.get("delta") or "").startswith("-")
            arrow = ("\u25bc " if down else "\u25b2 ") if card.get("delta") else ""
            add_text(slide, arrow + card["small"], left + 0.2, top + 1.15, 2.55, 0.4, size=11, bold=True,
                     color=RED if down else NAVY)
    if data.get("notes"):
        commentary_box(slide, data["notes"], 5.5 if two_rows else 4.0, 1.4)


def add_table_slide(prs, data, footer, number):
    slide = blank_slide(prs, data["title"], footer, eyebrow_for(data, number), data.get("headline", ""), number)
    rows, cols = len(data["rows"]) + 1, len(data["header"])
    table = slide.shapes.add_table(rows, cols, Inches(0.6), Inches(1.9), Inches(12.1), Inches(0.3 * rows)).table
    for r, row in enumerate([data["header"]] + data["rows"]):
        table.rows[r].height = Inches(0.3)
        for c, text in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(text)
            para = cell.text_frame.paragraphs[0]
            para.runs[0].font.size = Pt(12)
            cell.fill.solid()
            if r == 0:
                para.runs[0].font.bold = True
                para.runs[0].font.color.rgb = WHITE
                cell.fill.fore_color.rgb = NAVY
            else:
                para.runs[0].font.color.rgb = INK
                cell.fill.fore_color.rgb = LIGHT if r % 2 else PAGE_BG
    finish(slide, data, notes_top=None)


def add_chart_slide(prs, data, footer, number):
    slide = blank_slide(prs, data["title"], footer, eyebrow_for(data, number), data.get("headline", ""), number)
    ctype = data["chart_type"]
    cats = list(data["categories"])
    series = [(n, list(v)) for n, v in data["series"]]
    if ctype in ("bar", "stacked bar"):  # bar charts draw bottom-up, so flip to keep the order
        cats = cats[::-1]
        series = [(n, v[::-1]) for n, v in series]
    chart_data = CategoryChartData()
    chart_data.categories = cats
    for name, vals in series:
        chart_data.add_series(name, [None if pd.isna(v) else float(v) for v in vals])
    has_notes = bool(data["notes"])
    height = 3.3 if has_notes else 4.7
    card_shape(slide, 0.6, 1.9, 12.1, height + 0.2, accent=False)
    chart = slide.shapes.add_chart(PPTX_TYPES[ctype], Inches(0.75), Inches(2.0),
                                   Inches(11.8), Inches(height), chart_data).chart
    chart.has_title = False
    chart.has_legend = len(series) > 1 or ctype in ("pie", "doughnut")
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(12)
    plot = chart.plots[0]
    if ctype in ("pie", "doughnut"):
        for i in range(len(cats)):
            pt = plot.series[0].points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = SERIES_COLORS[i % len(SERIES_COLORS)]
    else:
        for i, ser in enumerate(plot.series):
            color = SERIES_COLORS[i % len(SERIES_COLORS)]
            if ctype in ("line",):
                ser.format.line.color.rgb = color
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = color
    if data["show_labels"]:
        plot.has_data_labels = True
        plot.data_labels.number_format = data["number_format"]
        plot.data_labels.number_format_is_linked = False
        plot.data_labels.font.size = Pt(12)
    if ctype not in ("pie", "doughnut"):
        for axis in (chart.category_axis, chart.value_axis):
            axis.tick_labels.font.size = Pt(12)
        chart.value_axis.tick_labels.number_format = data["number_format"]
        chart.value_axis.tick_labels.number_format_is_linked = False
        chart.value_axis.has_major_gridlines = bool(data["gridlines"])
        all_values = [v for _, vals in series for v in vals if v is not None and not pd.isna(v)]
        if ctype in KEEP_ZERO_TYPES and all_values and min(all_values) >= 0:
            chart.value_axis.minimum_scale = 0  # keep bars honest
    finish(slide, data, notes_top=5.5)


def write_pptx(result):
    """Turn the slide data from build_slides into PowerPoint bytes."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9
    prs.slide_height = Inches(7.5)
    footer = result["footer"]
    number = 0
    for data in result["slides"]:
        kind = data["kind"]
        if kind == "title":
            add_title_slide(prs, data)
            continue
        if kind == "skipped":
            continue
        number += 1
        if kind == "kpi":
            add_kpi_slide(prs, data, footer, number)
        elif kind in ("table", "quality"):
            add_table_slide(prs, data, footer, number)
        elif kind == "chart":
            add_chart_slide(prs, data, footer, number)
    buffer = io.BytesIO()
    prs.save(buffer)
    return buffer.getvalue()


# ===========================================================================
# STREAMLIT UI  (everything above works without Streamlit)
# ===========================================================================

THEME_CSS = """
<style>
.stApp { background: #F3F6F4; }
h1, h2, h3, h4 { color: #1A2B25; }
.eyebrow { color: #0E7C61; font-weight: 700; font-size: 0.8rem; letter-spacing: 0.06em; }
.note { border-left: 4px solid #0E7C61; background: #FFFFFF; padding: 0.6rem 0.9rem; border-radius: 4px; margin-top: 0.5rem; }
div[data-testid="stMetric"] { background: #FFFFFF; border: 1px solid #DDE5E1; border-left: 4px solid #0E7C61;
  padding: 0.6rem 0.9rem; border-radius: 6px; }
div[data-testid="stVerticalBlockBorderWrapper"] { background: #FFFFFF; }
.stButton > button, .stDownloadButton > button { background: #0E7C61; color: #FFFFFF; border: 0; }
.stButton > button:hover, .stDownloadButton > button:hover { background: #0A5F4A; color: #FFFFFF; }
section[data-testid="stSidebar"] { background: #E6EEEA; }
</style>
"""

NO_X = "(none: summary card)"
NO_COL = "(none)"


@st.cache_data(show_spinner=False)
def cached_sheet_names(data):
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner="Reading file...")
def cached_load(data, name, sheet):
    return load_file(io.BytesIO(data), name, sheet)


@st.cache_data(show_spinner=False)
def cached_profile(_df, sig):
    return profile_columns(_df)


@st.cache_data(show_spinner=False)
def cached_date_check(_df, sig, col):
    """Status, message and first/last dates of a date column."""
    converted, status, msg = convert_column(_df[col], "date")
    valid = converted.dropna()
    lo = valid.min().date() if len(valid) else None
    hi = valid.max().date() if len(valid) else None
    return status, msg, lo, hi


@st.cache_data(show_spinner="Building slides...")
def cached_build(_df, _info, sig, settings_json, specs_json, today_iso):
    return build_slides(_df, _info, json.loads(settings_json), json.loads(specs_json), date.fromisoformat(today_iso))


def pick_index(options, value, default=0):
    return options.index(value) if value in options else default


def filter_editor(filters, df, prefix):
    """Edit a list of row filters in place."""
    cols = list(df.columns)
    for f in list(filters):
        kp = f"{prefix}|{f['id']}"
        c1, c2, c3, c4 = st.columns([3, 2, 5, 1])
        f["col"] = c1.selectbox("Column", cols, index=pick_index(cols, f["col"]),
                                key=f"{kp}|col", label_visibility="collapsed")
        f["mode"] = c2.selectbox("Mode", FILTER_MODES, index=pick_index(FILTER_MODES, f["mode"]),
                                 key=f"{kp}|mode", label_visibility="collapsed")
        if f["mode"] == "not blank":
            c3.caption("Rows where this column is empty are left out.")
        else:
            counts = category_key(df[f["col"]]).value_counts().head(200)
            options = list(counts.index)
            f["values"] = c3.multiselect("Values", options, default=[v for v in f["values"] if v in options],
                                         key=f"{kp}|vals|{f['col']}", label_visibility="collapsed")
        if c4.button("✕", key=f"{kp}|del"):
            filters.remove(f)
            st.rerun()
    if st.button("Add row filter", key=f"{prefix}|add"):
        filters.append(new_filter(col=cols[0]))
        st.rerun()
    return filters


def chart_card(idx, spec, df, info, settings, total):
    """Draw the controls for one chart. Returns 'up', 'down', 'copy', 'delete' or None."""
    cols = list(df.columns)
    blocked = set(settings["blocked"])
    kp = f"c|{spec['id']}|"

    def tag(c):
        return f"{c}  [{info[c]['kind']}{', blocked' if c in blocked else ''}]"

    with st.container(border=True):
        head, toggle = st.columns([8, 2])
        head_box = head.empty()
        spec["include"] = toggle.checkbox("Include", value=spec["include"], key=kp + "include")

        with st.expander("Chart settings"):
            x_options = [NO_X] + cols
            x = st.selectbox("X axis (what to group by)", x_options,
                             index=pick_index(x_options, spec["x"]),
                             format_func=lambda c: c if c == NO_X else tag(c), key=kp + "x")
            spec["x"] = None if x == NO_X else x
            xkind = "date" if spec["x"] and info[spec["x"]]["kind"] == "date" else "category"

            c1, c2 = st.columns(2)
            spec["agg"] = c1.selectbox("Calculation", AGGS, index=pick_index(AGGS, spec["agg"]), key=kp + "agg")
            if spec["agg"] != "count":
                spec["y"] = c2.selectbox("Y column (what to calculate)", cols, index=pick_index(cols, spec["y"]),
                                         format_func=tag, key=kp + "y")
            if spec["x"] is None:
                spec["chart_type"] = "kpi"
            else:
                types = chart_types_for(xkind, info[spec["x"]]["nunique"])
                spec["chart_type"] = st.selectbox("Chart type", types, index=pick_index(types, spec["chart_type"]),
                                                  key=kp + "type")
                problem = column_problem(spec["x"], info, blocked)
                if problem:
                    st.warning(problem)
            if spec["agg"] != "count" and spec["y"] in blocked and spec["agg"] != "distinct count":
                st.warning(f"'{spec['y']}' is blocked. Only distinct count can use it.")

            spec["title"] = st.text_input("Title (leave empty for an automatic title)", spec["title"], key=kp + "title")

            if spec["x"] is not None:
                c1, c2, c3 = st.columns(3)
                if xkind == "date":
                    grains = ["report default"] + GRAINS
                    spec["time_grain"] = c1.selectbox("Time grain", grains,
                                                      index=pick_index(grains, spec["time_grain"]), key=kp + "grain")
                else:
                    spec["top_n"] = c1.number_input("Top N", 1, 50, int(spec["top_n"]), key=kp + "topn")
                    spec["other_bucket"] = c2.checkbox("Add an 'Other' bucket", spec["other_bucket"], key=kp + "other")
                    spec["show_blanks"] = c3.checkbox("Show blanks", spec["show_blanks"], key=kp + "blanks")
                spec["sort"] = st.selectbox("Sort", SORTS, index=pick_index(SORTS, spec["sort"]), key=kp + "sort")
                if spec["chart_type"] in SPLIT_TYPES:
                    split_options = [NO_COL] + [c for c in cols if c != spec["x"]]
                    split = st.selectbox("Split by (second category column)", split_options,
                                         index=pick_index(split_options, spec["split"]),
                                         format_func=lambda c: c if c == NO_COL else tag(c), key=kp + "split")
                    spec["split"] = None if split == NO_COL else split
                    if spec["split"]:
                        p = column_problem(spec["split"], info, blocked)
                        if p:
                            st.warning(p)
                if xkind == "category":
                    spec["rename"] = st.text_input("Rename labels (example: Y=Yes, N=No)", spec["rename"],
                                                   key=kp + "rename")

            c1, c2, c3 = st.columns(3)
            spec["number_format"] = c1.selectbox("Number format", NUMBER_FORMATS,
                                                 index=pick_index(NUMBER_FORMATS, spec["number_format"]),
                                                 key=kp + "fmt")
            if spec["x"] is not None:
                spec["show_labels"] = c2.checkbox("Data labels", spec["show_labels"], key=kp + "labels")
                spec["gridlines"] = c3.checkbox("Gridlines", spec["gridlines"], key=kp + "grid")
                spec["section"] = c1.text_input("Section label (e.g. Linkage campaign)", spec["section"], key=kp + "section")
                spec["headline"] = st.text_input("Headline under the title (optional)", spec["headline"], key=kp + "headline")
                spec["notes"] = st.text_area("Explanation shown under the chart (optional)", spec["notes"], key=kp + "notes", height=90)

            st.markdown("**Filter for this chart only**")
            filter_editor(spec["filters"], df, kp + "f")

            b1, b2, b3, b4 = st.columns(4)
            action = None
            if b1.button("Move up", key=kp + "up", disabled=idx == 0):
                action = "up"
            if b2.button("Move down", key=kp + "down", disabled=idx == total - 1):
                action = "down"
            if b3.button("Duplicate", key=kp + "copy"):
                action = "copy"
            if b4.button("Delete", key=kp + "del"):
                action = "delete"

        label = spec["title"] or auto_title(spec)
        head_box.markdown(f"**{idx + 1}. {label}**  \n"
                          f"{spec['chart_type'] if spec['x'] else 'summary card'}"
                          f"{'' if spec['include'] else '  (off)'}")
    return action


def vega_spec(slide):
    """A chart description for the preview, in the same order as the slide."""
    frame = slide["frame"]
    order = list(frame.index)
    long = frame.rename_axis("label").reset_index().melt(id_vars="label", var_name="series", value_name="value")
    ctype = slide["chart_type"]
    percent = slide["fmt_name"] == "percent of total"
    value_axis = {"field": "value", "type": "quantitative", "title": None,
                  "axis": {"format": ".0%" if percent else ",.2~f"}}
    label_axis = {"field": "label", "type": "nominal", "sort": order, "title": None}
    multi = frame.shape[1] > 1
    if ctype in ("pie", "doughnut"):
        spec = {"mark": {"type": "arc", "innerRadius": 70 if ctype == "doughnut" else 0},
                "encoding": {"theta": {"field": "value", "type": "quantitative"},
                             "color": {"field": "label", "type": "nominal", "sort": order}}}
    else:
        mark = {"line": {"type": "line", "point": True, "color": HEX_COLORS[0]},
                "area": {"type": "area", "color": HEX_COLORS[0]}}.get(ctype, {"type": "bar", "color": HEX_COLORS[0]})
        horizontal = ctype in ("bar", "stacked bar")
        enc = {"x": value_axis, "y": label_axis} if horizontal else {"x": label_axis, "y": value_axis}
        if multi:
            enc["color"] = {"field": "series", "type": "nominal", "legend": {"title": None}}
        spec = {"mark": mark, "encoding": enc}
    scale = {"range": HEX_COLORS}
    if "color" in spec["encoding"]:
        spec["encoding"]["color"]["scale"] = scale
    return long, spec


def show_slide_preview(slide):
    kind = slide["kind"]
    with st.container(border=True):
        if kind == "title":
            st.markdown(f"### {slide['title']}")
            st.caption(slide["subtitle"])
            return
        if slide.get("section"):
            st.markdown(f"<span class='eyebrow'>{slide['section'].upper()}</span>", unsafe_allow_html=True)
        st.markdown(f"#### {slide['title']}")
        if slide.get("headline"):
            st.markdown(f"**{slide['headline']}**")
        if kind == "skipped":
            for e in slide["errors"]:
                st.error(f"This chart is skipped: {e}")
        for w in slide["warnings"]:
            st.warning(w)
        if kind == "kpi":
            cols = st.columns(4)
            for i, card in enumerate(slide["cards"]):
                cols[i % 4].metric(card["label"], card["value"], card["delta"])
        elif kind in ("table", "quality"):
            st.dataframe(pd.DataFrame(slide["rows"], columns=slide["header"]), hide_index=True)
        elif kind == "chart":
            long, spec = vega_spec(slide)
            st.vega_lite_chart(long, spec, use_container_width=True)
            with st.expander("Data behind this chart"):
                st.dataframe(slide["frame"])
        for note in slide.get("footnotes", []):
            st.caption(note)
        if slide.get("notes"):
            st.markdown(f"<div class='note'><b>WHAT THIS SHOWS</b><br>{slide['notes']}</div>", unsafe_allow_html=True)


def main():
    st.set_page_config(page_title="Report Generator", layout="wide")
    st.markdown(THEME_CSS, unsafe_allow_html=True)
    st.title("Report Generator")
    st.write("Upload an Excel or CSV export, choose what goes in, check the preview, download a PowerPoint.")

    # Firebase is optional for report generation: a Firestore outage should not
    # prevent users from creating a PowerPoint.
    db = None
    save_report_fn = None
    firebase_error = None
    try:
        # Import lazily so missing Firebase packages/configuration do not stop
        # the report generator from running.
        from firebase_config import initialize_firebase
        from firestore_db import save_report as save_report_fn
        db = initialize_firebase()
    except Exception as error:
        firebase_error = error

    uploaded = st.file_uploader("Report file (CSV or Excel)", type=["csv", "xlsx"])
    if not uploaded:
        st.stop()

    data = uploaded.getvalue()
    sheet = None
    try:
        df = cached_load(data, uploaded.name, sheet)
    except Exception as error:
        st.error(f"Could not read that file: {error}")
        st.stop()
    if df.empty:
        st.error("That file has no rows.")
        st.stop()

    # Save each uploaded file once per Streamlit session, after it has been
    # successfully read. A changed file with the same name is saved again.
    if db is not None:
        report_key = f"{uploaded.name}|{hashlib.sha256(data).hexdigest()}"
        if st.session_state.get("saved_report_key") != report_key:
            try:
                save_report_fn(
                    db,
                    "Loan Report",
                    st.session_state.get("reporting_month", "Not specified"),
                    uploaded.name,
                )
                st.session_state["saved_report_key"] = report_key
                st.success("Report details saved to Firestore.")
            except Exception as error:
                st.warning(f"The report could not be saved to Firestore: {error}")
    elif firebase_error:
        st.warning(f"Firebase is unavailable; PowerPoint generation can continue. Details: {firebase_error}")

    sig = f"{uploaded.name}|{sheet}|{len(data)}"
    info = cached_profile(df, sig)

    # New file or sheet: start from the preset or from suggestions.
    if st.session_state.get("sig") != sig:
        settings, specs, preset_name = starting_config(df, info, uploaded.name)
        st.session_state.update(sig=sig, settings=settings, specs=specs, ver=0,
                                preset_name=preset_name, deck=None, cfg_n=0)
    settings = st.session_state["settings"]
    specs = st.session_state["specs"]
    ver = st.session_state["ver"]
    cols = list(df.columns)
    blocked = set(settings["blocked"])

    def tag(c):
        return f"{c}  [{info[c]['kind']}]"

    st.caption(f"{len(df):,} rows, {len(cols)} columns"
               + (f". Recognised as: {st.session_state['preset_name']}" if st.session_state["preset_name"] else ""))

    # Sidebar: save and load all settings
    with st.sidebar:
        st.header("Saved settings")
        st.download_button("Save settings (JSON)", config_to_json(settings, specs),
                           file_name="report_settings.json", mime="application/json")
        loaded = st.file_uploader("Load settings (JSON)", type=["json"], key=f"cfg_{st.session_state['cfg_n']}")
        if loaded:
            try:
                new_settings, new_specs = config_from_json(loaded.getvalue().decode("utf-8"))
            except Exception as error:
                st.error(f"Could not read that settings file: {error}")
            else:
                st.session_state.update(settings=new_settings, specs=new_specs, ver=ver + 1,
                                        cfg_n=st.session_state["cfg_n"] + 1)
                st.rerun()

    with st.sidebar:
        st.header("Not allowed on a slide")
        settings["blocked"] = st.multiselect(
            "Columns never allowed on a slide", cols, format_func=tag, key=f"{ver}|blocked",
            default=[c for c in settings["blocked"] if c in cols],
            help="Blocked columns cannot be used to group or split a chart. Names, IDs, phones and emails are ticked for you.")

    tab_rules, tab_report, tab_charts, tab_preview = st.tabs(
        ["1. Data rules", "2. Report settings", "3. Charts", "4. Preview and download"])

    # ---- Tab 1: what can enter the report ----
    with tab_rules:
        st.caption("These rules apply before anything else. Slides only ever show totals and counts.")
        st.markdown("**Date range**")
        date_options = [NO_COL] + cols
        chosen = st.selectbox("Date column", date_options, index=pick_index(date_options, settings["date_col"]),
                              format_func=lambda c: c if c == NO_COL else tag(c), key=f"{ver}|datecol")
        settings["date_col"] = None if chosen == NO_COL else chosen
        if settings["date_col"]:
            status, msg, lo, hi = cached_date_check(df, sig, settings["date_col"])
            if status == "block":
                st.error(f"{msg} Dates will be ignored.")
            else:
                if status == "warn":
                    st.warning(msg)
                settings["date_preset"] = st.selectbox(
                    "Show data for", DATE_PRESETS, index=pick_index(DATE_PRESETS, settings["date_preset"]),
                    key=f"{ver}|preset")
                if settings["date_preset"] == "Custom":
                    start = max(to_date(settings["start"]) or lo, lo)
                    end = min(to_date(settings["end"]) or hi, hi)
                    picked = st.date_input("From and to", (start, max(start, end)), min_value=lo, max_value=hi,
                                           key=f"{ver}|range")
                    if len(picked) == 2:
                        settings["start"], settings["end"] = picked[0].isoformat(), picked[1].isoformat()
                st.caption(f"Data runs from {lo:%d %b %Y} to {hi:%d %b %Y}.")
        st.divider()
        c1, c2 = st.columns(2)
        settings["min_group"] = c1.number_input(
            "Smallest group allowed on a slide (0 = off)", 0, 100, int(settings["min_group"]), key=f"{ver}|mingroup",
            help="Groups with fewer records than this are merged into Other or hidden, to protect individuals.")
        settings["suppress_mode"] = c2.selectbox(
            "Small groups are", SUPPRESS_MODES, index=pick_index(SUPPRESS_MODES, settings["suppress_mode"]),
            key=f"{ver}|suppress")
        st.markdown("**Row filters** (apply to every chart)")
        st.caption("Example: Status, include only, Approved and Disbursed. Or exclude test accounts.")
        filter_editor(settings["filters"], df, f"{ver}|gf")

    # ---- Tab 2: report-level settings ----
    with tab_report:
        c1, c2 = st.columns(2)
        settings["title"] = c1.text_input("Report title", settings["title"], key=f"{ver}|title")
        settings["currency"] = c2.text_input("Currency label (optional)", settings["currency"], key=f"{ver}|currency")
        settings["footer"] = st.text_input("Footer on every slide (optional, {date} becomes today)",
                                           settings["footer"], key=f"{ver}|footer")
        settings["quality_slide"] = st.checkbox("Add a data quality slide at the end",
                                                settings["quality_slide"], key=f"{ver}|quality")

        st.markdown("**Summary slide text**")
        settings["summary_headline"] = st.text_input("Headline under the summary title (optional)",
                                                     settings["summary_headline"], key=f"{ver}|sumhead")
        settings["summary_notes"] = st.text_area("Explanation under the summary cards (optional)",
                                                 settings["summary_notes"], key=f"{ver}|sumnotes", height=90)
        st.markdown("**Dates** (column and range are chosen in the sidebar)")
        if settings["date_col"]:
            status, msg, lo, hi = cached_date_check(df, sig, settings["date_col"])
            if status != "block":
                c1, c2 = st.columns(2)
                settings["grain"] = c1.selectbox("Time grain for trend charts", GRAINS,
                                                 index=pick_index(GRAINS, settings["grain"]), key=f"{ver}|grain")
                settings["undated"] = c2.selectbox("Rows with no date", UNDATED_MODES,
                                                   index=pick_index(UNDATED_MODES, settings["undated"]),
                                                   key=f"{ver}|undated")
                settings["fill_empty"] = st.checkbox("Show empty periods as zero", settings["fill_empty"],
                                                     key=f"{ver}|fill")
                settings["compare_prev"] = st.checkbox(
                    "Compare summary cards to the previous period (same length, just before)",
                    settings["compare_prev"], key=f"{ver}|compare")
                with st.expander("Week and fiscal year options"):
                    settings["week_start"] = st.selectbox("Week starts on", DAY_NAMES,
                                                          index=pick_index(DAY_NAMES, settings["week_start"]),
                                                          key=f"{ver}|weekstart")
                    month = st.selectbox("Fiscal year starts in", MONTH_NAMES,
                                         index=int(settings["fiscal_start"]) - 1, key=f"{ver}|fiscal")
                    settings["fiscal_start"] = MONTH_NAMES.index(month) + 1

    # ---- Tab 3: the chart list ----
    with tab_charts:
        c1, c2, c3 = st.columns([1, 1, 4])
        if c1.button("Suggest charts"):
            st.session_state["specs"] = suggest_specs(df, info, settings)
            st.rerun()
        if c2.button("Add chart"):
            specs.append(new_spec())
            st.rerun()
        c3.caption("Suggest charts replaces the list with automatic ones. Each chart can be switched off, moved, copied or deleted.")
        if not specs:
            st.info("No charts yet. Click Suggest charts or Add chart.")
        for idx, spec in enumerate(specs):
            action = chart_card(idx, spec, df, info, settings, len(specs))
            if action == "up":
                specs[idx - 1], specs[idx] = specs[idx], specs[idx - 1]
            elif action == "down":
                specs[idx + 1], specs[idx] = specs[idx], specs[idx + 1]
            elif action == "copy":
                clone = copy.deepcopy(spec)
                clone["id"] = new_id()
                for f in clone["filters"]:
                    f["id"] = new_id()
                specs.insert(idx + 1, clone)
            elif action == "delete":
                specs.pop(idx)
            if action:
                st.rerun()

    # ---- Tab 4: preview and download ----
    with tab_preview:
        key_settings = json.dumps(settings, sort_keys=True, default=str)
        key_specs = json.dumps(specs, sort_keys=True, default=str)
        result = cached_build(df, info, sig, key_settings, key_specs, date.today().isoformat())
        for w in result["warnings"]:
            st.warning(w)
        with st.expander("What went into this report"):
            st.dataframe(pd.DataFrame(result["stats"], columns=["Check", "Value"]), hide_index=True)
        n_slides = sum(1 for s in result["slides"] if s["kind"] != "skipped")
        st.metric("Slides in the deck", n_slides)
        for slide in result["slides"]:
            show_slide_preview(slide)

        if n_slides:
            deck_key = key_settings + key_specs + sig
            if st.button("Build PowerPoint"):
                st.session_state["deck"] = (deck_key, write_pptx(result))
            deck = st.session_state.get("deck")
            if deck and deck[0] == deck_key:
                st.download_button(
                    "Download PowerPoint", deck[1], file_name="report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            elif deck:
                st.caption("Settings changed since the last build. Click Build PowerPoint again.")


if __name__ == "__main__":
    main()
