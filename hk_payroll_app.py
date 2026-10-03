# -*- coding: utf-8 -*-
"""
HK Payroll & Attendance Helper — Streamlit single-file app v2.4.2
Monthly flow:
  1. Download CSV template -> fill sales/OT/shortfall -> upload (one row per employee)
  2. App auto-pulls attendance + master data -> calculates everything
  3. Output A: per-employee JSON block for Grok (payslip generation)
  4. Output B: summary CSV block (one row per employee-month) -> saved to running ledger
Run:  streamlit run hk_payroll_app_v2_4_2.py
Data persisted in ./payroll_data/ (next to this script) as CSV files.

v2.4.2 (channel sales consolidation + T-2 commission):
  - Channel sales (SEVENS / MEAT DEVIL / PATTY) moved OUT of per-employee input into a new
    "Channel Sales" tab: one consolidated figure per channel per month (channel_sales.csv).
  - Monthly commission now auto-pulls the T-2 month: payroll for month M uses channel
    sales of month M-2 (bonus shows up 2 months later automatically, no re-entry).
  - Batch CSV template drops the per-employee sales columns; old templates that still have
    them keep working (non-zero values are treated as a manual override for that row).
  - Single + Export pages show which sales month is being used; warning if not yet recorded.
  - Summary ledger rows gain a sales_month column for audit (old rows default to blank).

v2.4.1 (env compatibility):
  - FIX crash on old Streamlit (< 1.19, e.g. anaconda Python 3.7 stacks): st.rerun() did not
    exist yet. Shim falls back to st.experimental_rerun() (or a no-op on very old versions)
    so the app keeps running. RECOMMENDED long-term fix: Python >= 3.10 + latest Streamlit
    (conda create -n payroll python=3.11 pandas openpyxl; pip install streamlit).

v2.4 (new on v2.3):
  - Employee master: open-name monthly SUBSIDY and DEDUCTION (custom label + fixed HKD
    amount) applied automatically every month.
  - Monthly calculation (batch CSV, single, export): extra open-name subsidy / deduction
    inputs for one-off adjustments in that month.
  - Each shows as its own payslip line with its custom label; totals updated;
    MPF is still calculated on basic salary only.

v2.3 (pay logic change):
  - Month shortfall is now ALWAYS carried into outstanding hours: the team gets the
    FULL month basic; the shortfall is recovered later (or via the manual
    "shortfall hours to deduct" input at hour-value rate) - never auto-pro-rated.
  - shortfall_rule values (pay_full_recover_next_month / pay_actual / none) now all
    behave the same: full basic + carry. "none" = full pay. Field kept for record.
  - Batch + single calculation pages: help text updated to say so.

v2.2 (fixes on v2.1):
  - FIX crash in batch results table (list + tuple)
  - FIX outstanding hours double-counting when the page re-ran after Save
    (balance now carried by the monthly ledger; re-saving a month is idempotent)
  - FIX employee-form edits being overwritten on submit; end-date box now works inside the form
  - FIX closing of previous version (blank entry_end was read back as "nan")
  - FIX duplicate attendance rows (same employee+date now replaces, never doubles hours)
  - Delete: per-row attendance delete, ledger-month delete, confirmations on destructive actions
  - CSV import: tolerant of Excel date formats / encodings, row-level error messages
  - 2025/2026 default holidays corrected to the gazette
  - Atomic file writes + one-click backup zip (sidebar)
  - Continuous-contract: 4 CONSECUTIVE weeks >= 17h
"""

import io
import json
import math
import os
import re
import calendar
import zipfile
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# --- old-Streamlit compatibility (v2.4.1) -------------------------------------
# st.rerun() only exists on Streamlit >= 1.19. Older stacks (e.g. anaconda + Python 3.7)
# have st.experimental_rerun() instead; on very old versions neither exists, so we fall
# back to a no-op (page still loads; flash messages appear on the next interaction).
if not hasattr(st, "rerun"):
    st.rerun = getattr(st, "experimental_rerun", None) or (lambda: None)
# -------------------------------------------------------------------------------

_SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = str(_SCRIPT_DIR / "payroll_data")
os.makedirs(DATA_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# HK General Holidays (defaults — verify against gazette each year; editable in app)
# NB: once holidays.json exists it overrides these defaults (use "Load defaults" to add).
# ---------------------------------------------------------------------------
DEFAULT_HOLIDAYS = {
    2025: ["2025-01-01", "2025-01-29", "2025-01-30", "2025-01-31", "2025-04-04",
           "2025-04-18", "2025-04-19", "2025-04-21", "2025-05-01", "2025-05-05",
           "2025-05-31", "2025-07-01", "2025-10-01", "2025-10-07", "2025-10-29",
           "2025-12-25", "2025-12-26"],
    2026: ["2026-01-01", "2026-02-17", "2026-02-18", "2026-02-19", "2026-04-03",
           "2026-04-04", "2026-04-06", "2026-04-07", "2026-05-01", "2026-05-25",
           "2026-06-19", "2026-07-01", "2026-09-26", "2026-10-01", "2026-10-19",
           "2026-12-25", "2026-12-26"],
}


def load_holidays():
    path = os.path.join(DATA_DIR, "holidays.json")
    try:
        with open(path) as f:
            return set(json.load(f))
    except (OSError, ValueError):
        return {d for ds in DEFAULT_HOLIDAYS.values() for d in ds}


def save_holidays(h):
    path = os.path.join(DATA_DIR, "holidays.json")
    with open(path + ".tmp", "w") as f:
        json.dump(sorted(h), f)
    os.replace(path + ".tmp", path)


HOLIDAYS = load_holidays()

# ---------------------------------------------------------------------------
# Column definitions
# ---------------------------------------------------------------------------
NUMERIC_EMP_COLS = ["basic_monthly", "hourly_rate", "hours_per_day", "travel_subsidy",
                    "subsidy_amount", "deduction_amount",
                    "ot_rate", "comm_sevens_pct", "comm_meatdevil_pct", "comm_patty_pct",
                    "mpf_employee_pct", "mpf_employer_pct", "outstanding_hours"]
BOOL_EMP_COLS = ["probation", "mpf_started"]
EMP_COLS = [
    "name", "join_date", "entry_start", "entry_end", "change_reason",
    "package_type", "basic_monthly", "hourly_rate",
    "hours_per_day", "work_pattern", "travel_subsidy",
    "subsidy_name", "subsidy_amount", "deduction_name", "deduction_amount",
    "ot_rate",
    "comm_sevens_pct", "comm_meatdevil_pct", "comm_patty_pct",
    "probation", "shortfall_rule",
    "mpf_employee_pct", "mpf_employer_pct", "mpf_started", "mpf_start_date", "mpf_enrollment_note",
    "outstanding_hours", "notes",
]
ATT_COLS = ["employee", "date", "start", "end", "break_hours",
            "raw_hours", "rounded_hours", "notes"]
HISTORY_COLS = ["timestamp", "employee", "action", "field", "old_value", "new_value", "note"]
INPUT_COLS = ["name", "year", "month",
              "ot_hours", "shortfall_hours", "other_adjustment",
              "extra_subsidy_name", "extra_subsidy_amount",
              "extra_deduction_name", "extra_deduction_amount"]
SALES_COLS = ["year", "month", "sevens_sales", "meatdevil_sales", "patty_sales"]
SUMMARY_COLS = ["year", "month", "employee", "package_type",
                "expected_hours", "actual_hours", "shortfall",
                "outstanding_before", "shortfall_deducted_hours", "outstanding_after",
                "basic_pay", "hourly_pay", "ot_pay",
                "comm_sevens", "comm_meatdevil", "comm_patty", "comm_total", "sales_month",
                "travel", "shortfall_deduction", "other_adjustment",
                "subsidy_name", "subsidy_amount", "extra_subsidy_name", "extra_subsidy_amount",
                "deduction_name", "deduction_amount", "extra_deduction_name", "extra_deduction_amount",
                "total_before_mpf",
                "mpf_employee", "mpf_employer", "total_after_mpf", "saved_at"]
MONEY_COLS = ["basic_pay", "hourly_pay", "ot_pay", "comm_sevens", "comm_meatdevil",
              "comm_patty", "comm_total", "travel", "shortfall_deduction",
              "other_adjustment", "subsidy_amount", "extra_subsidy_amount",
              "deduction_amount", "extra_deduction_amount",
              "total_before_mpf", "mpf_employee", "mpf_employer",
              "total_after_mpf"]
HOUR_COLS = ["expected_hours", "actual_hours", "shortfall", "outstanding_before",
             "shortfall_deducted_hours", "outstanding_after"]
NUM_COLS = (set(NUMERIC_EMP_COLS) | set(MONEY_COLS) | set(HOUR_COLS)
            | {"year", "month", "break_hours", "raw_hours", "rounded_hours"})

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def safe_float(val, default=0.0):
    try:
        if val is None or pd.isna(val):
            return default
        return float(val)
    except (TypeError, ValueError):
        return default


def truthy(v):
    return str(v).strip().lower() in ("true", "1", "1.0", "yes", "y", "t")


def hkd(x):
    return f"{safe_float(x):,.2f}"


def round15(h):
    if h is None or pd.isna(h):
        return 0.0
    return math.floor(h * 4 + 0.5) / 4


def norm_date(v):
    """'' for blank; ISO string for valid; ValueError if unparseable.
    ISO input is trusted; anything else (e.g. 1/8/2026 from Excel) is read day-first."""
    s = str(v).strip()
    if s in ("", "nan", "None", "NaT"):
        return ""
    try:
        if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s):
            return pd.to_datetime(s[:10]).date().isoformat()
        return pd.to_datetime(s, dayfirst=True).date().isoformat()
    except Exception:
        raise ValueError(f"bad date '{s}'")


def to_date(s, fallback=None):
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except ValueError:
        return fallback or date.today()


def read_csv_upload(f):
    """Read an uploaded CSV as all-strings; tries UTF-8 (Excel BOM ok), then Big5/GBK."""
    raw = f.getvalue()
    for enc in ("utf-8-sig", "cp950", "gbk"):
        try:
            df = pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, encoding=enc)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not decode the CSV — save it as UTF-8 CSV and retry.")


def show_df(df, **kw):
    try:
        st.dataframe(df, width="stretch", hide_index=True, **kw)
    except Exception:  # older Streamlit without width="stretch"
        st.dataframe(df, use_container_width=True, hide_index=True, **kw)


def flash(msg, kind="success"):
    st.session_state["_flash"] = (kind, msg)


# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------
def load_df(name, cols):
    """Load a CSV with predictable types: text -> str ('' when blank), numbers -> float,
    flags -> bool. Missing columns are added with defaults."""
    path = os.path.join(DATA_DIR, f"{name}.csv")
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except (FileNotFoundError, pd.errors.EmptyDataError):
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            if c in ("mpf_employee_pct", "mpf_employer_pct"):
                df[c] = "5.0"
            elif c in NUM_COLS:
                df[c] = "0"
            else:
                df[c] = ""
    df = df[cols].copy()
    for c in cols:
        if c in NUM_COLS:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        elif c in BOOL_EMP_COLS:
            df[c] = df[c].map(truthy)
    for c in ("year", "month"):
        if c in cols:
            df[c] = df[c].astype(int)
    return df.reset_index(drop=True)


def save_df(df, name):
    path = os.path.join(DATA_DIR, f"{name}.csv")
    df.to_csv(path + ".tmp", index=False, encoding="utf-8")
    os.replace(path + ".tmp", path)  # atomic: never leaves a half-written file


def log_history(employee, action, field, old_value, new_value, note=""):
    hist = load_df("employee_history", HISTORY_COLS)
    row = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "employee": employee, "action": action, "field": field,
           "old_value": "" if old_value is None else str(old_value),
           "new_value": "" if new_value is None else str(new_value), "note": note}
    save_df(pd.concat([hist, pd.DataFrame([row])], ignore_index=True), "employee_history")


def backup_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(Path(DATA_DIR).glob("*")):
            if f.suffix in (".csv", ".json"):
                z.write(f, f.name)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Calculation helpers
# ---------------------------------------------------------------------------
def workdays_in_month(year, month, pattern):
    full, half = 0, 0
    for d in range(1, calendar.monthrange(year, month)[1] + 1):
        dt = date(year, month, d)
        if dt.isoformat() in HOLIDAYS:
            continue
        wd = dt.weekday()
        if wd <= 4:
            full += 1
        elif wd == 5 and pattern == "5.5-day":
            half += 1
    return full, half


def expected_hours(year, month, hours_per_day, pattern):
    full, half = workdays_in_month(year, month, pattern)
    return round(full * hours_per_day + half * hours_per_day * 0.5, 2)


def parse_time(s):
    s = str(s).strip()
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def compute_raw_hours(start, end, brk):
    s, e = parse_time(start), parse_time(end)
    if not s or not e:
        return None
    sh = s.hour + s.minute / 60
    eh = e.hour + e.minute / 60
    if eh < sh:
        eh += 24
    return round(eh - sh - (brk or 0), 2)


def month_attendance(emp, year, month):
    att = load_df("attendance", ATT_COLS)
    if att.empty:
        return att
    d = pd.to_datetime(att["date"], errors="coerce")
    mask = (att["employee"].str.strip() == str(emp).strip()) & (d.dt.year == year) & (d.dt.month == month)
    return att[mask].copy()


def parse_wework_export(uploaded_file, known_names=None, default_break=0.0):
    """Parse 企业微信 上下班打卡_日报 export (.xlsx, sheet 1). Returns (df, skipped_notes)."""
    raw = pd.read_excel(io.BytesIO(uploaded_file.getvalue()), sheet_name=0, header=None)
    if raw.shape[1] < 56:
        raise ValueError(f"Unexpected layout: {raw.shape[1]} columns (need 56+). "
                         "Is this the 上下班打卡_日报 export?")
    rows, skipped = [], []
    EMPTY = ("--", "nan", "", "None")
    for _, r in raw.iloc[4:].iterrows():
        name = str(r[1]).strip()
        punches = str(r[9]).strip()
        if punches in EMPTY:
            continue
        if known_names is not None and name not in known_names:
            skipped.append(f"{name}: not in employee master")
            continue
        try:
            d = norm_date(str(r[0])[:10].replace("/", "-"))
        except ValueError:
            skipped.append(f"{name}: unreadable date '{r[0]}'")
            continue
        on1, off1 = str(r[49]).strip(), str(r[51]).strip()
        on2, off2 = str(r[53]).strip(), str(r[55]).strip()
        segments = []
        for on, off in ((on1, off1), (on2, off2)):
            if on not in EMPTY and off not in EMPTY:
                h = compute_raw_hours(on, off, 0)
                if h:
                    segments.append(h)
        start = on1 if on1 not in EMPTY else on2
        end = off2 if off2 not in EMPTY else off1
        if not segments:
            rows.append({"employee": name, "date": d, "start": start, "end": end,
                         "break_hours": 0.0, "raw_hours": 0.0, "rounded_hours": 0.0,
                         "notes": f"single punch ({punches}x) – verify manually"})
            continue
        total = sum(segments)
        raw_h = max(round(total - default_break, 2), 0.0) if default_break else round(total, 2)
        rows.append({"employee": name, "date": d, "start": start, "end": end,
                     "break_hours": default_break, "raw_hours": raw_h,
                     "rounded_hours": round15(raw_h),
                     "notes": "2 segments" if len(segments) > 1 else ""})
    return pd.DataFrame(rows, columns=ATT_COLS), skipped


def add_attendance(new):
    """Add rows; an existing row for the same employee+date is REPLACED (never double-counted).
    Returns (added, replaced)."""
    att = load_df("attendance", ATT_COLS)
    new = new.drop_duplicates(["employee", "date"], keep="last")   # last line wins within a batch
    replaced = 0
    if not att.empty and not new.empty:
        keys = set(zip(new["employee"], new["date"]))
        dup = pd.Series([(e, d) in keys for e, d in zip(att["employee"], att["date"])],
                        index=att.index, dtype=bool)
        replaced = int(dup.sum())
        att = att[~dup]
    save_df(pd.concat([att, new], ignore_index=True), "attendance")
    return len(new), replaced


def hour_value_rate(e, exp):
    """Hourly value: explicit hourly_rate, else basic_monthly / expected hours."""
    r = safe_float(e["hourly_rate"])
    if r <= 0 and safe_float(e["basic_monthly"]) > 0 and exp > 0:
        r = safe_float(e["basic_monthly"]) / exp
    return r


def get_employee_version(emp, name, year, month):
    """Employee row (dict) effective for year-month: entry_start <= month end and
    (entry_end blank or >= month start); latest entry_start wins. Falls back to latest version."""
    if emp is None or emp.empty:
        return None
    rows = emp[emp["name"].astype(str).str.strip() == str(name).strip()].copy()
    if rows.empty:
        return None
    first = date(year, month, 1).isoformat()
    last = date(year, month, calendar.monthrange(year, month)[1]).isoformat()
    rows["_es"] = rows["entry_start"].astype(str).str[:10].replace("", "1900-01-01")
    rows["_ee"] = rows["entry_end"].astype(str).str[:10].replace("", "9999-12-31")
    active = rows[(rows["_es"] <= last) & (rows["_ee"] >= first)]
    chosen = (active if not active.empty else rows).sort_values("_es").iloc[-1]
    return chosen.drop(labels=["_es", "_ee"], errors="ignore").to_dict()


def current_version_idx(emp, name):
    """Index of the open version (blank entry_end); else the latest version."""
    rows = emp[emp["name"].astype(str).str.strip() == str(name).strip()]
    if rows.empty:
        return None
    open_rows = rows[rows["entry_end"].astype(str).str.strip() == ""]
    return (open_rows if not open_rows.empty else rows).sort_values("entry_start").index[-1]


def ledger_before(summ, ver, name, y, m):
    """Outstanding hours at the START of a month.
    Same month already saved -> its stored 'before' (so re-calc / re-save never double counts).
    Else the previous saved month's 'after'. Else the master's outstanding_hours (opening balance)."""
    if not summ.empty:
        rows = summ[summ["employee"] == name]
        if not rows.empty:
            k = rows["year"] * 12 + rows["month"]
            same = rows[k == y * 12 + m]
            if not same.empty:
                return safe_float(same.iloc[-1]["outstanding_before"])
            prior = k[k < y * 12 + m]
            if not prior.empty:
                return safe_float(rows.loc[prior.idxmax(), "outstanding_after"])
    return safe_float(ver.get("outstanding_hours"))


def shift_month(y, m, delta):
    """Shift a year-month by delta months (negative = earlier)."""
    k = y * 12 + (m - 1) + delta
    return k // 12, k % 12 + 1


def get_channel_sales(y, m):
    """Consolidated channel sales recorded for a year-month.
    Returns (sevens, meatdevil, patty, found)."""
    df = load_df("channel_sales", SALES_COLS)
    row = df[(df["year"] == y) & (df["month"] == m)]
    if row.empty:
        return 0.0, 0.0, 0.0, False
    r = row.iloc[-1]
    return (safe_float(r["sevens_sales"]), safe_float(r["meatdevil_sales"]),
            safe_float(r["patty_sales"]), True)


def save_channel_sales(y, m, sevens, meat, patty):
    """Upsert one consolidated channel-sales row for a year-month."""
    df = load_df("channel_sales", SALES_COLS)
    row = pd.DataFrame([{"year": y, "month": m, "sevens_sales": sevens,
                         "meatdevil_sales": meat, "patty_sales": patty}], columns=SALES_COLS)
    if df.empty:
        out = row
    else:
        hit = (df["year"] == y) & (df["month"] == m)
        out = pd.concat([df[~hit], row], ignore_index=True)
    save_df(out, "channel_sales")


def compute_monthly(e, y, m, sales_sevens=None, sales_meat=None, sales_patty=None,
                    ot_hours=0.0, shortfall_hours=0.0, other_adjust=0.0, outstanding_before=None,
                    extra_subsidy_name="", extra_subsidy_amount=0.0,
                    extra_deduction_name="", extra_deduction_amount=0.0):
    """Single source of truth for monthly pay. Returns (summary_dict, grok_payload_dict).
    Commission channel sales (v2.4.2): T-2 rule - payroll month M auto-pulls consolidated
    channel sales of month M-2 from the Channel Sales tab. Pass explicit sales_* values
    to override per call (batch CSV sales columns still work as manual override).
    shortfall_hours handled like OT: explicit hours x rate, own payslip line.
    OT uses ot_rate when set (>0), otherwise hour_value_rate.
    Month shortfall is ALWAYS carried into outstanding hours (full month basic paid;
    the team prefers full payroll). It is deducted only via the manual shortfall_hours
    input, at hour value rate. All shortfall_rule values now behave the same
    (pay_full_recover_next_month / pay_actual / none = full pay + carry); the field is
    kept for record only.
    Open-name adjustments (v2.4): master subsidy_name/subsidy_amount and deduction_name/
    deduction_amount apply EVERY month; extra_subsidy_* / extra_deduction_* are one-off for
    the month being calculated. All four appear as their own payslip lines. MPF stays on basic.
    hybrid_prestart: pre-start hourly split is not auto-computed — use other_adjustment."""
    name = e["name"]
    exp = expected_hours(y, m, safe_float(e["hours_per_day"], 8), e["work_pattern"])
    a = month_attendance(name, y, m)
    actual = round(float(a["rounded_hours"].sum()), 2) if not a.empty else 0.0
    short = round(exp - actual, 2) if e["package_type"] in ("fixed_monthly", "hybrid_prestart") else 0.0
    short = max(short, 0.0)
    outstanding_before = (safe_float(e.get("outstanding_hours", 0)) if outstanding_before is None
                          else safe_float(outstanding_before))
    rate = hour_value_rate(e, exp)
    ot_rate = safe_float(e.get("ot_rate"))
    if ot_rate <= 0:
        ot_rate = rate

    # Channel sales for commission: T-2 rule (payroll month M uses sales of month M-2).
    sy, sm = shift_month(y, m, -2)
    auto_sales = sales_sevens is None or sales_meat is None or sales_patty is None
    if auto_sales:
        sales_sevens, sales_meat, sales_patty, sales_found = get_channel_sales(sy, sm)
        sales_source = (f"channel sales T-2 ({calendar.month_name[sm]} {sy})"
                        if sales_found
                        else f"NOT RECORDED ({calendar.month_name[sm]} {sy}) - commission 0")
    else:
        sales_found = True
        sales_source = "manual override (per-employee input)"

    basic_pay, hourly_pay = 0.0, 0.0
    if e["package_type"] == "hourly":
        hourly_pay = round(actual * safe_float(e["hourly_rate"]), 2)
    else:
        # FULL month basic, always - the month's shortfall is carried into outstanding
        # hours instead of being pro-rated off the basic (v2.3 rule; see docstring).
        basic_pay = safe_float(e["basic_monthly"])

    ot_pay = round(ot_hours * ot_rate, 2)
    travel = safe_float(e["travel_subsidy"])
    comm_sevens = round(sales_sevens * safe_float(e["comm_sevens_pct"]) / 100, 2)
    comm_meat = round(sales_meat * safe_float(e["comm_meatdevil_pct"]) / 100, 2)
    comm_patty = round(sales_patty * safe_float(e["comm_patty_pct"]) / 100, 2)
    comm_total = round(comm_sevens + comm_meat + comm_patty, 2)
    shortfall_deduction = round(shortfall_hours * rate, 2)
    subsidy_name = str(e.get("subsidy_name", "")).strip()
    subsidy_amt = safe_float(e.get("subsidy_amount"))
    deduction_name = str(e.get("deduction_name", "")).strip()
    deduction_amt = safe_float(e.get("deduction_amount"))
    ex_sub_name = str(extra_subsidy_name).strip()
    ex_sub_amt = safe_float(extra_subsidy_amount)
    ex_ded_name = str(extra_deduction_name).strip()
    ex_ded_amt = safe_float(extra_deduction_amount)
    subsidies_total = round(subsidy_amt + ex_sub_amt, 2)
    deductions_total = round(deduction_amt + ex_ded_amt, 2)

    total_before_mpf = round(basic_pay + hourly_pay + ot_pay + comm_total + travel
                             + subsidies_total + other_adjust
                             - shortfall_deduction - deductions_total, 2)

    mpf_emp_amt = mpf_er_amt = 0.0
    mpf_on = truthy(e.get("mpf_started"))
    if mpf_on:
        mpf_emp_amt = round(basic_pay * safe_float(e["mpf_employee_pct"]) / 100, 2)
        mpf_er_amt = round(basic_pay * safe_float(e["mpf_employer_pct"]) / 100, 2)
    total_after_mpf = round(total_before_mpf - mpf_emp_amt, 2)
    carry = short   # v2.3: month's shortfall always carried into outstanding hours
    outstanding_after = round(max(outstanding_before + carry - shortfall_hours, 0), 2)

    summary = {"year": y, "month": m, "employee": name, "package_type": e["package_type"],
               "expected_hours": exp, "actual_hours": actual, "shortfall": short,
               "outstanding_before": outstanding_before,
               "shortfall_deducted_hours": shortfall_hours,
               "outstanding_after": outstanding_after,
               "basic_pay": basic_pay, "hourly_pay": hourly_pay, "ot_pay": ot_pay,
               "comm_sevens": comm_sevens, "comm_meatdevil": comm_meat, "comm_patty": comm_patty,
               "comm_total": comm_total, "sales_month": f"{sy}-{sm:02d}", "travel": travel,
               "shortfall_deduction": shortfall_deduction, "other_adjustment": other_adjust,
               "subsidy_name": subsidy_name, "subsidy_amount": subsidy_amt,
               "extra_subsidy_name": ex_sub_name, "extra_subsidy_amount": ex_sub_amt,
               "deduction_name": deduction_name, "deduction_amount": deduction_amt,
               "extra_deduction_name": ex_ded_name, "extra_deduction_amount": ex_ded_amt,
               "total_before_mpf": total_before_mpf,
               "mpf_employee": mpf_emp_amt, "mpf_employer": mpf_er_amt,
               "total_after_mpf": total_after_mpf,
               "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M")}

    hybrid_note = ""
    if e["package_type"] == "hybrid_prestart":
        hybrid_note = ("hybrid_prestart: full basic is paid; the month's shortfall is carried "
                       "into outstanding hours; put any pre-start hourly portion in "
                       "other_adjustment if needed.")

    payload = {"employee": name, "period": f"{calendar.month_name[m]} {y}", "year": y, "month": m,
               "package_type": e["package_type"],
               "basic_monthly": safe_float(e["basic_monthly"]), "hourly_rate": safe_float(e["hourly_rate"]),
               "ot_rate": round(ot_rate, 2),
               "travel_subsidy": travel, "hour_value_rate": round(rate, 2),
               "work_pattern": e["work_pattern"], "hours_per_day": safe_float(e["hours_per_day"]),
               "expected_hours": exp, "actual_rounded_hours": actual, "shortfall_hours": short,
               "outstanding_hours_before": outstanding_before,
               "shortfall_deducted_hours": shortfall_hours,
               "outstanding_hours_after": outstanding_after,
               "commission": {"sevens_pct": safe_float(e["comm_sevens_pct"]),
                              "meatdevil_pct": safe_float(e["comm_meatdevil_pct"]),
                              "patty_pct": safe_float(e["comm_patty_pct"]),
                              "sevens_sales": sales_sevens, "meatdevil_sales": sales_meat,
                              "patty_sales": sales_patty},
               "sales_month_used": f"{calendar.month_name[sm]} {sy}",
               "sales_source": sales_source,
               "ot_hours": ot_hours,
               "calculation": {"basic_pay": basic_pay, "hourly_pay": hourly_pay, "ot_pay": ot_pay,
                               "comm_sevens": comm_sevens, "comm_meatdevil": comm_meat,
                               "comm_patty": comm_patty, "comm_total": comm_total,
                               "travel_subsidy": travel, "shortfall_deduction": shortfall_deduction,
                               "other_adjustment": other_adjust,
                               "subsidy_name": subsidy_name, "subsidy_amount": subsidy_amt,
                               "extra_subsidy_name": ex_sub_name, "extra_subsidy_amount": ex_sub_amt,
                               "deduction_name": deduction_name, "deduction_amount": deduction_amt,
                               "extra_deduction_name": ex_ded_name, "extra_deduction_amount": ex_ded_amt,
                               "subsidies_total": subsidies_total, "deductions_total": deductions_total,
                               "total_before_mpf": total_before_mpf,
                               "mpf_employee": mpf_emp_amt, "mpf_employer": mpf_er_amt,
                               "total_after_mpf": total_after_mpf},
               "probation": truthy(e["probation"]),
               "shortfall_rule": e["shortfall_rule"],
               "mpf": {"employee_pct": safe_float(e["mpf_employee_pct"]),
                       "employer_pct": safe_float(e["mpf_employer_pct"]),
                       "started": mpf_on, "basis": "basic salary only",
                       "start_date": str(e.get("mpf_start_date", "")),
                       "enrollment_note": str(e.get("mpf_enrollment_note", ""))},
               "notes": str(e.get("notes", "")),
               "hybrid_note": hybrid_note,
               "shop_address": "Shop B, Basehall 03, LG/F, 1 Connaught Pl, Central"}
    return summary, payload


def save_months(summaries):
    """Upsert ledger rows (employee+year+month) and, where the saved month is the employee's
    latest in the ledger, push outstanding_after to the master's current version."""
    summ = load_df("monthly_summary", SUMMARY_COLS)
    new = pd.DataFrame(summaries)[SUMMARY_COLS].drop_duplicates(["employee", "year", "month"], keep="last")
    keys = set(zip(new["employee"], new["year"].astype(int), new["month"].astype(int)))
    old_hit = pd.Series([(r.employee, int(r.year), int(r.month)) in keys for r in summ.itertuples()],
                        index=summ.index, dtype=bool)
    summ = pd.concat([summ[~old_hit], new], ignore_index=True)
    save_df(summ, "monthly_summary")

    emp = load_df("employees", EMP_COLS)
    changed = False
    for s in new.to_dict("records"):
        rows = summ[summ["employee"] == s["employee"]]
        if s["year"] * 12 + s["month"] != (rows["year"] * 12 + rows["month"]).max():
            continue  # an older month re-saved: don't touch the current balance
        idx = current_version_idx(emp, s["employee"])
        if idx is None:
            continue
        old = safe_float(emp.loc[idx, "outstanding_hours"])
        if old != s["outstanding_after"]:
            emp.loc[idx, "outstanding_hours"] = s["outstanding_after"]
            changed = True
            log_history(s["employee"], "update", "outstanding_hours", old, s["outstanding_after"],
                        f"Saved month {int(s['year'])}-{int(s['month']):02d}: shortfall {s['shortfall']}, "
                        f"deducted {s['shortfall_deducted_hours']}h")
    if changed:
        save_df(emp, "employees")


def delete_months(keys):
    """Delete ledger rows [(employee, year, month)]. If the deleted month is the employee's latest,
    the master's outstanding balance is rolled back to that month's 'before'."""
    summ = load_df("monthly_summary", SUMMARY_COLS)
    emp = load_df("employees", EMP_COLS)
    for n, y, m in sorted(keys, key=lambda t: (t[1], t[2]), reverse=True):
        hit = summ[(summ["employee"] == n) & (summ["year"] == y) & (summ["month"] == m)]
        if hit.empty:
            continue
        k = summ["year"] * 12 + summ["month"]
        later = ((summ["employee"] == n) & (k > y * 12 + m)).any()
        if not later:
            idx = current_version_idx(emp, n)
            if idx is not None:
                ob = safe_float(hit.iloc[0]["outstanding_before"])
                log_history(n, "update", "outstanding_hours", emp.loc[idx, "outstanding_hours"], ob,
                            f"Ledger month {y}-{m:02d} deleted — balance rolled back")
                emp.loc[idx, "outstanding_hours"] = ob
        summ = summ.drop(hit.index)
    save_df(summ, "monthly_summary")
    save_df(emp, "employees")


def payslip_table(summary):
    rows = [("Basic salary", "basic_pay"), ("Hourly pay", "hourly_pay"),
            ("OT pay", "ot_pay"), ("Commission – SEVENS", "comm_sevens"),
            ("Commission – MEAT DEVIL", "comm_meatdevil"), ("Commission – PATTY", "comm_patty"),
            ("Travel subsidy", "travel"), ("Shortfall deduction", "shortfall_deduction"),
            ("Other adjustment", "other_adjustment")]
    if safe_float(summary.get("subsidy_amount")):
        rows.append((f"Subsidy – {summary.get('subsidy_name') or 'monthly'}", "subsidy_amount"))
    if safe_float(summary.get("extra_subsidy_amount")):
        rows.append((f"Extra subsidy – {summary.get('extra_subsidy_name') or 'this month'}", "extra_subsidy_amount"))
    if safe_float(summary.get("deduction_amount")):
        rows.append((f"Deduction – {summary.get('deduction_name') or 'monthly'}", "deduction_amount"))
    if safe_float(summary.get("extra_deduction_amount")):
        rows.append((f"Extra deduction – {summary.get('extra_deduction_name') or 'this month'}", "extra_deduction_amount"))
    rows += [("TOTAL before MPF", "total_before_mpf"),
            ("MPF employee (on basic)", "mpf_employee"), ("MPF employer (on basic)", "mpf_employer"),
            ("TOTAL after MPF", "total_after_mpf")]
    df = pd.DataFrame({"Item": [r[0] for r in rows],
                       "Amount (HKD)": [hkd(summary[k]) for _, k in rows]})

    def hl(row):
        if row["Item"].startswith("TOTAL"):
            return ["font-weight: 700; background-color: #f0f4ff;"] * len(row)
        if "deduction" in row["Item"].lower():
            return ["color: #b00020;"] * len(row)
        return [""] * len(row)
    return df.style.apply(hl, axis=1)


def normalize_holiday_dates(text):
    out = set()
    for d in re.findall(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", str(text)):
        try:
            out.add(datetime.strptime(d.replace("/", "-"), "%Y-%m-%d").date().isoformat())
        except ValueError:
            pass
    return out


_pm = date.today().replace(day=1) - timedelta(days=1)   # payroll is normally run for last month
DEF_Y, DEF_M = _pm.year, _pm.month

# ---------------------------------------------------------------------------
# Shell: sidebar + page routing
# ---------------------------------------------------------------------------
st.set_page_config(page_title="HK Payroll Helper", page_icon="🧾", layout="wide")
st.markdown("""
<style>
[data-testid="stSidebar"] h1 {font-size: 1.4rem; margin-bottom: 0;}
</style>
""", unsafe_allow_html=True)

PAGES = {
    "👥 1. Employee Master": "emp", "🕐 2. Attendance": "att",
    "🛒 3. Channel Sales": "sales", "💰 4. Monthly Calculation": "calc",
    "📈 5. Continuous Contract": "cc", "📤 6. Export / Payslip Block": "exp",
    "📒 7. Monthly Summary": "sum", "🕑 8. History": "hist",
}
with st.sidebar:
    st.title("🧾 HK Payroll")
    st.caption("永盈茶餐厅 / SEVENS")
    st.divider()
    page = st.radio("Go to", list(PAGES))
    st.divider()
    st.download_button("⬇ Backup all data (zip)", backup_zip(),
                       f"payroll_backup_{date.today().isoformat()}.zip", mime="application/zip")
    st.caption(f"Data folder: {DATA_DIR}")
pg = PAGES[page]

if "_flash" in st.session_state:
    _kind, _msg = st.session_state.pop("_flash")
    getattr(st, _kind)(_msg)

# =========================== 1. EMPLOYEE MASTER ===========================
FORM_DEFAULTS = {
    "ef_name": "", "ef_join": date.today(), "ef_pkg": "fixed_monthly",
    "ef_entry_start": date.today(), "ef_entry_end": "", "ef_reason": "",
    "ef_basic": 0.0, "ef_hr": 0.0, "ef_hpd": 8.0, "ef_pattern": "5-day",
    "ef_travel": 0.0, "ef_ot": 0.0, "ef_cs": 0.0, "ef_cm": 0.0, "ef_cp": 0.0,
    "ef_prob": False, "ef_rule": "none",
    "ef_sub_name": "", "ef_sub_amt": 0.0, "ef_ded_name": "", "ef_ded_amt": 0.0,
    "ef_mpf_emp": 5.0, "ef_mpf_er": 5.0, "ef_mpf_started": False,
    "ef_mpf_start": "", "ef_mpf_note": "", "ef_outstanding": 0.0, "ef_notes": "",
}


def _load_preset():
    """Selectbox callback: fill the form ONCE when a person is chosen (not on every rerun)."""
    st.session_state.update(FORM_DEFAULTS)
    sel = st.session_state.get("load_emp")
    if not sel or sel == "(new employee)":
        return
    rows = load_df("employees", EMP_COLS)
    rows = rows[rows["name"] == sel].sort_values("entry_start")
    if rows.empty:
        return
    p = rows.iloc[-1].to_dict()
    st.session_state.update({
        "ef_name": p["name"], "ef_join": to_date(p["join_date"]), "ef_pkg": p["package_type"] or "fixed_monthly",
        "ef_basic": safe_float(p["basic_monthly"]), "ef_hr": safe_float(p["hourly_rate"]),
        "ef_hpd": safe_float(p["hours_per_day"], 8), "ef_pattern": p["work_pattern"] or "5-day",
        "ef_travel": safe_float(p["travel_subsidy"]), "ef_ot": safe_float(p["ot_rate"]),
        "ef_cs": safe_float(p["comm_sevens_pct"]), "ef_cm": safe_float(p["comm_meatdevil_pct"]),
        "ef_cp": safe_float(p["comm_patty_pct"]), "ef_prob": truthy(p["probation"]),
        "ef_rule": p["shortfall_rule"] or "pay_full_recover_next_month",
        "ef_mpf_emp": safe_float(p["mpf_employee_pct"], 5), "ef_mpf_er": safe_float(p["mpf_employer_pct"], 5),
        "ef_mpf_started": truthy(p["mpf_started"]), "ef_mpf_start": str(p["mpf_start_date"])[:10],
        "ef_mpf_note": p["mpf_enrollment_note"],
        "ef_sub_name": p["subsidy_name"], "ef_sub_amt": safe_float(p["subsidy_amount"]),
        "ef_ded_name": p["deduction_name"], "ef_ded_amt": safe_float(p["deduction_amount"]),
        "ef_outstanding": safe_float(p["outstanding_hours"]),
        "ef_notes": p["notes"],
    })


if pg == "emp":
    st.header("👥 Employee Master")
    st.caption("One row per person per version — packages, rates, MPF, commission %, outstanding hours. "
               "All changes are logged in History.")
    emp = load_df("employees", EMP_COLS)
    display_cols = ["name", "join_date", "entry_start", "entry_end", "change_reason",
                    "package_type", "basic_monthly", "hourly_rate",
                    "hours_per_day", "work_pattern", "travel_subsidy", "ot_rate",
                    "comm_sevens_pct", "comm_meatdevil_pct", "comm_patty_pct",
                    "subsidy_name", "subsidy_amount", "deduction_name", "deduction_amount",
                    "probation", "mpf_started", "mpf_start_date", "outstanding_hours", "shortfall_rule"]
    show_df(emp[display_cols])

    t_view, t_edit, t_hol = st.tabs(["📂 CSV Import / Export / Delete", "✏️ Add / Edit Form", "🗓️ Public Holidays"])

    # ---------- Tab: CSV ----------
    with t_view:
        st.markdown("**Bulk maintenance** — export, edit in Excel, re-import. Import matches name + entry_start (upsert).")
        ce1, ce2, ce3 = st.columns([1, 2, 1])
        ce1.download_button("⬇ Export employees CSV",
                            emp.to_csv(index=False).encode("utf-8-sig"),
                            "employees_export.csv", mime="text/csv", disabled=emp.empty)
        ce1.download_button("⬇ Blank template (headers)",
                            pd.DataFrame(columns=EMP_COLS).to_csv(index=False).encode("utf-8-sig"),
                            "employees_template.csv", mime="text/csv")
        up_emp = ce2.file_uploader("Import employees CSV", type=["csv"], key="empcsv")
        replace_all = ce3.checkbox("Replace entire master",
                                   help="Anyone NOT in the uploaded file is deleted.")
        if up_emp:
            try:
                df_i = read_csv_upload(up_emp)
            except Exception as ex:
                st.error(f"Could not read file: {ex}")
                df_i = None
            if df_i is not None:
                if "entry_start" not in df_i.columns:
                    df_i["entry_start"] = df_i["join_date"] if "join_date" in df_i.columns else ""
                    st.info("No entry_start column found — defaulted to join_date for all rows.")
                for c, default in (("entry_end", ""), ("change_reason", "csv import")):
                    if c not in df_i.columns:
                        df_i[c] = default
                missing = [c for c in EMP_COLS if c not in df_i.columns]
                st.caption(f"{len(df_i)} row(s) read.")
                if missing:
                    st.error(f"Missing columns: {missing}. Export first to get the exact headers.")
                elif st.button("Run employee CSV import", type="primary"):
                    created = updated = 0
                    errs = []
                    for n, row in df_i.iterrows():
                        nm = str(row["name"]).strip()
                        if not nm:
                            continue
                        try:
                            vals = {}
                            for c in EMP_COLS:
                                v = row[c]
                                if c in NUMERIC_EMP_COLS:
                                    vals[c] = safe_float(v)
                                elif c in BOOL_EMP_COLS:
                                    vals[c] = truthy(v)
                                elif c in ("join_date", "entry_start", "entry_end", "mpf_start_date"):
                                    vals[c] = norm_date(v)
                                else:
                                    vals[c] = str(v)
                            vals["name"] = nm
                            if not vals["entry_start"]:
                                vals["entry_start"] = vals["join_date"]
                            if not vals["entry_start"]:
                                raise ValueError("entry_start and join_date both blank")
                        except ValueError as ex:
                            errs.append(f"Row {n + 2} ({nm}): {ex} — skipped")
                            continue
                        existing = emp[(emp["name"] == nm) & (emp["entry_start"] == vals["entry_start"])]
                        if existing.empty:
                            created += 1
                            log_history(nm, "create", "ALL", "", f"csv_import v{vals['entry_start']}",
                                        "New version via CSV import")
                        else:
                            updated += 1
                            log_history(nm, "csv_import", "ALL", f"version {vals['entry_start']}",
                                        "see new values", "Version replaced via CSV import")
                            emp = emp.drop(existing.index)
                        emp = pd.concat([emp, pd.DataFrame([vals])], ignore_index=True)
                    if replace_all:
                        keep_names = {str(x).strip() for x in df_i["name"]}
                        for d in sorted(set(emp.loc[~emp["name"].isin(keep_names), "name"])):
                            log_history(d, "delete", "ALL", "exists", "deleted", "Removed by CSV replace-all")
                        emp = emp[emp["name"].isin(keep_names)]
                    save_df(emp.reset_index(drop=True), "employees")
                    for er in errs:
                        st.error(er)
                    flash(f"Import done: {created} created, {updated} updated"
                          + (", replace-all applied" if replace_all else "")
                          + (f" · {len(errs)} row(s) skipped (see History for the rest)" if errs else ""))
                    if not errs:
                        st.rerun()
                    else:
                        st.warning(f"Import done with {len(errs)} skipped row(s) — fix them and re-import.")
        st.divider()
        st.markdown("**Danger zone — delete a version**")
        if not emp.empty:
            opts = {i: f"{r['name']} | {r['entry_start'] or '?'} → {r['entry_end'] or 'current'}"
                    for i, r in emp.iterrows()}
            dz1, dz2, dz3 = st.columns([2, 1, 1])
            del_idx = dz1.selectbox("Version to delete", list(opts), format_func=opts.get, key="del_ver")
            sure = dz2.checkbox("I'm sure", key=f"del_ver_ok_{st.session_state.get('del_n', 0)}")
            if dz3.button("Delete version (history kept)"):
                if not sure:
                    st.warning("Tick 'I'm sure' first.")
                else:
                    r = emp.loc[del_idx]
                    log_history(r["name"], "delete", "ALL", f"version {r['entry_start']}", "deleted",
                                "Version removed from master")
                    save_df(emp.drop(del_idx).reset_index(drop=True), "employees")
                    st.session_state["del_n"] = st.session_state.get("del_n", 0) + 1
                    flash(f"Deleted {opts[del_idx]}")
                    st.rerun()

    # ---------- Tab: form ----------
    with t_edit:
        st.caption("Pick an existing employee to load their latest values, change what changed, then Save. "
                   "Same entry-start date = edit that version in place · later date = NEW version "
                   "(previous version is closed the day before).")
        for k, v in FORM_DEFAULTS.items():
            st.session_state.setdefault(k, v)
        st.selectbox("Load employee into form",
                     ["(new employee)"] + sorted(emp["name"].unique().tolist()),
                     key="load_emp", on_change=_load_preset)

        with st.form("emp_form"):
            st.markdown("**Identity & package**")
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("Name *", key="ef_name")
            join = c2.date_input("Join date", key="ef_join")
            pkg = c3.selectbox("Package type", ["fixed_monthly", "hourly", "hybrid_prestart"], key="ef_pkg",
                               help="hybrid_prestart: tracks shortfall like fixed; put pre-start hourly pay in Other adjustment.")
            st.markdown("**Versioning** — promotions & rate changes create a NEW entry; old entries stay untouched.")
            v1, v2 = st.columns(2)
            entry_start = v1.date_input("Entry start (effective date) *", key="ef_entry_start")
            entry_end_txt = v2.text_input("Entry end (YYYY-MM-DD, blank = current)", key="ef_entry_end")
            change_reason = st.text_input("Change reason (e.g. Promotion, hourly rate change)", key="ef_reason")
            c4, c5, c6 = st.columns(3)
            basic = c4.number_input("Basic monthly (HKD)", min_value=0.0, step=100.0, key="ef_basic")
            hr = c5.number_input("Hourly rate (HKD)", min_value=0.0, step=1.0, key="ef_hr")
            hpd = c6.number_input("Expected hours / day", min_value=0.0, step=0.5, key="ef_hpd")
            c7, c8, c9 = st.columns(3)
            pattern = c7.selectbox("Work pattern", ["5-day", "5.5-day"], key="ef_pattern")
            travel = c8.number_input("Travel subsidy (monthly HKD)", min_value=0.0, step=50.0, key="ef_travel")
            ot = c9.number_input("OT rate (HKD/hour, 0 = use hour value)", min_value=0.0, step=1.0, key="ef_ot")
            st.markdown("**Open-name monthly adjustments** — fixed HKD, applied every month")
            c20, c21, c22, c23 = st.columns(4)
            sub_name = c20.text_input("Subsidy label", key="ef_sub_name", placeholder="e.g. Meal allowance")
            sub_amt = c21.number_input("Subsidy HKD/mo", min_value=0.0, step=50.0, key="ef_sub_amt")
            ded_name = c22.text_input("Deduction label", key="ef_ded_name", placeholder="e.g. Advance repayment")
            ded_amt = c23.number_input("Deduction HKD/mo", min_value=0.0, step=50.0, key="ef_ded_amt")
            st.markdown("**Commission % by channel** (of that channel's sales)")
            c10, c11, c12 = st.columns(3)
            comm_s = c10.number_input("SEVENS %", min_value=0.0, step=0.1, format="%.2f", key="ef_cs")
            comm_m = c11.number_input("MEAT DEVIL %", min_value=0.0, step=0.1, format="%.2f", key="ef_cm")
            comm_p = c12.number_input("PATTY %", min_value=0.0, step=0.1, format="%.2f", key="ef_cp")
            st.markdown("**Status & shortfall rule**")
            c13, c14 = st.columns(2)
            prob = c13.checkbox("On probation", key="ef_prob")
            rule = c14.selectbox("Shortfall rule", ["none", "pay_full_recover_next_month", "pay_actual"],
                                 key="ef_rule",
                                 help="v2.3: all three rules now mean the same thing - FULL month basic, and the "
                                      "month's shortfall is carried into outstanding hours. Deduct only via the "
                                      "'Shortfall hours to deduct' input on the Calculation page (x hour value rate). "
                                      "This field is kept for record.")
            st.markdown("**MPF** — calculated on basic salary only (default 5% / 5%)")
            c15, c16, c17 = st.columns(3)
            mpf_emp = c15.number_input("Employee MPF %", min_value=0.0, step=0.5, key="ef_mpf_emp")
            mpf_er = c16.number_input("Employer MPF %", min_value=0.0, step=0.5, key="ef_mpf_er")
            mpf_started = c17.checkbox("MPF started / enrolled", key="ef_mpf_started")
            c18, c19 = st.columns(2)
            mpf_start_txt = c18.text_input("MPF start date (YYYY-MM-DD, optional)", key="ef_mpf_start")
            mpf_note = c19.text_input("Enrollment confirmation note", key="ef_mpf_note",
                                      placeholder="e.g. Enrolled 2026-08-15, scheme ABC")
            outstanding = st.number_input("Outstanding / shortfall hours carried forward (opening balance)",
                                          min_value=0.0, step=0.25, key="ef_outstanding")
            notes = st.text_area("Notes (continuous-contract status, special rules)", key="ef_notes")
            submitted = st.form_submit_button("💾 Save version", type="primary")

        if submitted:
            try:
                ee = norm_date(entry_end_txt)
                mpf_sd = norm_date(mpf_start_txt)
                err = ""
            except ValueError as ex:
                ee = mpf_sd = ""
                err = str(ex)
            es = entry_start.isoformat()
            name = name.strip()
            if not name:
                st.error("Name is required.")
            elif err:
                st.error(f"Date problem: {err}. Use YYYY-MM-DD.")
            elif ee and ee < es:
                st.error("Entry end is before entry start.")
            else:
                new_vals = {"name": name, "join_date": join.isoformat(),
                            "entry_start": es, "entry_end": ee, "change_reason": change_reason,
                            "package_type": pkg,
                            "basic_monthly": basic, "hourly_rate": hr, "hours_per_day": hpd,
                            "work_pattern": pattern, "travel_subsidy": travel, "ot_rate": ot,
                            "comm_sevens_pct": comm_s, "comm_meatdevil_pct": comm_m, "comm_patty_pct": comm_p,
                            "probation": prob, "shortfall_rule": rule,
                            "mpf_employee_pct": mpf_emp, "mpf_employer_pct": mpf_er,
                            "mpf_started": mpf_started, "mpf_start_date": mpf_sd,
                            "mpf_enrollment_note": mpf_note,
                            "subsidy_name": sub_name.strip(), "subsidy_amount": sub_amt,
                            "deduction_name": ded_name.strip(), "deduction_amount": ded_amt,
                            "outstanding_hours": outstanding, "notes": notes}
                emp = load_df("employees", EMP_COLS)
                same_version = emp[(emp["name"] == name) & (emp["entry_start"] == es)]
                closed_note = ""
                if not same_version.empty:
                    old_row = same_version.iloc[0]
                    for field, new_v in new_vals.items():
                        if str(old_row.get(field, "")) != str(new_v):
                            log_history(name, "update", field, old_row.get(field, ""), new_v,
                                        f"Edit of version starting {es}")
                    emp = emp.drop(same_version.index)
                else:
                    mine = emp[emp["name"] == name]
                    # close earlier open version(s) the day before the new one starts
                    prev = mine[(mine["entry_start"] < es) & (mine["entry_end"] == "")]
                    close_d = (entry_start - timedelta(days=1)).isoformat()
                    for idx in prev.index:
                        emp.loc[idx, "entry_end"] = close_d
                        log_history(name, "update", "entry_end", "", close_d,
                                    f"Closed by new version starting {es}: {change_reason or 'no reason given'}")
                    if not prev.empty:
                        closed_note = f" Previous version closed at {close_d}."
                    # historical insert: end this version before the next later one
                    later = mine[mine["entry_start"] > es]
                    if not ee and not later.empty:
                        new_vals["entry_end"] = (to_date(later["entry_start"].min()) - timedelta(days=1)).isoformat()
                    log_history(name, "create_version", "entry_start", "", es,
                                f"New version: {change_reason or 'no reason given'}")
                emp = pd.concat([emp, pd.DataFrame([new_vals])], ignore_index=True)
                save_df(emp, "employees")
                flash(f"Saved version for {name} effective {es}.{closed_note}")
                st.rerun()

    # ---------- Tab: holidays ----------
    with t_hol:
        st.markdown(f"**{len(HOLIDAYS)} holidays currently loaded**")
        st.caption("Paste the gazette list in any format (YYYY-MM-DD or YYYY/M/D, comma or line separated), "
                   "or upload a CSV with dates anywhere in the file.")
        paste_h = st.text_area("Paste holiday dates", height=110,
                               placeholder="2027-01-01\n2027-02-06, 2027-02-07, 2027-02-08")
        if st.button("Add pasted holidays") and paste_h.strip():
            found = normalize_holiday_dates(paste_h)
            new = {d for d in found if d not in HOLIDAYS}
            HOLIDAYS |= new
            save_holidays(HOLIDAYS)
            flash(f"Added {len(new)} new holiday date(s) ({len(found)} recognised).")
            st.rerun()
        up_h = st.file_uploader("Or upload holidays CSV", type=["csv"], key="holcsv")
        if up_h:
            found = normalize_holiday_dates(up_h.getvalue().decode("utf-8-sig", errors="ignore"))
            if st.button(f"Add {len(found)} date(s) from CSV"):
                new = {d for d in found if d not in HOLIDAYS}
                HOLIDAYS |= new
                save_holidays(HOLIDAYS)
                flash(f"Added {len(new)} new holiday date(s).")
                st.rerun()
        hy1, hy2 = st.columns([1, 2])
        def_year = hy1.selectbox("Built-in HK defaults", sorted(DEFAULT_HOLIDAYS), key="defyear")
        if hy2.button(f"Load {def_year} defaults"):
            new = {d for d in DEFAULT_HOLIDAYS[def_year] if d not in HOLIDAYS}
            HOLIDAYS |= set(DEFAULT_HOLIDAYS[def_year])
            save_holidays(HOLIDAYS)
            flash(f"Loaded {len(DEFAULT_HOLIDAYS[def_year])} holidays for {def_year} ({len(new)} new).")
            st.rerun()
        rem = st.multiselect("Remove holiday(s)", sorted(HOLIDAYS))
        if st.button("Remove selected") and rem:
            for d in rem:
                HOLIDAYS.discard(d)
            save_holidays(HOLIDAYS)
            st.rerun()

# =========================== 2. ATTENDANCE ===========================
elif pg == "att":
    st.header("🕐 Attendance")
    st.caption("Import the 企业微信 daily report, or enter manually. Hours rounded to the nearest 15 minutes. "
               "One record per employee per day — re-importing a day replaces it.")
    emp = load_df("employees", EMP_COLS)
    if emp.empty:
        st.warning("Add employees first (page 1).")
        st.stop()

    t_imp, t_man, t_all = st.tabs(["📥 Import WeChat xlsx", "✍️ Manual Entry", "📋 All Records / Delete"])

    with t_imp:
        st.markdown("**Step 1** — upload 上下班打卡_日报 (.xlsx)")
        up = st.file_uploader("WeChat Work daily report", type=["xlsx"])
        if up:
            db = st.number_input("Default break per day (hours, subtracted before rounding)",
                                 min_value=0.0, step=0.5, value=0.0)
            try:
                recs, skipped = parse_wework_export(up, known_names=set(emp["name"]), default_break=db)
            except Exception as ex:
                st.error(f"Could not read this file: {ex}")
                recs, skipped = pd.DataFrame(columns=ATT_COLS), []
            if recs.empty:
                st.warning("No punch records found in this file.")
                for s in skipped:
                    st.caption("⏭️ Skipped: " + s)
            else:
                recs = recs.reset_index(drop=True)
                flag_idx = recs.index[recs["notes"].str.contains("single punch", na=False)]
                if len(flag_idx) > 0:
                    st.warning(f"**Step 2 — review:** {len(flag_idx)} single-punch day(s). "
                               "Fill in the missing Start or End time (leave as-is if absent):")
                    edited = st.data_editor(
                        recs.loc[flag_idx, ["employee", "date", "start", "end", "break_hours", "notes"]],
                        num_rows="fixed",
                        column_config={
                            "employee": st.column_config.TextColumn(disabled=True),
                            "date": st.column_config.TextColumn(disabled=True),
                            "start": st.column_config.TextColumn("Start (HH:MM)"),
                            "end": st.column_config.TextColumn("End (HH:MM)"),
                            "notes": st.column_config.TextColumn(disabled=True)})
                    for i, row in edited.iterrows():
                        raw = compute_raw_hours(row["start"], row["end"], safe_float(row["break_hours"]))
                        if raw and raw > 0:
                            recs.loc[i, ["start", "end", "break_hours"]] = [row["start"], row["end"],
                                                                            safe_float(row["break_hours"])]
                            recs.loc[i, "raw_hours"] = raw
                            recs.loc[i, "rounded_hours"] = round15(raw)
                            recs.loc[i, "notes"] = "corrected after review"
                    fixed = int((recs.loc[flag_idx, "notes"] == "corrected after review").sum())
                    st.caption(f"Reviewed: {fixed}/{len(flag_idx)} corrected, "
                               f"{len(flag_idx) - fixed} still single-punch (imported as 0h, flagged).")
                for s in skipped:
                    st.caption("⏭️ Skipped: " + s)
                st.markdown(f"**Step 3 — import {len(recs)} record(s):**")
                show_df(recs)
                if st.button(f"Import {len(recs)} record(s)", type="primary"):
                    added, replaced = add_attendance(recs)
                    flash(f"Imported {added} record(s); {replaced} existing day(s) replaced.")
                    st.rerun()

    with t_man:
        names_set = set(emp["name"])
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Paste rows** — one per line:  \n`Employee | YYYY-MM-DD | Start | End | Break(h) | Notes`")
            paste = st.text_area("Paste attendance rows", height=180,
                                 placeholder="Natalie | 2026-09-03 | 09:00 | 18:00 | 1 |")
            if st.button("Add pasted rows"):
                good, bad = [], []
                for line in paste.strip().splitlines():
                    if not line.strip():
                        continue
                    parts = [p.strip() for p in line.split("|")]
                    try:
                        if len(parts) < 5:
                            raise ValueError("need at least 5 fields")
                        e2, d, s, en, b = parts[0], norm_date(parts[1]), parts[2], parts[3], float(parts[4])
                        if not d:
                            raise ValueError("date missing")
                        if e2 not in names_set:
                            raise ValueError("unknown employee")
                        raw = compute_raw_hours(s, en, b)
                        if raw is None:
                            raise ValueError("bad time")
                    except ValueError as ex:
                        bad.append(f"{line}  →  {ex}")
                        continue
                    good.append({"employee": e2, "date": d, "start": s, "end": en, "break_hours": b,
                                 "raw_hours": raw, "rounded_hours": round15(raw),
                                 "notes": parts[5] if len(parts) > 5 else ""})
                if good:
                    added, replaced = add_attendance(pd.DataFrame(good, columns=ATT_COLS))
                    st.success(f"Added {added} row(s) ({replaced} replaced).")
                for b in bad:
                    st.error(f"Skipped: {b}")
        with c2:
            st.markdown("**Single entry**")
            with st.form("att_form"):
                e2 = st.selectbox("Employee", sorted(names_set))
                d = st.date_input("Date", value=date.today())
                s = st.text_input("Start", value="09:00")
                en = st.text_input("End", value="18:00")
                b = st.number_input("Break (h)", min_value=0.0, step=0.5)
                n = st.text_input("Notes")
                if st.form_submit_button("Add record", type="primary"):
                    raw = compute_raw_hours(s, en, b)
                    if raw is None:
                        st.error("Could not parse times. Use HH:MM.")
                    else:
                        add_attendance(pd.DataFrame([{
                            "employee": e2, "date": d.isoformat(), "start": s, "end": en,
                            "break_hours": b, "raw_hours": raw,
                            "rounded_hours": round15(raw), "notes": n}], columns=ATT_COLS))
                        st.success(f"Saved: raw {raw}h → rounded {round15(raw)}h (replaces any existing record that day)")

    with t_all:
        att = load_df("attendance", ATT_COLS)
        if att.empty:
            st.info("No attendance records yet.")
        else:
            att["_ym"] = att["date"].str[:7]
            f1, f2 = st.columns(2)
            fname = f1.selectbox("Employee", ["(all)"] + sorted(att["employee"].unique().tolist()))
            fmon = f2.selectbox("Month", ["(all)"] + sorted(att["_ym"].unique().tolist(), reverse=True))
            view = att
            if fname != "(all)":
                view = view[view["employee"] == fname]
            if fmon != "(all)":
                view = view[view["_ym"] == fmon]
            view = view.drop(columns="_ym").sort_values(["employee", "date"])
            st.caption(f"{len(view)} record(s) shown · total rounded hours {view['rounded_hours'].sum():.2f}")
            ver = st.session_state.get("att_ver", 0)
            tbl = view.copy()
            tbl.insert(0, "delete", False)
            edited = st.data_editor(
                tbl, key=f"att_edit_{fname}_{fmon}_{ver}", num_rows="fixed", hide_index=True,
                disabled=ATT_COLS,
                column_config={"delete": st.column_config.CheckboxColumn("🗑 delete", default=False)})
            ticked = edited.index[edited["delete"]]
            d1, d2, d3 = st.columns([1, 1, 1])
            if d1.button(f"🗑️ Delete {len(ticked)} ticked record(s)", disabled=len(ticked) == 0):
                save_df(att.drop(columns="_ym").drop(ticked).reset_index(drop=True), "attendance")
                st.session_state["att_ver"] = ver + 1
                flash(f"Deleted {len(ticked)} record(s).")
                st.rerun()
            sure = d2.checkbox(f"I'm sure — delete all {len(view)} shown", key=f"att_sure_{ver}")
            if d3.button("🗑️ Delete ALL shown records", disabled=view.empty):
                if not sure:
                    st.warning("Tick 'I'm sure' first.")
                else:
                    save_df(att.drop(columns="_ym").drop(view.index).reset_index(drop=True), "attendance")
                    st.session_state["att_ver"] = ver + 1
                    flash(f"Deleted {len(view)} record(s).")
                    st.rerun()

# =========================== 3. CHANNEL SALES ===========================
elif pg == "sales":
    st.header("🛒 Channel Sales")
    st.caption("One consolidated sales figure per channel per month. Monthly commission "
               "auto-pulls the T-2 month: payroll for month M uses these figures from "
               "month M-2 (the bonus shows up 2 months later automatically).")
    cs = load_df("channel_sales", SALES_COLS).sort_values(["year", "month"], ascending=False)
    t_in, t_all = st.tabs(["✍️ Enter / Update", "📋 All Months"])
    with t_in:
        c1, c2 = st.columns(2)
        sy_in = int(c1.number_input("Year", 2024, 2040, DEF_Y, key="csy"))
        sm_in = int(c2.selectbox("Month", list(range(1, 13)), index=DEF_M - 1, key="csm"))
        existing = cs[(cs["year"] == sy_in) & (cs["month"] == sm_in)] if not cs.empty else pd.DataFrame()
        d_sev = d_meat = d_pat = 0.0
        if not existing.empty:
            d_sev, d_meat, d_pat = (safe_float(existing.iloc[0]["sevens_sales"]),
                                    safe_float(existing.iloc[0]["meatdevil_sales"]),
                                    safe_float(existing.iloc[0]["patty_sales"]))
            st.caption(f"Existing entry for {calendar.month_name[sm_in]} {sy_in} — saving overwrites it.")
        s1, s2, s3 = st.columns(3)
        v_sev = s1.number_input("SEVENS sales (HKD)", min_value=0.0, step=1000.0, value=d_sev, key="cs_sev")
        v_meat = s2.number_input("MEAT DEVIL sales (HKD)", min_value=0.0, step=1000.0, value=d_meat, key="cs_meat")
        v_pat = s3.number_input("PATTY sales (HKD)", min_value=0.0, step=1000.0, value=d_pat, key="cs_pat")
        py_, pm_ = shift_month(sy_in, sm_in, +2)
        st.info(f"These figures will be used for the payroll of {calendar.month_name[pm_]} {py_} (T+2).")
        if st.button("💾 Save channel sales", type="primary"):
            save_channel_sales(sy_in, sm_in, v_sev, v_meat, v_pat)
            flash(f"Saved {calendar.month_name[sm_in]} {sy_in}: SEVENS {hkd(v_sev)} / "
                  f"MEAT DEVIL {hkd(v_meat)} / PATTY {hkd(v_pat)}")
            st.rerun()
    with t_all:
        if cs.empty:
            st.info("No channel sales recorded yet.")
        else:
            show_df(cs.assign(month=cs["month"].map(lambda m: calendar.month_name[int(m)])))
            st.download_button("⬇ Download channel sales (CSV)",
                               cs.to_csv(index=False).encode("utf-8-sig"),
                               "channel_sales.csv", mime="text/csv")
            st.divider()
            st.markdown("**Delete a month**")
            labels = {f"{int(r.year)}-{int(r.month):02d}": (int(r.year), int(r.month))
                      for r in cs.itertuples()}
            pick = st.multiselect("Month(s) to delete", list(labels))
            sure = st.checkbox("I'm sure", key=f"cs_del_{st.session_state.get('cs_n', 0)}")
            if st.button("🗑️ Delete selected", disabled=not pick):
                if not sure:
                    st.warning("Tick 'I'm sure' first.")
                else:
                    df = load_df("channel_sales", SALES_COLS)
                    for yy, mm_ in [labels[p] for p in pick]:
                        df = df[~((df["year"] == yy) & (df["month"] == mm_))]
                    save_df(df.reset_index(drop=True), "channel_sales")
                    st.session_state["cs_n"] = st.session_state.get("cs_n", 0) + 1
                    flash(f"Deleted {len(pick)} month(s).")
                    st.rerun()

# =========================== 4. MONTHLY CALCULATION ===========================
elif pg == "calc":
    st.header("💰 Monthly Calculation")
    st.caption("Batch: one CSV for the whole team → Grok blocks + summary CSV. Single mode for quick checks.")
    emp = load_df("employees", EMP_COLS)
    if emp.empty:
        st.warning("Add employees first (page 1).")
        st.stop()
    summ_ledger = load_df("monthly_summary", SUMMARY_COLS)

    t_batch, t_single = st.tabs(["👥 Batch — whole team (recommended)", "👤 Single employee"])

    # ---------------- BATCH ----------------
    with t_batch:
        st.markdown("**1 · Get the template** (pre-filled with everyone) → fill OT / shortfall / adjustments → "
                    "**2 · upload** → **3 · review** → **4 · outputs & save**.")
        st.caption("Sales are NOT entered here anymore — commission auto-pulls channel sales (T-2) from the "
                   "🛒 Channel Sales tab. shortfall_hours = hours to recover from outstanding this month "
                   "(× hour value rate). Delete rows for people not paid this month.")
        c1, c2 = st.columns(2)
        tmpl = pd.DataFrame(
            [{"name": n, "year": DEF_Y, "month": DEF_M, "ot_hours": 0, "shortfall_hours": 0,
              "other_adjustment": 0, "extra_subsidy_name": "", "extra_subsidy_amount": 0,
              "extra_deduction_name": "", "extra_deduction_amount": 0}
             for n in sorted(emp["name"].unique())], columns=INPUT_COLS)
        c1.download_button("⬇ 1. Download input template (CSV)",
                           tmpl.to_csv(index=False).encode("utf-8-sig"),
                           "monthly_input_template.csv", mime="text/csv")
        upcsv = c2.file_uploader("2. Upload filled monthly input CSV", type=["csv"], key="batchcsv")

        if upcsv:
            try:
                inp = read_csv_upload(upcsv)
            except Exception as ex:
                st.error(f"Could not read file: {ex}")
                inp = None
            missing_cols = [c for c in INPUT_COLS if inp is not None and c not in inp.columns]
            if inp is not None and missing_cols:
                # v2.4: templates saved before v2.4 lack only the new extra-adjustment columns -> default them
                legacy = {"extra_subsidy_name", "extra_subsidy_amount",
                          "extra_deduction_name", "extra_deduction_amount"}
                if set(missing_cols) <= legacy:
                    for c in missing_cols:
                        inp[c] = "" if c.endswith("_name") else "0"
                    st.info("Template predates v2.4 — extra subsidy/deduction columns defaulted to blank/0.")
                    missing_cols = []
            if inp is not None and missing_cols:
                st.error(f"Missing columns: {missing_cols}. Use the template headers.")
            elif inp is not None:
                summaries, payloads, errors, seen = [], [], [], set()
                for n, row in inp.iterrows():
                    nm = str(row["name"]).strip()
                    if not nm:
                        continue
                    try:
                        y, m = int(float(row["year"])), int(float(row["month"]))
                        if not (1 <= m <= 12 and 2000 <= y <= 2100):
                            raise ValueError
                    except (ValueError, TypeError):
                        errors.append(f"Row {n + 2} ({nm}): bad year/month — skipped")
                        continue
                    if (nm, y, m) in seen:
                        errors.append(f"Row {n + 2} ({nm} {y}-{m:02d}): duplicate row — skipped")
                        continue
                    seen.add((nm, y, m))
                    ver = get_employee_version(emp, nm, y, m)
                    if ver is None:
                        errors.append(f"Row {n + 2} ({nm}): not in employee master — skipped")
                        continue
                    # pre-v2.4.2 templates carry per-employee sales columns: non-zero values act
                    # as a manual override; otherwise commission auto-pulls channel sales T-2
                    has_sales_cols = {"sales_sevens", "sales_meatdevil", "sales_patty"} <= set(inp.columns)
                    manual_sales = (has_sales_cols and (safe_float(row["sales_sevens"]) > 0
                                    or safe_float(row["sales_meatdevil"]) > 0
                                    or safe_float(row["sales_patty"]) > 0))
                    s, p = compute_monthly(
                        ver, y, m,
                        sales_sevens=(safe_float(row["sales_sevens"]) if manual_sales else None),
                        sales_meat=(safe_float(row["sales_meatdevil"]) if manual_sales else None),
                        sales_patty=(safe_float(row["sales_patty"]) if manual_sales else None),
                        ot_hours=safe_float(row["ot_hours"]),
                        shortfall_hours=safe_float(row["shortfall_hours"]),
                        other_adjust=safe_float(row["other_adjustment"]),
                        extra_subsidy_name=row.get("extra_subsidy_name", ""),
                        extra_subsidy_amount=safe_float(row.get("extra_subsidy_amount")),
                        extra_deduction_name=row.get("extra_deduction_name", ""),
                        extra_deduction_amount=safe_float(row.get("extra_deduction_amount")),
                        outstanding_before=ledger_before(summ_ledger, ver, nm, y, m))
                    summaries.append(s)
                    payloads.append(p)
                for er in errors:
                    st.error(er)
                if summaries:
                    res_df = pd.DataFrame(summaries)[SUMMARY_COLS[:-1]]
                    st.success(f"✅ Calculated {len(summaries)} employee(s)")
                    show_df(res_df, column_config={
                        c: st.column_config.NumberColumn(format="%.2f") for c in MONEY_COLS + HOUR_COLS})
                    zero_att = [s["employee"] for s in summaries if s["actual_hours"] == 0]
                    if zero_att:
                        st.warning("No attendance found for: " + ", ".join(zero_att)
                                   + " — import attendance first if that's unexpected.")
                    if any("NOT RECORDED" in p["sales_source"] for p in payloads):
                        missing_m = sorted({p["sales_month_used"] for p in payloads
                                            if "NOT RECORDED" in p["sales_source"]})
                        st.warning("Channel sales not recorded for: " + ", ".join(missing_m)
                                   + " — commission is 0. Enter them in the 🛒 Channel Sales tab.")
                    tots = res_df[["total_before_mpf", "total_after_mpf"]].sum()
                    m1, m2 = st.columns(2)
                    m1.metric("Team total before MPF", f"HK$ {tots['total_before_mpf']:,.2f}")
                    m2.metric("Team total after MPF", f"HK$ {tots['total_after_mpf']:,.2f}")

                    if st.button("💾 Save to Monthly Summary & update outstanding balances", type="primary"):
                        save_months(summaries)
                        st.success("Saved. See 📒 Monthly Summary and 🕑 History. "
                                   "(Re-saving the same month replaces it — safe to repeat.)")

                    st.divider()
                    o1, o2 = st.tabs(["📤 Output 1 — Grok blocks (payslips)",
                                      "📤 Output 2 — Summary CSV (records)"])
                    with o1:
                        st.caption("Copy each JSON block into the chat to generate that employee's payslip.")
                        for p in payloads:
                            with st.expander(f"📋 {p['employee']} — {p['period']} · "
                                             f"total after MPF HK$ {p['calculation']['total_after_mpf']:,.2f}"):
                                st.code(json.dumps(p, indent=2, ensure_ascii=False), language="json")
                    with o2:
                        st.caption("One row per employee-month — keep for your ledger or Excel.")
                        csv_block = pd.DataFrame(summaries)[SUMMARY_COLS].to_csv(index=False).encode("utf-8-sig")
                        st.download_button("⬇ Download summary CSV block", csv_block,
                                           "monthly_summary_block.csv", mime="text/csv")

    # ---------------- SINGLE ----------------
    with t_single:
        c1, c2, c3 = st.columns(3)
        ename = c1.selectbox("Employee", sorted(emp["name"].unique().tolist()))
        y = int(c2.number_input("Year", 2024, 2040, DEF_Y))
        m = int(c3.selectbox("Month", list(range(1, 13)), index=DEF_M - 1))
        e = get_employee_version(emp, ename, y, m)
        if e is None:
            st.error(f"No employee version found for {ename}.")
            st.stop()
        st.caption(
            f"Using version effective {str(e.get('entry_start', ''))[:10] or '?'} → "
            f"{str(e.get('entry_end', ''))[:10] or 'current'}"
            + (f" · {e.get('change_reason')}" if e.get("change_reason") else ""))
        if e.get("package_type") == "hybrid_prestart":
            st.info("hybrid_prestart: full basic is paid and the month's shortfall is carried into "
                    "outstanding hours; put any pre-start hourly portion in Other adjustment if needed.")

        st.markdown("**Monthly inputs**")
        sy2, sm2 = shift_month(y, m, -2)
        _sv, _mv, _pv, _sf = get_channel_sales(sy2, sm2)
        st.info(f"Commission auto-pulls channel sales (T-2): {calendar.month_name[sm2]} {sy2} — "
                f"SEVENS {hkd(_sv)} / MEAT DEVIL {hkd(_mv)} / PATTY {hkd(_pv)}"
                + ("" if _sf else "  ·  ⚠️ not recorded yet — commission will be 0"))
        i1, i2, i3 = st.columns(3)
        ot_hours = i1.number_input("OT hours (× ot_rate if set, else hour value)", min_value=0.0, step=0.5)
        shortfall_hours = i2.number_input("Shortfall hours to deduct (× rate)", min_value=0.0, step=0.25)
        other_adjust = i3.number_input("Other adjustment (±HKD)", step=50.0)
        j1, j2 = st.columns(2)
        ex_sub_name = j1.text_input("Extra subsidy label (optional)")
        ex_sub_amt = j2.number_input("Extra subsidy (HKD)", min_value=0.0, step=50.0)
        j3, j4 = st.columns(2)
        ex_ded_name = j3.text_input("Extra deduction label (optional)")
        ex_ded_amt = j4.number_input("Extra deduction (HKD)", min_value=0.0, step=50.0)

        summary, payload = compute_monthly(
            e, y, m,
            ot_hours=ot_hours, shortfall_hours=shortfall_hours, other_adjust=other_adjust,
            extra_subsidy_name=ex_sub_name, extra_subsidy_amount=ex_sub_amt,
            extra_deduction_name=ex_ded_name, extra_deduction_amount=ex_ded_amt,
            outstanding_before=ledger_before(summ_ledger, e, ename, y, m))
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Expected hours", summary["expected_hours"])
        k2.metric("Actual rounded hours", summary["actual_hours"])
        k3.metric("Month shortfall (h)", summary["shortfall"])
        k4.metric("Outstanding after", summary["outstanding_after"])

        st.markdown("**Payslip summary**")
        show_df(payslip_table(summary))

        oa, ob = st.columns(2)
        oa.markdown("**Grok JSON**")
        oa.code(json.dumps(payload, indent=2, ensure_ascii=False), language="json")
        ob.markdown("**Summary row (CSV)**")
        ob.code(pd.DataFrame([summary])[SUMMARY_COLS].to_csv(index=False), language="text")

        if st.button("💾 Save this month to summary", type="primary"):
            save_months([summary])
            st.success("Saved. (Re-saving the same month replaces it — safe to repeat.)")

        a = month_attendance(ename, y, m)
        if not a.empty:
            with st.expander("Daily breakdown"):
                show_df(a[["date", "start", "end", "break_hours", "raw_hours", "rounded_hours", "notes"]])

# =========================== 4. CONTINUOUS CONTRACT ===========================
elif pg == "cc":
    st.header("📈 Continuous-Contract Helper")
    st.caption("Weekly hours vs the 17h threshold and rolling 4-week totals vs 68h — "
               "evidence aid only, confirm status under the EO.")
    emp = load_df("employees", EMP_COLS)
    if emp.empty:
        st.warning("Add employees first.")
        st.stop()
    ename = st.selectbox("Employee", sorted(emp["name"].unique().tolist()))
    att = load_df("attendance", ATT_COLS)
    att = att[att["employee"] == ename].copy()
    att["date"] = pd.to_datetime(att["date"], errors="coerce").dt.date
    att = att.dropna(subset=["date"]).sort_values("date")
    if att.empty:
        st.info("No attendance records for this employee.")
        st.stop()
    start_mon = att["date"].min() - timedelta(days=att["date"].min().weekday())
    end_mon = att["date"].max() + timedelta(days=(7 - att["date"].max().weekday()) % 7 or 7)
    weeks, cur = [], start_mon
    while cur <= end_mon:
        nxt = cur + timedelta(days=7)
        wk = att[(att["date"] >= cur) & (att["date"] < nxt)]["rounded_hours"].sum()
        weeks.append({"week_start": str(cur), "hours": round(wk, 2)})
        cur = nxt
    wdf = pd.DataFrame(weeks)
    wdf[">=17h"] = wdf["hours"] >= 17
    wdf["4 consecutive wks >=17h"] = wdf[">=17h"].astype(int).rolling(4, min_periods=4).sum() == 4
    wdf["rolling_4wk"] = wdf["hours"].rolling(4, min_periods=4).sum().round(2)
    wdf["4wk>=68h"] = wdf["rolling_4wk"] >= 68

    k1, k2, k3 = st.columns(3)
    k1.metric("Weeks ≥ 17h", int(wdf[">=17h"].sum()))
    k2.metric("Rolling 4-wk windows ≥ 68h", int(wdf["4wk>=68h"].sum()))
    k3.metric("First record", str(att["date"].min()))
    if wdf["4 consecutive wks >=17h"].any() or wdf["4wk>=68h"].any():
        st.success("Evidence suggests continuous-contract conditions may be met — review under the EO.")
    else:
        st.info("Conditions not yet evidenced by attendance data.")

    st.bar_chart(wdf.set_index("week_start")[["hours"]])
    st.caption("Bars = rounded hours per week (Mon–Sun). Tests: 4 consecutive weeks ≥ 17h, "
               "or any rolling 4-week total ≥ 68h.")
    show_df(wdf)

# =========================== 5. EXPORT / PAYSLIP BLOCK ===========================
elif pg == "exp":
    st.header("📤 Export / Payslip Block")
    st.caption("One employee's payload — same output as batch mode.")
    emp = load_df("employees", EMP_COLS)
    if emp.empty:
        st.warning("Add employees first.")
        st.stop()
    summ_ledger = load_df("monthly_summary", SUMMARY_COLS)
    ename = st.selectbox("Employee", sorted(emp["name"].unique().tolist()))
    c1, c2 = st.columns(2)
    y = int(c1.number_input("Year", 2024, 2040, DEF_Y, key="expy"))
    m = int(c2.selectbox("Month", list(range(1, 13)), index=DEF_M - 1, key="expm"))
    e = get_employee_version(emp, ename, y, m)
    if e is None:
        st.error(f"No employee version found for {ename}.")
        st.stop()
    st.caption(f"Using version effective {str(e.get('entry_start', ''))[:10] or '?'} → "
               f"{str(e.get('entry_end', ''))[:10] or 'current'}")
    st.markdown("**Monthly inputs**")
    xsy, xsm = shift_month(y, m, -2)
    _xsv, _xmv, _xpv, _xsf = get_channel_sales(xsy, xsm)
    st.info(f"Commission auto-pulls channel sales (T-2): {calendar.month_name[xsm]} {xsy} — "
            f"SEVENS {hkd(_xsv)} / MEAT DEVIL {hkd(_xmv)} / PATTY {hkd(_xpv)}"
            + ("" if _xsf else "  ·  ⚠️ not recorded yet — commission will be 0"))
    xi1, xi2, xi3 = st.columns(3)
    x_ot = xi1.number_input("OT hours (× ot_rate if set)", min_value=0.0, step=0.5, key="xot")
    x_sh = xi2.number_input("Shortfall hours to deduct", min_value=0.0, step=0.25, key="xsh")
    x_adj = xi3.number_input("Other adjustment (HKD)", step=50.0, key="xadj")
    xj1, xj2 = st.columns(2)
    x_sub_name = xj1.text_input("Extra subsidy label (optional)", key="xsublbl")
    x_sub_amt = xj2.number_input("Extra subsidy (HKD)", min_value=0.0, step=50.0, key="xsub")
    xj3, xj4 = st.columns(2)
    x_ded_name = xj3.text_input("Extra deduction label (optional)", key="xdedlbl")
    x_ded_amt = xj4.number_input("Extra deduction (HKD)", min_value=0.0, step=50.0, key="xded")
    summary, payload = compute_monthly(
        e, y, m,
        ot_hours=x_ot, shortfall_hours=x_sh, other_adjust=x_adj,
        extra_subsidy_name=x_sub_name, extra_subsidy_amount=x_sub_amt,
        extra_deduction_name=x_ded_name, extra_deduction_amount=x_ded_amt,
        outstanding_before=ledger_before(summ_ledger, e, ename, y, m))
    tj, tc = st.tabs(["JSON payload (paste into Grok)", "Summary row (CSV)"])
    with tj:
        js = json.dumps(payload, indent=2, ensure_ascii=False)
        st.code(js, language="json")
        st.download_button("⬇ Download JSON", js.encode("utf-8"), f"payslip_{ename}_{y}-{m:02d}.json",
                           mime="application/json")
    with tc:
        row_csv = pd.DataFrame([summary])[SUMMARY_COLS].to_csv(index=False)
        st.code(row_csv, language="text")
        st.download_button("⬇ Download CSV row", row_csv.encode("utf-8-sig"),
                           f"summary_{ename}_{y}-{m:02d}.csv", mime="text/csv")

# =========================== 6. MONTHLY SUMMARY ===========================
elif pg == "sum":
    st.header("📒 Monthly Summary — running ledger")
    st.caption("One row per employee per month. Re-saving a month replaces its row — no duplicates.")
    summ = load_df("monthly_summary", SUMMARY_COLS)
    if summ.empty:
        st.info("Nothing saved yet — run a calculation on page 3 and click Save.")
    else:
        summ = summ.sort_values(["year", "month", "employee"])
        f1, f2, f3 = st.columns(3)
        yr = f1.selectbox("Year", ["(all)"] + sorted(summ["year"].unique().tolist()))
        mo = f2.selectbox("Month", ["(all)"] + sorted(summ["month"].unique().tolist()))
        em = f3.selectbox("Employee", ["(all)"] + sorted(summ["employee"].unique().tolist()))
        view = summ.copy()
        if yr != "(all)":
            view = view[view["year"] == int(yr)]
        if mo != "(all)":
            view = view[view["month"] == int(mo)]
        if em != "(all)":
            view = view[view["employee"] == em]
        tot = view[["total_before_mpf", "total_after_mpf"]].sum()
        k1, k2, k3 = st.columns(3)
        k1.metric("Rows (filtered)", len(view))
        k2.metric("Total before MPF", f"HK$ {tot['total_before_mpf']:,.2f}")
        k3.metric("Total after MPF", f"HK$ {tot['total_after_mpf']:,.2f}")
        show_df(view)
        st.download_button("⬇ Download full summary (CSV)", summ.to_csv(index=False).encode("utf-8-sig"),
                           "monthly_summary_all.csv", mime="text/csv")

        st.divider()
        st.markdown("**Delete saved month(s)** — if it was the employee's latest month, "
                    "their outstanding-hours balance is rolled back.")
        labels = {f"{r.employee} | {int(r.year)}-{int(r.month):02d}": (r.employee, int(r.year), int(r.month))
                  for r in view.itertuples()}
        pick = st.multiselect("Rows to delete", list(labels))
        d1, d2 = st.columns([1, 2])
        sure = d1.checkbox("I'm sure", key=f"sum_del_ok_{st.session_state.get('sum_n', 0)}")
        if d2.button("🗑️ Delete selected month(s)", disabled=not pick):
            if not sure:
                st.warning("Tick 'I'm sure' first.")
            else:
                delete_months([labels[p] for p in pick])
                st.session_state["sum_n"] = st.session_state.get("sum_n", 0) + 1
                flash(f"Deleted {len(pick)} ledger row(s).")
                st.rerun()

# =========================== 8. HISTORY ===========================
else:
    st.header("🕑 History Log")
    st.caption("Every change to master data, CSV imports, and outstanding-balance updates.")
    hist = load_df("employee_history", HISTORY_COLS)
    if hist.empty:
        st.info("No history recorded yet.")
    else:
        f1, f2 = st.columns([1, 3])
        filt = f1.selectbox("Filter by employee", ["(all)"] + sorted(hist["employee"].unique().tolist()))
        view = hist if filt == "(all)" else hist[hist["employee"] == filt]
        f2.metric("Entries (filtered)", len(view))
        show_df(view.sort_values("timestamp", ascending=False))