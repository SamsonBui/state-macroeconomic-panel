#!/usr/bin/env python3
"""
build_state_datasets.py
─────────────────────────────────────────────────────────────────────────────
Build all state-level economic datasets for the 27 target states.

Outputs (in ./output):
  • real_gdp_per_capita_by_state.csv
  • personal_income_per_capita_by_state.csv
  • monthly_unemployment_rate_by_state.csv
  • unemployment_rate_by_state.csv
  • homeownership_rate_by_state.csv
  • source_notes.csv
  • data_quality_report.csv

Professor-ready cleaned copies are written to ./cleaned_output/ (originals
in ./output/ are never overwritten by the cleanup step).

HTTP downloads use Scrapling (curl_cffi browser impersonation) to avoid FRED
Akamai blocks on plain Python requests.

Run:
  python3 build_state_datasets.py
─────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import calendar
import logging
import os
import random
import re
import sys
import time
from io import BytesIO, StringIO
from pathlib import Path

import pandas as pd

# Quiet Scrapling per-request INFO logs (they look like the script is stuck).
logging.basicConfig(level=logging.WARNING)
for _logger in ("scrapling", "scrapling.core.utils"):
    logging.getLogger(_logger).setLevel(logging.WARNING)

# Scrapling — curl_cffi-based fetcher that bypasses FRED TLS fingerprint blocks
SCRAPLING_ROOT = Path("/Users/samsonbui/Documents/Github-2026/Scrapling")
if str(SCRAPLING_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRAPLING_ROOT))

from scrapling.fetchers import FetcherSession  # noqa: E402

# ── Configuration ─────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
CLEANED_OUTPUT_DIR = SCRIPT_DIR / "cleaned_output"

START_DATE = "2010-01-01"
PANEL_START_YEAR = 2010
PANEL_END_YEAR = 2025
PANEL_END_QUARTER = 4
PANEL_END_MONTH = 12

HOMEOWNERSHIP_FILE = SCRIPT_DIR / "tab3_state05_2026_hmr.xlsx"
STATES_FILE = SCRIPT_DIR / "List of States.xlsx"

FRED_API_KEY = os.environ.get("FRED_API_KEY", None)
FRED_CACHE_DIR = OUTPUT_DIR / ".fred_cache"
FRED_CACHE_DIR.mkdir(exist_ok=True)

BLS_LANRDERR_URL = "https://www.bls.gov/web/laus/lanrderr.xlsx"
BLS_LANRDERR_CACHE = FRED_CACHE_DIR / "lanrderr.xlsx"

# Only AR and IL publish quarterly {ABBR}OPCI on FRED; others use OTOT/POP.
OPCI_STATES = frozenset({"AR", "IL"})

# Per-request pacing (seconds). Set FRED_PAUSE=0 to disable.
FRED_PAUSE = float(os.environ.get("FRED_PAUSE", "0.05"))

TARGET_ABBRS = frozenset({
    "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "IL", "MA", "MD", "ME",
    "MN", "NH", "NJ", "NM", "NY", "OH", "OR", "PA", "RI", "TX", "UT", "VA",
    "VT", "WA", "WI",
})

STATE_MAP = {
    "AR": "Arkansas",
    "AZ": "Arizona",
    "CA": "California",
    "CO": "Colorado",
    "CT": "Connecticut",
    "DC": "District of Columbia",
    "DE": "Delaware",
    "FL": "Florida",
    "IL": "Illinois",
    "MA": "Massachusetts",
    "MD": "Maryland",
    "ME": "Maine",
    "MN": "Minnesota",
    "NH": "New Hampshire",
    "NJ": "New Jersey",
    "NM": "New Mexico",
    "NY": "New York",
    "OH": "Ohio",
    "OR": "Oregon",
    "PA": "Pennsylvania",
    "RI": "Rhode Island",
    "TX": "Texas",
    "UT": "Utah",
    "VA": "Virginia",
    "VT": "Vermont",
    "WA": "Washington",
    "WI": "Wisconsin",
}

QTR_END_MONTH = {1: 3, 2: 6, 3: 9, 4: 12}

# Track failed FRED series across sections
FAILED_SERIES: list[str] = []


def log(msg: str) -> None:
    print(msg, flush=True)


def _fred_cache_path(series_id: str) -> Path:
    return FRED_CACHE_DIR / f"{series_id}.csv"


def quarter_end_date(year: int, quarter: int) -> str:
    month = QTR_END_MONTH[quarter]
    day = calendar.monthrange(year, month)[1]
    return f"{year}-{month:02d}-{day}"


def month_end_date(year: int, month: int) -> str:
    day = calendar.monthrange(year, month)[1]
    return f"{year}-{month:02d}-{day}"


def add_year_quarter(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["year"] = out["date"].dt.year
    out["quarter"] = out["date"].dt.quarter
    return out


def filter_quarterly_panel(df: pd.DataFrame) -> pd.DataFrame:
    """Keep 2010Q1 through 2025Q4 only."""
    mask = (
        (df["year"] > PANEL_START_YEAR)
        | ((df["year"] == PANEL_START_YEAR) & (df["quarter"] >= 1))
    ) & (
        (df["year"] < PANEL_END_YEAR)
        | ((df["year"] == PANEL_END_YEAR) & (df["quarter"] <= PANEL_END_QUARTER))
    )
    return df.loc[mask].copy()


def filter_monthly_panel(df: pd.DataFrame) -> pd.DataFrame:
    """Keep 2010M1 through 2025M12 only."""
    mask = (
        (df["year"] > PANEL_START_YEAR)
        | ((df["year"] == PANEL_START_YEAR) & (df["month"] >= 1))
    ) & (
        (df["year"] < PANEL_END_YEAR)
        | ((df["year"] == PANEL_END_YEAR) & (df["month"] <= PANEL_END_MONTH))
    )
    return df.loc[mask].copy()


def fill_interior_monthly_gaps(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    """
    Expand to the full 2010M1–2025M12 grid for every target state and linearly
    interpolate interior missing values (e.g. FRED/BLS blank for 2025M10).
    """
    expected: list[dict[str, object]] = []
    for abbr, state_name in STATE_MAP.items():
        for year in range(PANEL_START_YEAR, PANEL_END_YEAR + 1):
            month_end = PANEL_END_MONTH if year == PANEL_END_YEAR else 12
            month_start = 1 if year > PANEL_START_YEAR else 1
            for month in range(month_start, month_end + 1):
                expected.append({
                    "state_abbr": abbr,
                    "state": state_name,
                    "year": year,
                    "month": month,
                })

    grid = pd.DataFrame(expected)
    merged = grid.merge(
        df[["state_abbr", "year", "month", value_col]],
        on=["state_abbr", "year", "month"],
        how="left",
    )

    filled_parts: list[pd.DataFrame] = []
    fill_rows: list[tuple[str, int, int]] = []
    for abbr, grp in merged.groupby("state_abbr", sort=False):
        g = grp.sort_values(["year", "month"]).copy()
        missing = g[value_col].isna()
        g[value_col] = g[value_col].interpolate(method="linear", limit_area="inside")
        newly_filled = missing & g[value_col].notna()
        for row in g.loc[newly_filled, ["year", "month"]].itertuples(index=False):
            fill_rows.append((abbr, int(row.year), int(row.month)))
        filled_parts.append(g)

    out = pd.concat(filled_parts, ignore_index=True)
    if fill_rows:
        sample = ", ".join(f"{a} {y}M{m}" for a, y, m in fill_rows[:3])
        extra = f" (+{len(fill_rows) - 3} more)" if len(fill_rows) > 3 else ""
        log(
            f"  Filled {len(fill_rows)} missing monthly unemployment value(s) via "
            f"linear interpolation (FRED CSV blank: {sample}{extra})."
        )
    if out[value_col].isna().any():
        still_missing = out[out[value_col].isna()][["state_abbr", "year", "month"]]
        raise RuntimeError(
            "Unemployment panel still has gaps after interpolation:\n"
            f"{still_missing.head(10).to_string(index=False)}"
        )
    return out


def finalize_quarterly(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    """Filter target states, panel range, dedupe, quarter-end dates."""
    out = df[df["state_abbr"].isin(TARGET_ABBRS)].copy()
    out = filter_quarterly_panel(out)
    out["date"] = out.apply(
        lambda r: quarter_end_date(int(r["year"]), int(r["quarter"])), axis=1
    )
    out = (
        out.sort_values(["state_abbr", "year", "quarter"])
        .drop_duplicates(subset=["state_abbr", "date"], keep="last")
        .reset_index(drop=True)
    )
    if out[value_col].isna().any():
        pass  # keep blanks per requirements
    return out


def parse_ci_half_width(ci) -> float | None:
    """Parse BLS 'low – high' 90% confidence interval text to half-width (pp)."""
    if ci is None or (isinstance(ci, float) and pd.isna(ci)):
        return None
    text = str(ci).strip()
    match = re.match(r"([\d.]+)\s*[–\-]\s*([\d.]+)", text)
    if not match:
        return None
    low, high = float(match.group(1)), float(match.group(2))
    return (high - low) / 2


def fetch_bls_unemployment_rate_moe(session: FetcherSession) -> pd.DataFrame:
    """
    Latest BLS LAUS model-based 90% confidence interval half-width by state.

    Source: lanrderr.xlsx (current release month). Used as reference to scale
    margin of error across the historical panel (see attach_scaled_unemployment_moe).
    """
    import openpyxl

    if BLS_LANRDERR_CACHE.exists():
        wb = openpyxl.load_workbook(BLS_LANRDERR_CACHE, data_only=True)
    else:
        resp = session.get(BLS_LANRDERR_URL, timeout=30)
        BLS_LANRDERR_CACHE.write_bytes(resp.body)
        wb = openpyxl.load_workbook(BytesIO(resp.body), data_only=True)

    ws = wb.active
    name_to_abbr = {v: k for k, v in STATE_MAP.items()}
    rows: list[dict] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] not in name_to_abbr:
            continue
        abbr = name_to_abbr[row[0]]
        if abbr not in TARGET_ABBRS:
            continue
        rate, ci = row[1], row[2]
        half = parse_ci_half_width(ci)
        if half is None or rate is None or float(rate) == 0:
            continue
        rows.append({"state_abbr": abbr, "ref_rate": float(rate), "ref_moe": half})
    wb.close()

    if len(rows) < len(TARGET_ABBRS):
        log(
            f"  WARNING: BLS unemployment MOE matched {len(rows)}/"
            f"{len(TARGET_ABBRS)} target states"
        )
    return pd.DataFrame(rows)


def attach_scaled_unemployment_moe(
    df: pd.DataFrame,
    rate_col: str,
    moe_ref: pd.DataFrame,
) -> pd.DataFrame:
    """Scale latest BLS 90% MOE to each row's unemployment rate (percentage points)."""
    out = df.merge(moe_ref, on="state_abbr", how="left")
    out["margin_of_error"] = out[rate_col] * (out["ref_moe"] / out["ref_rate"])
    return out.drop(columns=["ref_rate", "ref_moe"])


def finalize_monthly(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    out = df[df["state_abbr"].isin(TARGET_ABBRS)].copy()
    out = filter_monthly_panel(out)
    out["date"] = out.apply(
        lambda r: month_end_date(int(r["year"]), int(r["month"])), axis=1
    )
    out = (
        out.sort_values(["state_abbr", "year", "month"])
        .drop_duplicates(subset=["state_abbr", "date"], keep="last")
        .reset_index(drop=True)
    )
    return out


# ── FRED fetcher ──────────────────────────────────────────────────────────────
def _parse_csv_response(resp_text: str, series_id: str, start: str) -> pd.DataFrame:
    df = pd.read_csv(StringIO(resp_text))
    if df.shape[1] < 2:
        raise ValueError(f"Unexpected FRED response for {series_id}: {df.head()}")
    df.columns = ["date", "value"]
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna(subset=["date"]).copy()
    df = df[df["date"] >= pd.Timestamp(start)].reset_index(drop=True)
    if df.empty:
        raise ValueError(f"No data returned for {series_id} after {start}")
    return df


def _fetch_fred_csv(
    session: FetcherSession,
    series_id: str,
    start: str,
    timeout: int = 20,
) -> pd.DataFrame:
    cache_path = _fred_cache_path(series_id)
    if cache_path.exists():
        return _parse_csv_response(cache_path.read_text(encoding="utf-8"), series_id, start)

    url = (
        f"https://fred.stlouisfed.org/graph/fredgraph.csv"
        f"?id={series_id}"
        f"&cosd={start}"
    )
    resp = session.get(url, timeout=timeout, stealthy_headers=True)
    if resp.status == 404:
        raise RuntimeError(f"Series not found (404): {series_id}")
    if resp.status != 200:
        raise RuntimeError(f"HTTP {resp.status} for {series_id}")
    text = resp.body.decode("utf-8")
    if not text.strip():
        raise ValueError(f"Empty response body for {series_id}")
    cache_path.write_text(text, encoding="utf-8")
    return _parse_csv_response(text, series_id, start)


def _fetch_fred_api(series_id: str, start: str, timeout: int = 30) -> pd.DataFrame:
    if not FRED_API_KEY:
        raise RuntimeError("No FRED_API_KEY set — cannot use API fallback")
    url = (
        "https://api.stlouisfed.org/fred/series/observations"
        f"?series_id={series_id}"
        f"&observation_start={start}"
        f"&file_type=json"
        f"&api_key={FRED_API_KEY}"
    )
    with FetcherSession(impersonate="chrome", stealthy_headers=True, timeout=timeout) as session:
        resp = session.get(url)
    if resp.status != 200:
        raise RuntimeError(f"FRED API HTTP {resp.status} for {series_id}")
    obs = resp.json().get("observations", [])
    if not obs:
        raise ValueError(f"No observations from API for {series_id}")
    df = pd.DataFrame(obs)[["date", "value"]]
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna(subset=["date"]).copy()
    df = df[df["date"] >= pd.Timestamp(start)].reset_index(drop=True)
    if df.empty:
        raise ValueError(f"No data from API for {series_id} after {start}")
    return df


def fetch_fred(
    session: FetcherSession,
    series_id: str,
    start: str = START_DATE,
    max_retries: int = 2,
) -> pd.DataFrame:
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            return _fetch_fred_csv(session, series_id, start)
        except Exception as e:
            last_error = e
            if "404" in str(e) or "not found" in str(e).lower():
                break
            if attempt < max_retries:
                wait = min(2 ** attempt + random.uniform(0, 1), 8)
                log(f"     retry {attempt}/{max_retries} for {series_id}: {e} ({wait:.1f}s)")
                time.sleep(wait)

    if FRED_API_KEY:
        log(f"     CSV failed; trying FRED API for {series_id} …")
        try:
            return _fetch_fred_api(series_id, start)
        except Exception as api_err:
            print(f"     API fallback also failed for {series_id}: {api_err}")
            last_error = api_err

    raise RuntimeError(f"Failed to download {series_id} after {max_retries} attempts: {last_error}")


def fetch_state_series(
    session: FetcherSession,
    suffix: str,
    label: str,
    pause: float | None = None,
) -> tuple[pd.DataFrame, list[str]]:
    """Download {ABBR}{suffix} for every target state."""
    if pause is None:
        pause = FRED_PAUSE
    frames: list[pd.DataFrame] = []
    failed: list[str] = []
    total = len(STATE_MAP)
    log(f"\n{'=' * 65}")
    log(label)
    log(f"Series pattern : {{ABBR}}{suffix}")
    log(f"States         : {total}")
    log(f"{'=' * 65}")

    for i, (abbr, state_name) in enumerate(STATE_MAP.items(), start=1):
        series_id = f"{abbr}{suffix}"
        try:
            df = fetch_fred(session, series_id, start=START_DATE)
            df["state_abbr"] = abbr
            df["state"] = state_name
            frames.append(df)
            last = df["date"].max().strftime("%Y-%m")
            cached = " (cache)" if _fred_cache_path(series_id).exists() else ""
            log(f"  [{i:2d}/{total}] ✓  {abbr:2s}  ({series_id})  {len(df)} obs  through {last}{cached}")
        except Exception as exc:
            log(f"  [{i:2d}/{total}] ✗  {abbr:2s}  ({series_id})  ERROR: {exc}")
            failed.append(series_id)
            FAILED_SERIES.append(series_id)
        if pause > 0:
            time.sleep(pause)

    if not frames:
        raise RuntimeError(f"No {suffix} data downloaded.")
    return pd.concat(frames, ignore_index=True), failed


# ── Dataset builders ──────────────────────────────────────────────────────────
def try_fetch_fred(
    session: FetcherSession,
    series_id: str,
    start: str = START_DATE,
) -> pd.DataFrame | None:
    """Single-attempt fetch; returns None if the series does not exist."""
    try:
        return _fetch_fred_csv(session, series_id, start)
    except Exception as exc:
        if "404" in str(exc) or "not found" in str(exc).lower():
            return None
        raise


def population_from_fred(pop_raw: pd.DataFrame) -> pd.DataFrame:
    pop = pop_raw.copy()
    pop["year"] = pop["date"].dt.year
    pop = pop.rename(columns={"value": "pop_thousands"})
    pop["population"] = pop["pop_thousands"] * 1000
    pop_annual = (
        pop.sort_values(["state_abbr", "year", "date"])
        .groupby(["state_abbr", "year"], as_index=False)
        .agg(population=("population", "last"))
    )
    pop_annual = pop_annual.sort_values(["state_abbr", "year"])
    pop_annual["population"] = pop_annual.groupby("state_abbr")["population"].ffill()
    return pop_annual


def build_real_gdp_per_capita(
    session: FetcherSession,
    pop_annual: pd.DataFrame,
) -> pd.DataFrame:
    gdp_raw, gdp_failed = fetch_state_series(
        session, "RQGSP", "Real GDP by state (FRED {ABBR}RQGSP)"
    )

    gdp = add_year_quarter(gdp_raw)
    gdp = gdp.rename(columns={"value": "real_gdp_millions_chained_2017_dollars"})

    merged = gdp.merge(pop_annual, on=["state_abbr", "year"], how="left")
    merged["real_gdp_per_capita"] = (
        merged["real_gdp_millions_chained_2017_dollars"] * 1_000_000 / merged["population"]
    )

    out = merged[
        [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "real_gdp_millions_chained_2017_dollars",
            "population",
            "real_gdp_per_capita",
        ]
    ].copy()
    out = finalize_quarterly(out, "real_gdp_per_capita")
    out["margin_of_error"] = pd.NA
    out = out[
        [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "date",
            "real_gdp_per_capita",
            "margin_of_error",
            "real_gdp_millions_chained_2017_dollars",
            "population",
        ]
    ]
    if gdp_failed:
        log(f"  WARNING: RQGSP failed series: {gdp_failed}")
    return out


def build_personal_income_per_capita(
    session: FetcherSession,
    pop_annual: pd.DataFrame,
) -> pd.DataFrame:
    """
    Prefer FRED {ABBR}OPCI (quarterly per capita). Only AR and IL expose OPCI on
    FRED; other states fall back to {ABBR}OTOT / annual population.
    """
    frames: list[pd.DataFrame] = []
    opc_states: list[str] = []
    otot_states: list[str] = []
    failed: list[str] = []

    total = len(STATE_MAP)
    log(f"\n{'=' * 65}")
    log("Personal income per capita (FRED {ABBR}OPCI or OTOT/POP)")
    log(f"States         : {total}  (OPCI direct: {', '.join(sorted(OPCI_STATES))})")
    log(f"{'=' * 65}")

    for i, (abbr, state_name) in enumerate(STATE_MAP.items(), start=1):
        if abbr in OPCI_STATES:
            opc_id = f"{abbr}OPCI"
            opc_df = try_fetch_fred(session, opc_id)
            if opc_df is not None and len(opc_df) >= 4:
                opc_df["state_abbr"] = abbr
                opc_df["state"] = state_name
                opc_df = add_year_quarter(opc_df)
                opc_df = opc_df.rename(columns={"value": "personal_income_per_capita"})
                frames.append(opc_df)
                opc_states.append(abbr)
                last = opc_df["date"].max().strftime("%Y-%m")
                log(f"  [{i:2d}/{total}] ✓  {abbr:2s}  ({opc_id})  direct OPCI  through {last}")
                if FRED_PAUSE > 0:
                    time.sleep(FRED_PAUSE)
                continue

        otot_id = f"{abbr}OTOT"
        try:
            otot_df = fetch_fred(session, otot_id, start=START_DATE)
            otot_df["state_abbr"] = abbr
            otot_df["state"] = state_name
            otot_df = add_year_quarter(otot_df)
            otot_df = otot_df.rename(columns={"value": "personal_income_millions"})
            merged = otot_df.merge(pop_annual, on=["state_abbr", "year"], how="left")
            merged["personal_income_per_capita"] = (
                merged["personal_income_millions"] * 1_000_000 / merged["population"]
            )
            frames.append(merged)
            otot_states.append(abbr)
            last = merged["date"].max().strftime("%Y-%m")
            log(f"  [{i:2d}/{total}] ✓  {abbr:2s}  ({otot_id} + POP)  computed  through {last}")
        except Exception as exc:
            log(f"  [{i:2d}/{total}] ✗  {abbr:2s}  ERROR: {exc}")
            failed.append(abbr)
            FAILED_SERIES.append(otot_id)
        if FRED_PAUSE > 0:
            time.sleep(FRED_PAUSE)

    if not frames:
        raise RuntimeError("No personal income data downloaded.")

    if opc_states:
        log(f"  OPCI direct: {len(opc_states)} states ({', '.join(opc_states)})")
    if otot_states:
        log(f"  OTOT/POP computed: {len(otot_states)} states")
    if failed:
        log(f"  WARNING: personal income failed: {failed}")

    raw = pd.concat(frames, ignore_index=True)
    out = raw[
        ["state", "state_abbr", "year", "quarter", "personal_income_per_capita"]
    ].copy()
    out = finalize_quarterly(out, "personal_income_per_capita")
    out["margin_of_error"] = pd.NA
    return out[
        [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "date",
            "personal_income_per_capita",
            "margin_of_error",
        ]
    ]


def build_unemployment(session: FetcherSession) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Monthly + quarterly unemployment; keeps complete quarters only."""
    all_monthly: list[pd.DataFrame] = []
    failed: list[str] = []

    total = len(STATE_MAP)
    log(f"\n{'=' * 65}")
    log("Unemployment rate by state (FRED BLS LAUS)")
    log("Series pattern : {ABBR}UR   e.g. ARUR, WAUR, WIUR")
    log(f"States         : {total}")
    log(f"{'=' * 65}")

    for i, (abbr, state_name) in enumerate(STATE_MAP.items(), start=1):
        series_id = f"{abbr}UR"
        try:
            df = fetch_fred(session, series_id, start=START_DATE)
            if df.empty:
                raise ValueError(f"No data returned after {START_DATE}")
            df["state"] = state_name
            df["state_abbr"] = abbr
            all_monthly.append(df)
            last = df["date"].max().strftime("%Y-%m")
            log(f"  [{i:2d}/{total}] ✓  {abbr:2s}  ({series_id})  {len(df)} months  through {last}")
        except Exception as exc:
            log(f"  [{i:2d}/{total}] ✗  {abbr:2s}  ({series_id})  ERROR: {exc}")
            failed.append(abbr)
            FAILED_SERIES.append(series_id)
        if FRED_PAUSE > 0:
            time.sleep(FRED_PAUSE)

    if not all_monthly:
        raise RuntimeError("No unemployment data downloaded.")

    moe_ref = fetch_bls_unemployment_rate_moe(session)
    log(f"  BLS unemployment MOE reference: {len(moe_ref)} states (lanrderr.xlsx)")

    monthly = pd.concat(all_monthly, ignore_index=True)
    monthly["year"] = monthly["date"].dt.year
    monthly["month"] = monthly["date"].dt.month
    monthly = monthly.rename(columns={"value": "unemployment_rate"})
    monthly = monthly[monthly["state_abbr"].isin(TARGET_ABBRS)].copy()

    monthly = fill_interior_monthly_gaps(
        monthly[["state", "state_abbr", "year", "month", "unemployment_rate"]],
        "unemployment_rate",
    )
    monthly = attach_scaled_unemployment_moe(monthly, "unemployment_rate", moe_ref)
    monthly["margin_of_error"] = monthly["margin_of_error"].round(4)
    log(
        f"  Monthly panel: {len(monthly)} rows "
        f"({monthly['state_abbr'].nunique()} states × "
        f"{len(monthly) // monthly['state_abbr'].nunique()} months)"
    )

    monthly_out = finalize_monthly(monthly, "unemployment_rate")

    # Aggregate quarterly from the filled monthly panel (complete quarters only).
    panel_monthly = monthly.copy()
    panel_monthly["quarter"] = panel_monthly["month"].apply(lambda m: (m - 1) // 3 + 1)
    quarterly = (
        panel_monthly.groupby(["state", "state_abbr", "year", "quarter"], as_index=False)
        .agg(
            unemployment_rate=("unemployment_rate", "mean"),
            margin_of_error=("margin_of_error", "mean"),
        )
        .round(4)
    )
    month_counts = (
        panel_monthly.groupby(["state_abbr", "year", "quarter"])["month"]
        .count()
        .reset_index()
        .rename(columns={"month": "n_months"})
    )
    quarterly = quarterly.merge(
        month_counts, on=["state_abbr", "year", "quarter"], how="left"
    )
    quarterly = quarterly[quarterly["n_months"] == 3].drop(columns=["n_months"])

    quarterly_out = quarterly[
        ["state", "state_abbr", "year", "quarter", "unemployment_rate", "margin_of_error"]
    ].copy()
    quarterly_out = finalize_quarterly(quarterly_out, "unemployment_rate")

    return monthly_out, quarterly_out, failed


def build_homeownership() -> pd.DataFrame:
    from filter_homeownership_to_tidy import parse_homeownership

    print(f"\n{'=' * 65}")
    print("Homeownership rate by state (Census HVS)")
    print(f"Source file    : {HOMEOWNERSHIP_FILE}")
    print(f"{'=' * 65}")

    if not HOMEOWNERSHIP_FILE.exists():
        raise FileNotFoundError(f"Homeownership file not found: {HOMEOWNERSHIP_FILE}")

    states_path = STATES_FILE if STATES_FILE.exists() else HOMEOWNERSHIP_FILE
    records = parse_homeownership(HOMEOWNERSHIP_FILE, states_path)
    if not records:
        raise ValueError(f"No homeownership records parsed from {HOMEOWNERSHIP_FILE}")

    home = pd.DataFrame(records)
    home = home.rename(columns={
        "state_name": "state",
        "homeownership_rate_pct": "homeownership_rate",
    })
    home["quarter"] = home["quarter"].str.replace("Q", "", regex=False).astype(int)
    home = home[home["state_abbr"].isin(TARGET_ABBRS)].copy()
    home = filter_quarterly_panel(home)
    home["date"] = home.apply(
        lambda r: quarter_end_date(int(r["year"]), int(r["quarter"])), axis=1
    )
    out = home[
        ["state", "state_abbr", "year", "quarter", "date", "homeownership_rate", "margin_of_error"]
    ].sort_values(["state_abbr", "year", "quarter"])
    out = out.drop_duplicates(subset=["state_abbr", "date"], keep="last").reset_index(drop=True)
    print(f"  ✓  {len(out)} rows, {out['state_abbr'].nunique()} states, "
          f"{out['year'].min()}Q{out.loc[out['year']==out['year'].min(), 'quarter'].min()}"
          f" – {out['year'].max()}Q{out.loc[out['year']==out['year'].max(), 'quarter'].max()}")
    return out


def build_source_notes() -> pd.DataFrame:
    notes = [
        {
            "variable": "real_gdp_per_capita",
            "source": "U.S. Bureau of Economic Analysis via FRED ({ABBR}RQGSP, {ABBR}POP)",
            "frequency": "quarterly",
            "seasonal_adjustment": "seasonally adjusted annual rate (real GDP)",
            "units": "chained 2017 dollars per person",
            "date_range": "2010Q1-2025Q4",
            "output_file": "real_gdp_per_capita_by_state.csv",
            "notes": (
                "Real GDP in millions of chained 2017 dollars divided by annual "
                "population (thousands × 1000 from FRED POP). Annual population "
                "is merged to all quarters in the same year; latest available "
                "population is forward-filled within each state when missing. "
                "margin_of_error: BEA does not publish sampling margins of error "
                "for state GDP; column is blank."
            ),
        },
        {
            "variable": "personal_income_per_capita",
            "source": "U.S. Bureau of Economic Analysis via FRED ({ABBR}OPCI)",
            "frequency": "quarterly",
            "seasonal_adjustment": "seasonally adjusted annual rate",
            "units": "current dollars per person",
            "date_range": "2010Q1-2025Q4",
            "output_file": "personal_income_per_capita_by_state.csv",
            "notes": (
                "Primary: FRED {ABBR}OPCI where available (AR, IL). "
                "Other states: quarterly total personal income ({ABBR}OTOT, "
                "millions of dollars SAAR) divided by annual population "
                "({ABBR}POP thousands × 1000), with population forward-filled "
                "within each state when missing. "
                "margin_of_error: BEA does not publish sampling margins of error "
                "for state personal income; column is blank."
            ),
        },
        {
            "variable": "monthly_unemployment_rate",
            "source": "U.S. Bureau of Labor Statistics LAUS via FRED ({ABBR}UR)",
            "frequency": "monthly",
            "seasonal_adjustment": "seasonally adjusted",
            "units": "percent",
            "date_range": "2010M1-2025M12",
            "output_file": "monthly_unemployment_rate_by_state.csv",
            "notes": (
                "Statewide unemployment rate via Scrapling/FRED CSV ({ABBR}UR). "
                "2025M10 was blank in the FRED download for all 27 states; "
                "values were linearly interpolated from adjacent months "
                "(September and November 2025). "
                "margin_of_error: 90% confidence interval half-width from BLS "
                "lanrderr.xlsx (latest LAUS release), scaled to each month's "
                "rate as rate × (ref_moe / ref_rate); units are percentage points."
            ),
        },
        {
            "variable": "quarterly_unemployment_rate",
            "source": "U.S. Bureau of Labor Statistics LAUS via FRED ({ABBR}UR)",
            "frequency": "quarterly",
            "seasonal_adjustment": "seasonally adjusted (average of monthly SA rates)",
            "units": "percent",
            "date_range": "2010Q1-2025Q4",
            "output_file": "unemployment_rate_by_state.csv",
            "notes": (
                "Arithmetic mean of three monthly unemployment rates; "
                "complete quarters only. Includes 2025Q4 using interpolated "
                "2025M10 (see monthly unemployment notes). "
                "margin_of_error: mean of the three monthly margins of error "
                "in each quarter (percentage points)."
            ),
        },
        {
            "variable": "homeownership_rate",
            "source": "U.S. Census Bureau Housing Vacancy Survey, Table 3",
            "frequency": "quarterly",
            "seasonal_adjustment": "not seasonally adjusted",
            "units": "percent of occupied housing units owner-occupied",
            "date_range": "2010Q1-2025Q4",
            "output_file": "homeownership_rate_by_state.csv",
            "notes": (
                f"Parsed from local file {HOMEOWNERSHIP_FILE.name}; "
                "2026Q1 excluded from final panel for consistency. "
                "margin_of_error: Census HVS Table 3 published margin of error "
                "(percentage points, 90% confidence level)."
            ),
        },
    ]
    return pd.DataFrame(notes)


def assess_file(
    path: Path,
    value_col: str,
    is_monthly: bool = False,
) -> dict:
    df = pd.read_csv(path)
    states = sorted(df["state_abbr"].unique())
    target = set(TARGET_ABBRS)
    present = set(states)
    missing_states = sorted(target - present)
    extra_states = sorted(present - target)
    dups = int(df.duplicated(subset=["state_abbr", "date"]).sum())
    missing_vals = int(df[value_col].isna().sum()) if value_col in df.columns else 0

    return {
        "file": path.name,
        "rows": len(df),
        "unique_states": len(states),
        "min_date": df["date"].min(),
        "max_date": df["date"].max(),
        "missing_values": missing_vals,
        "duplicate_state_date_rows": dups,
        "all_27_states_present": present == target,
        "missing_states": ",".join(missing_states) if missing_states else "",
        "extra_states": ",".join(extra_states) if extra_states else "",
        "_state_list": states,
        "_is_monthly": is_monthly,
    }


def print_quality_report(reports: list[dict]) -> None:
    print(f"\n{'=' * 65}")
    print("QUALITY CHECKS")
    print(f"{'=' * 65}")
    for r in reports:
        print(f"\n--- {r['file']} ---")
        print(f"  Row count          : {r['rows']}")
        print(f"  Unique states      : {r['unique_states']}")
        print(f"  State list         : {r['_state_list']}")
        print(f"  All 27 present     : {r['all_27_states_present']}")
        print(f"  Date range         : {r['min_date']}  →  {r['max_date']}")
        print(f"  Missing values     : {r['missing_values']}")
        print(f"  Duplicate rows     : {r['duplicate_state_date_rows']}")
        if r["missing_states"]:
            print(f"  Missing states     : {r['missing_states']}")
        if r["extra_states"]:
            print(f"  Extra states       : {r['extra_states']}")


# ── Cleaned output (read output/, write cleaned_output/) ─────────────────────

SOURCE_NOTES_COLUMNS = [
    "variable",
    "source",
    "frequency",
    "seasonal_adjustment",
    "units",
    "date_range",
    "output_file",
    "notes",
]

DATA_FILE_SPECS: dict[str, dict] = {
    "real_gdp_per_capita_by_state.csv": {
        "columns": [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "date",
            "real_gdp_per_capita",
            "margin_of_error",
            "real_gdp_millions_chained_2017_dollars",
            "population",
        ],
        "optional_columns": [
            "real_gdp_millions_chained_2017_dollars",
            "population",
            "margin_of_error",
        ],
        "value_col": "real_gdp_per_capita",
        "numeric_cols": [
            "real_gdp_per_capita",
            "margin_of_error",
            "real_gdp_millions_chained_2017_dollars",
            "population",
        ],
        "is_monthly": False,
        "expected_rows": 1728,
        "sort_cols": ["state_abbr", "year", "quarter"],
    },
    "personal_income_per_capita_by_state.csv": {
        "columns": [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "date",
            "personal_income_per_capita",
            "margin_of_error",
        ],
        "optional_columns": ["margin_of_error"],
        "value_col": "personal_income_per_capita",
        "numeric_cols": ["personal_income_per_capita", "margin_of_error"],
        "is_monthly": False,
        "expected_rows": 1728,
        "sort_cols": ["state_abbr", "year", "quarter"],
    },
    "monthly_unemployment_rate_by_state.csv": {
        "columns": [
            "state",
            "state_abbr",
            "year",
            "month",
            "date",
            "unemployment_rate",
            "margin_of_error",
        ],
        "optional_columns": [],
        "value_col": "unemployment_rate",
        "numeric_cols": ["unemployment_rate", "margin_of_error"],
        "is_monthly": True,
        "expected_rows": 5184,
        "sort_cols": ["state_abbr", "year", "month"],
    },
    "unemployment_rate_by_state.csv": {
        "columns": [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "date",
            "unemployment_rate",
            "margin_of_error",
        ],
        "optional_columns": [],
        "value_col": "unemployment_rate",
        "numeric_cols": ["unemployment_rate", "margin_of_error"],
        "is_monthly": False,
        "expected_rows": 1728,
        "sort_cols": ["state_abbr", "year", "quarter"],
    },
    "homeownership_rate_by_state.csv": {
        "columns": [
            "state",
            "state_abbr",
            "year",
            "quarter",
            "date",
            "homeownership_rate",
            "margin_of_error",
        ],
        "optional_columns": [],
        "value_col": "homeownership_rate",
        "numeric_cols": ["homeownership_rate", "margin_of_error"],
        "is_monthly": False,
        "expected_rows": 1728,
        "sort_cols": ["state_abbr", "year", "quarter"],
    },
}


def _standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = (
        out.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", "_", regex=True)
    )
    drop_cols = [c for c in out.columns if not c or c.startswith("unnamed")]
    if drop_cols:
        out = out.drop(columns=drop_cols)
    return out


def _strip_string_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.select_dtypes(include=["object", "string"]).columns:
        out[col] = out[col].astype(str).str.strip()
        out[col] = out[col].replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
    return out


def _normalize_state_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "state_abbr" not in out.columns:
        raise ValueError("Missing required column: state_abbr")
    out["state_abbr"] = out["state_abbr"].astype(str).str.strip().str.upper()
    out["state"] = out["state_abbr"].map(STATE_MAP)
    unknown = sorted(out.loc[out["state"].isna(), "state_abbr"].unique())
    if unknown:
        raise ValueError(f"Unknown state abbreviations: {unknown}")
    return out


def _ensure_iso_dates(df: pd.DataFrame, is_monthly: bool) -> pd.DataFrame:
    out = df.copy()
    if "date" in out.columns:
        out["date"] = pd.to_datetime(out["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    elif is_monthly:
        out["date"] = out.apply(
            lambda r: month_end_date(int(r["year"]), int(r["month"])), axis=1
        )
    else:
        out["date"] = out.apply(
            lambda r: quarter_end_date(int(r["year"]), int(r["quarter"])), axis=1
        )
    return out


def _clean_dataframe(df: pd.DataFrame, spec: dict) -> pd.DataFrame:
    out = _standardize_column_names(df)
    out = _strip_string_columns(out)
    out = _normalize_state_fields(out)

    out = out[out["state_abbr"].isin(TARGET_ABBRS)].copy()

    out["year"] = pd.to_numeric(out["year"], errors="coerce").astype("Int64")
    if spec["is_monthly"]:
        out["month"] = pd.to_numeric(out["month"], errors="coerce").astype("Int64")
        out = out.dropna(subset=["year", "month"])
        out["year"] = out["year"].astype(int)
        out["month"] = out["month"].astype(int)
        out = filter_monthly_panel(out)
    else:
        out["quarter"] = pd.to_numeric(out["quarter"], errors="coerce").astype("Int64")
        out = out.dropna(subset=["year", "quarter"])
        out["year"] = out["year"].astype(int)
        out["quarter"] = out["quarter"].astype(int)
        out = filter_quarterly_panel(out)

    out = _ensure_iso_dates(out, is_monthly=spec["is_monthly"])

    for col in spec.get("numeric_cols", []):
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")

    out = (
        out.sort_values(spec["sort_cols"])
        .drop_duplicates(subset=["state_abbr", "date"], keep="last")
        .reset_index(drop=True)
    )

    missing_optional = [c for c in spec.get("optional_columns", []) if c not in out.columns]
    if missing_optional:
        log(f"  WARNING: missing optional columns (omitted): {', '.join(missing_optional)}")

    final_cols = [c for c in spec["columns"] if c in out.columns]
    return out[final_cols]


def _clean_source_notes(df: pd.DataFrame) -> pd.DataFrame:
    out = _standardize_column_names(df)
    out = _strip_string_columns(out)
    final_cols = [c for c in SOURCE_NOTES_COLUMNS if c in out.columns]
    extra_cols = [c for c in out.columns if c not in final_cols]
    return out[final_cols + extra_cols]


def _assess_cleaned_file(path: Path, spec: dict) -> dict:
    df = pd.read_csv(path)
    value_col = spec["value_col"]
    states = sorted(df["state_abbr"].unique())
    target = set(TARGET_ABBRS)
    present = set(states)
    missing_states = sorted(target - present)
    extra_states = sorted(present - target)
    dups = int(df.duplicated(subset=["state_abbr", "date"]).sum())
    missing_vals = int(df[value_col].isna().sum()) if value_col in df.columns else 0
    rows = len(df)
    expected_rows = spec["expected_rows"]

    all_27 = present == target
    no_dups = dups == 0
    no_extra = not extra_states
    row_ok = rows == expected_rows

    status = "ok" if all_27 and no_dups and no_extra and row_ok and not missing_states else "review"

    return {
        "file": path.name,
        "rows": rows,
        "unique_states": len(states),
        "min_date": df["date"].min(),
        "max_date": df["date"].max(),
        "missing_values": missing_vals,
        "duplicate_state_date_rows": dups,
        "all_27_states_present": all_27,
        "missing_states": ",".join(missing_states) if missing_states else "",
        "extra_states": ",".join(extra_states) if extra_states else "",
        "status": status,
    }


def run_cleanup_step(input_dir: Path | None = None, output_dir: Path | None = None) -> None:
    """
    Read assembled CSVs from output/, clean/standardize, write to cleaned_output/.
    Never modifies files in the original output/ folder.
    """
    input_dir = input_dir or OUTPUT_DIR
    output_dir = output_dir or CLEANED_OUTPUT_DIR
    output_dir.mkdir(exist_ok=True)

    log(f"\n{'=' * 65}")
    log("CLEANUP: standardizing CSVs for Stata/R import")
    log(f"Reading from  : {input_dir}")
    log(f"Writing to    : {output_dir}")
    log(f"{'=' * 65}")

    cleaned_reports: list[dict] = []

    for filename, spec in DATA_FILE_SPECS.items():
        src = input_dir / filename
        if not src.exists():
            raise FileNotFoundError(f"Expected input file not found: {src}")
        raw = pd.read_csv(src)
        cleaned = _clean_dataframe(raw, spec)
        dest = output_dir / filename
        cleaned.to_csv(dest, index=False, na_rep="")
        report = _assess_cleaned_file(dest, spec)
        cleaned_reports.append(report)
        log(f"  Cleaned: {filename} ({len(cleaned)} rows)")

    src_notes = input_dir / "source_notes.csv"
    if not src_notes.exists():
        raise FileNotFoundError(f"Expected input file not found: {src_notes}")
    notes_cleaned = _clean_source_notes(pd.read_csv(src_notes))
    notes_cleaned.to_csv(output_dir / "source_notes.csv", index=False, na_rep="")

    report_df = pd.DataFrame(cleaned_reports)
    report_df.to_csv(output_dir / "data_quality_report.csv", index=False, na_rep="")

    log(f"\n{'=' * 65}")
    log(f"Cleaned files written to {output_dir.name}/")
    log(f"{'=' * 65}")
    for r in cleaned_reports:
        log(f"\n  {r['file']}")
        log(f"    rows            : {r['rows']}")
        log(f"    unique states   : {r['unique_states']}")
        log(f"    date range      : {r['min_date']} → {r['max_date']}")
        log(f"    missing values  : {r['missing_values']}")
        log(f"    duplicates      : {r['duplicate_state_date_rows']}")
        log(f"    status          : {r['status']}")


def main() -> None:
    global FAILED_SERIES
    FAILED_SERIES = []
    unemployment_failed: list[str] = []

    log("=" * 65)
    log("Building state economic datasets")
    log(f"Output directory : {OUTPUT_DIR}")
    log(f"Panel range      : {PANEL_START_YEAR} – {PANEL_END_YEAR}")
    log(f"FRED cache       : {FRED_CACHE_DIR}")
    log(f"HTTP client      : Scrapling FetcherSession (curl_cffi / Chrome)")
    log("=" * 65)

    with FetcherSession(
        impersonate="chrome",
        stealthy_headers=True,
        timeout=20,
        retries=2,
        retry_delay=2,
    ) as fred_session:
        pop_raw, pop_failed = fetch_state_series(
            fred_session, "POP", "Annual population by state (FRED {ABBR}POP)"
        )
        pop_annual = population_from_fred(pop_raw)
        if pop_failed:
            print(f"  WARNING: POP failed series: {pop_failed}")

        gdp_out = build_real_gdp_per_capita(fred_session, pop_annual)
        pi_out = build_personal_income_per_capita(fred_session, pop_annual)
        monthly_out, quarterly_out, unemployment_failed = build_unemployment(fred_session)

    home_out = build_homeownership()
    source_notes = build_source_notes()

    outputs = {
        "real_gdp_per_capita_by_state.csv": gdp_out,
        "personal_income_per_capita_by_state.csv": pi_out,
        "monthly_unemployment_rate_by_state.csv": monthly_out,
        "unemployment_rate_by_state.csv": quarterly_out,
        "homeownership_rate_by_state.csv": home_out,
        "source_notes.csv": source_notes,
    }

    print(f"\n{'=' * 65}")
    print("Saving CSV files")
    print(f"{'=' * 65}")
    for name, df in outputs.items():
        path = OUTPUT_DIR / name
        df.to_csv(path, index=False)
        print(f"  Saved: {path}")

    value_cols = {
        "real_gdp_per_capita_by_state.csv": "real_gdp_per_capita",
        "personal_income_per_capita_by_state.csv": "personal_income_per_capita",
        "monthly_unemployment_rate_by_state.csv": "unemployment_rate",
        "unemployment_rate_by_state.csv": "unemployment_rate",
        "homeownership_rate_by_state.csv": "homeownership_rate",
    }
    reports = [
        assess_file(
            OUTPUT_DIR / name,
            value_cols[name],
            is_monthly=(name == "monthly_unemployment_rate_by_state.csv"),
        )
        for name in value_cols
    ]
    report_df = pd.DataFrame([
        {k: v for k, v in r.items() if not k.startswith("_")}
        for r in reports
    ])
    report_path = OUTPUT_DIR / "data_quality_report.csv"
    report_df.to_csv(report_path, index=False)
    print(f"  Saved: {report_path}")

    print_quality_report(reports)

    all_ready = all(r["all_27_states_present"] for r in reports)
    no_dups = all(r["duplicate_state_date_rows"] == 0 for r in reports)
    no_extra = all(not r["extra_states"] for r in reports)

    print(f"\n{'=' * 65}")
    print("SUBMISSION STATUS")
    print(f"{'=' * 65}")
    if FAILED_SERIES:
        print(f"  ⚠ Failed FRED series ({len(FAILED_SERIES)}): {FAILED_SERIES}")
    if unemployment_failed:
        print(f"  ⚠ Unemployment download failures: {unemployment_failed}")
    if all_ready and no_dups and no_extra and not FAILED_SERIES:
        print("  ✓ All outputs appear submission-ready for 2010–2025 panel.")
    else:
        issues = []
        if not all_ready:
            issues.append("not all 27 states in every file")
        if not no_dups:
            issues.append("duplicate state-date rows detected")
        if not no_extra:
            issues.append("non-target states present")
        if FAILED_SERIES:
            issues.append("some FRED series failed")
        print(f"  ⚠ Review before submission: {', '.join(issues)}.")

    run_cleanup_step()

    print("\nDone.")


if __name__ == "__main__":
    main()
