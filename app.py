"""
Product Comparison Tool
=======================
Compares the financial outcome of two electricity pricing products:
  - Product A: Weighted average (hourly Belpex × hourly volume)
  - Product B: Arithmetic average (monthly avg Belpex × monthly volume)

Both products use the formula: Y = A·Belpex + B (€/MWh)
Input: 3-column file (Date | Time | Value in kW or kWh), CSV or Excel.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO
from pathlib import Path
import datetime

# ─── Paths ────────────────────────────────────────────────────────────────────
BELPEX_PATH = Path(__file__).parent / "data" / "Day ahead Belgium from 2015.csv"

# ─── Color conventions (consistent across all charts) ───────────────────────
# WA product (Weighted Average) → blue
# AA product (Arithmetic Average) → orange
# These apply to main charts, decomposition bars, and sensitivity curves.
COL_WA   = "#1f77b4"  # blue  — Weighted Average pricing
COL_AA   = "#ff7f0e"  # orange — Arithmetic Average pricing
COL_MIXED = "#9467bd"  # purple — Mixed (baseload click) product

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Product Comparison",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Data helpers ─────────────────────────────────────────────────────────────

@st.cache_data
def load_belpex() -> pd.Series:
    """
    Load hourly Belpex day-ahead prices.
    Index: tz-naive Brussels local timestamps (matches consumption file local time).
    Values: EUR/MWh.
    """
    df = pd.read_csv(BELPEX_PATH)
    # Use local-time column → tz-naive, matches vol series index
    df["ts"] = pd.to_datetime(df["Datetime (Local)"], dayfirst=False).dt.floor("h")
    df["price"] = pd.to_numeric(df["Price (EUR/MWhe)"], errors="coerce")
    # Deduplicate (Oct fall-back: same local hour appears twice)
    return df.drop_duplicates("ts").set_index("ts")["price"]


# ─── Parsing helpers (ported from EAP data_processing / utils) ────────────────

_DT_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%d.%m.%Y %H:%M:%S",
    "%d.%m.%Y %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
]


def _clean_numeric(series: pd.Series) -> pd.Series:
    """Handle both EU (1.234,56) and US (1,234.56) decimal formats."""
    if series.dtype in (float, int, "float64", "int64"):
        return series.astype(float)

    def _parse(val):
        if pd.isna(val):
            return np.nan
        s = str(val).strip()
        if "," in s and "." in s:
            if s.rfind(".") > s.rfind(","):
                s = s.replace(",", "")          # US thousands
            else:
                s = s.replace(".", "").replace(",", ".")  # EU decimal
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return np.nan

    return series.apply(_parse)


def _parse_datetimes(date_col: pd.Series, time_col: pd.Series) -> pd.Series:
    """Combine date+time columns and try multiple format patterns."""
    if pd.api.types.is_datetime64_any_dtype(date_col):
        combined = date_col.dt.strftime("%Y-%m-%d") + " " + time_col.astype(str).str.strip()
    else:
        combined = date_col.astype(str).str.strip() + " " + time_col.astype(str).str.strip()

    for fmt in _DT_FORMATS:
        result = pd.to_datetime(combined, format=fmt, errors="coerce")
        if result.notna().sum() > 0:
            return result

    return pd.to_datetime(combined, dayfirst=True, errors="coerce")


def _detect_granularity(df: pd.DataFrame) -> str:
    diffs = df.index.to_series().diff().dropna()
    if not len(diffs):
        return "unknown"
    med = diffs.median()
    if med == pd.Timedelta("15min"):
        return "15min"
    if med == pd.Timedelta("1h"):
        return "60min"
    return "unknown"


def _drop_dst_phantoms(df: pd.DataFrame) -> tuple:
    """Remove nonexistent local timestamps from spring-forward DST gap."""
    if df.index.tz is not None:
        return df, 0
    try:
        test = df.index.tz_localize("Europe/Brussels", ambiguous=False, nonexistent="NaT")
        mask = test.isna()
        n = int(mask.sum())
        return df[~mask], n
    except Exception:
        return df, 0


def parse_consumption_file(uploaded, unit: str) -> tuple:
    """
    Parse a 3-column consumption file (Date | Time | Value).
    Supports CSV (comma or semicolon) and Excel (.xlsx / .xls).
    unit: 'kW' or 'kWh'
    Returns (hourly_mwh: pd.Series | None, warnings: list[str], info: dict)
    Series index: tz-naive local timestamps floored to hour; values: MWh/h.
    """
    warns = []
    info = {}
    name = uploaded.name
    ext = name.rsplit(".", 1)[-1].lower()
    content = uploaded.read()

    # ── Load raw ──────────────────────────────────────────────────────────────
    try:
        if ext in ("xlsx", "xls"):
            df_raw = pd.read_excel(BytesIO(content), header=None,
                                   engine="openpyxl" if ext == "xlsx" else "xlrd")
        else:
            for sep in (",", ";", "\t"):
                try:
                    df_raw = pd.read_csv(BytesIO(content), sep=sep, header=None,
                                         encoding="utf-8-sig")
                    if df_raw.shape[1] >= 3:
                        break
                except Exception:
                    continue
            else:
                for sep in (",", ";"):
                    try:
                        df_raw = pd.read_csv(BytesIO(content), sep=sep, header=None,
                                             encoding="latin-1")
                        if df_raw.shape[1] >= 3:
                            break
                    except Exception:
                        continue
    except Exception as e:
        return None, [f"Could not load file: {e}"], {}

    if df_raw.shape[1] < 3:
        return None, [
            f"Expected at least 3 columns (Date, Time, Value), found {df_raw.shape[1]}. "
            "Check separator (comma/semicolon) and file format."
        ], {}

    # Drop header rows until we hit a parseable datetime in col 0
    start_row = 0
    for i, val in enumerate(df_raw.iloc[:, 0]):
        test = pd.to_datetime(str(val), dayfirst=True, errors="coerce")
        if pd.notna(test):
            start_row = i
            break
    df_raw = df_raw.iloc[start_row:].reset_index(drop=True)

    # ── Parse datetime ────────────────────────────────────────────────────────
    dt = _parse_datetimes(df_raw.iloc[:, 0], df_raw.iloc[:, 1])
    null_dt = dt.isna().sum()
    if null_dt > len(dt) * 0.5:
        return None, [
            f"{null_dt}/{len(dt)} timestamps could not be parsed. "
            "Verify date format (expected DD.MM.YYYY or YYYY-MM-DD) and time column."
        ], {}
    if null_dt:
        warns.append(f"{null_dt} rows had unparseable timestamps and were dropped.")

    # ── Parse values ──────────────────────────────────────────────────────────
    values = _clean_numeric(df_raw.iloc[:, 2])
    null_v = values.isna().sum()
    if null_v:
        warns.append(f"{null_v} non-numeric values were set to NaN and will be interpolated.")

    df = pd.DataFrame({"value": values.values}, index=pd.DatetimeIndex(dt.values))
    df = df[dt.notna().values].sort_index()  # .values strips integer index → no alignment error

    # ── Duplicates (Oct DST fall-back) ────────────────────────────────────────
    dups = df.index.duplicated().sum()
    if dups:
        warns.append(f"{dups} duplicate timestamps removed (DST fall-back — kept first).")
        df = df[~df.index.duplicated(keep="first")]

    # ── Spring-forward phantoms ───────────────────────────────────────────────
    df, phantoms = _drop_dst_phantoms(df)
    if phantoms:
        warns.append(f"{phantoms} nonexistent timestamps in DST spring-forward gap removed.")

    # ── Granularity ───────────────────────────────────────────────────────────
    gran = _detect_granularity(df)
    info["granularity"] = gran
    if gran == "unknown":
        warns.append(
            "Could not determine data granularity (expected 15-min or hourly). "
            "Assuming 15-min."
        )
        gran = "15min"

    # ── Interpolate NaN values ────────────────────────────────────────────────
    if df["value"].isna().any():
        df["value"] = df["value"].interpolate(method="time", limit=8)
        still_nan = df["value"].isna().sum()
        if still_nan:
            warns.append(f"{still_nan} values remain NaN after interpolation (large gap); set to 0.")
            df["value"] = df["value"].fillna(0)

    # ── Negative values ───────────────────────────────────────────────────────
    neg = (df["value"] < 0).sum()
    if neg:
        warns.append(
            f"{neg} negative values detected (net injection periods). "
            "Included as-is — check if injection should be excluded."
        )

    # ── Unit → MWh per QH or per hour ────────────────────────────────────────
    # kW × period_h = kWh; kWh stays as-is; then /1000 → MWh
    period_h = 0.25 if gran == "15min" else 1.0
    if unit == "kW":
        df["mwh"] = df["value"] * period_h / 1000
    else:  # kWh
        df["mwh"] = df["value"] / 1000

    # ── Gap detection ─────────────────────────────────────────────────────────
    expected_periods = (df.index[-1] - df.index[0]) / pd.Timedelta(period_h, unit="h") + 1
    actual_periods = len(df)
    gap_periods = int(expected_periods - actual_periods)
    if gap_periods > 0:
        info["gap_periods"] = gap_periods
        warns.append(
            f"Data has {gap_periods} missing {gran} periods "
            f"({gap_periods * period_h:.1f} hours). Missing values treated as 0 after reindex."
        )
        full_idx = pd.date_range(df.index[0], df.index[-1], freq=gran)
        df = df.reindex(full_idx).fillna(0)

    # ── Aggregate QH → hourly MWh ─────────────────────────────────────────────
    df["ts_h"] = df.index.floor("h")
    hourly = df.groupby("ts_h")["mwh"].sum()
    hourly.index.name = "ts_local"

    info["rows"] = len(df)
    info["hours"] = len(hourly)
    info["period_start"] = hourly.index.min()
    info["period_end"] = hourly.index.max()
    info["total_mwh"] = hourly.sum()

    return hourly, warns, info


def classify_peak(index: pd.DatetimeIndex) -> np.ndarray:
    """True = peak hour: Mon–Fri 08:00–19:59 (local time, tz-naive)."""
    return (index.dayofweek < 5) & (index.hour >= 8) & (index.hour < 20)


def _monthly_cost(
    df_h: pd.DataFrame,
    monthly: pd.DataFrame,
    product_type: str,
    a_p: float, b_p: float,
    a_d: float, b_d: float,
) -> pd.Series:
    """Compute monthly cost Series for a single product."""
    if product_type == "Weighted":
        price_h = np.where(
            df_h["is_peak"].values,
            a_p * df_h["belpex"].values + b_p,
            a_d * df_h["belpex"].values + b_d,
        )
        tmp = pd.DataFrame({"month": df_h["month"].values, "cost": price_h * df_h["vol_mwh"].values})
        return tmp.groupby("month")["cost"].sum()
    else:  # Arithmetic
        return (
            (a_p * monthly["belpex_avg_peak"].fillna(0) + b_p) * monthly["vol_peak"]
            + (a_d * monthly["belpex_avg_dal"].fillna(0) + b_d) * monthly["vol_dal"]
        )


def compute_monthly(
    vol: pd.Series,
    belpex: pd.Series,
    type_1: str, a_1_p: float, b_1_p: float, a_1_d: float, b_1_d: float,
    type_2: str, a_2_p: float, b_2_p: float, a_2_d: float, b_2_d: float,
) -> tuple:
    """Compute monthly costs for two products (Weighted or Arithmetic)."""
    df = vol.to_frame("vol_mwh").join(belpex.rename("belpex"), how="left")

    missing = int(df["belpex"].isna().sum())
    neg_belpex = int((df["belpex"] < 0).sum())
    if missing:
        df["belpex"] = df["belpex"].fillna(df["belpex"].median())

    df["is_peak"] = classify_peak(df.index)
    df["vol_peak"] = df["vol_mwh"].where(df["is_peak"], 0.0)
    df["vol_dal"]  = df["vol_mwh"].where(~df["is_peak"], 0.0)
    df["belpex_peak"] = df["belpex"].where(df["is_peak"])
    df["belpex_dal"]  = df["belpex"].where(~df["is_peak"])
    df["month"] = df.index.to_period("M")

    monthly = df.groupby("month").agg(
        vol_mwh        =("vol_mwh",    "sum"),
        vol_peak       =("vol_peak",   "sum"),
        vol_dal        =("vol_dal",    "sum"),
        belpex_avg     =("belpex",     "mean"),
        belpex_avg_peak=("belpex_peak","mean"),
        belpex_avg_dal =("belpex_dal", "mean"),
    )

    monthly["cost_1"]      = _monthly_cost(df, monthly, type_1, a_1_p, b_1_p, a_1_d, b_1_d)
    monthly["cost_2"]      = _monthly_cost(df, monthly, type_2, a_2_p, b_2_p, a_2_d, b_2_d)
    monthly["eff_price_1"] = monthly["cost_1"] / monthly["vol_mwh"].replace(0, np.nan)
    monthly["eff_price_2"] = monthly["cost_2"] / monthly["vol_mwh"].replace(0, np.nan)
    monthly["delta"]       = monthly["cost_1"] - monthly["cost_2"]

    # Negative price hours for Product 1 (informational warning)
    price_1_h = np.where(
        df["is_peak"].values,
        a_1_p * df["belpex"].values + b_1_p,
        a_1_d * df["belpex"].values + b_1_d,
    )
    df["price_1"] = price_1_h
    monthly["neg_price_hours"] = df.groupby("month")["price_1"].apply(lambda x: (x < 0).sum())

    return monthly, df, missing, neg_belpex


# ── Baseload helpers ─────────────────────────────────────────────────────────────────

def _peak_frac_monthly(df_h: pd.DataFrame) -> pd.Series:
    """Fraction of hours in each month that fall in peak (Mon–Fri 08:00–19:59)."""
    return df_h.groupby("month")["is_peak"].mean()


def compute_baseload_mixed(
    monthly: pd.DataFrame,
    baseload_pct: float,
    clicked_price: float,
    wa_cost_col: str = "cost_1",
) -> tuple:
    """
    Split the Weighted product cost into:
    - Baseload slice: flat MWh/month = min_monthly_vol × pct%, priced at the
      user-defined clicked_price (€/MWh) — fully fixed, Belpex-independent.
    - Swing slice: residual volume, proportional share of the WA product cost.
    wa_cost_col: column in monthly that holds the Weighted product cost.
    Returns (cost_mixed Series, baseload_mwh scalar).
    """
    min_monthly_mwh = monthly["vol_mwh"].min()
    bl_mwh = min_monthly_mwh * baseload_pct / 100.0
    if bl_mwh == 0.0:
        return monthly[wa_cost_col].copy(), 0.0
    bl_cost = pd.Series(bl_mwh * clicked_price, index=monthly.index)
    swing_frac = (monthly["vol_mwh"] - bl_mwh) / monthly["vol_mwh"].replace(0, np.nan)
    swing_cost = monthly[wa_cost_col] * swing_frac.clip(lower=0.0)
    return (bl_cost + swing_cost), bl_mwh


def sweep_baseload(
    monthly: pd.DataFrame,
    clicked_price: float,
    wa_cost_col: str = "cost_1",
    ref_cost_col: str = "cost_2",
) -> pd.DataFrame:
    """Sweep baseload % 0→100 step 5. wa_cost_col = Weighted product, ref_cost_col = other product."""
    cost_ref_total = monthly[ref_cost_col].sum()
    cost_wa_pure = monthly[wa_cost_col].sum()
    rows = []
    for pct in range(0, 101, 5):
        mixed, bl_mwh = compute_baseload_mixed(monthly, pct, clicked_price, wa_cost_col)
        total = mixed.sum()
        rows.append({
            "pct": pct,
            "bl_mwh": bl_mwh,
            "total_cost": total,
            "delta_vs_p2": total - cost_ref_total,
            "savings_vs_pure_wa": cost_wa_pure - total,
        })
    return pd.DataFrame(rows)


def sweep_baseload_spread(
    monthly: pd.DataFrame,
    df_h: pd.DataFrame,
    a_1_p: float, b_1_p: float, a_1_d: float, b_1_d: float,
    a_2_p: float, b_2_p: float, a_2_d: float, b_2_d: float,
    type_2: str,
    clicked_price: float,
    spread_multipliers: list,
    wa_cost_col: str = "cost_1",
    ref_cost_col: str = "cost_2",
) -> dict:
    """
    Re-derive Belpex with scaled intra-month deviations:
        belpex_adj_h = monthly_mean + mult × (belpex_h − monthly_mean)
    Monthly averages are preserved; only within-month spread changes.
    The clicked baseload price is fixed regardless of spread — only the swing
    (floating WA) portion responds to volatility changes.
    Returns dict[multiplier → sweep DataFrame].
    """
    monthly_mean_h = df_h.groupby("month")["belpex"].transform("mean")
    results = {}
    for mult in spread_multipliers:
        df_adj = df_h.copy()
        df_adj["belpex"] = monthly_mean_h + mult * (df_h["belpex"] - monthly_mean_h)
        monthly_adj = monthly.copy()
        monthly_adj["belpex_avg_peak"] = (
            df_adj[df_adj["is_peak"]].groupby("month")["belpex"].mean()
        )
        monthly_adj["belpex_avg_dal"] = (
            df_adj[~df_adj["is_peak"]].groupby("month")["belpex"].mean()
        )
        price_1_h = np.where(
            df_adj["is_peak"].values,
            a_1_p * df_adj["belpex"].values + b_1_p,
            a_1_d * df_adj["belpex"].values + b_1_d,
        )
        tmp = pd.DataFrame({"month": df_adj["month"].values, "cost": price_1_h * df_adj["vol_mwh"].values})
        monthly_adj["cost_1"] = tmp.groupby("month")["cost"].sum()
        monthly_adj["cost_2"] = _monthly_cost(df_adj, monthly_adj, type_2, a_2_p, b_2_p, a_2_d, b_2_d)
        results[mult] = sweep_baseload(monthly_adj, clicked_price, wa_cost_col, ref_cost_col)
    return results


def build_excel(monthly: pd.DataFrame,
                type_1, a_1_p, b_1_p, a_1_d, b_1_d,
                type_2, a_2_p, b_2_p, a_2_d, b_2_d) -> bytes:
    total_1 = monthly["cost_1"].sum()
    total_2 = monthly["cost_2"].sum()
    delta = total_1 - total_2
    total_vol = monthly["vol_mwh"].sum()
    ref_1 = "Belpex_h" if type_1 == "Weighted" else "avg_Belpex"
    ref_2 = "Belpex_h" if type_2 == "Weighted" else "avg_Belpex"

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary = pd.DataFrame(
            {
                "Product": [f"Product 1 ({type_1})", f"Product 2 ({type_2})", "Delta (1 − 2)"],
                "Peak formula": [
                    f"{a_1_p}·{ref_1} + {b_1_p}",
                    f"{a_2_p}·{ref_2} + {b_2_p}",
                    "",
                ],
                "Off-peak formula": [
                    f"{a_1_d}·{ref_1} + {b_1_d}",
                    f"{a_2_d}·{ref_2} + {b_2_d}",
                    "",
                ],
                "Total Volume (MWh)": [f"{total_vol:,.1f}", f"{total_vol:,.1f}", ""],
                "Total Cost (€)": [
                    f"{total_1:,.2f}",
                    f"{total_2:,.2f}",
                    f"{delta:+,.2f}",
                ],
            }
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)

        out = monthly[
            ["vol_mwh", "vol_peak", "vol_dal", "belpex_avg_peak", "belpex_avg_dal",
             "eff_price_1", "cost_1", "eff_price_2", "cost_2", "delta"]
        ].copy()
        out.index = out.index.astype(str)
        out.columns = [
            "Volume (MWh)", "Peak Vol (MWh)", "Off-peak Vol (MWh)",
            "Belpex Avg Peak", "Belpex Avg Off-peak",
            f"Eff. Price P1 ({type_1}) (€/MWh)",
            f"Cost P1 ({type_1}) (€)",
            f"Eff. Price P2 ({type_2}) (€/MWh)",
            f"Cost P2 ({type_2}) (€)",
            "Delta P1−P2 (€)",
        ]
        out.to_excel(writer, sheet_name="Monthly Breakdown", index_label="Month")

        params = pd.DataFrame(
            {
                "Parameter": [
                    "Product 1 — Type", "Product 1 — Peak formula", "Product 1 — Off-peak formula",
                    "Product 2 — Type", "Product 2 — Peak formula", "Product 2 — Off-peak formula",
                    "Export date",
                ],
                "Value": [
                    type_1,
                    f"{a_1_p} × {ref_1} + {b_1_p}",
                    f"{a_1_d} × {ref_1} + {b_1_d}",
                    type_2,
                    f"{a_2_p} × {ref_2} + {b_2_p}",
                    f"{a_2_d} × {ref_2} + {b_2_d}",
                    datetime.date.today().isoformat(),
                ],
            }
        )
        params.to_excel(writer, sheet_name="Parameters", index=False)

    return buf.getvalue()


# ─── UI Layout ────────────────────────────────────────────────────────────────

st.title("Product Comparison")
st.caption("Weighted average vs. arithmetic average pricing — financial delta for a given consumption profile.")
st.info(
    "Upload a client's quarterly or annual consumption file to compute the financial difference between "
    "a **Weighted Average (WA)** and an **Arithmetic Average (AA)** electricity contract. "
    "WA prices each hour individually against the Belpex spot price; AA prices at the monthly average — "
    "the cost delta depends entirely on the client's load shape relative to the price curve."
)

# ── Top row: file upload + unit ───────────────────────────────────────────────
with st.container():
    col_file, col_unit = st.columns([3, 1])
    with col_file:
        uploaded = st.file_uploader(
            "Consumption file",
            type=["csv", "xlsx", "xls"],
            help=(
                "3-column format: Date (YYYY-MM-DD) | Time (HH:MM) | Value (kW or kWh). "
                "CSV (comma or semicolon) or Excel (.xlsx/.xls). "
                "15-min QH and hourly granularity both supported."
            ),
        )
    with col_unit:
        unit = st.radio(
            "Value unit",
            ["kW", "kWh"],
            index=0,
            horizontal=True,
            help="kW = instantaneous power (most common). kWh = energy per interval.",
        )

# ── Parameters block ──────────────────────────────────────────────────────────
st.divider()
st.subheader("Product parameters")
st.caption(
    "Formula: **Y = Slope × Belpex + Offset** (€/MWh)  │  "
    "Peak = Mon–Fri 08:00–19:59  │  Off-peak = all other hours"
)

same_params = st.checkbox(
    "Same slope & offset for both products",
    value=True,
    help="When checked, Product 2 uses the same slope and offset as Product 1. The product type (Weighted / Arithmetic) is always selected independently.",
)

# Header row
_h0, _h1, _h2, _h3, _h4, _h5 = st.columns([2, 1.5, 1, 1, 1, 1])
_h1.markdown("**Type**")
_h2.markdown("**Index coeff. (peak)**")
_h3.markdown("**Fixed adder (peak)**")
_h4.markdown("**Index coeff. (off-peak)**")
_h5.markdown("**Fixed adder (off-peak)**")

# Product 1 row
_w0, _w1, _w2, _w3, _w4, _w5 = st.columns([2, 1.5, 1, 1, 1, 1])
_w0.markdown("**Product 1**")
with _w1: type_1 = st.selectbox("type_1", ["Weighted", "Arithmetic"], key="type_1", label_visibility="collapsed")
with _w2: a_1_p = st.number_input("1_ap", value=1.0, step=0.01, format="%.4f", key="a_1_p", label_visibility="collapsed", help="Multiplier applied to the Belpex spot price (e.g. 1.05 = 5% uplift). Typical range: 0.90–1.10.")
with _w3: b_1_p = st.number_input("1_bp", value=0.0, step=0.5, format="%.2f", key="b_1_p", label_visibility="collapsed", help="Fixed add-on in €/MWh, independent of Belpex (can be negative). Typical range: −5 to +5.")
with _w4: a_1_d = st.number_input("1_ad", value=1.0, step=0.01, format="%.4f", key="a_1_d", label_visibility="collapsed", help="Multiplier applied to the Belpex spot price for off-peak hours. Typical range: 0.90–1.10.")
with _w5: b_1_d = st.number_input("1_bd", value=0.0, step=0.5, format="%.2f", key="b_1_d", label_visibility="collapsed", help="Fixed add-on in €/MWh for off-peak hours (can be negative). Typical range: −5 to +5.")

# Product 2 row
_a0, _a1, _a2, _a3, _a4, _a5 = st.columns([2, 1.5, 1, 1, 1, 1])
_a0.markdown("**Product 2**")
with _a1: type_2 = st.selectbox("type_2", ["Weighted", "Arithmetic"], index=1, key="type_2", label_visibility="collapsed")
if same_params:
    a_2_p, b_2_p, a_2_d, b_2_d = a_1_p, b_1_p, a_1_d, b_1_d
    _a2.caption("← same as P1")
else:
    with _a2: a_2_p = st.number_input("2_ap", value=1.0, step=0.01, format="%.4f", key="a_2_p", label_visibility="collapsed", help="Multiplier applied to the Belpex spot price (e.g. 1.05 = 5% uplift). Typical range: 0.90–1.10.")
    with _a3: b_2_p = st.number_input("2_bp", value=0.0, step=0.5, format="%.2f", key="b_2_p", label_visibility="collapsed", help="Fixed add-on in €/MWh, independent of Belpex (can be negative). Typical range: −5 to +5.")
    with _a4: a_2_d = st.number_input("2_ad", value=1.0, step=0.01, format="%.4f", key="a_2_d", label_visibility="collapsed", help="Multiplier applied to the Belpex spot price for off-peak hours. Typical range: 0.90–1.10.")
    with _a5: b_2_d = st.number_input("2_bd", value=0.0, step=0.5, format="%.2f", key="b_2_d", label_visibility="collapsed", help="Fixed add-on in €/MWh for off-peak hours (can be negative). Typical range: −5 to +5.")

st.divider()

# ── Guard: no file ────────────────────────────────────────────────────────────
if uploaded is None:
    with st.expander("How does this work?", expanded=False):
        st.markdown(
            """
**Weighted pricing**
- Each hour: `price_h = Slope × Belpex_h + Offset` — separately for peak and off-peak
- Cost captures the timing of consumption: consuming in expensive hours costs more

**Arithmetic pricing**
- Each month: `price = Slope × monthly_avg(Belpex) + Offset` — flat for peak/off-peak blocks
- No benefit or penalty for consumption timing within the month

**Delta = Weighted − Arithmetic**
- Positive → weighted is more expensive (consumption skewed toward high-price hours)
- Negative → arithmetic is more expensive (consumption skewed toward off-peak)

**Peak hours:** Mon–Fri 08:00–19:59 │ **Off-peak hours:** all other times
            """
        )
    st.stop()

# ── Parse file ────────────────────────────────────────────────────────────────
file_key = f"{uploaded.name}_{uploaded.size}_{unit}"
if file_key != st.session_state.get("_file_key"):
    with st.spinner("Parsing file…"):
        vol, parse_warns, parse_info = parse_consumption_file(uploaded, unit)
    if vol is None:
        st.error("Could not parse the file:\n\n" + "\n".join(parse_warns))
        st.stop()
    st.session_state["_file_key"] = file_key
    st.session_state["_vol"] = vol
    st.session_state["_parse_warns"] = parse_warns
    st.session_state["_parse_info"] = parse_info

vol = st.session_state["_vol"]
parse_info = st.session_state.get("_parse_info", {})
parse_warns = st.session_state.get("_parse_warns", [])

# ── Data quality badge (collapsed; red only if warnings exist) ────────────────
gran_label = {"15min": "15-min QH", "60min": "Hourly"}.get(parse_info.get("granularity", ""), "unknown")
summary_line = (
    f"**{uploaded.name}** | "
    f"Period: {parse_info.get('period_start', vol.index.min()).strftime('%b %Y')} – "
    f"{parse_info.get('period_end', vol.index.max()).strftime('%b %Y')} | "
    f"Granularity: {gran_label} | "
    f"Total: {parse_info.get('total_mwh', vol.sum()):,.1f} MWh"
)

if parse_warns:
    with st.expander(f"⚠ Data quality — {len(parse_warns)} notice(s)", expanded=False):
        for w in parse_warns:
            st.warning(w)
    st.info(summary_line)
else:
    st.success(summary_line + " | No data quality issues")

with st.expander("Preview uploaded data", expanded=False):
    _preview = vol.head(5).reset_index()
    _preview.columns = ["Timestamp", "MWh/h"]
    _preview["MWh/h"] = _preview["MWh/h"].map("{:.4f}".format)
    st.dataframe(_preview, use_container_width=True, hide_index=True)
belpex = load_belpex()

st.caption(
    f"Belpex source: EPEX Spot Belgium day-ahead │ "
    f"Data through **{belpex.index.max().strftime('%B %Y')}**"
)

if vol.index.max() > belpex.index.max():
    st.warning(
        f"Belpex data ends at {belpex.index.max().strftime('%b %Y')}. "
        "Hours beyond that use a median fallback — results are approximate."
    )

monthly, df_h, missing, neg_belpex = compute_monthly(
    vol, belpex,
    type_1, a_1_p, b_1_p, a_1_d, b_1_d,
    type_2, a_2_p, b_2_p, a_2_d, b_2_d,
)
df_h_stored = df_h

calc_warns = []
if missing:
    calc_warns.append(
        f"{missing} hours had no Belpex price match (period median substituted). "
        "Check that the file period is covered by the Belpex dataset (from 2015)."
    )
if neg_belpex:
    calc_warns.append(
        f"{neg_belpex} hours have negative Belpex prices. "
        "Verify the formula result is commercially correct for these periods."
    )
if calc_warns:
    with st.expander(f"⚠ Calculation notices — {len(calc_warns)} item(s)", expanded=False):
        for w in calc_warns:
            st.warning(w)

# ── Results ───────────────────────────────────────────────────────────────────
label_1 = f"Product 1 ({type_1})"
label_2 = f"Product 2 ({type_2})"

total_1 = monthly["cost_1"].sum()
total_2 = monthly["cost_2"].sum()
delta = total_1 - total_2
total_vol = monthly["vol_mwh"].sum()
delta_pct = (delta / total_2 * 100) if total_2 else 0

# ── KPI row ───────────────────────────────────────────────────────────────────
k1, k2, k3, k4 = st.columns(4)
k1.metric(label_1, f"€ {total_1:,.0f}")
k2.metric(label_2, f"€ {total_2:,.0f}")
k3.metric(
    "Delta (1 − 2)",
    f"€ {delta:+,.0f}",
    delta=f"{delta_pct:+.1f}%",
    delta_color="inverse",
)
k4.metric("Total Volume", f"{total_vol:,.1f} MWh")

_cheaper = label_2 if delta > 0 else label_1
_pricier = label_1 if delta > 0 else label_2
_delta_abs = abs(delta)
_delta_pct_abs = abs(delta_pct)
if delta != 0:
    st.success(
        f"**{_cheaper}** is cheaper by **€ {_delta_abs:,.0f}** ({_delta_pct_abs:.1f}%) "
        f"for this consumption profile — compared to **{_pricier}**."
    )
else:
    st.info("Both products yield identical total costs for this consumption profile.")

st.divider()

# ── Charts ────────────────────────────────────────────────────────────────────
months_str = [str(m) for m in monthly.index]

# Side-by-side monthly costs
fig_cost = go.Figure()
fig_cost.add_bar(
    name=label_1,
    x=months_str,
    y=monthly["cost_1"],
    marker_color=COL_WA,
)
fig_cost.add_bar(
    name=label_2,
    x=months_str,
    y=monthly["cost_2"],
    marker_color=COL_AA,
)
fig_cost.update_layout(
    barmode="group",
    title=f"Monthly Cost — {label_1} vs {label_2}",
    xaxis_title="Month",
    yaxis_title="Cost (€)",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    height=380,
    margin=dict(t=60),
)
st.plotly_chart(fig_cost, use_container_width=True)

# Monthly delta
delta_colors = ["#d62728" if d > 0 else "#2ca02c" for d in monthly["delta"]]
fig_delta = go.Figure()
fig_delta.add_bar(
    name="Delta P1 − P2",
    x=months_str,
    y=monthly["delta"],
    marker_color=delta_colors,
    text=[f"€ {d:+,.0f}" for d in monthly["delta"]],
    textposition="auto",
    textfont=dict(size=11),
)
fig_delta.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
_d_max = monthly["delta"].max()
_d_min = monthly["delta"].min()
_y_max = _d_max * 1.5 if _d_max > 0 else abs(_d_min) * 0.35
_y_min = _d_min * 1.5 if _d_min < 0 else -abs(_d_max) * 0.35
fig_delta.update_layout(
    title=f"Monthly Delta — {label_1} minus {label_2}  (red = P1 more expensive, green = P2 more expensive)",
    xaxis_title="Month",
    yaxis_title="Delta (€)",
    yaxis=dict(range=[_y_min, _y_max]),
    showlegend=False,
    height=340,
    margin=dict(t=60),
)
st.plotly_chart(fig_delta, use_container_width=True)

# ── Monthly table ─────────────────────────────────────────────────────────────
st.subheader("Monthly Breakdown")

neg_price_total = monthly["neg_price_hours"].sum()
if neg_price_total:
    st.caption(
        f"Note: {int(neg_price_total)} hour(s) had a negative computed price "
        "(negative Belpex with small B). Included at face value."
    )

display = pd.DataFrame(
    {
        "Month": months_str,
        "Volume (MWh)": monthly["vol_mwh"].map("{:,.1f}".format),
        "Peak Vol (MWh)": monthly["vol_peak"].map("{:,.1f}".format),
        "Off-peak Vol (MWh)": monthly["vol_dal"].map("{:,.1f}".format),
        "Belpex Peak (€/MWh)": monthly["belpex_avg_peak"].map(
            lambda x: f"{x:.2f}" if pd.notna(x) else "—"
        ),
        "Belpex Off-peak (€/MWh)": monthly["belpex_avg_dal"].map(
            lambda x: f"{x:.2f}" if pd.notna(x) else "—"
        ),
        f"Eff. Price P1 ({type_1}) (€/MWh)": monthly["eff_price_1"].map(
            lambda x: f"{x:.2f}" if pd.notna(x) else "—"
        ),
        f"Cost P1 ({type_1}) (€)": monthly["cost_1"].map("{:,.0f}".format),
        f"Eff. Price P2 ({type_2}) (€/MWh)": monthly["eff_price_2"].map(
            lambda x: f"{x:.2f}" if pd.notna(x) else "—"
        ),
        f"Cost P2 ({type_2}) (€)": monthly["cost_2"].map("{:,.0f}".format),
        "Delta P1−P2 (€)": monthly["delta"].map("{:+,.0f}".format),
    }
)
st.table(display.astype(object))
# ── Baseload Click Sensitivity ─────────────────────────────────────────────────────────
st.divider()
_has_wa = type_1 == "Weighted" or type_2 == "Weighted"
if _has_wa and df_h_stored is not None:
    # Resolve which product is Weighted and which is the reference
    if type_1 == "Weighted":
        _wa_cost_col, _ref_cost_col = "cost_1", "cost_2"
        _wa_label, _ref_label = label_1, label_2
        _wa_total, _ref_total = total_1, total_2
        _wa_eff_price = total_1 / monthly["vol_mwh"].sum() if monthly["vol_mwh"].sum() > 0 else 60.0
        _sw_a1p, _sw_b1p, _sw_a1d, _sw_b1d = a_1_p, b_1_p, a_1_d, b_1_d
        _sw_a2p, _sw_b2p, _sw_a2d, _sw_b2d = a_2_p, b_2_p, a_2_d, b_2_d
        _sw_type2 = type_2
    else:  # type_2 == "Weighted"
        _wa_cost_col, _ref_cost_col = "cost_2", "cost_1"
        _wa_label, _ref_label = label_2, label_1
        _wa_total, _ref_total = total_2, total_1
        _wa_eff_price = total_2 / monthly["vol_mwh"].sum() if monthly["vol_mwh"].sum() > 0 else 60.0
        _sw_a1p, _sw_b1p, _sw_a1d, _sw_b1d = a_2_p, b_2_p, a_2_d, b_2_d
        _sw_a2p, _sw_b2p, _sw_a2d, _sw_b2d = a_1_p, b_1_p, a_1_d, b_1_d
        _sw_type2 = type_1

    st.subheader("Baseload Click Sensitivity")
    _min_monthly_mwh = monthly["vol_mwh"].min()
    st.caption(
        f"Baseload click on **{_wa_label}** (Weighted product). "
        f"Anchored to minimum monthly volume: **{_min_monthly_mwh:,.1f} MWh/month**. "
        f"Baseload slice → locked at a fixed clicked price (€/MWh). "
        f"Swing slice → remains fully Belpex-exposed (Weighted hourly). "
        f"Reference: **{_ref_label}**."
    )

    _col_slider, _col_price = st.columns([3, 1])
    with _col_slider:
        bl_pct = st.slider(
            "Baseload volume click (%)",
            min_value=0, max_value=100, value=50, step=5,
            key="bl_pct",
            help=f"% of the minimum monthly volume ({_min_monthly_mwh:,.1f} MWh) to lock in at the clicked price. Remaining volume stays fully Belpex-exposed (Weighted).",
        )
    with _col_price:
        if st.session_state.get("_reset_clicked"):
            st.session_state["clicked_price"] = round(_wa_eff_price, 2)
            st.session_state["_reset_clicked"] = False
        clicked_price = st.number_input(
            "Clicked price (€/MWh)",
            min_value=0.0, max_value=500.0,
            value=st.session_state.get("clicked_price", round(_wa_eff_price, 2)),
            step=0.5,
            format="%.2f",
            key="clicked_price",
            help="The fixed price (€/MWh) locked in for the baseload volume. Fully decoupled from Belpex movements.",
        )
        if st.button("↺ Reset to WA price", key="reset_clicked_btn", help=f"Reset to the current WA effective price: €{_wa_eff_price:.2f}/MWh"):
            st.session_state["_reset_clicked"] = True
            st.rerun()

    _cost_mixed_sel, _bl_mwh_sel = compute_baseload_mixed(
        monthly, bl_pct, clicked_price, _wa_cost_col
    )
    _total_mixed = _cost_mixed_sel.sum()
    _savings_vs_wa = _wa_total - _total_mixed
    _delta_vs_p2 = _total_mixed - _ref_total

    _bm1, _bm2, _bm3, _bm4 = st.columns(4)
    _bm1.metric("Baseload MWh/month", f"{_bl_mwh_sel:,.1f} MWh")
    _bm2.metric(f"{_wa_label} Mixed total cost", f"€ {_total_mixed:,.0f}")
    _bm3.metric(
        f"Savings vs pure {_wa_label}",
        f"€ {_savings_vs_wa:+,.0f}",
        delta=f"{_savings_vs_wa / _wa_total * 100:+.1f}%" if _wa_total else None,
        delta_color="normal",
    )
    _bm4.metric(
        f"Delta vs {_ref_label}",
        f"€ {_delta_vs_p2:+,.0f}",
        delta=f"{_delta_vs_p2 / _ref_total * 100:+.1f}%" if _ref_total else None,
        delta_color="inverse",
    )

    # Break-even price
    if _bl_mwh_sel > 0:
        _swing_frac_ann = ((monthly["vol_mwh"] - _bl_mwh_sel) / monthly["vol_mwh"].replace(0, np.nan)).clip(lower=0.0)
        _swing_cost_ann = (monthly[_wa_cost_col] * _swing_frac_ann).sum()
        _bl_months = len(monthly)
        _breakeven_price = (_ref_total - _swing_cost_ann) / (_bl_mwh_sel * _bl_months)
    else:
        _breakeven_price = None

    _bm5_col, _, _, _ = st.columns(4)
    if _breakeven_price is not None:
        _be_margin = _breakeven_price - clicked_price
        _bm5_col.metric(
            f"Break-even price @ {bl_pct}%",
            f"€ {_breakeven_price:.2f}/MWh",
            delta=f"{_be_margin:+.2f} vs clicked" if _be_margin else None,
            delta_color="normal",
            help=f"Maximum clicked price at {bl_pct}% baseload volume where the mixed product still costs less than pure {_ref_label}. "
                 "Positive margin = you have headroom above your clicked price. Negative = you are already clicking above break-even.",
        )
    else:
        _bm5_col.metric("Break-even price", "N/A (0% click)")

    # ── Chart D: Monthly decomposition ──────────────────────────────────────────
    st.subheader("Monthly Cost Decomposition")
    _col_baseload = "#08519c"  # dark navy — always a darker shade of COL_WA (blue), WA product regardless of slot
    st.caption(
        f"**Dark navy (bottom)** = baseload slice ({_bl_mwh_sel:,.1f} MWh/month, identical every month) "
        f"priced at the fixed clicked price of **€{clicked_price:.2f}/MWh** — fully decoupled from Belpex. "
        f"**Blue (top)** = swing slice (volume above baseload) priced at the floating Weighted hourly rate. "
        f"**White dotted line** = pure {_ref_label} reference cost per month. "
        "Because the bottom block is constant, seasonal variation comes entirely from the swing (top) portion."
    )
    _bl_cost_m = pd.Series(_bl_mwh_sel * clicked_price, index=monthly.index)
    _swing_cost_m = _cost_mixed_sel - _bl_cost_m
    _months_d = [str(m) for m in monthly.index]
    _fig_d = go.Figure()
    # Baseload at BOTTOM — dark shade of WA product color
    _fig_d.add_bar(
        name=f"Baseload slice (fixed, {_bl_mwh_sel:,.1f} MWh/month)",
        x=_months_d, y=_bl_cost_m.values,
        marker_color=_col_baseload,
    )
    # Swing on TOP — standard WA product color
    _fig_d.add_bar(
        name=f"Swing slice (Weighted hourly)",
        x=_months_d, y=_swing_cost_m.values,
        marker_color=COL_WA,
    )
    _fig_d.add_scatter(
        name=f"Pure {_ref_label}",
        x=_months_d, y=monthly[_ref_cost_col].values,
        mode="lines+markers",
        line=dict(color="white", width=1.5, dash="dot"), marker=dict(size=4),
    )
    _fig_d.update_layout(
        barmode="stack",
        title=f"{bl_pct}% Baseload Click — {_bl_mwh_sel:,.1f} MWh/month priced at Arithmetic",
        xaxis_title="Month", yaxis_title="Cost (€)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=400, margin=dict(t=60),
    )
    st.plotly_chart(_fig_d, use_container_width=True)

    # ── Chart C: Cost convergence curve ─────────────────────────────────────────
    st.subheader("Cost Convergence Curve")
    st.caption(
        f"At the clicked price of **€{clicked_price:.2f}/MWh**, how does locking in more baseload volume affect total cost? "
        f"**Blue dashed** = pure {_wa_label} (no click). **Orange dashed** = pure {_ref_label} (target). "
        f"**Solid purple** = mixed cost at your clicked price. "
        f"The intersection with the orange line is the % at which the mixed product costs the same as pure {_ref_label}."
    )
    _sweep = sweep_baseload(monthly, clicked_price, _wa_cost_col, _ref_cost_col)
    _fig_c = go.Figure()
    # Pure WA reference
    _fig_c.add_scatter(
        x=[0, 100], y=[_wa_total, _wa_total],
        mode="lines", name=f"Pure {_wa_label} (no click)",
        line=dict(color=COL_WA, width=2, dash="dash"),
    )
    # Pure ref product reference
    _fig_c.add_scatter(
        x=[0, 100], y=[_ref_total, _ref_total],
        mode="lines", name=f"Pure {_ref_label} (target)",
        line=dict(color=COL_AA, width=2, dash="dash"),
    )
    # Main curve at clicked price
    _fig_c.add_scatter(
        x=_sweep["pct"], y=_sweep["total_cost"],
        mode="lines+markers", name=f"Clicked: €{clicked_price:.2f}/MWh",
        line=dict(color=COL_MIXED, width=2.5), marker=dict(size=5),
    )
    # Vertical line at selected %
    _fig_c.add_vline(
        x=bl_pct, line_dash="dash", line_color="#ff7f0e", line_width=1.5,
        annotation_text=f"  Selected: {bl_pct}%", annotation_position="top right",
    )
    # Crossover annotation
    _crossover = _sweep[_sweep["delta_vs_p2"] <= 0]["pct"]
    if len(_crossover) > 0 and _crossover.iloc[0] > 0:
        _fig_c.add_vline(
            x=_crossover.iloc[0], line_dash="dot", line_color="gray", line_width=1,
            annotation_text=f"  {_ref_label} parity ~{_crossover.iloc[0]}%",
            annotation_position="top left",
        )
    _fig_c.update_layout(
        xaxis=dict(title="Baseload % clicked", ticksuffix="%"),
        yaxis=dict(title="Total annual cost (€)", tickformat=",.0f"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=420, margin=dict(t=40),
    )
    st.plotly_chart(_fig_c, use_container_width=True)
    with st.expander("Full sweep data", expanded=False):
        _sd = _sweep.copy()
        _sd.columns = ["Baseload %", "BL MWh/month", f"Total {_wa_label} mixed (€)", f"Delta vs {_ref_label} (€)", f"Savings vs pure {_wa_label} (€)"]
        for _col in [f"Total {_wa_label} mixed (€)", f"Delta vs {_ref_label} (€)", f"Savings vs pure {_wa_label} (€)"]:
            _sd[_col] = _sd[_col].map(lambda v: f"€ {v:+,.0f}")
        _sd["BL MWh/month"] = _sd["BL MWh/month"].map("{:,.1f}".format)
        st.table(_sd.astype(object))

    # ── Chart E: Spread sensitivity ──────────────────────────────────────────────
    st.subheader("Spread Sensitivity")
    st.caption(
        f"The clicked baseload price (€{clicked_price:.2f}/MWh) is **fixed regardless of market conditions**. "
        "Only the swing portion (floating Weighted) is exposed to Belpex volatility. "
        "Each scenario re-scales intra-month price deviations using: "
        "`Belpex\_adj(h) = μ\_month + k × (Belpex(h) − μ\_month)` "
        "where μ\_month is the calendar-month average and k is the multiplier. "
        "This preserves the monthly mean exactly — so the Arithmetic product cost is **identical across all three scenarios** by construction. "
        "Negative-price hours are handled correctly: under Bull (×2.0) they go more negative; under Bear (×0.5) they compress toward the monthly mean. "
        "**Limitation:** this is a linear mean-preserving spread, not a stochastic simulation. "
        "It does not model changes in load-price correlation or profile shape across scenarios. "
        "Use it for directional sensitivity, not as a probabilistic forecast. "
        "Lines below the gray zero line mean the mixed product is cheaper than the Arithmetic reference at that baseload %."
    )
    _spread_mults = [0.5, 1.0, 2.0]
    _spread_labels = ["Bear (×0.5)", "Historical (×1.0)", "Bull (×2.0)"]
    with st.spinner("Computing spread scenarios…"):
        _sweep_spread = sweep_baseload_spread(
            monthly, df_h_stored,
            _sw_a1p, _sw_b1p, _sw_a1d, _sw_b1d,
            _sw_a2p, _sw_b2p, _sw_a2d, _sw_b2d,
            _sw_type2, clicked_price, _spread_mults,
            _wa_cost_col, _ref_cost_col,
        )
    _colors_e = ["#aec7e8", "#2171b5", "#e6550d"]
    _fig_e = go.Figure()
    for _mult, _col_e, _lbl_e in zip(_spread_mults, _colors_e, _spread_labels):
        _df_sw = _sweep_spread[_mult]
        _fig_e.add_scatter(
            x=_df_sw["pct"], y=_df_sw["delta_vs_p2"],
            mode="lines+markers", name=_lbl_e,
            line=dict(color=_col_e, width=2.5 if _mult == 1.0 else 1.5),
            marker=dict(size=4),
        )
    _fig_e.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
    _fig_e.add_vline(
        x=bl_pct, line_dash="dash", line_color="#ff7f0e", line_width=1.5,
        annotation_text=f"  Selected: {bl_pct}%",
    )
    _fig_e.update_layout(
        xaxis=dict(title="Baseload % clicked", ticksuffix="%"),
        yaxis=dict(title=f"Delta vs {_ref_label} (€)  [negative = mixed cheaper]", tickformat=",.0f"),
        legend=dict(title="Scenario", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=420, margin=dict(t=40),
    )
    st.plotly_chart(_fig_e, use_container_width=True)
elif not _has_wa:
    st.info("Baseload sensitivity requires at least one product set to Weighted.")
# ── Export ────────────────────────────────────────────────────────────────────
xlsx = build_excel(monthly, type_1, a_1_p, b_1_p, a_1_d, b_1_d, type_2, a_2_p, b_2_p, a_2_d, b_2_d)
_file_stem = Path(uploaded.name).stem if uploaded else "comparison"
_export_name = f"{_file_stem}_{datetime.date.today().strftime('%Y-%m-%d')}_comparison.xlsx"
st.download_button(
    label="Download Excel Report",
    data=xlsx,
    file_name=_export_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=False,
)
