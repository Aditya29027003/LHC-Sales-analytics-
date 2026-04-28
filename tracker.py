"""
Sales Analytics Tracker - LHC Logistics
=======================================
A numbers-first dashboard for sales reports.

Run:
    streamlit run tracker.py

Drops in any Excel/CSV file with the same columns as the LHC sales report.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# =============================================================================
# Page config & global styling
# =============================================================================
st.set_page_config(
    page_title="Sales Analytics Tracker",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_FILE = "Sales Report_24_04_2024.xlsx"
DEFAULT_ORDERS_FILE = "Secured orders.xlsx"

EXPECTED_COLUMNS = [
    "Billing Type", "Billing Document", "Billing Date", "Customer",
    "Product Id", "Product Desc", "Quantity (Actual)", "Net Price",
    "Net Sales Volume", "Division", "Sales Empl. Name",
    "Manufacturer Name", "Product Group", "Tax Amount",
    "Cost (Actual)", "Profit Margin", "Profit Margin Ratio", "City Name",
]

ORDERS_EXPECTED_COLUMNS = [
    "Sales Order", "Document Date", "Name of Customer", "Product Id",
    "Product Desc", "Order Quantity", "Net Price", "Net Value",
    "Delivered Quantity", "Net Value (Delivered)", "Pending Quantity",
    "Net Value (Pending)", "Division", "Sales Empl. Name",
    "Manufacturer Name", "City Name", "Product Group",
    "Requested Delivery Date",
]

# -----------------------------------------------------------------------------
# Sub-team mapping
# -----------------------------------------------------------------------------
# Some divisions are organised into smaller sub-teams. Add new divisions here
# in the same shape and the Departments tab will automatically show a sub-team
# breakdown when that division is selected.
#
# Format:  Division name (exact match) -> { Sub-team name -> [employee, ...] }
# Employees not listed for a mapped division are bucketed under
# "Other / Unassigned" so nothing silently disappears.
SUB_TEAM_MAP: dict[str, dict[str, list[str]]] = {
    "WSH (LB)": {
        "Dermatology": [
            "Vishnu T.V", "Sathish Kumar", "Abel Forbes", "Kit Aranas",
        ],
        "Life Sciences": [
            "Piyush Yadav", "Ramadan Youssef", "Mohammed Kamal",
            "Gopinath Sarangapani", "Ayyaz Khan",
        ],
        "Medical Division": [
            "Aldrich Gaupo Torres", "Shameem Arakkal Hydros",
            "Awaad Othmaan", "Mohammed Thwaha",
        ],
    },
    # Add other divisions here later, e.g.:
    # "MedDivision (L9)": { "Team A": [...], "Team B": [...] },
}


def assign_sub_team(division: str, employee: str) -> str:
    """Return the sub-team name for (division, employee).

    Returns "-" when the division has no sub-team configuration, or
    "Other / Unassigned" when the employee isn't listed in any sub-team
    of a mapped division.
    """
    mapping = SUB_TEAM_MAP.get(str(division))
    if not mapping:
        return "-"
    for team, members in mapping.items():
        if str(employee) in members:
            return team
    return "Other / Unassigned"


# -----------------------------------------------------------------------------
# Department colour palette
# -----------------------------------------------------------------------------
# Each division gets a stable colour used in all charts so a team is always
# recognised by the same shade. Add new divisions here as they appear.
DEPARTMENT_COLORS: dict[str, str] = {
    "MedDivision (L9)":     "#4F8DF7",  # blue
    "Simulation (LA)":      "#F59E0B",  # amber
    "Life Science (L8)":    "#10B981",  # emerald
    "WSH (LB)":             "#A855F7",  # violet
    "Derma HF (L3)":        "#EF4444",  # red
    "Derma Skincare (L4)":  "#EC4899",  # pink
    "Derma Equipment (L5)": "#F97316",  # orange
    "EduTech (L7)":         "#14B8A6",  # teal
    "Admin (L1)":           "#64748B",  # slate
}
DEFAULT_DEPT_COLOR = "#94A3B8"  # neutral fallback for unknown divisions

# Series colours (used wherever a chart compares Billed vs Pending etc.)
SERIES_COLORS = {
    "Billed sales":  "#4F8DF7",  # blue
    "Pending":       "#F59E0B",  # amber
    "Delivered":     "#10B981",  # emerald
    "Total order value": "#A855F7",
    "Net sales":     "#4F8DF7",
    "Profit":        "#10B981",
}

# Aging bucket colours (red = bad, green = healthy)
BUCKET_COLORS = {
    "Overdue 90+ days":     "#B91C1C",
    "Overdue 30-90 days":   "#EF4444",
    "Overdue 0-30 days":    "#F97316",
    "Due within 30 days":   "#EAB308",
    "Due in 30-90 days":    "#10B981",
    "Due 90+ days out":     "#0EA5E9",
    "Unknown":              "#94A3B8",
}


def dept_color(division: str) -> str:
    """Return the colour for a division (or a neutral fallback)."""
    return DEPARTMENT_COLORS.get(str(division), DEFAULT_DEPT_COLOR)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.6rem; padding-bottom: 2.5rem; max-width: 1500px; }

    /* KPI metric cards - use Streamlit theme tokens so they auto-adapt
       to dark and light modes. We tint with a translucent accent so the
       card is clearly visible against any page background. */
    div[data-testid="stMetric"] {
        background: color-mix(in srgb, var(--primary-color, #4F8DF7) 8%, var(--secondary-background-color, rgba(127,127,127,0.08)));
        border: 1px solid color-mix(in srgb, var(--primary-color, #4F8DF7) 22%, transparent);
        padding: 14px 16px 12px 16px;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.55rem;
        font-weight: 700;
        color: var(--text-color, inherit);
    }
    div[data-testid="stMetricLabel"],
    div[data-testid="stMetricLabel"] p,
    div[data-testid="stMetricLabel"] label {
        color: var(--text-color, inherit) !important;
        opacity: 0.78;
        font-size: 0.82rem !important;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        font-weight: 600;
    }
    div[data-testid="stMetricDelta"] { font-size: 0.85rem; opacity: 0.85; }

    /* Pills (period & active-filter summary) */
    .pill {
        display: inline-block;
        padding: 3px 12px;
        border-radius: 999px;
        background: color-mix(in srgb, var(--primary-color, #4F8DF7) 18%, transparent);
        color: var(--text-color, inherit);
        font-size: 0.80rem;
        margin-right: 6px;
        margin-bottom: 4px;
        font-weight: 500;
        border: 1px solid color-mix(in srgb, var(--primary-color, #4F8DF7) 28%, transparent);
    }
    .muted {
        color: var(--text-color, inherit);
        opacity: 0.65;
        font-size: 0.85rem;
    }

    /* Tighter tabs */
    button[data-baseweb="tab"] { padding: 8px 16px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# Data loading & cleaning
# =============================================================================
@st.cache_data(show_spinner="Reading file...")
def load_data(source, name: str) -> pd.DataFrame:
    """Load CSV or Excel file and normalize columns/types."""
    name_lower = name.lower()
    if name_lower.endswith(".csv"):
        try:
            df = pd.read_csv(source)
        except UnicodeDecodeError:
            if hasattr(source, "seek"):
                source.seek(0)
            df = pd.read_csv(source, encoding="latin-1")
    else:
        df = pd.read_excel(source)

    df.columns = [str(c).strip() for c in df.columns]

    if "Billing Date" in df.columns:
        df["Billing Date"] = pd.to_datetime(df["Billing Date"], errors="coerce")

    numeric_cols = [
        "Quantity (Actual)", "Net Price", "Net Sales Volume",
        "Tax Rate", "Tax Amount", "Cost (Actual)", "Profit Margin",
        "Profit Margin Ratio",
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    text_cols = ["Division", "Sales Empl. Name", "Customer", "Product Group",
                 "Manufacturer Name", "Billing Type", "Product Desc",
                 "Product Id", "City Name"]
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    if "City Name" in df.columns:
        df["City Name"] = df["City Name"].replace(
            {"nan": "Unknown", "": "Unknown", "None": "Unknown"}
        )

    if "Division" in df.columns and "Sales Empl. Name" in df.columns:
        df["Sub-team"] = [
            assign_sub_team(d, e)
            for d, e in zip(df["Division"], df["Sales Empl. Name"])
        ]

    return df


@st.cache_data(show_spinner="Reading secured orders...")
def load_orders_data(source, name: str) -> pd.DataFrame:
    """Load secured-orders CSV/Excel and normalize it.

    Renames `Name of Customer` -> `Customer` so it lines up with the
    sales report. Adds the same `Sub-team` column for consistency.
    """
    name_lower = name.lower()
    if name_lower.endswith(".csv"):
        try:
            df = pd.read_csv(source)
        except UnicodeDecodeError:
            if hasattr(source, "seek"):
                source.seek(0)
            df = pd.read_csv(source, encoding="latin-1")
    else:
        df = pd.read_excel(source)

    df.columns = [str(c).strip() for c in df.columns]

    # Align customer column with the sales report
    if "Name of Customer" in df.columns and "Customer" not in df.columns:
        df = df.rename(columns={"Name of Customer": "Customer"})

    for c in ["Document Date", "Requested Delivery Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    numeric_cols = [
        "Order Quantity", "Net Price", "Net Value",
        "Delivered Quantity", "Net Value (Delivered)",
        "Pending Quantity", "Net Value (Pending)",
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    text_cols = ["Division", "Sales Empl. Name", "Customer", "Product Group",
                 "Manufacturer Name", "Product Desc", "Product Id",
                 "City Name", "Sales Order"]
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    if "City Name" in df.columns:
        df["City Name"] = df["City Name"].replace(
            {"nan": "Unknown", "": "Unknown", "None": "Unknown"}
        )

    if "Division" in df.columns and "Sales Empl. Name" in df.columns:
        df["Sub-team"] = [
            assign_sub_team(d, e)
            for d, e in zip(df["Division"], df["Sales Empl. Name"])
        ]

    # Open / closed status flags
    if "Net Value (Pending)" in df.columns:
        df["Is Open"] = df["Net Value (Pending)"].fillna(0) > 0

    return df


# =============================================================================
# Formatting helpers
# =============================================================================
def fmt_money(v, currency: str = "AED") -> str:
    try:
        if pd.isna(v):
            return "-"
        return f"{currency} {v:,.2f}"
    except Exception:
        return str(v)


def fmt_int(v) -> str:
    try:
        return f"{int(round(float(v))):,}"
    except Exception:
        return str(v)


def fmt_pct(v) -> str:
    try:
        return f"{v * 100:.1f}%" if abs(v) <= 1.5 else f"{v:.1f}%"
    except Exception:
        return str(v)


def col_money(label: str, currency: str):
    return st.column_config.NumberColumn(label, format=f"{currency} %,.2f")


def col_pct(label: str):
    return st.column_config.NumberColumn(label, format="%,.2f%%")


def col_int(label: str):
    return st.column_config.NumberColumn(label, format="%,d")


def build_col_config(df: pd.DataFrame, currency: str,
                     money: list[str] | None = None,
                     pct: list[str] | None = None,
                     ints: list[str] | None = None) -> dict:
    cfg: dict = {}
    for c in money or []:
        if c in df.columns:
            cfg[c] = col_money(c, currency)
    for c in pct or []:
        if c in df.columns:
            cfg[c] = col_pct(c)
    for c in ints or []:
        if c in df.columns:
            cfg[c] = col_int(c)
    return cfg


def kpis(d: pd.DataFrame) -> dict:
    if d.empty:
        return dict(sales=0, cost=0, profit=0, qty=0, tax=0, ratio=0.0,
                    docs=0, customers=0, products=0, employees=0)
    sales = float(d["Net Sales Volume"].sum()) if "Net Sales Volume" in d else 0.0
    cost = float(d["Cost (Actual)"].sum()) if "Cost (Actual)" in d else 0.0
    profit = float(d["Profit Margin"].sum()) if "Profit Margin" in d else (sales - cost)
    qty = float(d["Quantity (Actual)"].sum()) if "Quantity (Actual)" in d else 0.0
    tax = float(d["Tax Amount"].sum()) if "Tax Amount" in d else 0.0
    ratio = (profit / sales) if sales else 0.0
    return dict(
        sales=sales, cost=cost, profit=profit, qty=qty, tax=tax, ratio=ratio,
        docs=int(d["Billing Document"].nunique()) if "Billing Document" in d else 0,
        customers=int(d["Customer"].nunique()) if "Customer" in d else 0,
        products=int(d["Product Id"].nunique()) if "Product Id" in d else 0,
        employees=int(d["Sales Empl. Name"].nunique()) if "Sales Empl. Name" in d else 0,
    )


def order_kpis(d: pd.DataFrame) -> dict:
    """Compute headline metrics for the secured-orders dataframe."""
    if d is None or d.empty:
        return dict(total_value=0.0, delivered_value=0.0, pending_value=0.0,
                    pending_qty=0.0, n_orders=0, n_open_orders=0,
                    n_lines=0, n_open_lines=0, customers=0, employees=0,
                    delivered_pct=0.0, pending_pct=0.0)
    total = float(d["Net Value"].sum()) if "Net Value" in d else 0.0
    delivered = float(d["Net Value (Delivered)"].sum()) \
        if "Net Value (Delivered)" in d else 0.0
    pending = float(d["Net Value (Pending)"].sum()) \
        if "Net Value (Pending)" in d else max(total - delivered, 0.0)
    pending_qty = float(d["Pending Quantity"].sum()) \
        if "Pending Quantity" in d else 0.0
    open_mask = d["Net Value (Pending)"] > 0 if "Net Value (Pending)" in d \
        else pd.Series([False] * len(d), index=d.index)
    return dict(
        total_value=total,
        delivered_value=delivered,
        pending_value=pending,
        pending_qty=pending_qty,
        n_orders=int(d["Sales Order"].nunique()) if "Sales Order" in d else 0,
        n_open_orders=int(d.loc[open_mask, "Sales Order"].nunique())
        if "Sales Order" in d else 0,
        n_lines=len(d),
        n_open_lines=int(open_mask.sum()),
        customers=int(d["Customer"].nunique()) if "Customer" in d else 0,
        employees=int(d["Sales Empl. Name"].nunique())
        if "Sales Empl. Name" in d else 0,
        delivered_pct=(delivered / total * 100) if total else 0.0,
        pending_pct=(pending / total * 100) if total else 0.0,
    )


# =============================================================================
# Header
# =============================================================================
st.title("Sales Analytics Tracker")
st.caption("LHC Logistics - numbers-first sales performance dashboard")


# =============================================================================
# Sidebar: data source + filters
# =============================================================================
st.sidebar.header("Data source")

uploaded = st.sidebar.file_uploader(
    "Upload sales report (Excel or CSV)",
    type=["xlsx", "xls", "csv"],
    help="Use any file with the same column structure as the original LHC sales report.",
)

df_full: pd.DataFrame | None = None
source_name: str | None = None

if uploaded is not None:
    try:
        df_full = load_data(uploaded, uploaded.name)
        source_name = uploaded.name
    except Exception as e:
        st.sidebar.error(f"Could not read uploaded file:\n{e}")
elif Path(DEFAULT_FILE).exists():
    try:
        df_full = load_data(DEFAULT_FILE, DEFAULT_FILE)
        source_name = DEFAULT_FILE
    except Exception as e:
        st.sidebar.error(f"Could not read default file:\n{e}")

if df_full is None or df_full.empty:
    st.info(
        "Upload a sales report file in the sidebar to begin. "
        "It must contain the columns shown in the original LHC report."
    )
    st.stop()

missing = [c for c in EXPECTED_COLUMNS if c not in df_full.columns]
if missing:
    st.warning(
        "Some expected columns are missing - related sections may be empty.\n\n"
        f"Missing: {', '.join(missing)}"
    )

st.sidebar.markdown(
    f"<span class='muted'>Loaded: <b>{source_name}</b></span>",
    unsafe_allow_html=True,
)
st.sidebar.markdown(
    f"<span class='muted'>{len(df_full):,} rows in file</span>",
    unsafe_allow_html=True,
)

# --- Secured orders (optional second file) ----------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("Secured orders (optional)")

uploaded_orders = st.sidebar.file_uploader(
    "Upload secured orders (Excel/CSV)",
    type=["xlsx", "xls", "csv"],
    key="orders_uploader",
    help=(
        "An open-orders / secured-orders export with columns like "
        "Sales Order, Document Date, Net Value, Net Value (Pending), "
        "Division, Sales Empl. Name."
    ),
)

df_orders_full: pd.DataFrame | None = None
orders_source: str | None = None

if uploaded_orders is not None:
    try:
        df_orders_full = load_orders_data(uploaded_orders, uploaded_orders.name)
        orders_source = uploaded_orders.name
    except Exception as e:
        st.sidebar.error(f"Could not read orders file:\n{e}")
elif Path(DEFAULT_ORDERS_FILE).exists():
    try:
        df_orders_full = load_orders_data(DEFAULT_ORDERS_FILE, DEFAULT_ORDERS_FILE)
        orders_source = DEFAULT_ORDERS_FILE
    except Exception as e:
        st.sidebar.error(f"Could not read default orders file:\n{e}")

if df_orders_full is not None and not df_orders_full.empty:
    st.sidebar.markdown(
        f"<span class='muted'>Orders: <b>{orders_source}</b></span>",
        unsafe_allow_html=True,
    )
    st.sidebar.markdown(
        f"<span class='muted'>{len(df_orders_full):,} order lines</span>",
        unsafe_allow_html=True,
    )
else:
    st.sidebar.markdown(
        "<span class='muted'>No secured-orders file loaded - upload "
        "one to unlock the Pipeline tab.</span>",
        unsafe_allow_html=True,
    )

currency = st.sidebar.text_input("Currency label", value="AED", max_chars=6).strip() or "AED"

show_charts = st.sidebar.checkbox(
    "Show summary charts",
    value=True,
    help="Toggle the headline bar charts on/off (Departments tab, "
    "Pipeline tab, etc.). The numeric tables stay visible regardless.",
)

st.sidebar.markdown("---")
st.sidebar.header("Filters")

# Date range
# Use the UNION of date ranges across the sales file and (when present)
# the secured-orders file, so "All time" really covers everything.
_date_minmax: list[date] = []
if "Billing Date" in df_full.columns and df_full["Billing Date"].notna().any():
    _date_minmax.append(df_full["Billing Date"].min().date())
    _date_minmax.append(df_full["Billing Date"].max().date())
if (df_orders_full is not None
        and "Document Date" in df_orders_full.columns
        and df_orders_full["Document Date"].notna().any()):
    _date_minmax.append(df_orders_full["Document Date"].min().date())
    _date_minmax.append(df_orders_full["Document Date"].max().date())

if _date_minmax:
    min_date = min(_date_minmax)
    max_date = max(_date_minmax)
else:
    min_date = max_date = date.today()

quick = st.sidebar.selectbox(
    "Quick period",
    ["All time", "This month", "Last 30 days", "Last 90 days", "Year to date", "Custom"],
    index=0,
)

if quick == "All time":
    start_d, end_d = min_date, max_date
elif quick == "This month":
    start_d = max_date.replace(day=1)
    end_d = max_date
elif quick == "Last 30 days":
    start_d = max(max_date - timedelta(days=30), min_date)
    end_d = max_date
elif quick == "Last 90 days":
    start_d = max(max_date - timedelta(days=90), min_date)
    end_d = max_date
elif quick == "Year to date":
    start_d = max(date(max_date.year, 1, 1), min_date)
    end_d = max_date
else:
    start_d, end_d = min_date, max_date

date_range = st.sidebar.date_input(
    "Billing date range",
    value=(start_d, end_d),
    min_value=min_date,
    max_value=max_date,
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_d, end_d = date_range
elif isinstance(date_range, date):
    start_d = end_d = date_range


def msel(label: str, col: str) -> list[str]:
    if col not in df_full.columns:
        return []
    options = sorted(df_full[col].dropna().astype(str).unique().tolist())
    return st.sidebar.multiselect(label, options=options, default=[])


f_div = msel("Department / Division", "Division")

# Sub-team filter only shown when at least one mapped sub-team has rows.
f_subteam: list[str] = []
if "Sub-team" in df_full.columns:
    available_subteams = sorted([
        s for s in df_full["Sub-team"].dropna().astype(str).unique()
        if s and s != "-"
    ])
    if available_subteams:
        f_subteam = st.sidebar.multiselect(
            "Sub-team", options=available_subteams, default=[]
        )

f_emp = msel("Sales Employee", "Sales Empl. Name")
f_pg = msel("Product Group", "Product Group")
f_man = msel("Manufacturer", "Manufacturer Name")
f_city = msel("City", "City Name")
f_btype = msel("Billing Type", "Billing Type")
f_cust = msel("Customer", "Customer")

include_credits = st.sidebar.checkbox(
    "Include credit memos & cancellations",
    value=True,
    help="Uncheck to view only invoices.",
)

# How should the sidebar date range apply to the secured-orders file?
# Default to "All open orders" so the Pipeline view reflects the full
# backlog of unfulfilled work regardless of when orders were booked,
# which matches how "pipeline" is normally read.
if df_orders_full is not None and not df_orders_full.empty:
    orders_date_mode = st.sidebar.radio(
        "Pipeline period filter",
        options=["All open orders (no date filter)",
                 "Filter by Document Date",
                 "Filter by Requested Delivery Date"],
        index=0,
        help=(
            "Pipeline normally means 'all unfulfilled work'. Switch to a "
            "date-filtered mode if you want to measure orders booked "
            "(or due) inside the selected period only."
        ),
    )
else:
    orders_date_mode = "All open orders (no date filter)"

# Apply filters
df = df_full.copy()
if "Billing Date" in df.columns:
    mask = (df["Billing Date"].dt.date >= start_d) & (df["Billing Date"].dt.date <= end_d)
    df = df[mask]


def apply_in(d: pd.DataFrame, col: str, vals: list[str]) -> pd.DataFrame:
    if vals and col in d.columns:
        return d[d[col].astype(str).isin(vals)]
    return d


df = apply_in(df, "Division", f_div)
df = apply_in(df, "Sub-team", f_subteam)
df = apply_in(df, "Sales Empl. Name", f_emp)
df = apply_in(df, "Product Group", f_pg)
df = apply_in(df, "Manufacturer Name", f_man)
df = apply_in(df, "City Name", f_city)
df = apply_in(df, "Billing Type", f_btype)
df = apply_in(df, "Customer", f_cust)

if not include_credits and "Billing Type" in df.columns:
    df = df[df["Billing Type"].str.contains("Invoice", case=False, na=False)]

# --- Apply the same filters to the orders dataframe ------------------------
# Date filter for orders is applied to "Document Date" (when the order
# was booked). Defaults to the same period the user picked.
df_orders = None
if df_orders_full is not None and not df_orders_full.empty:
    df_orders = df_orders_full.copy()

    if orders_date_mode == "Filter by Document Date" \
            and "Document Date" in df_orders.columns:
        m_o = (
            (df_orders["Document Date"].dt.date >= start_d)
            & (df_orders["Document Date"].dt.date <= end_d)
        )
        df_orders = df_orders[m_o | df_orders["Document Date"].isna()]
    elif orders_date_mode == "Filter by Requested Delivery Date" \
            and "Requested Delivery Date" in df_orders.columns:
        m_o = (
            (df_orders["Requested Delivery Date"].dt.date >= start_d)
            & (df_orders["Requested Delivery Date"].dt.date <= end_d)
        )
        df_orders = df_orders[m_o | df_orders["Requested Delivery Date"].isna()]
    # else: "All open orders (no date filter)" -> no date filtering applied

    df_orders = apply_in(df_orders, "Division", f_div)
    df_orders = apply_in(df_orders, "Sub-team", f_subteam)
    df_orders = apply_in(df_orders, "Sales Empl. Name", f_emp)
    df_orders = apply_in(df_orders, "Product Group", f_pg)
    df_orders = apply_in(df_orders, "Manufacturer Name", f_man)
    df_orders = apply_in(df_orders, "City Name", f_city)
    df_orders = apply_in(df_orders, "Customer", f_cust)

st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<span class='muted'>Filtered sales rows: <b>{len(df):,}</b></span>",
    unsafe_allow_html=True,
)
if df_orders is not None:
    st.sidebar.markdown(
        f"<span class='muted'>Filtered order lines: <b>{len(df_orders):,}</b></span>",
        unsafe_allow_html=True,
    )


# =============================================================================
# KPI block
# =============================================================================
m = kpis(df)
m_all = kpis(df_full)

period_label = f"{start_d.strftime('%d %b %Y')} -> {end_d.strftime('%d %b %Y')}"
days_n = (end_d - start_d).days + 1

filter_pills = ""
for label, vals in [
    ("Division", f_div), ("Sub-team", f_subteam),
    ("Employee", f_emp), ("Product group", f_pg),
    ("Manufacturer", f_man), ("City", f_city),
    ("Billing type", f_btype), ("Customer", f_cust),
]:
    if vals:
        v_disp = vals[0] if len(vals) == 1 else f"{len(vals)} selected"
        filter_pills += f"<span class='pill'>{label}: {v_disp}</span>"

_pipeline_pill = ""
if df_orders is not None and not df_orders.empty:
    if orders_date_mode == "All open orders (no date filter)":
        _pipeline_pill = (
            "<span class='pill'>Pipeline: all open orders</span>"
        )
    elif orders_date_mode == "Filter by Document Date":
        _pipeline_pill = (
            "<span class='pill'>Pipeline: orders booked in period</span>"
        )
    else:
        _pipeline_pill = (
            "<span class='pill'>Pipeline: deliveries due in period</span>"
        )

st.markdown(
    f"<span class='pill'>Period: {period_label}</span>"
    f"<span class='pill'>{days_n} days</span>"
    f"<span class='pill'>{len(df):,} line items</span>"
    + _pipeline_pill
    + filter_pills,
    unsafe_allow_html=True,
)
st.write("")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Net sales", fmt_money(m["sales"], currency))
c2.metric("Profit margin", fmt_money(m["profit"], currency),
          delta=fmt_pct(m["ratio"]), delta_color="off")
c3.metric("Cost", fmt_money(m["cost"], currency))
c4.metric("Tax collected", fmt_money(m["tax"], currency))
c5.metric("Units sold", fmt_int(m["qty"]))

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Invoices / docs", fmt_int(m["docs"]))
c2.metric("Customers active", fmt_int(m["customers"]))
c3.metric("Products sold", fmt_int(m["products"]))
c4.metric("Sales employees", fmt_int(m["employees"]))
avg_doc = (m["sales"] / m["docs"]) if m["docs"] else 0
c5.metric("Avg invoice value", fmt_money(avg_doc, currency))

# --- Pipeline KPI row (only when an orders file is loaded) -----------------
mo = order_kpis(df_orders) if df_orders is not None else None
if mo and mo["n_lines"] > 0:
    st.markdown(
        "<div style='margin-top: 0.6rem; color: var(--text-color, inherit); "
        "opacity: 0.65; font-size: 0.78rem; letter-spacing: 0.06em; "
        "text-transform: uppercase;'>Secured orders pipeline</div>",
        unsafe_allow_html=True,
    )
    p1, p2, p3, p4, p5 = st.columns(5)
    p1.metric("Pipeline value (pending)",
              fmt_money(mo["pending_value"], currency))
    p2.metric("Total order value",
              fmt_money(mo["total_value"], currency),
              delta=f"{mo['delivered_pct']:.1f}% delivered",
              delta_color="off")
    p3.metric("Open orders", fmt_int(mo["n_open_orders"]),
              delta=f"of {mo['n_orders']:,} total",
              delta_color="off")
    p4.metric("Combined expected (billed + pending)",
              fmt_money(m["sales"] + mo["pending_value"], currency))
    avg_pending = (mo["pending_value"] / mo["n_open_orders"]) \
        if mo["n_open_orders"] else 0
    p5.metric("Avg open-order value", fmt_money(avg_pending, currency))

st.markdown("<hr style='margin: 1.2rem 0; border-color: #e6e8ed;'>", unsafe_allow_html=True)


# =============================================================================
# Tabs
# =============================================================================
_tab_labels = ["Overview"]
if df_orders is not None and not df_orders.empty:
    _tab_labels.append("Pipeline")
_tab_labels += ["Departments", "Employees", "Customers", "Products",
                "Geography", "Time trend", "Raw data"]
tabs = st.tabs(_tab_labels)
_tab_idx = {label: i for i, label in enumerate(_tab_labels)}


# ---------- Tab: Overview --------------------------------------------------
with tabs[_tab_idx["Overview"]]:
    st.subheader("At-a-glance summary")

    if df.empty:
        st.info("No data for current filters.")
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Top 5 departments by net sales**")
            if "Division" in df.columns:
                t = (df.groupby("Division")
                     .agg(**{"Net sales": ("Net Sales Volume", "sum"),
                             "Profit": ("Profit Margin", "sum"),
                             "Invoices": ("Billing Document", "nunique")})
                     .sort_values("Net sales", ascending=False).head(5))
                t["Margin %"] = np.where(t["Net sales"] != 0,
                                         t["Profit"] / t["Net sales"] * 100, 0)
                t = t.reset_index()
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Profit"],
                        pct=["Margin %"], ints=["Invoices"],
                    ),
                )

        with col2:
            st.markdown("**Top 5 sales employees**")
            if "Sales Empl. Name" in df.columns:
                t = (df.groupby("Sales Empl. Name")
                     .agg(**{"Net sales": ("Net Sales Volume", "sum"),
                             "Profit": ("Profit Margin", "sum"),
                             "Invoices": ("Billing Document", "nunique")})
                     .sort_values("Net sales", ascending=False).head(5))
                t["Margin %"] = np.where(t["Net sales"] != 0,
                                         t["Profit"] / t["Net sales"] * 100, 0)
                t = t.reset_index().rename(columns={"Sales Empl. Name": "Employee"})
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Profit"],
                        pct=["Margin %"], ints=["Invoices"],
                    ),
                )

        st.write("")
        col3, col4 = st.columns(2)

        with col3:
            st.markdown("**Top 5 customers**")
            if "Customer" in df.columns:
                t = (df.groupby("Customer")
                     .agg(**{"Net sales": ("Net Sales Volume", "sum"),
                             "Profit": ("Profit Margin", "sum"),
                             "Invoices": ("Billing Document", "nunique")})
                     .sort_values("Net sales", ascending=False).head(5).reset_index())
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Profit"], ints=["Invoices"],
                    ),
                )

        with col4:
            st.markdown("**Top 5 product groups**")
            if "Product Group" in df.columns:
                t = (df.groupby("Product Group")
                     .agg(**{"Net sales": ("Net Sales Volume", "sum"),
                             "Profit": ("Profit Margin", "sum"),
                             "Units": ("Quantity (Actual)", "sum")})
                     .sort_values("Net sales", ascending=False).head(5).reset_index())
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Profit"], ints=["Units"],
                    ),
                )

        st.write("")
        st.markdown("**Headline numbers**")
        gross_margin_pct = (m["profit"] / m["sales"] * 100) if m["sales"] else 0
        avg_unit_price = (m["sales"] / m["qty"]) if m["qty"] else 0
        avg_lines_per_doc = (len(df) / m["docs"]) if m["docs"] else 0
        avg_sales_per_emp = (m["sales"] / m["employees"]) if m["employees"] else 0
        avg_sales_per_cust = (m["sales"] / m["customers"]) if m["customers"] else 0

        h1, h2, h3, h4, h5 = st.columns(5)
        h1.metric("Gross margin %", f"{gross_margin_pct:.2f}%")
        h2.metric("Avg unit price", fmt_money(avg_unit_price, currency))
        h3.metric("Avg lines per invoice", f"{avg_lines_per_doc:.2f}")
        h4.metric("Avg sales per employee", fmt_money(avg_sales_per_emp, currency))
        h5.metric("Avg sales per customer", fmt_money(avg_sales_per_cust, currency))


# ---------- Tab: Pipeline (secured orders) ---------------------------------
if "Pipeline" in _tab_idx:
    with tabs[_tab_idx["Pipeline"]]:
        st.subheader("Secured orders pipeline")
        st.caption(
            "All numbers below are based on the secured-orders file. The "
            "date range filters by **Document Date** (when the order was "
            "booked). Pending = order value still to deliver."
        )

        if df_orders is None or df_orders.empty:
            st.info("No secured-order rows match the current filters.")
        else:
            # Pipeline-only view toggle
            only_open = st.checkbox(
                "Show only open / pending order lines",
                value=True,
                help="When ticked, fully-delivered lines are hidden so you "
                "see exactly what's still to come in.",
            )
            d_o = df_orders[df_orders["Net Value (Pending)"] > 0] \
                if only_open and "Net Value (Pending)" in df_orders.columns \
                else df_orders

            if d_o.empty:
                st.info("No order lines after applying the open-only filter.")
            else:
                pipe = order_kpis(d_o)

                # Top KPIs (pipeline-specific)
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Pipeline value (pending)",
                          fmt_money(pipe["pending_value"], currency))
                k2.metric("Total order value",
                          fmt_money(pipe["total_value"], currency))
                k3.metric("Delivered so far",
                          fmt_money(pipe["delivered_value"], currency),
                          delta=f"{pipe['delivered_pct']:.1f}% of orders",
                          delta_color="off")
                k4.metric("Open orders", fmt_int(pipe["n_open_orders"]))
                k5.metric("Pending units", fmt_int(pipe["pending_qty"]))

                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Customers w/ pipeline",
                          fmt_int(pipe["customers"]))
                k2.metric("Employees w/ pipeline",
                          fmt_int(pipe["employees"]))
                avg_open = (pipe["pending_value"] / pipe["n_open_orders"]) \
                    if pipe["n_open_orders"] else 0
                k3.metric("Avg pending / order",
                          fmt_money(avg_open, currency))
                avg_line = (pipe["pending_value"] / pipe["n_open_lines"]) \
                    if pipe["n_open_lines"] else 0
                k4.metric("Avg pending / line",
                          fmt_money(avg_line, currency))
                # Conversion rate: delivered / total
                k5.metric("Order conversion %",
                          f"{pipe['delivered_pct']:.1f}%")

                st.markdown("---")

                # ---- Pipeline sub-tabs --------------------------------------
                _pipe_labels = [
                    "By department",
                    "By sub-team",
                    "By employee",
                    "By dept + employee",
                    "Combined (billed + pipeline)",
                    "By customer",
                    "By product / manufacturer",
                    "Aging",
                    "Time trend",
                    "Raw orders",
                ]
                pipe_tabs = st.tabs(_pipe_labels)
                _p = {label: i for i, label in enumerate(_pipe_labels)}

                # ---- By department ------------------------------------------
                with pipe_tabs[_p["By department"]]:
                    if "Division" in d_o.columns:
                        t = (d_o.groupby("Division")
                             .agg(**{
                                 "Total order value": ("Net Value", "sum"),
                                 "Delivered": ("Net Value (Delivered)", "sum"),
                                 "Pending": ("Net Value (Pending)", "sum"),
                                 "Pending units": ("Pending Quantity", "sum"),
                                 "Open orders": ("Sales Order", "nunique"),
                                 "Customers": ("Customer", "nunique"),
                                 "Employees": ("Sales Empl. Name", "nunique"),
                             }).reset_index())
                        t["Delivered %"] = np.where(
                            t["Total order value"] != 0,
                            t["Delivered"] / t["Total order value"] * 100, 0)
                        t["% of pipeline"] = np.where(
                            pipe["pending_value"],
                            t["Pending"] / pipe["pending_value"] * 100, 0)
                        t = t.sort_values("Pending", ascending=False)

                        total = pd.DataFrame([{
                            "Division": "-- TOTAL --",
                            "Total order value": t["Total order value"].sum(),
                            "Delivered": t["Delivered"].sum(),
                            "Pending": t["Pending"].sum(),
                            "Pending units": t["Pending units"].sum(),
                            "Open orders": int(d_o["Sales Order"].nunique()),
                            "Customers": int(d_o["Customer"].nunique()),
                            "Employees": int(d_o["Sales Empl. Name"].nunique()),
                            "Delivered %": (
                                t["Delivered"].sum() / t["Total order value"].sum() * 100
                            ) if t["Total order value"].sum() else 0,
                            "% of pipeline": 100.0,
                        }])
                        out = pd.concat([t, total], ignore_index=True)

                        if show_charts and not t.empty:
                            ch = t[["Division", "Pending"]].copy()
                            ch["_color"] = ch["Division"].map(dept_color)
                            st.markdown("**Pending pipeline by department**")
                            st.bar_chart(
                                ch, x="Division", y="Pending",
                                color="_color", horizontal=True, height=320,
                            )

                        st.dataframe(
                            out, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                out, currency,
                                money=["Total order value", "Delivered", "Pending"],
                                pct=["Delivered %", "% of pipeline"],
                                ints=["Pending units", "Open orders",
                                      "Customers", "Employees"],
                            ),
                        )

                        if show_charts and not t.empty:
                            with st.expander("More: Delivered vs Pending per department"):
                                stack = t[["Division", "Delivered", "Pending"]].copy()
                                st.bar_chart(
                                    stack, x="Division",
                                    y=["Delivered", "Pending"], stack=True,
                                    color=[SERIES_COLORS["Delivered"],
                                           SERIES_COLORS["Pending"]],
                                    horizontal=True, height=340,
                                )

                # ---- By sub-team --------------------------------------------
                with pipe_tabs[_p["By sub-team"]]:
                    if "Sub-team" not in d_o.columns:
                        st.info("No sub-team data.")
                    else:
                        st.markdown(
                            "Sub-team rollup. Divisions without a configured "
                            "mapping show as '-'."
                        )
                        t = (d_o.groupby(["Division", "Sub-team"])
                             .agg(**{
                                 "Total order value": ("Net Value", "sum"),
                                 "Delivered": ("Net Value (Delivered)", "sum"),
                                 "Pending": ("Net Value (Pending)", "sum"),
                                 "Open orders": ("Sales Order", "nunique"),
                                 "Customers": ("Customer", "nunique"),
                                 "Employees": ("Sales Empl. Name", "nunique"),
                             }).reset_index())
                        t["Delivered %"] = np.where(
                            t["Total order value"] != 0,
                            t["Delivered"] / t["Total order value"] * 100, 0)
                        t = t.sort_values(
                            ["Division", "Pending"],
                            ascending=[True, False],
                        )
                        st.dataframe(
                            t, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                t, currency,
                                money=["Total order value", "Delivered", "Pending"],
                                pct=["Delivered %"],
                                ints=["Open orders", "Customers", "Employees"],
                            ),
                        )

                # ---- By employee --------------------------------------------
                with pipe_tabs[_p["By employee"]]:
                    sort_pipe = st.selectbox(
                        "Sort employees by",
                        options=["Pending", "Total order value",
                                 "Delivered", "Open orders"],
                        index=0, key="pipe_emp_sort",
                    )
                    group_cols_e = ["Sales Empl. Name"]
                    if "Sub-team" in d_o.columns:
                        group_cols_e = ["Sales Empl. Name", "Sub-team"]
                    t = (d_o.groupby(group_cols_e, dropna=False)
                         .agg(**{
                             "Total order value": ("Net Value", "sum"),
                             "Delivered": ("Net Value (Delivered)", "sum"),
                             "Pending": ("Net Value (Pending)", "sum"),
                             "Open orders": ("Sales Order", "nunique"),
                             "Customers": ("Customer", "nunique"),
                             "Departments": ("Division", "nunique"),
                         }).reset_index())
                    t["Delivered %"] = np.where(
                        t["Total order value"] != 0,
                        t["Delivered"] / t["Total order value"] * 100, 0)
                    t["% of pipeline"] = np.where(
                        pipe["pending_value"],
                        t["Pending"] / pipe["pending_value"] * 100, 0)
                    t = t.sort_values(sort_pipe, ascending=False)
                    t = t.rename(columns={"Sales Empl. Name": "Employee"})
                    t.insert(0, "Rank", range(1, len(t) + 1))
                    st.dataframe(
                        t, hide_index=True, use_container_width=True,
                        column_config=build_col_config(
                            t, currency,
                            money=["Total order value", "Delivered", "Pending"],
                            pct=["Delivered %", "% of pipeline"],
                            ints=["Rank", "Open orders", "Customers", "Departments"],
                        ),
                    )

                # ---- By dept + employee (combined) --------------------------
                with pipe_tabs[_p["By dept + employee"]]:
                    st.markdown(
                        "Every employee grouped under their **department**. "
                        "Reps are ranked within their own department, with a "
                        "rolled-up department total at the bottom of each "
                        "group. Useful when you want to see how a team's reps "
                        "stack against each other rather than across the "
                        "whole company."
                    )
                    if "Division" not in d_o.columns or "Sales Empl. Name" not in d_o.columns:
                        st.info("Department or employee column missing.")
                    else:
                        sort_de = st.selectbox(
                            "Sort employees within each department by",
                            options=["Pending", "Total order value",
                                     "Delivered", "Open orders"],
                            index=0, key="pipe_dept_emp_sort",
                        )
                        group_cols_de = ["Division", "Sales Empl. Name"]
                        if "Sub-team" in d_o.columns:
                            group_cols_de.append("Sub-team")
                        t_de = (d_o.groupby(group_cols_de, dropna=False)
                                .agg(**{
                                    "Total order value": ("Net Value", "sum"),
                                    "Delivered": ("Net Value (Delivered)", "sum"),
                                    "Pending": ("Net Value (Pending)", "sum"),
                                    "Open orders": ("Sales Order", "nunique"),
                                    "Customers": ("Customer", "nunique"),
                                }).reset_index())
                        t_de["Delivered %"] = np.where(
                            t_de["Total order value"] != 0,
                            t_de["Delivered"] / t_de["Total order value"] * 100, 0)
                        t_de["% of pipeline"] = np.where(
                            pipe["pending_value"],
                            t_de["Pending"] / pipe["pending_value"] * 100, 0)

                        # Department-level totals for the % of dept share
                        dept_totals = (t_de.groupby("Division")["Pending"]
                                       .sum().to_dict())
                        t_de["% of dept pending"] = t_de.apply(
                            lambda r: (
                                r["Pending"] / dept_totals[r["Division"]] * 100
                            ) if dept_totals.get(r["Division"]) else 0,
                            axis=1,
                        )

                        # Sort: Division alpha, then chosen metric desc inside
                        t_de = t_de.sort_values(
                            ["Division", sort_de],
                            ascending=[True, False],
                        )
                        # Rank within each department
                        t_de.insert(0, "Rank in dept",
                                    t_de.groupby("Division").cumcount() + 1)

                        # Build a department-total summary row inserted after
                        # each department group, for at-a-glance reading
                        rows: list[pd.Series | dict] = []
                        for div, sub in t_de.groupby("Division", sort=False):
                            for _, r in sub.iterrows():
                                rows.append(r.to_dict())
                            rows.append({
                                "Rank in dept": "",
                                "Division": div,
                                "Sales Empl. Name": "-- Department total --",
                                "Sub-team": "",
                                "Total order value": sub["Total order value"].sum(),
                                "Delivered": sub["Delivered"].sum(),
                                "Pending": sub["Pending"].sum(),
                                "Open orders": int(
                                    d_o.loc[d_o["Division"] == div,
                                            "Sales Order"].nunique()
                                ),
                                "Customers": int(
                                    d_o.loc[d_o["Division"] == div,
                                            "Customer"].nunique()
                                ),
                                "Delivered %": (
                                    sub["Delivered"].sum()
                                    / sub["Total order value"].sum() * 100
                                ) if sub["Total order value"].sum() else 0,
                                "% of pipeline": (
                                    sub["Pending"].sum()
                                    / pipe["pending_value"] * 100
                                ) if pipe["pending_value"] else 0,
                                "% of dept pending": 100.0,
                            })
                        out_de = pd.DataFrame(rows)
                        out_de = out_de.rename(
                            columns={"Sales Empl. Name": "Employee",
                                     "Division": "Department"}
                        )

                        display_de = ["Rank in dept", "Department", "Employee"]
                        if "Sub-team" in out_de.columns:
                            display_de.append("Sub-team")
                        display_de += ["Pending", "Total order value",
                                       "Delivered", "Delivered %",
                                       "% of dept pending", "% of pipeline",
                                       "Open orders", "Customers"]
                        out_de = out_de[[c for c in display_de
                                         if c in out_de.columns]]

                        if show_charts and not t_de.empty:
                            dept_totals_chart = (
                                t_de.groupby("Division")["Pending"]
                                .sum().reset_index()
                                .sort_values("Pending", ascending=False)
                            )
                            dept_totals_chart["_color"] = (
                                dept_totals_chart["Division"].map(dept_color)
                            )
                            st.markdown("**Department pipeline totals**")
                            st.bar_chart(
                                dept_totals_chart,
                                x="Division", y="Pending",
                                color="_color", horizontal=True, height=300,
                            )

                        st.dataframe(
                            out_de, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                out_de, currency,
                                money=["Total order value", "Delivered", "Pending"],
                                pct=["Delivered %", "% of dept pending",
                                     "% of pipeline"],
                                ints=["Open orders", "Customers"],
                            ),
                            height=560,
                        )

                # ---- Combined (billed + pipeline) ---------------------------
                with pipe_tabs[_p["Combined (billed + pipeline)"]]:
                    _billed_scope = (
                        f"period {start_d.strftime('%d %b %Y')} - "
                        f"{end_d.strftime('%d %b %Y')}"
                    )
                    if orders_date_mode == "All open orders (no date filter)":
                        _pipe_scope = "all open orders (any date)"
                    elif orders_date_mode == "Filter by Document Date":
                        _pipe_scope = "orders booked in selected period"
                    else:
                        _pipe_scope = "deliveries due in selected period"
                    st.markdown(
                        f"**Billed sales** = invoiced sales for **{_billed_scope}**. "
                        f"**Pending** = open pipeline ({_pipe_scope}). "
                        "Division/employee filters apply to both. "
                        "Use the *Pipeline period filter* in the sidebar "
                        "to change how the date range applies to orders."
                    )
                    if "Sales Empl. Name" in df.columns and "Sales Empl. Name" in d_o.columns:
                        billed = (df.groupby("Sales Empl. Name")
                                  .agg(**{"Billed sales":
                                          ("Net Sales Volume", "sum")})
                                  .reset_index())
                        pending = (d_o.groupby("Sales Empl. Name")
                                   .agg(**{
                                       "Pending":
                                           ("Net Value (Pending)", "sum"),
                                       "Open orders":
                                           ("Sales Order", "nunique"),
                                   }).reset_index())
                        merged = billed.merge(
                            pending, on="Sales Empl. Name", how="outer"
                        ).fillna(0)
                        merged["Combined"] = (
                            merged["Billed sales"] + merged["Pending"]
                        )
                        merged["Pipeline mix %"] = np.where(
                            merged["Combined"] != 0,
                            merged["Pending"] / merged["Combined"] * 100, 0,
                        )
                        # Add sub-team
                        if "Sub-team" in df_full.columns:
                            emp_st = (
                                pd.concat([
                                    df_full[["Sales Empl. Name", "Sub-team"]],
                                    df_orders_full[["Sales Empl. Name", "Sub-team"]]
                                    if df_orders_full is not None
                                    else df[["Sales Empl. Name", "Sub-team"]],
                                ], ignore_index=True)
                                .drop_duplicates(subset=["Sales Empl. Name"])
                            )
                            merged = merged.merge(
                                emp_st, on="Sales Empl. Name", how="left",
                            )
                        merged = merged.sort_values("Combined", ascending=False)
                        merged = merged.rename(
                            columns={"Sales Empl. Name": "Employee"}
                        )
                        merged.insert(0, "Rank", range(1, len(merged) + 1))

                        display_cols = ["Rank", "Employee"]
                        if "Sub-team" in merged.columns:
                            display_cols.append("Sub-team")
                        display_cols += ["Billed sales", "Pending",
                                         "Combined", "Pipeline mix %",
                                         "Open orders"]
                        merged = merged[[c for c in display_cols
                                         if c in merged.columns]]
                        st.dataframe(
                            merged, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                merged, currency,
                                money=["Billed sales", "Pending", "Combined"],
                                pct=["Pipeline mix %"],
                                ints=["Rank", "Open orders"],
                            ),
                        )

                        # Combined by department too
                        st.markdown("**Combined by department**")
                        b2 = (df.groupby("Division")
                              .agg(**{"Billed sales":
                                      ("Net Sales Volume", "sum")})
                              .reset_index())
                        p2 = (d_o.groupby("Division")
                              .agg(**{
                                  "Pending":
                                      ("Net Value (Pending)", "sum"),
                                  "Open orders":
                                      ("Sales Order", "nunique"),
                              }).reset_index())
                        m2 = b2.merge(p2, on="Division", how="outer").fillna(0)
                        m2["Combined"] = m2["Billed sales"] + m2["Pending"]
                        m2["Pipeline mix %"] = np.where(
                            m2["Combined"] != 0,
                            m2["Pending"] / m2["Combined"] * 100, 0,
                        )
                        m2 = m2.sort_values("Combined", ascending=False)

                        total = pd.DataFrame([{
                            "Division": "-- TOTAL --",
                            "Billed sales": m2["Billed sales"].sum(),
                            "Pending": m2["Pending"].sum(),
                            "Open orders": m2["Open orders"].sum(),
                            "Combined": m2["Combined"].sum(),
                            "Pipeline mix %": (
                                m2["Pending"].sum() / m2["Combined"].sum() * 100
                            ) if m2["Combined"].sum() else 0,
                        }])
                        m2_out = pd.concat([m2, total], ignore_index=True)
                        st.dataframe(
                            m2_out, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                m2_out, currency,
                                money=["Billed sales", "Pending", "Combined"],
                                pct=["Pipeline mix %"],
                                ints=["Open orders"],
                            ),
                        )

                        if show_charts and not m2.empty:
                            st.markdown(
                                "**Billed vs pending per department**"
                            )
                            ch = m2[["Division", "Billed sales", "Pending"]].copy()
                            st.bar_chart(
                                ch, x="Division",
                                y=["Billed sales", "Pending"], stack=True,
                                color=[SERIES_COLORS["Billed sales"],
                                       SERIES_COLORS["Pending"]],
                                horizontal=True, height=340,
                            )

                # ---- By customer --------------------------------------------
                with pipe_tabs[_p["By customer"]]:
                    top_n_c = st.slider(
                        "Show top N customers",
                        10, 500, 50, 10, key="pipe_cust_n",
                    )
                    t = (d_o.groupby("Customer")
                         .agg(**{
                             "Total order value": ("Net Value", "sum"),
                             "Delivered": ("Net Value (Delivered)", "sum"),
                             "Pending": ("Net Value (Pending)", "sum"),
                             "Open orders": ("Sales Order", "nunique"),
                             "Last order": ("Document Date", "max"),
                         }).reset_index())
                    t["Delivered %"] = np.where(
                        t["Total order value"] != 0,
                        t["Delivered"] / t["Total order value"] * 100, 0)
                    t = t.sort_values("Pending", ascending=False).head(top_n_c)
                    t.insert(0, "Rank", range(1, len(t) + 1))

                    cfg = build_col_config(
                        t, currency,
                        money=["Total order value", "Delivered", "Pending"],
                        pct=["Delivered %"],
                        ints=["Rank", "Open orders"],
                    )
                    if "Last order" in t.columns:
                        cfg["Last order"] = st.column_config.DateColumn(
                            "Last order", format="DD MMM YYYY"
                        )
                    st.dataframe(
                        t, hide_index=True, use_container_width=True,
                        column_config=cfg,
                    )

                # ---- By product / manufacturer ------------------------------
                with pipe_tabs[_p["By product / manufacturer"]]:
                    sub = st.tabs(["Product group", "Manufacturer", "Product"])

                    with sub[0]:
                        if "Product Group" in d_o.columns:
                            t = (d_o.groupby("Product Group")
                                 .agg(**{
                                     "Total order value": ("Net Value", "sum"),
                                     "Pending": ("Net Value (Pending)", "sum"),
                                     "Open orders": ("Sales Order", "nunique"),
                                     "Distinct products":
                                         ("Product Id", "nunique"),
                                 }).reset_index()
                                 .sort_values("Pending", ascending=False))
                            st.dataframe(
                                t, hide_index=True, use_container_width=True,
                                column_config=build_col_config(
                                    t, currency,
                                    money=["Total order value", "Pending"],
                                    ints=["Open orders", "Distinct products"],
                                ),
                            )

                    with sub[1]:
                        if "Manufacturer Name" in d_o.columns:
                            t = (d_o.groupby("Manufacturer Name")
                                 .agg(**{
                                     "Total order value": ("Net Value", "sum"),
                                     "Pending": ("Net Value (Pending)", "sum"),
                                     "Open orders": ("Sales Order", "nunique"),
                                     "Distinct products":
                                         ("Product Id", "nunique"),
                                 }).reset_index()
                                 .rename(columns={"Manufacturer Name":
                                                  "Manufacturer"})
                                 .sort_values("Pending", ascending=False))
                            st.dataframe(
                                t, hide_index=True, use_container_width=True,
                                column_config=build_col_config(
                                    t, currency,
                                    money=["Total order value", "Pending"],
                                    ints=["Open orders", "Distinct products"],
                                ),
                            )

                    with sub[2]:
                        if "Product Id" in d_o.columns:
                            top_n_p = st.slider(
                                "Top N products",
                                10, 500, 50, 10, key="pipe_prod_n",
                            )
                            grp = ["Product Id"]
                            if "Product Desc" in d_o.columns:
                                grp.append("Product Desc")
                            t = (d_o.groupby(grp)
                                 .agg(**{
                                     "Total order value": ("Net Value", "sum"),
                                     "Pending": ("Net Value (Pending)", "sum"),
                                     "Pending units":
                                         ("Pending Quantity", "sum"),
                                     "Open orders": ("Sales Order", "nunique"),
                                 }).reset_index()
                                 .sort_values("Pending", ascending=False)
                                 .head(top_n_p))
                            st.dataframe(
                                t, hide_index=True, use_container_width=True,
                                column_config=build_col_config(
                                    t, currency,
                                    money=["Total order value", "Pending"],
                                    ints=["Pending units", "Open orders"],
                                ),
                            )

                # ---- Aging --------------------------------------------------
                with pipe_tabs[_p["Aging"]]:
                    st.markdown(
                        "Open order lines bucketed by age. Age is measured "
                        "from **Requested Delivery Date** when available, "
                        "else from **Document Date**. Negative age means "
                        "the order is past its requested delivery."
                    )
                    today = pd.Timestamp.today().normalize()
                    open_only = d_o[d_o["Net Value (Pending)"] > 0].copy() \
                        if "Net Value (Pending)" in d_o.columns else d_o.copy()

                    if open_only.empty:
                        st.info("No open order lines.")
                    else:
                        ref_date = open_only.get(
                            "Requested Delivery Date",
                            open_only.get("Document Date"),
                        )
                        if "Requested Delivery Date" in open_only.columns:
                            open_only["Days to delivery"] = (
                                open_only["Requested Delivery Date"] - today
                            ).dt.days
                        elif "Document Date" in open_only.columns:
                            open_only["Days to delivery"] = (
                                today - open_only["Document Date"]
                            ).dt.days * -1

                        def _bucket(d):
                            if pd.isna(d):
                                return "Unknown"
                            if d < -90:
                                return "Overdue 90+ days"
                            if d < -30:
                                return "Overdue 30-90 days"
                            if d < 0:
                                return "Overdue 0-30 days"
                            if d <= 30:
                                return "Due within 30 days"
                            if d <= 90:
                                return "Due in 30-90 days"
                            return "Due 90+ days out"

                        open_only["Bucket"] = open_only["Days to delivery"].apply(_bucket)
                        bucket_order = [
                            "Overdue 90+ days", "Overdue 30-90 days",
                            "Overdue 0-30 days", "Due within 30 days",
                            "Due in 30-90 days", "Due 90+ days out",
                            "Unknown",
                        ]
                        agg = (open_only.groupby("Bucket")
                               .agg(**{
                                   "Pending": ("Net Value (Pending)", "sum"),
                                   "Open orders": ("Sales Order", "nunique"),
                                   "Lines": ("Sales Order", "size"),
                                   "Customers": ("Customer", "nunique"),
                                   "Employees":
                                       ("Sales Empl. Name", "nunique"),
                               }).reset_index())
                        agg["Bucket"] = pd.Categorical(
                            agg["Bucket"], categories=bucket_order, ordered=True,
                        )
                        agg = agg.sort_values("Bucket")
                        agg["% of pipeline"] = np.where(
                            pipe["pending_value"],
                            agg["Pending"] / pipe["pending_value"] * 100, 0)

                        if show_charts and not agg.empty:
                            ch_age = agg[["Bucket", "Pending"]].copy()
                            ch_age["Bucket"] = ch_age["Bucket"].astype(str)
                            ch_age["_color"] = ch_age["Bucket"].map(
                                BUCKET_COLORS
                            ).fillna(DEFAULT_DEPT_COLOR)
                            st.markdown("**Pending pipeline by aging bucket**")
                            st.bar_chart(
                                ch_age, x="Bucket", y="Pending",
                                color="_color", height=300,
                            )

                        st.dataframe(
                            agg, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                agg, currency,
                                money=["Pending"],
                                pct=["% of pipeline"],
                                ints=["Open orders", "Lines",
                                      "Customers", "Employees"],
                            ),
                        )

                # ---- Time trend ---------------------------------------------
                with pipe_tabs[_p["Time trend"]]:
                    if "Document Date" not in d_o.columns:
                        st.info("No date data on orders.")
                    else:
                        gran = st.radio(
                            "Granularity",
                            ["Daily", "Weekly", "Monthly", "Quarterly"],
                            index=2, horizontal=True, key="pipe_gran",
                        )
                        d2 = d_o.copy()
                        if gran == "Daily":
                            d2["Period"] = d2["Document Date"].dt.date.astype(str)
                        elif gran == "Weekly":
                            d2["Period"] = (
                                d2["Document Date"].dt.to_period("W")
                                .apply(lambda p: p.start_time.strftime("%Y-%m-%d"))
                            )
                        elif gran == "Monthly":
                            d2["Period"] = (
                                d2["Document Date"].dt.to_period("M").astype(str)
                            )
                        else:
                            d2["Period"] = (
                                d2["Document Date"].dt.to_period("Q").astype(str)
                            )
                        t = (d2.groupby("Period")
                             .agg(**{
                                 "Total order value": ("Net Value", "sum"),
                                 "Delivered": ("Net Value (Delivered)", "sum"),
                                 "Pending": ("Net Value (Pending)", "sum"),
                                 "Open orders": ("Sales Order", "nunique"),
                             }).reset_index()
                             .sort_values("Period"))
                        if show_charts and not t.empty:
                            st.markdown(
                                "**Pending and delivered over time**"
                            )
                            try:
                                st.bar_chart(
                                    t, x="Period",
                                    y=["Pending", "Delivered"], stack=True,
                                    color=[SERIES_COLORS["Pending"],
                                           SERIES_COLORS["Delivered"]],
                                    height=300,
                                )
                            except Exception:
                                pass

                        st.dataframe(
                            t, hide_index=True, use_container_width=True,
                            column_config=build_col_config(
                                t, currency,
                                money=["Total order value", "Delivered", "Pending"],
                                ints=["Open orders"],
                            ),
                        )

                # ---- Raw orders ---------------------------------------------
                with pipe_tabs[_p["Raw orders"]]:
                    st.markdown(
                        f"<span class='muted'>{len(d_o):,} order lines after "
                        "filters. Click column headers to sort.</span>",
                        unsafe_allow_html=True,
                    )
                    cols = [c for c in [
                        "Sales Order", "Document Date",
                        "Requested Delivery Date", "Division", "Sub-team",
                        "Sales Empl. Name", "Customer", "City Name",
                        "Product Id", "Product Desc", "Product Group",
                        "Manufacturer Name", "Order Quantity", "Net Price",
                        "Net Value", "Delivered Quantity",
                        "Net Value (Delivered)", "Pending Quantity",
                        "Net Value (Pending)",
                    ] if c in d_o.columns]
                    show_orders = d_o[cols].copy()
                    cfg = build_col_config(
                        show_orders, currency,
                        money=["Net Price", "Net Value",
                               "Net Value (Delivered)", "Net Value (Pending)"],
                        ints=["Order Quantity", "Delivered Quantity",
                              "Pending Quantity"],
                    )
                    if "Document Date" in show_orders.columns:
                        cfg["Document Date"] = st.column_config.DateColumn(
                            "Document Date", format="DD MMM YYYY",
                        )
                    if "Requested Delivery Date" in show_orders.columns:
                        cfg["Requested Delivery Date"] = \
                            st.column_config.DateColumn(
                                "Requested Delivery Date",
                                format="DD MMM YYYY",
                            )
                    st.dataframe(
                        show_orders, hide_index=True,
                        use_container_width=True,
                        column_config=cfg, height=520,
                    )
                    csv_o = show_orders.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "Download filtered orders (CSV)",
                        data=csv_o,
                        file_name=(
                            f"orders_filtered_{start_d.isoformat()}"
                            f"_to_{end_d.isoformat()}.csv"
                        ),
                        mime="text/csv",
                    )


# ---------- Tab: Departments -----------------------------------------------
with tabs[_tab_idx["Departments"]]:
    st.subheader("Department / Division breakdown")
    if "Division" not in df.columns or df.empty:
        st.info("No data for current filters.")
    else:
        agg = (df.groupby("Division")
               .agg(**{
                   "Net sales": ("Net Sales Volume", "sum"),
                   "Cost": ("Cost (Actual)", "sum"),
                   "Profit": ("Profit Margin", "sum"),
                   "Tax": ("Tax Amount", "sum"),
                   "Units": ("Quantity (Actual)", "sum"),
                   "Invoices": ("Billing Document", "nunique"),
                   "Customers": ("Customer", "nunique"),
                   "Employees": ("Sales Empl. Name", "nunique"),
               }).reset_index())
        agg["Margin %"] = np.where(agg["Net sales"] != 0,
                                   agg["Profit"] / agg["Net sales"] * 100, 0)
        agg["% of total sales"] = np.where(
            m["sales"], agg["Net sales"] / m["sales"] * 100, 0
        )
        agg = agg.sort_values("Net sales", ascending=False)

        total_row = pd.DataFrame([{
            "Division": "-- TOTAL --",
            "Net sales": agg["Net sales"].sum(),
            "Cost": agg["Cost"].sum(),
            "Profit": agg["Profit"].sum(),
            "Tax": agg["Tax"].sum(),
            "Units": agg["Units"].sum(),
            "Invoices": int(df["Billing Document"].nunique()),
            "Customers": int(df["Customer"].nunique()),
            "Employees": int(df["Sales Empl. Name"].nunique()),
            "Margin %": (agg["Profit"].sum() / agg["Net sales"].sum() * 100)
            if agg["Net sales"].sum() else 0,
            "% of total sales": 100.0,
        }])
        out = pd.concat([agg, total_row], ignore_index=True)

        if show_charts and not agg.empty:
            chart_df = agg[["Division", "Net sales", "Profit"]].copy()
            chart_df["_color"] = chart_df["Division"].map(dept_color)
            st.markdown("**Net sales by department**")
            st.bar_chart(
                chart_df, x="Division", y="Net sales",
                color="_color", horizontal=True, height=320,
            )

        st.dataframe(
            out, hide_index=True, use_container_width=True,
            column_config=build_col_config(
                out, currency,
                money=["Net sales", "Cost", "Profit", "Tax"],
                pct=["Margin %", "% of total sales"],
                ints=["Units", "Invoices", "Customers", "Employees"],
            ),
        )

        if show_charts and not agg.empty:
            with st.expander("More charts: Profit and Margin %"):
                c_left, c_right = st.columns(2)
                with c_left:
                    st.markdown("**Profit by department**")
                    p_df = agg[["Division", "Profit"]].copy()
                    p_df["_color"] = p_df["Division"].map(dept_color)
                    st.bar_chart(
                        p_df, x="Division", y="Profit",
                        color="_color", horizontal=True, height=300,
                    )
                with c_right:
                    st.markdown("**Margin % by department**")
                    mg_df = agg[["Division", "Margin %"]].copy()
                    mg_df["_color"] = mg_df["Division"].map(dept_color)
                    st.bar_chart(
                        mg_df, x="Division", y="Margin %",
                        color="_color", horizontal=True, height=300,
                    )

        st.write("")
        st.markdown("**Drill down: select a department to see its employees**")
        div_pick = st.selectbox(
            "Department",
            options=["(all)"] + sorted(df["Division"].unique().tolist()),
        )
        sub = df if div_pick == "(all)" else df[df["Division"] == div_pick]

        # If a single division with a sub-team mapping is selected,
        # show the sub-team summary first, then the per-employee rows.
        has_sub_teams = (
            div_pick != "(all)"
            and div_pick in SUB_TEAM_MAP
            and "Sub-team" in sub.columns
            and not sub.empty
        )

        if has_sub_teams:
            st.markdown(f"**Sub-team summary - {div_pick}**")
            st_agg = (sub.groupby("Sub-team")
                      .agg(**{
                          "Net sales": ("Net Sales Volume", "sum"),
                          "Cost": ("Cost (Actual)", "sum"),
                          "Profit": ("Profit Margin", "sum"),
                          "Units": ("Quantity (Actual)", "sum"),
                          "Invoices": ("Billing Document", "nunique"),
                          "Customers": ("Customer", "nunique"),
                          "Employees": ("Sales Empl. Name", "nunique"),
                      }).reset_index())
            st_agg["Margin %"] = np.where(
                st_agg["Net sales"] != 0,
                st_agg["Profit"] / st_agg["Net sales"] * 100, 0,
            )
            div_sales = float(st_agg["Net sales"].sum())
            st_agg["% of dept sales"] = np.where(
                div_sales, st_agg["Net sales"] / div_sales * 100, 0,
            )
            st_agg = st_agg.sort_values("Net sales", ascending=False)

            st_total = pd.DataFrame([{
                "Sub-team": "-- TOTAL --",
                "Net sales": st_agg["Net sales"].sum(),
                "Cost": st_agg["Cost"].sum(),
                "Profit": st_agg["Profit"].sum(),
                "Units": st_agg["Units"].sum(),
                "Invoices": int(sub["Billing Document"].nunique()),
                "Customers": int(sub["Customer"].nunique()),
                "Employees": int(sub["Sales Empl. Name"].nunique()),
                "Margin %": (st_agg["Profit"].sum() / st_agg["Net sales"].sum() * 100)
                if st_agg["Net sales"].sum() else 0,
                "% of dept sales": 100.0,
            }])
            st_out = pd.concat([st_agg, st_total], ignore_index=True)

            st.dataframe(
                st_out, hide_index=True, use_container_width=True,
                column_config=build_col_config(
                    st_out, currency,
                    money=["Net sales", "Cost", "Profit"],
                    pct=["Margin %", "% of dept sales"],
                    ints=["Units", "Invoices", "Customers", "Employees"],
                ),
            )

            st.write("")
            st.markdown(f"**Employees by sub-team - {div_pick}**")
            sub_team_pick = st.selectbox(
                "Sub-team",
                options=["(all sub-teams)"] + st_agg["Sub-team"].tolist(),
            )
            sub2 = sub if sub_team_pick == "(all sub-teams)" \
                else sub[sub["Sub-team"] == sub_team_pick]

            t = (sub2.groupby(["Sub-team", "Sales Empl. Name"])
                 .agg(**{
                     "Net sales": ("Net Sales Volume", "sum"),
                     "Profit": ("Profit Margin", "sum"),
                     "Units": ("Quantity (Actual)", "sum"),
                     "Invoices": ("Billing Document", "nunique"),
                     "Customers": ("Customer", "nunique"),
                 }).reset_index())
            t["Margin %"] = np.where(t["Net sales"] != 0,
                                     t["Profit"] / t["Net sales"] * 100, 0)
            t = t.sort_values(["Sub-team", "Net sales"],
                              ascending=[True, False])
            t = t.rename(columns={"Sales Empl. Name": "Employee"})
            st.dataframe(
                t, hide_index=True, use_container_width=True,
                column_config=build_col_config(
                    t, currency,
                    money=["Net sales", "Profit"],
                    pct=["Margin %"],
                    ints=["Units", "Invoices", "Customers"],
                ),
            )

        elif "Sales Empl. Name" in sub.columns and not sub.empty:
            t = (sub.groupby(["Division", "Sales Empl. Name"])
                 .agg(**{
                     "Net sales": ("Net Sales Volume", "sum"),
                     "Profit": ("Profit Margin", "sum"),
                     "Units": ("Quantity (Actual)", "sum"),
                     "Invoices": ("Billing Document", "nunique"),
                     "Customers": ("Customer", "nunique"),
                 }).reset_index())
            t["Margin %"] = np.where(t["Net sales"] != 0,
                                     t["Profit"] / t["Net sales"] * 100, 0)
            t = t.sort_values(["Division", "Net sales"], ascending=[True, False])
            t = t.rename(columns={"Sales Empl. Name": "Employee"})
            st.dataframe(
                t, hide_index=True, use_container_width=True,
                column_config=build_col_config(
                    t, currency,
                    money=["Net sales", "Profit"],
                    pct=["Margin %"],
                    ints=["Units", "Invoices", "Customers"],
                ),
            )

            if div_pick == "(all)":
                st.caption(
                    "Tip: pick a single department above to see its sub-team "
                    "breakdown (configured for "
                    f"{', '.join(sorted(SUB_TEAM_MAP.keys())) or 'no departments yet'})."
                )


# ---------- Tab: Employees -------------------------------------------------
with tabs[_tab_idx["Employees"]]:
    st.subheader("Sales employee leaderboard")
    if "Sales Empl. Name" not in df.columns or df.empty:
        st.info("No data for current filters.")
    else:
        sort_by = st.selectbox(
            "Sort by",
            options=["Net sales", "Profit", "Margin %", "Invoices",
                     "Units", "Customers"],
            index=0,
        )

        has_subteam_col = "Sub-team" in df.columns

        group_cols = ["Sales Empl. Name"]
        if has_subteam_col:
            group_cols = ["Sales Empl. Name", "Sub-team"]

        agg = (df.groupby(group_cols, dropna=False)
               .agg(**{
                   "Net sales": ("Net Sales Volume", "sum"),
                   "Cost": ("Cost (Actual)", "sum"),
                   "Profit": ("Profit Margin", "sum"),
                   "Units": ("Quantity (Actual)", "sum"),
                   "Invoices": ("Billing Document", "nunique"),
                   "Customers": ("Customer", "nunique"),
                   "Departments": ("Division", "nunique"),
               }).reset_index())
        agg["Margin %"] = np.where(agg["Net sales"] != 0,
                                   agg["Profit"] / agg["Net sales"] * 100, 0)
        agg["Avg invoice value"] = np.where(
            agg["Invoices"] != 0, agg["Net sales"] / agg["Invoices"], 0
        )
        agg["% of period sales"] = np.where(
            m["sales"], agg["Net sales"] / m["sales"] * 100, 0
        )
        agg = agg.sort_values(sort_by, ascending=False)
        agg = agg.rename(columns={"Sales Empl. Name": "Employee"})
        agg.insert(0, "Rank", range(1, len(agg) + 1))

        if has_subteam_col:
            display_cols = ["Rank", "Employee", "Sub-team",
                            "Net sales", "Cost", "Profit", "Margin %",
                            "Avg invoice value", "Units", "Invoices",
                            "Customers", "Departments", "% of period sales"]
            agg = agg[[c for c in display_cols if c in agg.columns]]

        st.dataframe(
            agg, hide_index=True, use_container_width=True,
            column_config=build_col_config(
                agg, currency,
                money=["Net sales", "Cost", "Profit", "Avg invoice value"],
                pct=["Margin %", "% of period sales"],
                ints=["Rank", "Units", "Invoices", "Customers", "Departments"],
            ),
        )


# ---------- Tab: Customers -------------------------------------------------
with tabs[_tab_idx["Customers"]]:
    st.subheader("Customer breakdown")
    if "Customer" not in df.columns or df.empty:
        st.info("No data for current filters.")
    else:
        top_n = st.slider("Show top N customers", 10, 500, 50, 10)
        agg = (df.groupby("Customer")
               .agg(**{
                   "Net sales": ("Net Sales Volume", "sum"),
                   "Cost": ("Cost (Actual)", "sum"),
                   "Profit": ("Profit Margin", "sum"),
                   "Units": ("Quantity (Actual)", "sum"),
                   "Invoices": ("Billing Document", "nunique"),
                   "Last invoice": ("Billing Date", "max"),
               }).reset_index())
        agg["Margin %"] = np.where(agg["Net sales"] != 0,
                                   agg["Profit"] / agg["Net sales"] * 100, 0)
        agg["% of period sales"] = np.where(
            m["sales"], agg["Net sales"] / m["sales"] * 100, 0
        )
        agg = agg.sort_values("Net sales", ascending=False).head(top_n)
        agg.insert(0, "Rank", range(1, len(agg) + 1))

        cfg = build_col_config(
            agg, currency,
            money=["Net sales", "Cost", "Profit"],
            pct=["Margin %", "% of period sales"],
            ints=["Rank", "Units", "Invoices"],
        )
        if "Last invoice" in agg.columns:
            cfg["Last invoice"] = st.column_config.DateColumn(
                "Last invoice", format="DD MMM YYYY"
            )
        st.dataframe(agg, hide_index=True, use_container_width=True,
                     column_config=cfg)


# ---------- Tab: Products --------------------------------------------------
with tabs[_tab_idx["Products"]]:
    st.subheader("Products & manufacturers")
    if df.empty:
        st.info("No data for current filters.")
    else:
        sub_tabs = st.tabs(["By product group", "By manufacturer", "By individual product"])

        with sub_tabs[0]:
            if "Product Group" in df.columns:
                t = (df.groupby("Product Group")
                     .agg(**{
                         "Net sales": ("Net Sales Volume", "sum"),
                         "Cost": ("Cost (Actual)", "sum"),
                         "Profit": ("Profit Margin", "sum"),
                         "Units": ("Quantity (Actual)", "sum"),
                         "Invoices": ("Billing Document", "nunique"),
                         "Distinct products": ("Product Id", "nunique"),
                     }).reset_index())
                t["Margin %"] = np.where(t["Net sales"] != 0,
                                         t["Profit"] / t["Net sales"] * 100, 0)
                t = t.sort_values("Net sales", ascending=False)
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Cost", "Profit"],
                        pct=["Margin %"],
                        ints=["Units", "Invoices", "Distinct products"],
                    ),
                )

        with sub_tabs[1]:
            if "Manufacturer Name" in df.columns:
                t = (df.groupby("Manufacturer Name")
                     .agg(**{
                         "Net sales": ("Net Sales Volume", "sum"),
                         "Profit": ("Profit Margin", "sum"),
                         "Units": ("Quantity (Actual)", "sum"),
                         "Invoices": ("Billing Document", "nunique"),
                         "Distinct products": ("Product Id", "nunique"),
                     }).reset_index())
                t["Margin %"] = np.where(t["Net sales"] != 0,
                                         t["Profit"] / t["Net sales"] * 100, 0)
                t = t.sort_values("Net sales", ascending=False)
                t = t.rename(columns={"Manufacturer Name": "Manufacturer"})
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Profit"],
                        pct=["Margin %"],
                        ints=["Units", "Invoices", "Distinct products"],
                    ),
                )

        with sub_tabs[2]:
            if "Product Id" in df.columns:
                top_n_p = st.slider("Show top N products", 10, 500, 50, 10, key="topp")
                group_keys = ["Product Id"]
                if "Product Desc" in df.columns:
                    group_keys.append("Product Desc")
                t = (df.groupby(group_keys)
                     .agg(**{
                         "Net sales": ("Net Sales Volume", "sum"),
                         "Profit": ("Profit Margin", "sum"),
                         "Units": ("Quantity (Actual)", "sum"),
                         "Invoices": ("Billing Document", "nunique"),
                     }).reset_index())
                t["Margin %"] = np.where(t["Net sales"] != 0,
                                         t["Profit"] / t["Net sales"] * 100, 0)
                t["Avg unit price"] = np.where(
                    t["Units"] != 0, t["Net sales"] / t["Units"], 0
                )
                t = t.sort_values("Net sales", ascending=False).head(top_n_p)
                st.dataframe(
                    t, hide_index=True, use_container_width=True,
                    column_config=build_col_config(
                        t, currency,
                        money=["Net sales", "Profit", "Avg unit price"],
                        pct=["Margin %"],
                        ints=["Units", "Invoices"],
                    ),
                )


# ---------- Tab: Geography -------------------------------------------------
with tabs[_tab_idx["Geography"]]:
    st.subheader("Geography breakdown")
    if "City Name" not in df.columns or df.empty:
        st.info("No city data available for current filters.")
    else:
        t = (df.groupby("City Name")
             .agg(**{
                 "Net sales": ("Net Sales Volume", "sum"),
                 "Profit": ("Profit Margin", "sum"),
                 "Units": ("Quantity (Actual)", "sum"),
                 "Invoices": ("Billing Document", "nunique"),
                 "Customers": ("Customer", "nunique"),
             }).reset_index())
        t["Margin %"] = np.where(t["Net sales"] != 0,
                                 t["Profit"] / t["Net sales"] * 100, 0)
        t["% of period sales"] = np.where(
            m["sales"], t["Net sales"] / m["sales"] * 100, 0
        )
        t = t.sort_values("Net sales", ascending=False)
        t = t.rename(columns={"City Name": "City"})
        st.dataframe(
            t, hide_index=True, use_container_width=True,
            column_config=build_col_config(
                t, currency,
                money=["Net sales", "Profit"],
                pct=["Margin %", "% of period sales"],
                ints=["Units", "Invoices", "Customers"],
            ),
        )


# ---------- Tab: Time trend ------------------------------------------------
with tabs[_tab_idx["Time trend"]]:
    st.subheader("Performance over time")
    if "Billing Date" not in df.columns or df.empty:
        st.info("No date data for current filters.")
    else:
        granularity = st.radio(
            "Granularity",
            ["Daily", "Weekly", "Monthly", "Quarterly"],
            index=2, horizontal=True,
        )
        d = df.copy()
        if granularity == "Daily":
            d["Period"] = d["Billing Date"].dt.date.astype(str)
        elif granularity == "Weekly":
            d["Period"] = (d["Billing Date"].dt.to_period("W")
                           .apply(lambda p: p.start_time.strftime("%Y-%m-%d")))
        elif granularity == "Monthly":
            d["Period"] = d["Billing Date"].dt.to_period("M").astype(str)
        else:
            d["Period"] = d["Billing Date"].dt.to_period("Q").astype(str)

        t = (d.groupby("Period")
             .agg(**{
                 "Net sales": ("Net Sales Volume", "sum"),
                 "Cost": ("Cost (Actual)", "sum"),
                 "Profit": ("Profit Margin", "sum"),
                 "Units": ("Quantity (Actual)", "sum"),
                 "Invoices": ("Billing Document", "nunique"),
             }).reset_index())
        t["Margin %"] = np.where(t["Net sales"] != 0,
                                 t["Profit"] / t["Net sales"] * 100, 0)
        t = t.sort_values("Period")

        if show_charts and not t.empty:
            st.markdown("**Net sales and profit over time**")
            try:
                st.bar_chart(
                    t, x="Period", y=["Net sales", "Profit"],
                    color=[SERIES_COLORS["Net sales"],
                           SERIES_COLORS["Profit"]],
                    height=300,
                )
            except Exception:
                pass

        st.dataframe(
            t, hide_index=True, use_container_width=True,
            column_config=build_col_config(
                t, currency,
                money=["Net sales", "Cost", "Profit"],
                pct=["Margin %"],
                ints=["Units", "Invoices"],
            ),
        )


# ---------- Tab: Raw data --------------------------------------------------
with tabs[_tab_idx["Raw data"]]:
    st.subheader("Filtered transactions")
    st.markdown(
        f"<span class='muted'>{len(df):,} rows after filters. "
        "Click column headers to sort. Use the search icon (top right of the table) "
        "to find specific values.</span>",
        unsafe_allow_html=True,
    )

    show_cols = [c for c in [
        "Billing Date", "Billing Document", "Billing Type", "Division",
        "Sales Empl. Name", "Customer", "City Name",
        "Product Id", "Product Desc", "Product Group", "Manufacturer Name",
        "Quantity (Actual)", "Net Price", "Net Sales Volume",
        "Cost (Actual)", "Profit Margin", "Profit Margin Ratio",
        "Tax Amount",
    ] if c in df.columns]
    show = df[show_cols].copy()

    cfg = build_col_config(
        show, currency,
        money=["Net Price", "Net Sales Volume", "Cost (Actual)",
               "Profit Margin", "Tax Amount"],
        ints=["Quantity (Actual)"],
    )
    if "Profit Margin Ratio" in show.columns:
        show["Profit Margin Ratio"] = show["Profit Margin Ratio"] * 100
        cfg["Profit Margin Ratio"] = st.column_config.NumberColumn(
            "Profit Margin %", format="%,.2f%%",
            help="Profit margin as percent of net sales.",
        )
    if "Billing Date" in show.columns:
        cfg["Billing Date"] = st.column_config.DateColumn(
            "Billing Date", format="DD MMM YYYY"
        )

    st.dataframe(show, hide_index=True, use_container_width=True,
                 column_config=cfg, height=520)

    csv_bytes = show.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered data (CSV)",
        data=csv_bytes,
        file_name=f"sales_filtered_{start_d.isoformat()}_to_{end_d.isoformat()}.csv",
        mime="text/csv",
    )


# =============================================================================
# Footer
# =============================================================================
st.markdown("<hr style='margin-top: 2rem; border-color: #e6e8ed;'>", unsafe_allow_html=True)
_footer = (
    f"<span class='muted'>Sales source: <b>{source_name}</b> "
    f"({len(df_full):,} rows) &nbsp;&middot;&nbsp; "
    f"Net sales (all time): <b>{fmt_money(m_all['sales'], currency)}</b>"
)
if df_orders_full is not None and orders_source:
    mo_all = order_kpis(df_orders_full)
    _footer += (
        f" &nbsp;&middot;&nbsp; Orders source: <b>{orders_source}</b> "
        f"({len(df_orders_full):,} lines) &nbsp;&middot;&nbsp; "
        f"Pipeline (all time): <b>{fmt_money(mo_all['pending_value'], currency)}</b>"
    )
_footer += "</span>"
st.markdown(_footer, unsafe_allow_html=True)
