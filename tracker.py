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

EXPECTED_COLUMNS = [
    "Billing Type", "Billing Document", "Billing Date", "Customer",
    "Product Id", "Product Desc", "Quantity (Actual)", "Net Price",
    "Net Sales Volume", "Division", "Sales Empl. Name",
    "Manufacturer Name", "Product Group", "Tax Amount",
    "Cost (Actual)", "Profit Margin", "Profit Margin Ratio", "City Name",
]

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

currency = st.sidebar.text_input("Currency label", value="AED", max_chars=6).strip() or "AED"

st.sidebar.markdown("---")
st.sidebar.header("Filters")

# Date range
if "Billing Date" in df_full.columns and df_full["Billing Date"].notna().any():
    min_date = df_full["Billing Date"].min().date()
    max_date = df_full["Billing Date"].max().date()
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
df = apply_in(df, "Sales Empl. Name", f_emp)
df = apply_in(df, "Product Group", f_pg)
df = apply_in(df, "Manufacturer Name", f_man)
df = apply_in(df, "City Name", f_city)
df = apply_in(df, "Billing Type", f_btype)
df = apply_in(df, "Customer", f_cust)

if not include_credits and "Billing Type" in df.columns:
    df = df[df["Billing Type"].str.contains("Invoice", case=False, na=False)]

st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<span class='muted'>Filtered rows: <b>{len(df):,}</b></span>",
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
    ("Division", f_div), ("Employee", f_emp), ("Product group", f_pg),
    ("Manufacturer", f_man), ("City", f_city),
    ("Billing type", f_btype), ("Customer", f_cust),
]:
    if vals:
        v_disp = vals[0] if len(vals) == 1 else f"{len(vals)} selected"
        filter_pills += f"<span class='pill'>{label}: {v_disp}</span>"

st.markdown(
    f"<span class='pill'>Period: {period_label}</span>"
    f"<span class='pill'>{days_n} days</span>"
    f"<span class='pill'>{len(df):,} line items</span>"
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

st.markdown("<hr style='margin: 1.2rem 0; border-color: #e6e8ed;'>", unsafe_allow_html=True)


# =============================================================================
# Tabs
# =============================================================================
tabs = st.tabs([
    "Overview",
    "Departments",
    "Employees",
    "Customers",
    "Products",
    "Geography",
    "Time trend",
    "Raw data",
])


# ---------- Tab: Overview --------------------------------------------------
with tabs[0]:
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


# ---------- Tab: Departments -----------------------------------------------
with tabs[1]:
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

        st.dataframe(
            out, hide_index=True, use_container_width=True,
            column_config=build_col_config(
                out, currency,
                money=["Net sales", "Cost", "Profit", "Tax"],
                pct=["Margin %", "% of total sales"],
                ints=["Units", "Invoices", "Customers", "Employees"],
            ),
        )

        st.write("")
        st.markdown("**Drill down: select a department to see its employees**")
        div_pick = st.selectbox(
            "Department",
            options=["(all)"] + sorted(df["Division"].unique().tolist()),
        )
        sub = df if div_pick == "(all)" else df[df["Division"] == div_pick]

        if "Sales Empl. Name" in sub.columns and not sub.empty:
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


# ---------- Tab: Employees -------------------------------------------------
with tabs[2]:
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
        agg = (df.groupby("Sales Empl. Name")
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
with tabs[3]:
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
with tabs[4]:
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
with tabs[5]:
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
with tabs[6]:
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

        st.dataframe(
            t, hide_index=True, use_container_width=True,
            column_config=build_col_config(
                t, currency,
                money=["Net sales", "Cost", "Profit"],
                pct=["Margin %"],
                ints=["Units", "Invoices"],
            ),
        )

        with st.expander("Optional: tiny trend chart"):
            try:
                chart_df = t.set_index("Period")[["Net sales", "Profit"]]
                st.bar_chart(chart_df, height=260)
            except Exception:
                st.write("Chart unavailable.")


# ---------- Tab: Raw data --------------------------------------------------
with tabs[7]:
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
st.markdown(
    f"<span class='muted'>Source file: <b>{source_name}</b> &nbsp;&middot;&nbsp; "
    f"Total rows in file: <b>{len(df_full):,}</b> &nbsp;&middot;&nbsp; "
    f"Net sales (all time): <b>{fmt_money(m_all['sales'], currency)}</b></span>",
    unsafe_allow_html=True,
)
