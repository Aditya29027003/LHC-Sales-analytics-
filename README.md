# Sales Analytics Tracker

A numbers-first dashboard for the LHC Logistics sales report. Designed to give
you fast, sortable, drill-down breakdowns of departments, employees, customers,
products, manufacturers, cities, and time periods - with KPI tiles up top and
detailed tables below. Minimal charts, lots of numbers.

## What's in this folder

| File                              | Purpose                                            |
| --------------------------------- | -------------------------------------------------- |
| `Sales Report_24_04_2024.xlsx`    | Your source sales report (auto-loaded by default). |
| `tracker.py`                      | The Streamlit analytics app.                       |
| `requirements.txt`                | Python packages needed.                            |
| `run.bat`                         | Double-click to launch on Windows.                 |
| `README.md`                       | This file.                                         |

## First-time setup (only once)

1. Make sure Python 3.10+ is installed.
2. Open PowerShell in this folder and install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

## Launching the tracker

Either:

- **Double-click `run.bat`** - it will install missing packages and open the app
  in your default browser, or
- Run from PowerShell:

   ```powershell
   python -m streamlit run tracker.py
   ```

The app opens at `http://localhost:8501`. Press `Ctrl+C` in the terminal window
to stop it when you're done.

## Loading another sales file later

Use the **"Upload sales report"** box in the sidebar - any `.xlsx`, `.xls`, or
`.csv` file with the same column structure as the current LHC report will work.
The dashboard will refresh automatically with the new data. The original file
in this folder remains the default whenever no upload is provided.

### Required columns

The tracker expects the following column names (extra columns are ignored):

```
Billing Type, Billing Document, Billing Date, Customer,
Product Id, Product Desc, Quantity (Actual), Net Price,
Net Sales Volume, Division, Sales Empl. Name,
Manufacturer Name, Product Group, Tax Amount,
Cost (Actual), Profit Margin, Profit Margin Ratio, City Name
```

If a column is missing the related section will simply be empty - the rest of
the app keeps working.

## What you can do in the app

### Sidebar filters
- **Quick period** presets (This month, Last 30 days, Last 90 days, YTD) plus
  a custom date-range picker.
- **Multi-select filters** for Division, Sales Employee, Product Group,
  Manufacturer, City, Billing Type, Customer.
- Toggle to **include or exclude credit memos / cancellations**.
- Custom **currency label** (defaults to AED).

### Tabs
1. **Overview** - top-5 leaderboards across all dimensions and headline KPIs
   (gross margin %, average unit price, average sales per employee/customer,
   average lines per invoice).
2. **Departments** - per-division table with net sales, cost, profit, tax,
   units, invoices, customers, employees, margin %, and share of total sales,
   plus a department drill-down to the employee level.
3. **Employees** - sortable leaderboard with rank, all financials, average
   invoice value, % of period sales, and number of departments served.
4. **Customers** - top-N customers with full financials, last invoice date,
   margin %, and share of period sales.
5. **Products** - three sub-views: by product group, by manufacturer, by
   individual product (with average unit price).
6. **Geography** - city-level breakdown.
7. **Time trend** - daily / weekly / monthly / quarterly aggregation table,
   plus an optional small bar chart in an expander.
8. **Raw data** - filtered transaction-level data with sortable columns and a
   one-click CSV export of the current filtered view.

### KPI tiles (always visible at the top)
Net sales, profit margin (with margin %), cost, tax collected, units sold,
invoices, customers, products, sales employees, and average invoice value -
all computed against the currently filtered data.

## Tips
- Click any column header in a table to sort. Use the search icon in the
  table's top-right corner to filter rows.
- Combine filters: e.g. select **Division = MedDivision (L9)** + **Period =
  Last 30 days** to see how that one team performed last month, then drill to
  the Employees tab to rank reps within that scope.
- The **Raw data** tab's CSV download always reflects the current filters, so
  you can hand a focused slice to anyone else who needs it.
