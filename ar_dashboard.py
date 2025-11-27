import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime
import io
import os

st.set_page_config(page_title="AR Dashboard", layout="wide")


# ✅ MUST be the first Streamlit call


# ---------------------------------------------------------------
# 🏢 Company Header
# ---------------------------------------------------------------
COMPANY_NAME = "DEV IT SERV PVT LTD"

st.markdown(
    f"""
    <style>
        .header-container {{
            text-align: center;
            padding-top: 10px;
            padding-bottom: 5px;
        }}
        .company-title {{
            color: #0A66C2;
            font-size: 52px;
            font-weight: 800;
            letter-spacing: 1px;
            margin-bottom: 0px;
        }}
        .dashboard-subtitle {{
            color: #34495E;
            font-size: 28px;
            font-weight: 500;
            margin-top: 5px;
        }}
        .divider {{
            border: none;
            height: 2px;
            background-color: #B0B0B0;
            margin-top: 10px;
            margin-bottom: 20px;
            width: 95%;
            margin-left: auto;
            margin-right: auto;
        }}
    </style>

    <div class="header-container">
        <div class="company-title">{COMPANY_NAME}</div>
        <div class="dashboard-subtitle">Accounts Receivable Dashboard</div>
        <hr class="divider">
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------
# 🧾 Fix AgGrid column/row alignment and spacing
# ---------------------------------------------------------------
st.markdown(
    """
    <style>
        .ag-theme-balham .ag-header-cell-label {
            justify-content: center !important;
            text-align: center !important;
        }
        .ag-theme-balham .ag-cell {
            text-align: center !important;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px !important;
            padding: 6px 8px !important;
        }
        .ag-theme-balham .ag-header-cell {
            font-size: 17px !important;
            font-weight: 600 !important;
            padding: 8px !important;
        }
        .ag-theme-balham .ag-row:nth-child(even) {
            background-color: #F8F9F9 !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------
# 🧾 Fix Streamlit DataFrame column alignment & sizing
# ---------------------------------------------------------------
st.markdown(
    """
    <style>
        /* Align headers and body cells */
        .stDataFrame thead tr th {
            text-align: center !important;
            vertical-align: middle !important;
            font-size: 18px !important;
            padding: 8px 10px !important;
        }

        .stDataFrame tbody tr td {
            text-align: center !important;
            vertical-align: middle !important;
            font-size: 17px !important;
            padding: 6px 10px !important;
        }

        /* Make columns auto-fit width */
        div[data-testid="stDataFrame"] div[data-testid="stVerticalBlock"] {
            overflow-x: auto !important;
        }

        /* Fix column alignment in scrollable mode */
        [data-testid="stHorizontalBlock"] {
            align-items: stretch !important;
        }

        /* Optional: alternate row shading for readability */
        .stDataFrame tbody tr:nth-child(even) {
            background-color: #F8F9F9 !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------
def safe_currency(x):
    try:
        v = float(str(x).replace("₹", "").replace(",", "").strip())
        return f"₹{v:,.0f}"
    except:
        return "-"

def highlight_overdue(row):
    """Highlight overdue rows safely, even if key columns are missing."""
    overdue_col = f">{bucket_def[-1][1]} Days"
    overdue = row.get(overdue_col, 0)

    try:
        overdue_val = float(str(overdue).replace("₹", "").replace(",", "").strip())
    except Exception:
        overdue_val = 0.0

    key_name = None
    for possible in ["Customer Name", "Account Manager", "S.No"]:
        if possible in row.index:
            key_name = possible
            break

    key_value = str(row.get(key_name, "")).strip() if key_name else ""
    color = "background-color: #ffe6e6;" if (overdue_val > 0 and key_value != "Grand Total") else ""
    return [color] * len(row)

def find_header_row(raw, max_scan=10):
    keywords = ["customer", "invoice", "due", "amount", "payment"]
    for i in range(min(max_scan, len(raw))):
        row = [str(x).lower() for x in raw.iloc[i].fillna("")]
        hits = sum(any(k in cell for cell in row) for k in keywords)
        if hits >= 2:
            return i
    return 0

def bucket_category(days, buckets):
    """Categorize safely; never fails even if inputs are strings."""
    # Force numeric for the "days" value
    try:
        d = float(str(days).replace(",", "").strip())
    except Exception:
        d = 0.0

    # Build a clean numeric bucket list
    safe_buckets = []
    for label, limit in buckets:
        try:
            lim = float(str(limit).replace(",", "").strip())
        except Exception:
            lim = 0.0
        safe_buckets.append((label, lim))

    # Defensive comparison
    for label, limit in safe_buckets:
        try:
            if float(d) <= float(limit):
                return label
        except Exception as e:
            # Debugging aid – prints the pair that failed
            st.warning(f"⚠️ Comparison failed for days={days} (type {type(days)}) and limit={limit} (type {type(limit)})")
            return f"ErrorBucket({days})"

    # If all else fails, return the final bucket
    return f">{safe_buckets[-1][1]:.0f} Days"


# ---------------------------------------------------------------
# Excel Data Source
# ---------------------------------------------------------------
data_mode = st.radio("📊 Select data source mode:", ["Upload Excel", "Linked Excel File"], horizontal=True)

if data_mode == "Upload Excel":
    uploaded = st.file_uploader("📂 Upload AR Excel file", type=["xlsx", "xls"])
    if not uploaded:
        st.info("Upload your AR Excel file to view the dashboard.")
        st.stop()
    raw = pd.read_excel(uploaded, header=None, dtype=str)

else:
    st.info("Using linked Excel file path below:")
    default_path = r"C:\Users\Rinku.soni\OneDrive - DEV IT SERV\Desktop\AR Dashboard\08 Nov 2025 AR Report Final.xlsx"
    file_path = st.text_input("🔗 Excel file path:", default_path)

    if not os.path.exists(file_path):
        st.error(f"❌ File not found: {file_path}")
        st.stop()
    else:
        raw = pd.read_excel(file_path, header=None, dtype=str)
        # Convert raw Excel into a working dataframe
header_row = 0  # Adjust if your actual headers are not on the first row
df = pd.read_excel(uploaded if data_mode == "Upload Excel" else file_path, header=header_row)
df.columns = [str(c).strip() for c in df.columns]


# ---------------------------------------------------------------
# 🔍 Remove fully paid invoices (where Due Amount = 0)
# ---------------------------------------------------------------
if "Due Amount" in df.columns:
    df["Due Amount"] = pd.to_numeric(df["Due Amount"], errors="coerce").fillna(0)
    df = df[df["Due Amount"] > 0]  # ✅ keep only invoices with balance due


# ---------------------------------------------------------------
# Column mapping
# ---------------------------------------------------------------
def match_col(df, keywords):
    for c in df.columns:
        if any(k.lower() in c.lower() for k in keywords):
            return c
    return None

cust_col = match_col(df, ["customer", "client", "party"])
am_col = match_col(df, ["account manager", "owner", "manager"])
inv_amt_col = match_col(df, ["invoice amount", "amount", "invoice_amt"])
inv_date_col = match_col(df, ["invoice date", "invoice_date", "invoice dt"])
due_date_col = match_col(df, ["due date", "duedate", "due_dt"])
paid_amt_col = match_col(df, ["paid amount", "paidamount", "amount paid"])
due_amt_col = match_col(df, ["due amount", "dueamount", "amount due"])
status_col = match_col(df, ["payment status", "status", "paid/unpaid"])

mapping = {
    cust_col: "Customer Name",
    am_col: "Account Manager",
    inv_amt_col: "Invoice Amount",
    inv_date_col: "Invoice Date",
    due_date_col: "Due Date",
    paid_amt_col: "Paid Amount",
    due_amt_col: "Due Amount",
    status_col: "Payment Status"
}
mapping = {k: v for k, v in mapping.items() if k}
df.rename(columns=mapping, inplace=True)

# ---------------------------------------------------------------
# Clean & prepare data
# ---------------------------------------------------------------
df["Customer Name"] = df.get("Customer Name", "").astype(str).str.strip()
df = df[df["Customer Name"] != ""]

# Fix numeric fields
for c in ["Invoice Amount", "Paid Amount", "Due Amount"]:
    if c in df.columns:
        df[c] = (
            df[c]
            .astype(str)
            .str.replace("₹", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

df["Invoice Date"] = pd.to_datetime(df.get("Invoice Date"), errors="coerce")
df["Due Date"] = pd.to_datetime(df.get("Due Date"), errors="coerce")

# ---------------------------------------------------------------
# Dynamic Aging Bucket Setup
# ---------------------------------------------------------------
with st.expander("⚙️ Aging Bucket Configuration"):
    col_a, col_b, col_c, col_d = st.columns(4)
    bucket1 = col_a.number_input("Bucket 1 upper limit (days)", value=30)
    bucket2 = col_b.number_input("Bucket 2 upper limit (days)", value=60)
    bucket3 = col_c.number_input("Bucket 3 upper limit (days)", value=90)
    st.caption("Anything beyond the last limit goes to the '90+ Days' bucket automatically.")

bucket_def = [
    (f"0–{bucket1} Days", bucket1),
    (f"{bucket1+1}–{bucket2} Days", bucket2),
    (f"{bucket2+1}–{bucket3} Days", bucket3)
]

# ---- Force all bucket limits to float safely ----
bucket_def = [(label, float(str(limit).strip())) for label, limit in bucket_def]


# --- compute and clean "Days Overdue" safely ---
today = pd.Timestamp(datetime.now().date())
df["Days Overdue"] = (today - df["Due Date"]).dt.days
df["Days Overdue"] = (
    df["Days Overdue"]
    .astype(str)
    .str.replace(r"[^0-9\-]", "", regex=True)      # keep digits and minus
    .replace("", "0")
)
df["Days Overdue"] = pd.to_numeric(df["Days Overdue"], errors="coerce").fillna(0)
df["Days Overdue"] = df["Days Overdue"].clip(lower=0).astype(int)

# --- assign aging buckets safely (guaranteed numeric) ---
def bucket_category(days, buckets):
    """Return bucket label for given days; fully numeric safe."""
    try:
        d = float(str(days).replace(",", "").strip())
    except Exception:
        d = 0.0

    for label, limit in buckets:
        if d <= limit:
            return label
    return f">{buckets[-1][1]} Days"

df["Days Overdue"] = df["Days Overdue"].astype(str).str.replace(r"[^0-9\-]", "", regex=True)
df["Days Overdue"] = pd.to_numeric(df["Days Overdue"], errors="coerce").fillna(0)
df["Aging Bucket"] = df["Days Overdue"].apply(lambda x: bucket_category(x, bucket_def))


# ---------------------------------------------------------------
# Filters
# ---------------------------------------------------------------
col1, col2, col3 = st.columns(3)
if "Account Manager" in df.columns:
    am_list = sorted(df["Account Manager"].dropna().unique().tolist())
    selected_am = col1.selectbox("👤 Filter by Account Manager", ["All"] + am_list)
    if selected_am != "All":
        df = df[df["Account Manager"] == selected_am]

cust_list = sorted(df["Customer Name"].dropna().unique().tolist())
selected_cust = col2.selectbox("🏢 Filter by Customer", ["All"] + cust_list)
if selected_cust != "All":
    df = df[df["Customer Name"] == selected_cust]

view_mode = col3.radio("📊 View Mode", ["Customer-wise", "Account Manager Summary"], horizontal=True)

# -------------------- Robust pivot + display (force Customer Name) --------------------
# Ensure a Customer Name column exists (robust heuristics)
if "Customer Name" in df.columns and df["Customer Name"].notna().any():
    df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
else:
    candidate = None
    keys = ["customer name", "customer", "client", "party", "cust name", "party name"]
    for c in df.columns:
        low = str(c).strip().lower()
        if any(k in low for k in keys):
            candidate = c
            break
    if not candidate:
        # fallback: first non-numeric non-date column
        for c in df.columns:
            low = str(c).strip().lower()
            if any(k in low for k in ["date", "amount", "invoice", "due", "paid", "status", "terms", "no"]):
                continue
            sample = df[c].dropna().astype(str).head(20).tolist()
            if sample and any(not s.strip().replace(",", "").replace("₹", "").replace(".", "").isdigit() for s in sample):
                candidate = c
                break
    if candidate:
        df["Customer Name"] = df[candidate].astype(str).str.strip()
    else:
        # last resort: first non-empty column
        for c in df.columns:
            if df[c].dropna().astype(str).str.strip().astype(bool).any():
                df["Customer Name"] = df[c].astype(str).str.strip()
                candidate = c
                break

# If still missing or all blank, stop with diagnostic
if "Customer Name" not in df.columns or df["Customer Name"].dropna().astype(str).str.strip().eq("").all():
    st.error("❌ Could not detect any Customer Name values. Detected columns: " + ", ".join(map(str, df.columns.tolist())))
    # show a short sample to help debugging
    st.write("Sample header and first 5 rows:")
    st.write(df.head(5))
    st.stop()

# Ensure Invoice Amount numeric and Days Overdue numeric
df["Invoice Amount"] = df.get("Invoice Amount", 0).astype(str).str.replace("₹","", regex=False).str.replace(",","", regex=False).str.strip()
df["Invoice Amount"] = pd.to_numeric(df["Invoice Amount"], errors="coerce").fillna(0)

today = pd.Timestamp(datetime.now().date())
df["Days Overdue"] = (today - pd.to_datetime(df.get("Due Date"), errors="coerce")).dt.days
df["Days Overdue"] = pd.to_numeric(df["Days Overdue"], errors="coerce").fillna(0).clip(lower=0).astype(int)

# Recompute Aging Bucket based on your bucket_def (ensure bucket_def exists and is numeric)
# Force bucket_def limits to numeric if they exist
try:
    bucket_def = [(lbl, float(limit)) for lbl, limit in bucket_def]
except Exception:
    # fallback default buckets
    bucket_def = [("0–30 Days", 30.0), ("31–60 Days", 60.0), ("61–90 Days", 90.0)]

def bucket_category_safe(days, buckets):
    try:
        d = float(days)
    except Exception:
        d = 0.0
    for label, lim in buckets:
        if d <= lim:
            return label
    return f">{int(buckets[-1][1])} Days"

df["Aging Bucket"] = df["Days Overdue"].apply(lambda x: bucket_category_safe(x, bucket_def))

# Build pivot safely
index_col = "Customer Name" if view_mode == "Customer-wise" else "Account Manager"
if index_col not in df.columns:
    # If AM view but Account Manager missing, fall back to Customer Name
    if index_col == "Account Manager" and "Customer Name" in df.columns:
        index_col = "Customer Name"
    else:
        st.error(f"❌ Required index column '{index_col}' not found in data.")
        st.write("Detected columns:", df.columns.tolist())
        st.stop()

pivot = df.pivot_table(index=index_col, columns="Aging Bucket", values="Invoice Amount", aggfunc="sum", fill_value=0)

# Build column order from bucket_def
column_order = [b[0] for b in bucket_def] + [f">{int(bucket_def[-1][1])} Days"]
pivot = pivot.reindex(columns=column_order, fill_value=0)

# Add Grand Total row (only if not already present)
if "Grand Total" not in pivot.index:
    subtotal = pd.DataFrame(pivot.sum()).T
    subtotal.index = ["Grand Total"]
    pivot = pd.concat([pivot, subtotal])

# ensure numeric & compute totals
pivot = pivot.apply(pd.to_numeric, errors="coerce").fillna(0)
pivot["Total"] = pivot[column_order].sum(axis=1)

# Reset index -> Customer Name becomes a column
display_df = pivot.reset_index()

# Add S.No (1..N)
display_df.insert(0, "S.No", range(1, len(display_df) + 1))

# Format currency for display
for col in column_order + ["Total"]:
    if col in display_df.columns:
        display_df[col] = display_df[col].apply(lambda x: f"₹{x:,.0f}" if pd.notnull(x) else "-")

# Define highlight function (safe)
def highlight_overdue_safe(row):
    overdue_col = f">{int(bucket_def[-1][1])} Days"
    overdue_val = 0.0
    if overdue_col in row.index:
        try:
            overdue_val = float(str(row[overdue_col]).replace("₹","").replace(",","").strip())
        except Exception:
            overdue_val = 0.0
    key_value = ""
    for key in [index_col, "Customer Name", "Account Manager"]:
        if key in row.index:
            key_value = str(row[key])
            break
    color = "background-color: #ffe6e6;" if (overdue_val > 0 and key_value != "Grand Total") else ""
    return [color] * len(row)

# Display
if display_df.empty:
    st.warning("⚠️ No data to show after filtering.")
else:
    st.subheader(f"📅 {view_mode} Aging Buckets")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    
# ---------------------------------------------------------------
# 🔍 Invoice-level details (on Customer click)
# ---------------------------------------------------------------

st.markdown("### 🔍 Invoice-wise Details")

# Dropdown or clickable customer selection
customers = sorted(df["Customer Name"].dropna().unique().tolist())
selected_customer = st.selectbox(
    "Select Customer to view invoice details:",
    sorted(df["Customer Name"].unique())
)

if selected_customer:
    filtered_df = df[df["Customer Name"] == selected_customer].copy()

    # Convert numeric columns safely
    for col in ["Invoice Amount", "Paid Amount", "Due Amount"]:
        filtered_df[col] = pd.to_numeric(filtered_df[col], errors="coerce").fillna(0)

    # ✅ Calculate totals
    total_invoice = filtered_df["Invoice Amount"].sum()
    total_paid = filtered_df["Paid Amount"].sum()
    total_due = filtered_df["Due Amount"].sum()

    # Display subtotals neatly
    st.markdown(f"""
    <div style='font-size:18px; font-weight:600; text-align:right; margin-bottom:10px;'>
        Total Invoice Amount: <span style='color:#2E86C1'>₹{total_invoice:,.2f}</span><br>
        Total Paid: <span style='color:#239B56'>₹{total_paid:,.2f}</span><br>
        Total Due: <span style='color:#C0392B'>₹{total_due:,.2f}</span>
    </div>
    <hr style='border:1px solid #ccc; margin:10px 0;'>
    """, unsafe_allow_html=True)

    # ✅ Show invoice detail table
    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_pagination(enabled=True)
    gb.configure_default_column(resizable=True, sortable=True, filter=True)
    gridOptions = gb.build()

    AgGrid(filtered_df, gridOptions=gridOptions, theme="balham", height=400)


    # Ensure numeric formatting
    for col in ["Invoice Amount", "Paid Amount", "Due Amount"]:
        if col in filtered_df.columns:
            filtered_df[col] = (
                pd.to_numeric(filtered_df[col], errors="coerce")
                .fillna(0)
                .apply(lambda x: f"₹{x:,.2f}")
            )

    # Display nicely
    st.dataframe(
        filtered_df[
            [
                c
                for c in [
                    "Invoice No.",
                    "Invoice Date",
                    "Due Date",
                    "Invoice Amount",
                    "Paid Amount",
                    "Due Amount",
                    "Delay Days",
                ]
                if c in filtered_df.columns
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    # Optional: subtotal at the bottom
    total_due = pd.to_numeric(
        filtered_df["Due Amount"].replace("₹", "", regex=True).replace(",", "", regex=True),
        errors="coerce",
    ).sum()

    st.markdown(f"**Total Outstanding for {selected_customer}: ₹{total_due:,.2f}**")

# Download filtered/pivot as Excel (keeps Customer/AM column)
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    pivot.to_excel(writer, sheet_name="Aging_Pivot")
    df.to_excel(writer, sheet_name="Raw_Filtered_Data", index=False)
buf.seek(0)
st.download_button("⬇️ Download Filtered Report (Excel)", data=buf, file_name="Aging_Report_Filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# KPIs (safe numeric)
total = df["Invoice Amount"].sum()
paid = df.get("Paid Amount", 0)
if isinstance(paid, (pd.Series, pd.Index)):
    paid = pd.to_numeric(paid.astype(str).str.replace("₹","").str.replace(",",""), errors="coerce").fillna(0).sum()
else:
    try:
        paid = float(str(paid).replace("₹","").replace(",",""))
    except Exception:
        paid = 0.0
unpaid = total - paid
paid_ratio = (paid / total * 100) if total else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Receivable", f"₹{total:,.2f}")
c2.metric("Paid", f"₹{paid:,.2f}")
c3.metric("Unpaid", f"₹{unpaid:,.2f}")
c4.metric("Paid Ratio", f"{paid_ratio:.1f}%")
