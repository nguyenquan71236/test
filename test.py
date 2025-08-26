import pandas as pd
import numpy as np
import streamlit as st
import re
from time import time
from datetime import datetime
from io import BytesIO, StringIO

st.set_page_config(layout="wide")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="MTD")
    return output.getvalue()


def get_largest_sheet(file):
    # Read all sheet names
    xl = pd.ExcelFile(file)
    sheet_sizes = {}

    # Measure data size per sheet
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, nrows=500)  # limit for speed, adjust as needed
        sheet_sizes[sheet] = df.shape[0] * df.shape[1]  # total cells

    # Pick the sheet with the most cells
    largest_sheet = max(sheet_sizes, key=sheet_sizes.get)
    return largest_sheet

def get_columns_of_largest_sheet(file):
    xl = pd.ExcelFile(file)
    sheet_sizes = {}

    for sheet in xl.sheet_names:
        try:
            df_preview = xl.parse(sheet, nrows=500)
            sheet_sizes[sheet] = df_preview.shape[0] * df_preview.shape[1]
        except:
            sheet_sizes[sheet] = 0

    if not sheet_sizes:
        return []

    largest_sheet = max(sheet_sizes, key=sheet_sizes.get)
    df_header = xl.parse(largest_sheet, header = 4, nrows=0)

    # Clean column names: strip whitespace, drop empty ones
    cleaned_columns = [
        col.strip() for col in df_header.columns
        if isinstance(col, str) and col.strip() != ""
    ]

    return cleaned_columns

# === Streamlit UI ===
st.title("EPM Monthly Display Converter")
col1, col2 = st.columns(2)
columns_base = [
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType"
]

with col1:
    uploaded_files = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)

    check_uploaded_files = []

    for file in uploaded_files:
        match = re.search(r"(\d{4})M(\d+)", file.name)
        columns = get_columns_of_largest_sheet(file)
        check_uploaded_files.append({
            "File": file.name,
            "YEAR": int(match.group(1)) if match else None,
            "MONTH": int(match.group(2)) if match else None,
            "VALID": bool(match),
            "COLUMN CHECK" : columns == (columns_base + ["Amount", "Amount In EUR"])
        })

    check_uploaded_files = pd.DataFrame(check_uploaded_files)
    check_uploaded_files["CONSECUTIVE"] = False
    for i, row in check_uploaded_files.iterrows():
        year = row["YEAR"]
        month = row["MONTH"]
        if pd.isna(year) or pd.isna(month):
            continue
        if month == 1:
            check_uploaded_files.at[i, "CONSECUTIVE"] = True
        else:
            prev_month = month - 1
            next_month = month + 1
            same_year_months = check_uploaded_files[check_uploaded_files["YEAR"] == year]["MONTH"].tolist()
            if prev_month in same_year_months or next_month in same_year_months:
                check_uploaded_files.at[i, "CONSECUTIVE"] = True

    if check_uploaded_files.empty:
        st.info("📂 Please upload Excel files to begin")
    else:
        st.success(f"📄 {len(check_uploaded_files)} file(s) uploaded")

    run_btn = False
    valid_files = True
    CLOSING_M = len(check_uploaded_files)

    if not check_uploaded_files.empty:
        if (~check_uploaded_files["VALID"]).any():
            st.warning("⚠️ All files must have [yyyy]M[mm] in the name")
            valid_files = False
        if check_uploaded_files["YEAR"].nunique() != 1:
            st.warning("⚠️ All files must have the same year")
            valid_files = False
        if check_uploaded_files["MONTH"].min() != 1:
            st.warning("⚠️ Files must start from M1")
            valid_files = False
        if check_uploaded_files["MONTH"].duplicated().any():
            st.warning("⚠️ Files must have unique months")
            valid_files = False
        if (~check_uploaded_files["CONSECUTIVE"]).any():
            st.warning("⚠️ Months within a year must be consecutive")
            valid_files = False
        if (~check_uploaded_files["COLUMN CHECK"]).any():
            st.warning("⚠️ Columns Name must be consistent")
            valid_files = False
        if valid_files == False:
            st.dataframe(check_uploaded_files)
        if valid_files:
            CURRENCY = st.selectbox("Select currency amount:", ["LCC and EUR", "LCC only", "EUR only"])
            run_btn = st.button("🚀 Convert")

# === Run Conversion ===
if run_btn:
    with st.spinner("Preparing your file… this will just take a moment."):
        start_time = time()
        all_dfs = []
        
        for file in uploaded_files:
            match = re.search(r"(\d{4})M(\d+)", file.name)
            if match:
                year, month = int(match.group(1)), int(match.group(2))

                # Skip messy rows, read clean Excel directly
                df_excel = pd.read_excel(file, sheet_name=get_largest_sheet(file), engine="openpyxl", header=4, na_values=[], keep_default_na=False)
                df_excel = df_excel[columns_base + ["Amount", "Amount In EUR"]]
                # Optional: convert to CSV buffer in memory, then read_csv for faster processing
                csv_buffer = StringIO()
                df_excel.to_csv(csv_buffer, index=False, na_rep="None")
                csv_buffer.seek(0)
                df = pd.read_csv(csv_buffer,na_values=[], keep_default_na=False)

                df["YEAR"] = year
                df["MONTH"] = month
                for col in columns_base:
                    if col in df.columns:
                        df[col] = df[col].astype("category")
                all_dfs.append(df)

        df = pd.concat(all_dfs, ignore_index=True)

        read_file_time = time()
        read_file_time_amount = read_file_time - start_time
        with col1:
            st.text(f"read file time: {read_file_time_amount:.2f}s")
        
        df["MONTH+1"] = df["MONTH"] + 1

        columns_next = columns_base + ["YEAR", "MONTH+1"]
        columns_current = columns_base + ["YEAR", "MONTH"]

        df_next = df.groupby(columns_next, observed=True).agg({
            "Amount": "sum",
            "Amount In EUR": "sum"
        }).reset_index().rename(columns={
            "Amount": "Amount_Next",
            "Amount In EUR": "Amount In EUR_Next",
            "MONTH+1": "MONTH"
        })

        df = df.merge(df_next, how="outer", on=columns_current)
        df[["Amount", "Amount_Next", "Amount In EUR", "Amount In EUR_Next"]] = df[["Amount", "Amount_Next", "Amount In EUR", "Amount In EUR_Next"]].fillna(0)

        df["LCC AMOUNT"] = df["Amount"] - df["Amount_Next"]
        df["EUR AMOUNT"] = df["Amount In EUR"] - df["Amount In EUR_Next"]

        df = df.drop(columns=["Amount", "Amount In EUR", "Amount_Next", "Amount In EUR_Next", "MONTH+1"])
        df = df[df["MONTH"] <= CLOSING_M]
        df = df[~((df["EUR AMOUNT"] == 0) & (df["LCC AMOUNT"] == 0))]

        columns_final = columns_base + ["LCC AMOUNT", "EUR AMOUNT", "YEAR", "MONTH"]
        
        if CURRENCY == "LCC only":
            df_final = df[columns_final].drop(columns=["EUR AMOUNT"])
            df_final = df_final[df_final["LCC AMOUNT"] != 0]
        elif CURRENCY == "EUR only":
            df_final = df[columns_final].drop(columns=["LCC AMOUNT"])
            df_final = df_final[df_final["EUR AMOUNT"] != 0]
        else:
            df_final = df[columns_final]

        df_final = df_final.sort_values(by=["YEAR", "MONTH"])

        now = datetime.now()
        date_str = now.strftime("%y%m%d_%H%M")
        max_month = f"{CLOSING_M:02d}"
        currency_code = {
            "LCC only": "LCC",
            "EUR only": "EUR",
            "LCC and EUR": "LCCEUR"
        }[CURRENCY]

        output_filename = f"MTD{max_month}_{currency_code}_{date_str}.xlsx"

        process_file_time = time()
        process_file_time_amount = process_file_time - read_file_time

        with col1:
            st.text(f"process file time: {process_file_time_amount:.2f}s")

        excel_data = to_excel(df_final)

        export_file_time = time()
        export_file_time_amount = export_file_time - process_file_time
        with col1:
            st.text(f"export file time: {export_file_time_amount:.2f}s")
        
        elapsed_time = time() - start_time
        
    with col1:
        st.success(f"✅ Processing completed in {elapsed_time:.2f} seconds! Click below to download.")
        st.download_button(
                label="📥 Download Converted File",
                data=excel_data,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with col2:
            st.subheader("🔍 Summary Statistics")
            row1_col1, row1_col2 = st.columns(2)
            row1_col1.metric("Total Rows", f"{len(df_final):,}")
            row1_col2.metric("Date Range", f"M{df_final['MONTH'].min()} → M{df_final['MONTH'].max()}")
            
            row2_col1, row2_col2, row2_col3, row2_col4 = st.columns(4)
            row2_col1.metric("Total Entities", df_final["Entity"].nunique())
            row2_col2.metric("Total Clients", df_final[df_final["Client"].str.startswith("ABC_")]["Client"].nunique())
            row2_col3.metric("Total Suppliers", df_final[df_final["Supplier"].str.startswith("SUP")]["Supplier"].nunique())
            row2_col4.metric("Total Employees", df_final[df_final["Employee"].str.startswith("DNA_")]["Employee"].nunique())            


    with col2:
            st.subheader("📈Total Monthly Amount Trend")
            
            # Keep only relevant columns based on currency filter
            if CURRENCY == "LCC only":
                st.line_chart(df_final.groupby(["MONTH"]).agg({"LCC AMOUNT": "sum"}).reset_index().set_index("MONTH"))
            elif CURRENCY == "EUR only":
                st.line_chart(df_final.groupby(["MONTH"]).agg({"EUR AMOUNT": "sum"}).reset_index().set_index("MONTH"))
            else:
                st.line_chart(df_final.groupby(["MONTH"]).agg({"LCC AMOUNT": "sum","EUR AMOUNT": "sum"}).reset_index().set_index("MONTH"))
