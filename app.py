import re
from datetime import datetime
from io import BytesIO

import pdfplumber
import pandas as pd
import streamlit as st


def convert_date(date_text):
    return datetime.strptime(date_text.strip(), "%d.%m.%Y").strftime("%m/%d/%Y")


def parse_amount(value):
    if not value:
        return None

    clean = str(value).replace(",", "").strip()

    try:
        return float(clean)
    except ValueError:
        return None


def is_valid_invoice(value):
    return bool(value and re.match(r"^\d{8}B\d+$", value.strip()))


def is_valid_billing_doc(value):
    return bool(value and re.match(r"^\d{10}$", value.strip()))


def build_file_name(rows):
    dates = [row["_sort_date"] for row in rows]

    start_month = min(dates).strftime("%b %Y")
    end_month = max(dates).strftime("%b %Y")

    return f"{start_month} - {end_month}.xlsx"


def extract_required_rows(pdf_file):
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                for row in table:
                    if not row or len(row) < 4:
                        continue

                    invoice_no = str(row[0] or "").strip().replace("\n", "")
                    billing_doc_no = str(row[1] or "").strip().replace("\n", "")
                    date_text = str(row[2] or "").strip().replace("\n", "")
                    debit_text = str(row[3] or "").strip().replace("\n", "")

                    if "Invoice" in invoice_no or "Billing" in billing_doc_no:
                        continue

                    if not is_valid_invoice(invoice_no):
                        continue

                    if not is_valid_billing_doc(billing_doc_no):
                        continue

                    debit_value = parse_amount(debit_text)

                    if debit_value is None:
                        continue

                    try:
                        formatted_date = convert_date(date_text)
                        date_obj = datetime.strptime(date_text, "%d.%m.%Y")
                    except ValueError:
                        continue

                    rows.append({
                        "Date": formatted_date,
                        "Csh Rpt No / Invoice No": invoice_no,
                        "Cheque no. / Billing Doc No.": billing_doc_no,
                        "Debit": debit_value,
                        "_sort_date": date_obj
                    })

    return rows


def remove_duplicate_combinations(df):
    return df.drop_duplicates(
        subset=[
            "Csh Rpt No / Invoice No",
            "Cheque no. / Billing Doc No."
        ],
        keep="first"
    )


def group_and_sort_rows(rows):
    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df = df.sort_values(
        by=["Debit", "_sort_date"],
        ascending=[False, True]
    )

    df["MONTH"] = df["_sort_date"].dt.strftime("%B %Y")

    final_rows = []

    for month, month_group in df.groupby("MONTH", sort=False):
        debit_groups = month_group.sort_values(
            by="Debit",
            ascending=False
        ).groupby("Debit", sort=False)

        for debit, debit_group in debit_groups:
            debit_group = debit_group.sort_values(
                by="_sort_date",
                ascending=True
            )

            for _, row in debit_group.iterrows():
                final_rows.append({
                    "MONTH": month,
                    "GLOBAL SERIAL NUMBER": row["Cheque no. / Billing Doc No."],
                    "INVOICE NUMBER": row["Csh Rpt No / Invoice No"],
                    "INVOICE DATE": row["Date"],
                    "DEBIT": row["Debit"]
                })

            final_rows.append({
                "MONTH": "",
                "GLOBAL SERIAL NUMBER": "",
                "INVOICE NUMBER": "",
                "INVOICE DATE": "",
                "DEBIT": ""
            })

    return pd.DataFrame(final_rows)


def auto_adjust_column_width(worksheet):
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)


st.set_page_config(page_title="Fuel Statement Extractor", layout="wide")

st.title("Fuel Statement Extractor")
st.write("Upload one or more statement PDFs and generate an Excel file with month-wise sheets.")

uploaded_pdfs = st.file_uploader(
    "Upload Statement PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_pdfs:
    st.info(f"{len(uploaded_pdfs)} PDF file(s) selected.")

    if st.button("Extract Data"):
        all_rows = []

        for pdf_file in uploaded_pdfs:
            extracted_rows = extract_required_rows(pdf_file)
            all_rows.extend(extracted_rows)

        if not all_rows:
            st.warning("No valid rows found.")
        else:
            original_count = len(all_rows)

            df = pd.DataFrame(all_rows)
            df = remove_duplicate_combinations(df)

            duplicate_count = original_count - len(df)

            rows_after_duplicates = df.to_dict("records")
            file_name = build_file_name(rows_after_duplicates)

            st.success(f"Extracted {original_count} total valid rows.")
            st.info(f"Removed {duplicate_count} duplicate row(s).")
            st.info(f"Final row count: {len(df)}")
            st.info(f"Generated file name: {file_name}")

            df = df.sort_values(
                by=["_sort_date", "Debit"],
                ascending=[True, False]
            )

            df["MONTH"] = df["_sort_date"].dt.strftime("%B %Y")

            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for month, month_group in df.groupby("MONTH", sort=False):
                    month_rows = month_group.to_dict("records")
                    month_df = group_and_sort_rows(month_rows)

                    sheet_name = month[:31]

                    month_df.to_excel(
                        writer,
                        index=False,
                        sheet_name=sheet_name
                    )

                    worksheet = writer.sheets[sheet_name]
                    auto_adjust_column_width(worksheet)

            st.download_button(
                label="Download Excel",
                data=output.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )