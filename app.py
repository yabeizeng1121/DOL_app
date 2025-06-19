import streamlit as st
import pandas as pd
import pdfkit
from io import BytesIO
import tempfile
import os
from PyPDF2 import PdfMerger
import shutil

# Dynamically find wkhtmltopdf path
path_wkhtmltopdf = shutil.which("wkhtmltopdf")
if path_wkhtmltopdf is None:
    raise OSError("wkhtmltopdf not found in PATH")

config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

st.title("📦 UniUni Combined Bill of Lading PDF Generator")

if "last_excel" not in st.session_state:
    st.session_state.last_excel = None
if "last_template" not in st.session_state:
    st.session_state.last_template = None
if "last_date" not in st.session_state:
    st.session_state.last_date = None

ship_date = st.date_input("📅 Enter Ship Date (MM/DD/YYYY)")
uploaded_excel = st.file_uploader("📄 Upload Pickup Plan Excel File", type=["xlsx"])
uploaded_template = st.file_uploader("📄 Upload HTML Template File", type=["html"])

if (
    uploaded_excel != st.session_state.last_excel
    or uploaded_template != st.session_state.last_template
    or ship_date != st.session_state.last_date
):
    st.cache_data.clear()
    st.session_state.last_excel = uploaded_excel
    st.session_state.last_template = uploaded_template
    st.session_state.last_date = ship_date


@st.cache_data(show_spinner="⏳ Generating PDF...")
def generate_combined_pdf(excel_file, html_file, full_date, short_date):
    df = pd.read_excel(excel_file)
    required_columns = ["Address", "Phone", "Note", "DSP"]
    if not all(col in df.columns for col in required_columns):
        return None, f"❌ Excel is missing required columns: {required_columns}"

    html_template = html_file.read().decode("utf-8")

    with tempfile.TemporaryDirectory() as tmpdir:
        pdfs = []

        for idx, row in df.iterrows():
            address = str(row["Address"])
            phone = str(row["Phone"])
            note = str(row["Note"])
            dsp = str(row["DSP"]).replace("/", "_").replace("\\", "_")
            seq = idx + 1

            filled_html = html_template.replace(
                "SEA-[pickup address]+TEPHONE+NOTE",
                f"SEA - {address} | TEL: {phone} | Note: {note}",
            )
            filled_html = filled_html.replace(
                "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ", f"UNI-SEA-PICKUP-{full_date}-{seq}"
            )
            filled_html = filled_html.replace(
                "Carrier Name: GN GREENWHEELS INC.",
                f"Carrier Name: GN GREENWHEELS INC. - {dsp}",
            )
            filled_html = filled_html.replace("Ship_date", short_date)

            output_pdf_path = os.path.join(tmpdir, f"{seq}_{dsp}.pdf")
            pdfkit.from_string(
                filled_html,
                output_pdf_path,
                options={"enable-local-file-access": None},
                configuration=config,
            )

            with open(output_pdf_path, "rb") as f:
                pdfs.append(f.read())

        merged = PdfMerger()
        for i, pdf_bytes in enumerate(pdfs):
            temp_pdf_path = os.path.join(tmpdir, f"temp_{i}.pdf")
            with open(temp_pdf_path, "wb") as temp_pdf:
                temp_pdf.write(pdf_bytes)
            merged.append(temp_pdf_path)

        output = BytesIO()
        merged.write(output)
        merged.close()
        output.seek(0)
        return output, None


if uploaded_excel and uploaded_template and ship_date:
    full_date = ship_date.strftime("%m/%d/%Y")
    short_date = ship_date.strftime("%m/%d/%y")

    if st.button("🚀 Generate Combined PDF"):
        pdf_result, error_msg = generate_combined_pdf(
            uploaded_excel, uploaded_template, full_date, short_date
        )

        if error_msg:
            st.error(error_msg)
        else:
            st.success("✅ Combined PDF generated successfully!")
            st.download_button(
                label="📥 Download Combined PDF",
                data=pdf_result,
                file_name="All_BOLs_Combined.pdf",
                mime="application/pdf",
            )
