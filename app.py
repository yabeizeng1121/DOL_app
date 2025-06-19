import streamlit as st
import pandas as pd
import os
from io import BytesIO
import tempfile
import pdfkit
from pdfkit.configuration import Configuration


CONFIG = Configuration(wkhtmltopdf="/usr/bin/wkhtmltopdf")


st.title("📦 UniUni Combined Bill of Lading PDF Generator (HTML版)")

ship_date = st.date_input("📅 Enter Ship Date (MM/DD/YYYY)")
uploaded_excel = st.file_uploader("📄 Upload Pickup Plan Excel File", type=["xlsx"])
uploaded_template = st.file_uploader("📄 Upload HTML Template File", type=["html"])


@st.cache_data(show_spinner="⏳ Generating PDF...")
def generate_combined_pdf(excel_bytes, html_bytes, ship_date_str, ship_date_short_str):
    df = pd.read_excel(excel_bytes)
    required_cols = ["Address", "Phone", "Note", "DSP"]
    if not all(col in df.columns for col in required_cols):
        return None, f"❌ Missing columns: {required_cols}"

    html_template = html_bytes.read().decode("utf-8")

    pdfs = []
    with tempfile.TemporaryDirectory() as tmpdir:
        for idx, row in df.iterrows():
            address = str(row["Address"])
            phone = str(row["Phone"])
            note = str(row["Note"])
            dsp = str(row["DSP"]).replace("/", "_").replace("\\", "_")
            seq = idx + 1

            replacements = {
                "SEA-[pickup address]+TEPHONE+NOTE": f"SEA - {address} | TEL: {phone} | Note: {note}",
                "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ": f"UNI-SEA-PICKUP-{ship_date_str}-{seq}",
                "Carrier Name: GN GREENWHEELS INC.": f"Carrier Name: GN GREENWHEELS INC. - {dsp}",
                "Ship_date": ship_date_short_str,
            }

            filled_html = html_template
            for key, val in replacements.items():
                filled_html = filled_html.replace(key, val)

            output_pdf_path = os.path.join(tmpdir, f"{seq}_{dsp}.pdf")

            pdfkit.from_string(
                filled_html,
                output_pdf_path,
                options={"enable-local-file-access": None},
                configuration=CONFIG,
            )

            pdfs.append(output_pdf_path)

        # Merge PDFs
        from PyPDF2 import PdfMerger

        merger = PdfMerger()
        for pdf in pdfs:
            merger.append(pdf)

        output_pdf = BytesIO()
        merger.write(output_pdf)
        merger.close()
        output_pdf.seek(0)
        return output_pdf, None


if uploaded_excel and uploaded_template and ship_date:
    full_date = ship_date.strftime("%m/%d/%Y")
    short_date = ship_date.strftime("%m/%d/%y")

    if st.button("🚀 Generate Combined PDF"):
        pdf_result, error = generate_combined_pdf(
            uploaded_excel, uploaded_template, full_date, short_date
        )
        if error:
            st.error(error)
        else:
            st.success("✅ Combined PDF generated!")
            st.download_button(
                "📥 Download Combined PDF",
                pdf_result,
                file_name="All_BOLs_Combined.pdf",
                mime="application/pdf",
            )
