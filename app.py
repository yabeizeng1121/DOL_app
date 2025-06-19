import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import os
from xhtml2pdf import pisa

st.title("📦 UniUni Combined Bill of Lading PDF Generator (HTML Template)")

ship_date = st.date_input("📅 Enter Ship Date (MM/DD/YYYY)")
uploaded_excel = st.file_uploader("📄 Upload Pickup Plan Excel File", type=["xlsx"])
uploaded_template = st.file_uploader("📄 Upload HTML Template File", type=["html"])


@st.cache_data(show_spinner="⏳ Generating PDF...")
def generate_combined_pdf(excel_file, html_template_file, full_date, short_date):
    df = pd.read_excel(excel_file)
    required_columns = ["Address", "Phone", "Note", "DSP"]
    if not all(col in df.columns for col in required_columns):
        return None, f"❌ Excel is missing required columns: {required_columns}"

    template_str = html_template_file.read().decode("utf-8")
    pdf_files = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for idx, row in df.iterrows():
            address = str(row["Address"])
            phone = str(row["Phone"])
            note = str(row["Note"])
            dsp = str(row["DSP"]).replace("/", "_").replace("\\", "_")
            seq = idx + 1

            filled_html = (
                template_str.replace(
                    "SEA-[pickup address]+TEPHONE+NOTE",
                    f"SEA - {address} | TEL: {phone} | Note: {note}",
                )
                .replace(
                    "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ", f"UNI-SEA-PICKUP-{full_date}-{seq}"
                )
                .replace(
                    "Carrier Name: GN GREENWHEELS INC.",
                    f"Carrier Name: GN GREENWHEELS INC. - {dsp}",
                )
                .replace("Ship_date", short_date)
            )

            output_pdf_path = os.path.join(tmpdir, f"{seq}_{dsp}.pdf")
            with open(output_pdf_path, "wb") as pdf_file:
                pisa.CreatePDF(filled_html, dest=pdf_file)

            with open(output_pdf_path, "rb") as f:
                pdf_files.append(f.read())

        # Merge all PDFs
        from PyPDF2 import PdfMerger

        merger = PdfMerger()
        for pdf_bytes in pdf_files:
            merger.append(BytesIO(pdf_bytes))

        merged_pdf = BytesIO()
        merger.write(merged_pdf)
        merger.close()
        merged_pdf.seek(0)
        return merged_pdf, None


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
                "📥 Download Combined PDF",
                pdf_result,
                file_name="All_BOLs_Combined.pdf",
            )
