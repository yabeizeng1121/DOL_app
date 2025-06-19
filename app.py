import streamlit as st
import pandas as pd
from io import BytesIO
import os
import tempfile
from PyPDF2 import PdfMerger
from weasyprint import HTML

st.title("📦 UniUni Combined Bill of Lading PDF Generator (HTML Version)")

# Session cache
if "last_excel" not in st.session_state:
    st.session_state.last_excel = None
if "last_template" not in st.session_state:
    st.session_state.last_template = None
if "last_date" not in st.session_state:
    st.session_state.last_date = None

# Input UI
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
def generate_combined_pdf(
    excel_bytes, html_template_bytes, ship_date_str, ship_date_short_str
):
    df = pd.read_excel(excel_bytes)
    required_columns = ["Address", "Phone", "Note", "DSP"]
    if not all(col in df.columns for col in required_columns):
        return None, f"❌ Excel is missing required columns: {required_columns}"

    template = html_template_bytes.read().decode("utf-8")
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_paths = []

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

            html_filled = template
            for key, value in replacements.items():
                html_filled = html_filled.replace(key, value)

            pdf_path = os.path.join(tmpdir, f"{seq}_{dsp}_BOL.pdf")
            HTML(string=html_filled).write_pdf(pdf_path)
            pdf_paths.append(pdf_path)

        # Merge PDFs
        merger = PdfMerger()
        for path in pdf_paths:
            merger.append(path)

        output_pdf = BytesIO()
        merger.write(output_pdf)
        merger.close()
        output_pdf.seek(0)
        return output_pdf, None


# Main logic
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
