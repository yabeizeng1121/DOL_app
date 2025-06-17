import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import tempfile
import os
from docx2pdf import convert
from PyPDF2 import PdfMerger
import pythoncom


# Replace placeholders in both paragraphs and tables (including multi-run content)
def replace_all_text(doc, replacements):
    for para in doc.paragraphs:
        inline = para.runs
        full_text = "".join(run.text for run in inline)
        for key, value in replacements.items():
            if key in full_text:
                full_text = full_text.replace(key, value)
                for run in inline:
                    run.text = ""
                if inline:
                    inline[0].text = full_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    inline = para.runs
                    full_text = "".join(run.text for run in inline)
                    for key, value in replacements.items():
                        if key in full_text:
                            full_text = full_text.replace(key, value)
                            for run in inline:
                                run.text = ""
                            if inline:
                                inline[0].text = full_text


# Streamlit UI
st.title("📦 UniUni Combined Bill of Lading PDF Generator")

# Track uploads to invalidate cache
if "last_excel" not in st.session_state:
    st.session_state.last_excel = None
if "last_template" not in st.session_state:
    st.session_state.last_template = None
if "last_date" not in st.session_state:
    st.session_state.last_date = None

# Inputs
ship_date = st.date_input("📅 Enter Ship Date (MM/DD/YYYY)")
uploaded_excel = st.file_uploader("📄 Upload Pickup Plan Excel File", type=["xlsx"])
uploaded_template = st.file_uploader("📄 Upload Word Template File", type=["docx"])

# Clear cache when inputs change
if (
    uploaded_excel != st.session_state.last_excel
    or uploaded_template != st.session_state.last_template
    or ship_date != st.session_state.last_date
):
    st.cache_data.clear()
    st.session_state.last_excel = uploaded_excel
    st.session_state.last_template = uploaded_template
    st.session_state.last_date = ship_date


# Cached generator
@st.cache_data(show_spinner="⏳ Generating PDF...")
def generate_combined_pdf(
    excel_bytes, template_bytes, ship_date_str, ship_date_short_str
):
    df = pd.read_excel(excel_bytes)
    required_columns = ["Address", "Phone", "Note", "DSP"]
    if not all(col in df.columns for col in required_columns):
        return None, f"❌ Excel is missing required columns: {required_columns}"

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_paths = []
        total = len(df)

        for idx, row in df.iterrows():
            doc = Document(template_bytes)

            address = str(row["Address"])
            phone = str(row["Phone"])
            note = str(row["Note"])
            dsp = str(row["DSP"]).replace("/", "_").replace("\\", "_")
            seq = idx + 1

            insert_pickup_text = f"SEA - {address} | TEL: {phone} | Note: {note}"
            bol_number = f"UNI-SEA-PICKUP-{ship_date_str}-{seq}"
            new_carrier_text = f"Carrier Name: GN GREENWHEELS INC. - {dsp}"

            replacements = {
                "SEA-[pickup address]+TEPHONE+NOTE": insert_pickup_text,
                "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ": bol_number,
                "Carrier Name: GN GREENWHEELS INC.": new_carrier_text,
                "Ship_date": ship_date_short_str,
            }

            replace_all_text(doc, replacements)

            doc_path = os.path.join(tmpdir, f"{seq}_{dsp}_BOL.docx")
            pdf_path = doc_path.replace(".docx", ".pdf")
            doc.save(doc_path)

            try:
                pythoncom.CoInitialize()  # 💡 fix COM error
                convert(doc_path, pdf_path)
                pythoncom.CoUninitialize()
                pdf_paths.append(pdf_path)
                # st.info(f"✅ Page {seq} of {total} ready")
            except Exception as e:
                return None, f"❌ PDF conversion failed on page {seq}: {e}"

        merger = PdfMerger()
        for pdf in pdf_paths:
            merger.append(pdf)

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
            st.success("✅ Combined PDF generated successfully")
            st.download_button(
                "📥 Download Combined PDF",
                pdf_result,
                file_name="All_BOLs_Combined.pdf",
            )
