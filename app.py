import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import tempfile
import os
import subprocess
from PyPDF2 import PdfMerger


# Replace placeholders in paragraphs & tables
def replace_all_text(doc, replacements):
    for para in doc.paragraphs:
        runs = para.runs
        text = "".join(r.text for r in runs)
        for key, val in replacements.items():
            if key in text:
                text = text.replace(key, val)
                for r in runs:
                    r.text = ""
                if runs:
                    runs[0].text = text
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    runs = para.runs
                    text = "".join(r.text for r in runs)
                    for key, val in replacements.items():
                        if key in text:
                            text = text.replace(key, val)
                            for r in runs:
                                r.text = ""
                            if runs:
                                runs[0].text = text


st.set_page_config(page_title="UniUni BOL DOCX→PDF", layout="wide")
st.title("📄→📦 BOL Generator (DOCX→PDF via LibreOffice)")

ship_date = st.date_input("📅 Enter Ship Date")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
uploaded_docx = st.file_uploader("Upload Word Template (.docx)", type=["docx"])


@st.cache_data(show_spinner="⏳ Generating Combined PDF…")
def generate_combined_pdf(excel_bytes, docx_bytes, ship_date_str):
    df = pd.read_excel(BytesIO(excel_bytes))
    needed = {"Address", "Phone", "Note", "DSP"}
    missing = needed - set(df.columns)
    if missing:
        return None, f"❌ Excel missing columns: {sorted(missing)}"

    pdf_paths = []
    with tempfile.TemporaryDirectory() as tmpdir:
        total = len(df)
        # Save template once per run to disk
        tpl_path = os.path.join(tmpdir, "template.docx")
        with open(tpl_path, "wb") as f:
            f.write(docx_bytes)

        for idx, row in df.iterrows():
            seq = idx + 1
            # load and replace
            doc = Document(tpl_path)
            reps = {
                "SEA-[pickup address]+TEPHONE+NOTE": f"SEA - {row['Address']} | TEL: {row['Phone']} | Note: {row['Note']}",
                "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ": f"UNI-SEA-PICKUP-{ship_date_str}-{seq}",
                "Carrier Name: GN GREENWHEELS INC.": f"Carrier Name: GN GREENWHEELS INC. - {row['DSP']}",
                "Ship_date": ship_date_str,
            }
            replace_all_text(doc, reps)

            # write filled docx
            out_docx = os.path.join(tmpdir, f"{seq}.docx")
            out_pdf = out_docx.replace(".docx", ".pdf")
            doc.save(out_docx)

            # convert to PDF via LibreOffice
            try:
                subprocess.run(
                    [
                        "soffice",
                        "--headless",
                        "--convert-to",
                        "pdf",
                        out_docx,
                        "--outdir",
                        tmpdir,
                    ],
                    check=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.PIPE,
                )
            except subprocess.CalledProcessError as e:
                err = e.stderr.decode("utf-8", errors="ignore")
                return None, f"❌ Conversion failed on row {seq}: {err}"
            if not os.path.exists(out_pdf):
                return None, f"❌ PDF not created for row {seq}"

            pdf_paths.append(out_pdf)
            st.info(f"✅ Page {seq}/{total} done")

        # merge all PDFs
        merger = PdfMerger()
        for p in pdf_paths:
            merger.append(p)
        combined = BytesIO()
        merger.write(combined)
        merger.close()
        combined.seek(0)
        return combined, None


if ship_date and uploaded_excel and uploaded_docx:
    full_str = ship_date.strftime("%m/%d/%Y")
    if st.button("🚀 Generate Combined PDF"):
        pdf_bytes, err = generate_combined_pdf(
            uploaded_excel.read(), uploaded_docx.read(), full_str
        )
        if err:
            st.error(err)
        else:
            st.success("✅ Ready!")
            st.download_button(
                "📥 Download All BOLs PDF",
                data=pdf_bytes,
                file_name="All_BOLs_Combined.pdf",
                mime="application/pdf",
            )
