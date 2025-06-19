import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import tempfile
import os
import subprocess
from PyPDF2 import PdfMerger


# Helper to replace placeholders in paragraphs & tables
def replace_all_text(doc, reps):
    for para in doc.paragraphs:
        runs = para.runs
        txt = "".join(r.text for r in runs)
        for k, v in reps.items():
            if k in txt:
                txt = txt.replace(k, v)
                for r in runs:
                    r.text = ""
                runs[0].text = txt
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    runs = para.runs
                    txt = "".join(r.text for r in runs)
                    for k, v in reps.items():
                        if k in txt:
                            txt = txt.replace(k, v)
                            for r in runs:
                                r.text = ""
                            runs[0].text = txt


st.set_page_config(page_title="DOCX→PDF BOL", layout="wide")
st.title("📄→📦 UniUni BOL Generator")

ship_date = st.date_input("Enter Ship Date")
uploaded_xl = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
uploaded_docx = st.file_uploader("Upload Word Template (.docx)", type=["docx"])


@st.cache_data(show_spinner="⏳ Generating PDF…")
def make_bols(excel_bytes, docx_bytes, ship_date_str):
    df = pd.read_excel(BytesIO(excel_bytes))
    needed = {"Address", "Phone", "Note", "DSP"}
    if needed - set(df.columns):
        return None, f"Missing cols: {needed - set(df.columns)}"

    pdfs = []
    with tempfile.TemporaryDirectory() as tmp:
        tpl = os.path.join(tmp, "tpl.docx")
        with open(tpl, "wb") as f:
            f.write(docx_bytes)

        for i, row in df.iterrows():
            seq = i + 1
            doc = Document(tpl)
            reps = {
                "SEA-[pickup address]+TEPHONE+NOTE": f"SEA - {row['Address']} | TEL: {row['Phone']} | Note: {row['Note']}",
                "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ": f"UNI-SEA-PICKUP-{ship_date_str}-{seq}",
                "Carrier Name: GN GREENWHEELS INC.": f"Carrier Name: GN GREENWHEELS INC. - {row['DSP']}",
                "Ship_date": ship_date_str,
            }
            replace_all_text(doc, reps)

            docx_out = os.path.join(tmp, f"{seq}.docx")
            pdf_out = os.path.join(tmp, f"{seq}.pdf")
            doc.save(docx_out)

            # LibreOffice headless conversion
            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    docx_out,
                    "--outdir",
                    tmp,
                ],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
            )
            if not os.path.exists(pdf_out):
                return None, f"Conversion failed for row {seq}"
            pdfs.append(pdf_out)

        # merge
        merger = PdfMerger()
        for p in pdfs:
            merger.append(p)
        buf = BytesIO()
        merger.write(buf)
        buf.seek(0)
        return buf, None


if ship_date and uploaded_xl and uploaded_docx:
    date_str = ship_date.strftime("%m/%d/%Y")
    if st.button("Generate Combined PDF"):
        pdf_buf, err = make_bols(uploaded_xl.read(), uploaded_docx.read(), date_str)
        if err:
            st.error(err)
        else:
            st.success("Here you go!")
            st.download_button(
                "Download All BOLs",
                data=pdf_buf,
                file_name="All_BOLs.pdf",
                mime="application/pdf",
            )
