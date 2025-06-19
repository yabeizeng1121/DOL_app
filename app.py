import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import tempfile, os, re
from pathlib import Path
from subprocess import run, PIPE
from PyPDF2 import PdfMerger


def replace_all_text(doc: Document, reps: dict):
    """Find & replace in paragraphs & table cells."""
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


def convert_doc_to_pdf_native(doc_file: Path, outdir: Path, timeout: int = 60):
    """Call LibreOffice headless to convert DOCX→PDF."""
    try:
        proc = run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                str(outdir),
                str(doc_file),
            ],
            stdout=PIPE,
            stderr=PIPE,
            timeout=timeout,
            check=True,
        )
        out = proc.stdout.decode("utf-8")
        m = re.search(r"-> (.*?) using filter", out)
        if m:
            return Path(m.group(1)), None
        else:
            return None, "LibreOffice did not report output file"
    except Exception as e:
        return None, e


st.set_page_config(page_title="📄→📦 UniUni BOL", layout="wide")
st.title("UniUni BOL Generator (DOCX→PDF via LibreOffice)")

ship_date = st.date_input("Ship Date")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
uploaded_docx = st.file_uploader("Upload Word Template (.docx)", type=["docx"])


@st.cache_data(show_spinner="⏳ Generating PDF…")
def generate_pdf(excel_bytes: bytes, docx_bytes: bytes, date_str: str):
    df = pd.read_excel(BytesIO(excel_bytes))
    req = {"Address", "Phone", "Note", "DSP"}
    missing = req - set(df.columns)
    if missing:
        return None, f"Missing columns in Excel: {missing}"

    pdfs = []
    with tempfile.TemporaryDirectory() as tmp:
        tpl_path = Path(tmp) / "template.docx"
        tpl_path.write_bytes(docx_bytes)

        for i, row in df.iterrows():
            seq = i + 1
            doc = Document(tpl_path)
            reps = {
                "SEA-[pickup address]+TEPHONE+NOTE": f"SEA - {row['Address']} | TEL: {row['Phone']} | Note: {row['Note']}",
                "UNI-SEA-PICKUP-MM/DD/YYYY-SEQ": f"UNI-SEA-PICKUP-{date_str}-{seq}",
                "Carrier Name: GN GREENWHEELS INC.": f"Carrier Name: GN GREENWHEELS INC. - {row['DSP']}",
                "Ship_date": date_str,
            }
            replace_all_text(doc, reps)

            docx_out = Path(tmp) / f"{seq}.docx"
            pdf_out = Path(tmp) / f"{seq}.pdf"
            doc.save(docx_out)

            out_pdf, err = convert_doc_to_pdf_native(docx_out, Path(tmp))
            if err or not out_pdf or not out_pdf.exists():
                return None, f"Conversion failed on row {seq}: {err}"
            pdfs.append(str(out_pdf))
            st.info(f"✅ Row {seq} of {len(df)} done")

        # Merge them
        merger = PdfMerger()
        for p in pdfs:
            merger.append(p)
        buf = BytesIO()
        merger.write(buf)
        merger.close()
        buf.seek(0)
        return buf, None


if ship_date and uploaded_excel and uploaded_docx:
    ds = ship_date.strftime("%m/%d/%Y")
    if st.button("Generate Combined PDF"):
        pdf_buf, error = generate_pdf(uploaded_excel.read(), uploaded_docx.read(), ds)
        if error:
            st.error(error)
        else:
            st.success("✅ Here’s your merged PDF!")
            st.download_button(
                "Download All BOLs",
                data=pdf_buf,
                file_name="All_BOLs.pdf",
                mime="application/pdf",
            )
