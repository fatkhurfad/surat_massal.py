import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO
import zipfile

st.set_page_config(page_title="Generator Surat Massal", layout="centered")
st.title("📄 Generator Surat Massal Word (Arial 12 + Hyperlink)")

uploaded_template = st.file_uploader("Upload Template Word (.docx)", type="docx")
uploaded_excel = st.file_uploader("Upload Data Excel (.xlsx)", type="xlsx")

if uploaded_template and uploaded_excel:
    df = pd.read_excel(uploaded_excel)

    if not {'nama_penyelenggara', 'short_link'}.issubset(df.columns):
        st.error("Excel harus memiliki kolom: 'nama_penyelenggara' dan 'short_link'")
    else:
        if st.button("🔄 Generate Surat"):
            output_zip = BytesIO()
            with zipfile.ZipFile(output_zip, "w") as zf:
                for _, row in df.iterrows():
                    doc = Document(uploaded_template)

                    for p in doc.paragraphs:
                        for run in p.runs:
                            if "{{nama_penyelenggara}}" in run.text:
                                run.text = run.text.replace("{{nama_penyelenggara}}", row["nama_penyelenggara"])

                    for p in doc.paragraphs:
                        if "{{short_link}}" in p.text:
                            parts = p.text.split("{{short_link}}")
                            p.clear()
                            if parts[0]:
                                p.add_run(parts[0])
                            link_run = p.add_run(row["short_link"])
                            link_run.font.name = "Arial"
                            link_run.font.size = Pt(12)
                            link_run.font.color.rgb = RGBColor(0, 0, 255)
                            link_run.underline = True
                            if len(parts) > 1:
                                p.add_run(parts[1])

                    for p in doc.paragraphs:
                        for run in p.runs:
                            run.font.name = "Arial"
                            run.font.size = Pt(12)

                    buffer = BytesIO()
                    filename = f"{row['nama_penyelenggara'].replace('/', '-')}.docx"
                    doc.save(buffer)
                    zf.writestr(filename, buffer.getvalue())

            st.success("✅ Surat berhasil dibuat!")
            st.download_button(
                label="📥 Download ZIP Hasil",
                data=output_zip.getvalue(),
                file_name="surat_massal_output.zip",
                mime="application/zip"
            )