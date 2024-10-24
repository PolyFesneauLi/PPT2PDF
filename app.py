import streamlit as st
import tempfile
import base64
from spire.presentation import Presentation
from spire.presentation import FileFormat as pFileFormat
from spire.pdf import PdfDocument

import os
def ppt_to_pdf(ppt_path, pdf_path): # font error
    pdf = PdfDocument()
    # 创建Presentation对象
    presentation = Presentation()
    # 载入PowerPoint文件
    presentation.LoadFromFile(ppt_path)
    # 将PowerPoint文件转换为PDF
    presentation.SaveToFile(pdf_path, pFileFormat.PDF)
    presentation.Dispose()
    pdf.Close()

def embed_pdf(pdf_path):
    with open(pdf_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
        pdf_display = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf">'
        return pdf_display
def main():
    # web page title
    st.set_page_config(page_title="PowerPoint Presentation Generator",layout="wide")
    st.title("PPT2PDF")
    ppt = None
    # upload the file
    uploaded_file = st.file_uploader("Choose a file", type=['pptx'])
    if uploaded_file is not None:
        ppt = uploaded_file
        st.write("File uploaded successfully")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
            tmp_ppt.write(ppt.getbuffer())
            tmp_ppt_path = tmp_ppt.name
            tmp_pdf_path = tmp_ppt_path.replace(".pptx", ".pdf")
        ppt_to_pdf(tmp_ppt_path,tmp_pdf_path)
        st.markdown(embed_pdf(tmp_pdf_path), unsafe_allow_html=True)
        os.remove(tmp_ppt_path)
        os.remove(tmp_pdf_path)

    
if __name__ == "__main__":
    main()
