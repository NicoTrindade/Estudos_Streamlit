import streamlit as st
from funcoes import convert_pdf_to_txt_file, save_pages

st.set_page_config(layout="wide")

with st.sidebar:
    st.title("Conversão de PDF para CSV")
    uploaded_file = st.file_uploader("Coloque o seu arquivo PDF aqui")

if uploaded_file is not None:
    path = uploaded_file.read()
    text_data_f, nbPages = convert_pdf_to_txt_file(uploaded_file)

    zipPath = save_pages(text_data)
    
    totalPages = "Pages: "+str(nbPages)+" in total"
    st.info(totalPages)
    st.download_button("Download txt file", text_data_f)

    # download text data   
    with open(zipPath, "rb") as fp:
        btn = st.download_button(
            label="Download ZIP (txt)",
            data=fp,
            file_name="pdf_to_txt.zip",
            mime="application/zip"
        )
