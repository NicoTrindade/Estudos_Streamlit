import streamlit as st
import pdf2image
from PIL import Image
import pytesseract
from pytesseract import Output, TesseractError
from funcoes import convert_pdf_to_txt_file, save_pages

st.set_page_config(layout="wide")

with st.sidebar:
    st.title("Conversão de PDF para CSV")
    uploaded_file = st.file_uploader("Coloque o seu arquivo PDF aqui")

if uploaded_file is not None:
    path = uploaded_file.read()
    text_data_f, nbPages = convert_pdf_to_txt_file(uploaded_file)
    
    totalPages = "Pages: "+str(nbPages)+" in total"
    st.info(totalPages)
    st.download_button("Download txt file", text_data_f)
