import streamlit as st
import pandas as pd

# st.write("Hello world!")
# st.write("We love studying")

# st.write("Testando streamlit. By Nielson")

st.set_page_config(layout="wide")

with st.sidebar:
    st.title("Conversão de PDF para CSV")
    uploaded_file = st.file_uploader("Coloque o seu arquivo PDF aqui")