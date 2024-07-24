import streamlit as st
from zipfile import ZipFile
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import base64
#------- OCR ------------
import pdf2image
import pytesseract
from pytesseract import Output, TesseractError

@st.cache_data
def convert_pdf_to_txt_file(path):
    texts = []
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    # fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    
    file_pages = PDFPage.get_pages(path)
    nbPages = len(list(file_pages))
    for page in PDFPage.get_pages(path):
      interpreter.process_page(page)
      t = retstr.getvalue()
    # text = retstr.getvalue()

    # fp.close()
    device.close()
    retstr.close()
    return t, nbPages

@st.cache_data
def save_pages(pages):
  
  files = []
  for page in range(len(pages)):
    filename = "page_"+str(page)+".txt"
    with open("./file_pages/"+filename, 'w', encoding="utf-8") as file:
      file.write(pages[page])
      files.append(file.name)
  
  # create zipfile object
  # zipPath = './file_pages/pdf_to_txt.zip'
  # zipObj = ZipFile(zipPath, 'w')
  # for f in files:
  #   zipObj.write(f)
  # zipObj.close()
    
  # return zipPath

def DadosRetornoCSV(TAMANHO_DADO, DADO_INICIO, DADO_FINAL, TEXTO_COMPLETO):
  DADOS_RESLTANTE = TEXTO_COMPLETO[DADO_INICIO + TAMANHO_DADO : DADO_FINAL-1]
  return DADOS_RESLTANTE.replace("\n", " ").strip()
