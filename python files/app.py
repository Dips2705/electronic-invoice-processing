from jpeg_invoice import convertImage
from pdf_invoice import convertPDF
from webcam_invoice import convertWebcam
from PIL import Image 
import pytesseract 
import sys 
from pdf2image import convert_from_path 
import os 
import cv2
import numpy as np
import re  
import xlsxwriter 
import math
import pandas as pd
import datetime
import streamlit as st
import time
import base64
import pandas as pd
st.set_option('deprecation.showfileUploaderEncoding', False)


file_type = ""

st.title("Electronic Invoice Processing") 
st.title("Team Casecation")  
st.title("")
st.sidebar.subheader("Select your choice of invoice")
if st.sidebar.button("Convert from PDF"):
  st.header("Invoice processing from PDF")
  file_type=" a pdf"
if st.sidebar.button("Convert from JPEG"):
  st.header("Invoice processing from JPEG")
  file_type = " an image"


st.sidebar.subheader("Instructions")
if st.sidebar.button("Instruction Manual"):
  st.header("Instructions")

# if st.sidebar.button("PDF"):
  
# uploaded_pdf = st.file_uploader("Upload an invoice in PDF Format", type="pdf")

# if st.button("Execute"):
  
#   convertPDF(uploaded_pdf)
#   st.write("Converting...")
    
#   execute_bar = st.progress(0)
#   status_text = st.empty()
#   for percent_complete in range(100):
#       time.sleep(0.03)
#       execute_bar.progress(percent_complete + 1)
#       status_text.text("%d %s"%(percent_complete+1, '%'))
#   status_text.text('Done!')

uploaded_image = st.file_uploader("Upload"+file_type, type="jpg")

if uploaded_image is not None:
    img = Image.open(uploaded_image)
    st.image(img, caption="User Input", width=1000, use_column_width=True)
    img.save("image.jpg")

if st.button("Execute"):
  
  convertImage('./image.jpg')
  st.write("Converting...")
    
  execute_bar = st.progress(0)
  status_text = st.empty()
  for percent_complete in range(100):
      time.sleep(0.03)
      execute_bar.progress(percent_complete + 1)
      status_text.text("%d %s"%(percent_complete+1, '%'))
  status_text.text('Done!')

if st.button("Download Converted Excel File"): 
  # st.text("Right click on the below link and click on 'Save the link'")
  # st.text("Then save the file with extension .xlsx") 
  def get_table_download_link(xlsx_file):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    
    data = open(xlsx_file, 'rb').read()
    base64_encoded = base64.b64encode(data).decode('UTF-8')
    # b64 = base64.b64encode(csv.encode()).decode()  # some strings <-> bytes conversions necessary here
    href = f'<a href="data:file/xlsx;base64,{base64_encoded}" download="converted.xlsx">Download excel file</a>'
    return href
  st.markdown(get_table_download_link('/content/converted.xlsx'), unsafe_allow_html=True)

if st.button("End Session"):
  os.remove('/content/converted.xlsx')
  st.title("Thank you! Visit Again!")


# if st.button("Execute"):
#   convertWebcam()