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
st.set_option('deprecation.showfileUploaderEncoding', False)

from .jpeg_invoice import convertImage
from .pdf_invoice import convertPDF
from .webcam_invoice import convertWebcam


st.title("Electronic Invoice Processing")   

st.sidebar.subheader("Select your choice of invoice")
st.header("Electronic Invoice Processing")
if st.sidebar.button("PDF"):
  uploaded_pdf = st.file_uploader("Upload an invoice in PDF Format", type="pdf")

if st.button("Execute"):
  file_name = "/content/Sample_Invoice_2.pdf"
  convertPDF(file_name)
  st.write("Converting...")
    
  execute_bar = st.progress(0)
  status_text = st.empty()
  for percent_complete in range(100):
      time.sleep(0.03)
      execute_bar.progress(percent_complete + 1)
      status_text.text("%d %s"%(percent_complete+1, '%'))
  status_text.text('Done!')