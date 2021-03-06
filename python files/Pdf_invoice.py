
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


def convertPDF(pdf_file_path):
  black = (0,0,0)
  white = (255,255,255)
  threshold = (160,160,160)
  config = ('-l eng --oem 3 --psm 6 ')
  image_counter = 1

  PDF_file = pdf_file_path
  pages = convert_from_path(PDF_file, 700)

  for page in pages: 
    
      filename = "page_"+str(image_counter)+".jpg"
      page.save(filename, 'JPEG') 
      img = Image.open(filename).convert("LA")
      pixels = img.getdata()
      newPixels = []
      for pixel in pixels:
          if pixel < threshold:
              newPixels.append(black)
          else:
              newPixels.append(white)

      newImg = Image.new("RGB",img.size)
      newImg.putdata(newPixels)
      newImg.save("new_image"+str(image_counter)+".jpg",dpi=(300,300))
      image = cv2.imread("new_image"+str(image_counter)+".jpg")
      mask = np.zeros(image.shape, dtype=np.uint8)
      gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
      thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
      cnts = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
      cnts = cnts[0] if len(cnts) == 2 else cnts[1]
      for c in cnts:
          area = cv2.contourArea(c)
          if area < 10000:
              cv2.drawContours(mask, [c], -1, (255,255,255), -1)

      mask = cv2.cvtColor(mask, cv2.COLOR_BGR2GRAY)
      result = cv2.bitwise_and(image,image,mask=mask)
      result[mask==0] = (255,255,255)
      width = int(result.shape[1] * 1.5)
      height = int(result.shape[0] * 1.5)
      dsize = (width, height)
      output = cv2.resize(result, dsize)
      cv2.imwrite("ne_image"+str(image_counter)+".jpg", output)
      image_counter = image_counter + 1

  filelimit = image_counter-1
  outfile =  'out_text.txt'
  f = open(outfile, "a") 
  for i in range(1, filelimit + 1): 
      filename = "ne_image"+str(i)+".jpg"
      text = str(((pytesseract.image_to_string(Image.open(filename),config=config)))) 
      text = text.replace('-\n', '')     
      f.write(text)
      os.remove("page_"+str(i)+".jpg")
      os.remove("new_image"+str(i)+".jpg")
      os.remove("ne_image"+str(i)+".jpg")

  f.close()
  
  #The code is specifically for India and can be extended to other countries as well . This is because Seller's state can be extracted from Gstn no then he list will be enormous .
  final_file=str(datetime.datetime.now().strftime("%d-%m-%Y_%I-%M-%S_%p"))+'.'+'xlsx'
  workbook = xlsxwriter.Workbook(final_file) 
  outfile='out_text.txt'
  worksheet = workbook.add_worksheet() 
  filename=outfile
  list={'01':'JAMMU AND KASHMIR','02':'HIMACHAL PRADESH','03':'PUNJAB','04':'CHANDIGARH','05':'UTTARAKHAND','06':'HARYANA','07':'DELHI','08':'RAJASTHAN','09':'UTTAR PRADESH','10':'BIHAR',
      '11':'SIKKIM','12':'ARUNACHAL PRADESH','13':'NAGALAND','14':'MANIPUR','15':'MIZORAM','16':'TRIPURA','17':'MEGHALAYA','1/':'MEGHALAYA','18':'ASSAM','19':'WEST BENGAL','20':'JHARKHAND',
      '21':'ODISHA','22':'CHATTISGARH','23':'MADHYA PRADESH','24':'GUJARAT','25':'DAMAN AND DIU','26':'DADRA AND NAGAR HAVELI','27':'MAHARASHTRA','28':'ANDHRA PRADESH (old)','29':'KARNATAKA',
      '30':'GOA','31':'LAKSHWADEEP','32':'KERALA','33':'TAMIL NADU','34':'PUDUCHERRY','35':'ANDAMAN AND NICOBAR ISLANDS','36':'TELANGANA','37':'ANDHRA PRADESH (NEW)'}       

  merge_format = workbook.add_format({'bold': 0,'border': 1,'align': 'center','valign': 'bottom','fg_color': '#BDD7EE'})
  merge_format_1 = workbook.add_format({'bold': 0,'border': 1,'align': 'center','valign': 'vcenter'})
  merge_format_2 = workbook.add_format({'bold': 0,'border': 1,'align': 'center','valign': 'bottom','fg_color': '#FF9E01'})
  merge_format_3 = workbook.add_format({'bold': .5,'border': 1,'align': 'left','valign': 'bottom','fg_color': '#BDD7EE'})
  merge_format_4 = workbook.add_format({'bold': .5,'border': 1,'align': 'center','valign': 'bottom','fg_color': '#BDD7EE'})
  merge_format_5 = workbook.add_format({'bold': .5,'border': 1,'align': 'center','valign': 'vcenter'})

  worksheet.set_column( 'A:Z',14)
  worksheet.set_default_row(16)
  worksheet.merge_range('A1:S2', 'GST Invoice', merge_format_4)
  worksheet.merge_range('A3:D3', 'Seller State', merge_format)
  worksheet.merge_range('A4:D4', 'Seller ID', merge_format)
  worksheet.merge_range('A5:D5', 'Seller Name', merge_format)
  worksheet.merge_range('A6:D10', 'Seller Address', merge_format)
  worksheet.merge_range('A11:D11', 'Seller GSTIN Number', merge_format)
  worksheet.merge_range('A12:D12', 'Country of Origin', merge_format)
  worksheet.merge_range('A13:D13', 'Currency', merge_format)
  worksheet.merge_range('A14:D14', 'Description', merge_format)
  worksheet.merge_range('A15:D16', '', merge_format)
  worksheet.merge_range('J3:L3', 'Invoice Number', merge_format)
  worksheet.merge_range('J4:L4', 'Invoice Date', merge_format)
  worksheet.merge_range('J5:L5', 'Due Date', merge_format)
  worksheet.merge_range('J6:L6', 'Total Invoice amount entered by WH operator ', merge_format)
  worksheet.merge_range('J7:L7', 'Total Invoice Quantity entered by WH operator ', merge_format)
  worksheet.merge_range('J8:L8', 'Total TCS Collected', merge_format)
  worksheet.merge_range('J9:L9', 'Round Off Charges', merge_format)
  worksheet.merge_range('J10:L10', 'PO Number', merge_format_2)
  worksheet.merge_range('J11:L11', 'Invoice Items Total Amount', merge_format)
  worksheet.merge_range('J12:L12', 'Invoice Items Total Quantity', merge_format)
  worksheet.merge_range('J13:L13', 'Buyer GSTIN Number', merge_format_2)
  worksheet.merge_range('J14:L16', 'Ship to Address', merge_format_4)
  worksheet.write('A17', 'S.No',merge_format_3)
  worksheet.write('B17','Product ID',merge_format_3)
  worksheet.write('C17', 'SKU',merge_format_3)
  worksheet.write('D17', 'HSN',merge_format_3)
  worksheet.write('E17','Title',merge_format_3)
  worksheet.write('F17', 'Quantity',merge_format_3)
  worksheet.write('G17', 'Unit Price',merge_format_3)
  worksheet.write('H17','Excise Duty',merge_format_3)
  worksheet.write('I17', 'Discount Percent',merge_format_3)
  worksheet.write('J17','SGST Percent',merge_format_3)
  worksheet.write('K17','CGST Percent',merge_format_3)
  worksheet.write('L17', 'IGST Percent',merge_format_3)
  worksheet.write('M17', 'Cess Percent',merge_format_3)
  worksheet.write('N17','TCS Percent',merge_format_3)
  worksheet.write('O17','Total Amount',merge_format_3)
  worksheet.write('P17', 'APP %',merge_format_3)
  worksheet.write('J2','GST Invoice',merge_format_3)

  #Searching for line of @ for mail id . Then extracting mail id by specifying format of email id .
  j=0
  line_num=0
  specific_line=''
  file = open(filename, "r")
  search_phrase = '@'
  for line in file.readlines():
      line_num += 1
      if line.find(search_phrase) >= 0:
          j=line_num
          break
  file.close()
  file = open(filename)
  all_lines_variable = file.readlines()
  specific_line=all_lines_variable[j-1]
  email = re.findall('\S+@\S+', specific_line) 
  is_email=bool(email)
  if is_email==True:
      worksheet.merge_range('E4:I4',email[0],merge_format_1 )
  file.close()

  #GST number is a 15 digit number with first 2 digit as state code next 10 digit as PAN no .Second most GST number is Buyer's GST number. 
  #Most of the GST number in given invoice is wrong so format is somewhat changed . 
  #Second most GST number is buyer's GST number
  file = open(filename, "r")
  l=0
  inter_state=0
  content = file.read()
  #increaing accuracy
  pattern = "[0-9O/]{2}[A-Z]{5,6}[0-9O/]{3,4}[A-Z]{1}[A-Z0-9/]{3,4}[\s]"
  GST = re.findall(pattern, content)
  for i in range(0,len(GST)):
      GST[i]=re.sub('/','7',GST[i])
  for i in range(0,len(GST)):
      GST[i]=re.sub('O','0',GST[i])
  if len(GST)>0:
      seller_gst=GST[0]
      worksheet.merge_range('E11:I11',seller_gst,merge_format_1 )
      #Checking first 2 digits of GSTN No for Seller's state                    
      if seller_gst[:2] in list.keys():
          worksheet.merge_range('E3:I3',list[seller_gst[:2]],merge_format_1 )

  if len(GST)>1:
      buyer_gst=GST[1]
      worksheet.merge_range('M13:Q13', buyer_gst,merge_format_1)
      if buyer_gst[:2]==seller_gst[:2]:
          inter_state=inter_state+1
  file.close()

  #Seller's country of origin
  worksheet.merge_range('E12:I12', 'India',merge_format_1)

  #Finding Seller's description(PAN) . PAN no is a 10 digit number with combination of 5 letters , 4 digits , 1 letter 
  file = open(filename, "r")
  content = file.read()
  pattern = "[\s][A-Z]{5}[0-9]{4}[A-Z]{1}[\s]"
  pan = re.findall(pattern, content)
  is_pan=bool(pan)
  if is_pan == True:
      worksheet.merge_range('E14:I14',pan[0],merge_format_1 )
  file.close()

  #The occurance of first India signifies address of seller . The line above it signifies Name of seller .
  f = open(filename,'r')
  line_num = 0
  j=0
  search_phrase = "India"
  for line in f.readlines():
      line_num += 1
      if line.find(search_phrase) >= 0:
          j=line_num
          break
  file.close()
  s=''
  f = open(filename)
  all_lines_variable = f.readlines()
  s=s+all_lines_variable[j-1]
  s=s.split('India')[0]
  s=s+'India'
  worksheet.merge_range('E6:I10',s,merge_format_1)
  s=''
  for x in range(0, 4):
      s=s+all_lines_variable[j-2][x]
  if s=='Page':
      s=all_lines_variable[j-3] 
  else:
      s=all_lines_variable[j-2]
  file.close()
  pattern = "\d{1,2}[/.]\d{1,2}[/.]\d{4}"
  file = open(filename)
  content = file.read()
  dates = re.findall(pattern,content )
  is_date=bool(dates)
  if is_date==True:
      s=s.replace(dates[0],'')
      worksheet.merge_range('E5:I5', s,merge_format_1)
  else:
      worksheet.merge_range('E5:I5', s ,merge_format_1)
  file.close()

  #Line of last occurance of India and above line above it 
  f = open(filename,'r')
  line_num = 0
  j=0
  i=0
  k=0
  search_phrase = "India"
  for line in reversed(f.readlines()):
      line_num += 1
      if line.find(search_phrase) >= 0:
          j=line_num
          break
  file.close()
  s=''
  with open(filename) as f:
      for i, l in enumerate(f):
          pass
  k=(i+1-j)
  f.close()
  f = open(filename)
  all_lines_variable = f.readlines()
  if all_lines_variable[k-1] == ' ':
      s=s+all_lines_variable[k-1]
  else:
      s=s+all_lines_variable[k-2]
  s=s+all_lines_variable[k]
  worksheet.merge_range('M14:Q16' , s ,merge_format_1 )
  f.close()
      
  #As the code is for India the currency will be INR only
  worksheet.merge_range('E13:I13','INR',merge_format_1 )

  #Finding the line of "Inv No" then next number after it is invoice number
  file = open(filename,'r')
  line_num = 0
  j=0
  search_phrase1 = "Inv"
  search_phrase2="No"
  for line in file.readlines():
      line_num += 1
      if line.find(search_phrase1) >= 0:
          if line.find(search_phrase2) >= 0:
              j=line_num
              break
  file.close()
  file = open(filename)
  all_lines_variable = file.readlines()
  if j>0:
      s=all_lines_variable[j-1]
      s=s.split('Inv')
      pattern = "[\s][0-9A-Z]{0,10}[/-]{0,1}[0-9A-Z]{0,10}[/-]{0,1}[0-9A-Z]{0,10}[/-]{0,1}[0-9]{1,10}[\s]"
      number = re.findall(pattern, s[1])
      is_number=bool(number)
      if is_number==True:
          worksheet.merge_range('M3:Q3',number[0],merge_format_1)
      else:
          s=all_lines_variable[j]
          pattern = "[\s][0-9A-Z]{0,10}[0-9]{1,10}[\s]"
          number = re.findall(pattern, s)
          is_number=bool(number)
          if is_number==True:
              worksheet.merge_range('M3:Q3',number[0],merge_format_1)
  f.close()

  #Finding the line number of Inv Date. The immediate next date format after the word "Inv Date" will give the date 
  file = open(filename,'r')
  line_num = 0
  j=0
  specific_line=''
  search_phrase1='Inv'
  search_phrase2="Date"
  for line in file.readlines():
      line_num += 1
      if line.find(search_phrase1) >= 0:
          if line.find(search_phrase2) >= 0:
              j=line_num
              break
  file.close()
  file = open(filename)
  all_lines_variable = file.readlines()
  if j>0:
      specific_line=all_lines_variable[j-1]
      specific_line=specific_line.split('Inv')
      pattern = "\d{1,2}[/.-]\d{1,2}[/.-]\d{4}"
      invoice_dates = re.findall(pattern, specific_line[1])
      is_invoice_dates=bool(invoice_dates)
      if is_invoice_dates==True:
          worksheet.merge_range('M4:Q4', invoice_dates[0],merge_format_1)
      else:
          specific_line=all_lines_variable[j]
          pattern = "\d{1,2}[/.-]\d{1,2}[/.-]\d{4}"
          invoice_dates = re.findall(pattern, specific_line)
          if bool(invoice_dates)==True:
              worksheet.merge_range('M4:Q4', invoice_dates[0],merge_format_1)
  file.close()

  #Searching PO number . 
  file = open(filename, "r")
  l=0
  content = file.read()
  pattern = "[\s]PO[\d+]{1,100}[\s]"; 
  PO = re.findall(pattern, content)
  l= len(PO)
  if len(PO)>0:
      worksheet.merge_range('M10:Q10', PO[0],merge_format_1)
  file.close()

  #Finding line number of due date . The immediate next date after it will be due date 
  file = open(filename,'r')
  line_num = 0
  j=0
  specific_line=''
  search_phrase1 = "Due"
  search_phrase2="Date"
  for line in file.readlines():
      line_num += 1
      if line.find(search_phrase1) >= 0:
          if line.find(search_phrase2) >= 0:
              j=line_num
              break
  file.close()
  file = open(filename)
  if j>0:
      all_lines_variable = file.readlines()
      specific_line=all_lines_variable[j-1]
      specific_line=specific_line.split('Due')
      pattern = "\d{1,2}[/.]\d{1,2}[/.]\d{4}"
      dates = re.findall(pattern, specific_line[1])
      is_dates=bool(dates)
      if is_dates==True:
          worksheet.merge_range('M5:Q5', dates[0],merge_format_1)
      else:
          specific_line=all_lines_variable[j]
          pattern = "\d{1,2}[/.]\d{1,2}[/.]\d{4}"
          dates = re.findall(pattern, specific_line)
          is_dates=bool(dates)
          if is_dates==True:
              worksheet.merge_range('M5:Q5', dates[0],merge_format_1)
  file.close()

  file = open(filename,'r')
  line_num = 0
  j=0
  specific_line=''
  search_phrase1 = "Days"
  for line in file.readlines():
          words = line.lower().split()
          if 'days' in words[1:]:
              due_days= words[words.index('days')-1]
              j=j+1
              break
  if j>0:
      if is_invoice_dates==True:
          startdate=invoice_dates[0]
          startdate=startdate.replace(".", "/")
          startdate=startdate.replace("-", "/")
          startdate=datetime.datetime.strptime(startdate, "%d/%m/%Y").strftime("%m-%d-%Y")
          due_days=int(due_days)
          enddate = pd.to_datetime(startdate) + pd.DateOffset(days=due_days)
          enddate=enddate.strftime('%d-%m-%Y')
          worksheet.merge_range('M5:Q5', enddate,merge_format_1)
  file.close()

  # Finding the line where first price is written . Then finding the line where line is starting from total . Lines between them starting with a numbr give S_No
  file = open(filename,'r')
  line_num = 0
  j=0
  for line in file.readlines():
      line_num += 1
      pattern = "[\s][\d*]{0,5}[,]{0,1}[\d*]{0,5}[,]{0,1}[\d*]{0,5}[,]{0,1}[\d*]{0,5}[,]{0,1}[\d+]{1,9}[.][0-9]{2,3}[\s]";
      price=re.findall(pattern , line)
      is_price=bool(price)
      if is_price==True:
          j=line_num
          break
  file.close()
  file = open(filename,'r')
  line_num = 0
  k=0
  for line in file.readlines():
      line_num+=1
      is_empty=line.isspace()
      if is_empty!=True:
          first_word = line.split()[0]
          if len(line.split())>1:
              second_word=line.split()[1]
              if second_word=='Total':
                  k=line_num
                  break
          if first_word=='Total' :
              k=line_num
              break
  file.close()

  l=0
  list5=[]
  S_No=[]
  file = open(filename,'r')
  line=file.readlines()
  for i in range(j-1,k):
      pattern = "^[\d]{1,15}[\s]";
      S_No=re.findall(pattern , line[i])
      is_S_No=bool(S_No)
      if is_S_No==True:
          list5.append(i)
          l=l+1
  file.close()
  z=0
  A=1
  S_No=''
  for i in range(0,l):
      worksheet.write('A'+str(i+18) , A,merge_format_1 )
      A=A+1
  worksheet.merge_range('B'+str(i+18+1)+':'+'E'+str(i+18+1),'Line Total',merge_format_5)
  worksheet.write('A'+str(i+18+1),'',merge_format_1)

  #Finding the string in the line and next lines till line is strating from a number
  qty=0
  C=0
  Net_Amount=0.0
  Net_Unit_price=0.0
  Net_WH_Amount=0.0
  Total_TCS=0.0
  for b in list5:
          pattern = "[a-zA-Z]{1,30}[\s][a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}";
          Name1=[]
          Title=''
          file = open(filename,'r')
          line=file.readlines()
          Name=re.findall(pattern,line[b])
          is_Name=bool(Name)
          if is_Name==True:
              Title=Title+Name[0]
          Range1=list5[C]
          for i in range(Range1-1,Range1+3):
              if line[i][0].isdigit()!=True:
                  Name1=re.findall(pattern,line[i])
                  is_Name1=bool(Name1)
                  if is_Name1==True:
                      Title=Title+Name1[0]
          worksheet.write('E'+str(C+18),Title,merge_format_1)
          file.close()
          
          # In price table first is either S.No and Product ID before description . 
          # After description first is SKU and then HSN . We find the occurnce of them by searching the word name in index . 
          file = open(filename,'r')
          pattern = "[a-zA-Z]{1,30}[\s][a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}[a-zA-Z]{0,30}[\s]{0,1}";        Name1=[]
          file = open(filename,'r')
          line=file.readlines()
          SKU=[]
          HSN=[]
          Name1=re.findall(pattern,line[b])
          s=line[b]
          file.close()
          if bool(Name1)==True:
              s=s.split(Name1[0])
          j=0
          file = open(filename,'r')
          specific_line=''
          search_phrase1 = "SKU"
          for line1 in file.readlines():
              if line1.find(search_phrase1) >= 0:
                  j=j+1
                  break
          if j>0:
              pattern="[A-Z]{0,5}[\d]{1,15}[A-Z]{0,5}";
              SKU=re.findall(pattern , s[1])
              is_SKU=bool(SKU)
              if is_SKU==True:
                  worksheet.write('C'+str(18+C),SKU[0],merge_format_1)
                  s[1]=s[1].replace(SKU[0],'')
          file.close()
          j=0
          file = open(filename,'r')
          search_phrase1 = " HS"
          for line1 in file.readlines():
              if line1.find(search_phrase1) >= 0:
                  j=j+1
                  break
          if j>0:
              pattern="[\d]{4,16}";
              HSN=re.findall(pattern , s[1])
              is_HSN=bool(HSN)
              if is_HSN==True:
                  worksheet.write('D'+str(18+C),HSN[0],merge_format_1)
                  s[1]=s[1].replace(HSN[0],'')
          s[0]=s[0].lstrip(str(C+1))
          worksheet.write('B'+str(18+C),s[0],merge_format_1)
          file.close()

  #Extracting all prices in a table starting from the S_No =1 till last 
  # quantity * unit price * (1-discount%) = Taxable Amount       
  # (quantity * unit price) - discount = Taxable Amount         
  # quantity * unit price = Taxable Amount        
  # Checking all the numbers which satisfy this relation . The relation ehich gives least taxable amount is the ans because it will consists discount .        
  # The remaining number less than 29(Max GST rate is 28%) is CGST, SGST or IGST depending on transportation is interstate or not
  #Cess percent is applicable in time of VAT so it is 0.0
  #App% = IGST+CGST+SGST-discount%
  #Total Amount = Taxable amount * (1+(CGST+SGST)/IGST)
  # if the data in invoice is wrong everything will be printed 0 
          app=0.0
          pattern1 = "[\d*]{0,5}[,]{0,1}[.]{0,1}[\d*]{0,5}[,]{0,1}[.]{0,1}[\d*]{0,5}[,]{0,1}[.]{0,1}[\d*]{0,5}[,]{0,1}[.]{0,1}[\d+]{1,9}[.]{0,1}[0-9]{0,3}";                
          everything = re.findall(pattern1, s[1])
          for i in range (0,len(everything)):
              point=0
              length=len(everything[i])
              for j in range(0,length):
                  if everything[i][j]=='.':
                      point = point+1
              if point > 1: 
                  everything[i]=everything[i].split('.')
                  correct_number=''
                  for j in range(0,len(everything[i])-1):
                      correct_number=correct_number+everything[i][j]
                  correct_number=correct_number+'.'+everything[i][len(everything[i])-1]
                  everything[i]=correct_number
          for i in range(0,len(everything)):
              everything[i]=everything[i].translate({ord(','): None})
          for i in range(0,len(everything)):
              everything[i]=everything[i].translate({ord('\n'): None})
          for i in range(0,len(everything)):
              everything[i]=float(everything[i])
          def remove_values_from_list(the_list, val):
              return [value for value in the_list if value != val]
          everything = remove_values_from_list(everything, 0.0)
          IGST=0.0
          CGST=0.0
          SGST=0.0
          for i in range(0,len(everything)):
              if inter_state==0:
                  if everything[i]==5.0 or everything[i]==12.0 or everything[i]==18.0 or everything[i]==28.0:
                      IGST=everything[i]
              else:
                  if everything[i]==5.0 or everything[i]==12.0 or everything[i]==18.0 or everything[i]==28.0:
                      CGST=everything[i]/2
                      SGST=everything[i]/2
                  if everything[i]==2.5 or everything[i]==6.0 or everything[i]==9.0 or everything[i]==14.0: 
                      CGST=everything[i]
                      SGST=everything[i]
          everything = [x for x in everything if x!=CGST]
          everything = [x for x in everything if x!=SGST]
          everything = [x for x in everything if x!=IGST]
          if CGST>0.0 or IGST>0.0:
              worksheet.write('J'+str(18+C),CGST,merge_format_1)
              worksheet.write('K'+str(18+C),SGST,merge_format_1)
              worksheet.write('L'+str(18+C),IGST,merge_format_1)
          else:
              worksheet.write('J'+str(18+C),'Detected Wrong or wrong GST written',merge_format_1)
              worksheet.write('K'+str(18+C),'Detected Wrong or wrong GST written',merge_format_1)
              worksheet.write('L'+str(18+C),'Detected Wrong or wrong GST written',merge_format_1)
          app=app+CGST+IGST+SGST
          a=0 
          b=0
          c=0
          d=0
          list1=[]
          list2=[]
          list3=[]
          def removing_1_function ():
              global list1
              global list2
              global list3
              for i in range(0,len(everything)):
                  a=everything[i]
                  for m in range(0,len(everything)):
                      b=everything[m]
                      for k in range(0,len(everything)):
                          c=everything[k]
                          if a==c:
                              continue
                          if b==c:
                              continue
                          for n in range(0,len(everything)):
                              if n==k or n==m or n==i:
                                  continue
                              else:
                                  d=everything[n]
                                  form_a=a*b*(1-c)
                                  form_b=(a*b)-c
                                  form_c=a*b
                                  a1=max(a,b)
                                  b1=min(a,b)
                                  if CGST!=0.0:
                                      for y in range(0,len(everything)):
                                          if a1/CGST==y:
                                              a1=y
                                              break
                                  if form_a == d:
                                      list1=[a1,b1,c,d]
                                      if list1[0]>1 and list1[1]>1:
                                          return 
                                  if form_b==d :
                                      list2=[a1,b1,c,d]
                                      if list2[0]>1 and list2[1]>1:
                                          return
                                  if form_c==d:
                                      list3=[a1,b1,d]
                                      if list3[0]>1.0 and list3[1]>1.0:
                                          return
                    
          removing_1_function ()
          x=float('inf')
          y=float('inf')
          z=float('inf')
          taxable_amount=0.0
          discount_percent=0.0
          is_list1=bool(list1)
          if is_list1==True:
              x=list1[3]
          is_list2=bool(list2)
          if is_list2==True:
              y=list2[3]
          is_list3=bool(list3)
          if is_list3==True:
              z=list3[2]
          
          if (is_list1==True or is_list2==True or is_list3==True)  :
              if(x <= y and x <= z): 
                  worksheet.write('G'+str(18+C),list1[0],merge_format_1)
                  worksheet.write('F'+str(18+C),list1[1],merge_format_1)
                  worksheet.write('I'+str(18+C),list1[2],merge_format_1)
                  app=app-list1[2]
                  Net_Unit_price=Net_Unit_price+list1[0]
                  taxable_amount=list1[3]
                  Net_WH_Amount=Net_WH_Amount+taxable_amount
              elif(y <= x and y <= z): 
                  worksheet.write('G'+str(18+C),list2[0],merge_format_1)
                  worksheet.write('F'+str(18+C),list2[1],merge_format_1)
                  worksheet.write('I'+str(18+C),list2[2],merge_format_1)
                  taxable_amount=list2[3]
                  Net_Unit_price=Net_Unit_price+list2[0]
                  discount_percent=list2[2]/(list2[0]*list2[1])
                  app=app-discount_percent
                  Net_WH_Amount=Net_WH_Amount+taxable_amount
              else:
                  worksheet.write('G'+str(18+C),list3[0],merge_format_1)
                  worksheet.write('F'+str(18+C),list3[1],merge_format_1)
                  worksheet.write('I'+str(18+C),'0.0',merge_format_1)
                  taxable_amount=list3[2]
                  app=app-discount_percent
                  qty=qty+list3[1]
                  Net_Unit_price=Net_Unit_price+list3[0]
                  worksheet.write('I'+str(18+C),'0.0',merge_format_1)
                  Net_WH_Amount=Net_WH_Amount+taxable_amount
          worksheet.write('M'+str(18+C),'0.0',merge_format_1)
          
          #Finding the line number of excise duty . The next number after it gives excise duty percent 
          file = open(filename,'r')
          line_num = 0
          j=0
          specific_line=''
          search_phrase1 = "Excise"
          search_phrase2="Duty"
          for line in file.readlines():
              line_num += 1
              if line.find(search_phrase1) >= 0:
                  if line.find(search_phrase2) >= 0:
                      j=line_num
                      break
          file.close()
          file = open(filename)
          all_lines_variable = file.readlines()
          if j>0:
              specific_line=all_lines_variable[j-1]
              specific_line=specific_line.split('Duty')
              pattern = "\d{1,15}[.]{0,1}\d{0,3}"
              excise_duty = re.findall(pattern, specific_line[1])
              worksheet.write('H'+str(18+C),excise_duty[0],merge_format_1)
          else:
              worksheet.write('H'+str(18+C),'0.0',merge_format_1)
          file.close()

          #Finding the line number of TCS . The next number after it gives TCS percent .
          file = open(filename,'r')
          line_num = 0
          j=0
          specific_line=''
          search_phrase1 = "TCS"
          for line in file.readlines():
              line_num=line_num+1
              if line.find(search_phrase1) >= 0:
                  j=line_num
                  break
          file.close()
          file = open(filename)
          all_lines_variable = file.readlines()
          if j>0:
              specific_line=all_lines_variable[j-1]
              specific_line=specific_line.split('TCS')
              pattern = "\d{1,15}[.]{0,1}\d{0,3}"
              TCS = re.findall(pattern, specific_line[1])
              worksheet.write('N'+str(18+C),TCS[0],merge_format_1)
              Total_TCS=Total_TCS+(TCS[0]*taxable_amount)
          else:
              worksheet.write('N'+str(18+C),'0.0',merge_format_1)
          file.close()
          
          Total_amount=taxable_amount - discount_percent + (taxable_amount*app)/100
          worksheet.write('O'+str(C+18) , Total_amount ,merge_format_1 )
          Net_Amount=Net_Amount+Total_amount
          C=C+1
  worksheet.merge_range('M7:Q7' , qty ,merge_format_1 )
  worksheet.merge_range('M12:Q12' , qty ,merge_format_1 )
  worksheet.write('F'+str(18+C) , qty ,merge_format_5 )
  worksheet.write('G'+str(18+C),Net_Unit_price,merge_format_5)
  worksheet.merge_range('M11:Q11' , Net_Amount ,merge_format_1 )
  worksheet.write('O'+str(C+18) , Net_Amount ,merge_format_5 )
  worksheet.merge_range('M6:Q6' , Net_WH_Amount ,merge_format_1 )
  worksheet.merge_range('M8:Q8',Total_TCS,merge_format_1)
  fractional,whole=math.modf(Net_Amount)
  if fractional >= .50:
      whole=whole+1
      result = re.match(r"whole", filename)
      if result is not None :
          worksheet.merge_range('M9:Q9' , whole ,merge_format_1 )
  else:
      result = re.match(r"whole", filename)
      if result is not None :
          worksheet.merge_range('M9:Q9' , whole ,merge_format_1 )
  workbook.close()       
  os.remove(outfile)

# convertPDF("Sample_Invoice_2.pdf")