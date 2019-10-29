import math
import xlrd
import xlwt
import smtplib
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# set email parameters
server = smtplib.SMTP('documentdirection-ca.mail.protection.outlook.com', 25)
fromaddr = "chris@documentdirection.ca"
toaddr = "chris@documentdirection.ca"
msg = MIMEMultipart()
msg['From'] = "chris@documentdirection.ca"
msg['To'] = "chris@documentdirection.ca"
msg['Subject'] = "Expiring Software Licenses"
body = ""

#open LSAP Rewnewal workbook
loc = ("lsapRenewal.xlsx")
wb = xlrd.open_workbook(loc)
numberOfSheets = len(wb.sheet_names())
sheet = wb.sheet_by_index(0)
thisSheet = 0
endOfList = 100
today = int(sheet.cell_value(0, 0))

for y in range(0, numberOfSheets):
  for x in range(1, endOfList):
    try:
      companyName = sheet.cell_value(x, 1)
    except: 
      break
    # if companyName == "":
    #   print("done")
    #   break
    try:
      NumberOfLicenses = sheet.cell_value(x, 2)
      lsapExp = sheet.cell_value(x, 3)
    except:
      break

    try:
      daysUntilExpired = lsapExp - today
    except:
      continue

    if daysUntilExpired < 0:
      print(companyName +  "'s contract expired " + str((int(daysUntilExpired) *-1)) + " days ago")
      body = body + "<br>" + companyName +  "'s contract expired " + str((int(daysUntilExpired) *-1)) + " days ago <br>"
    elif daysUntilExpired > 0 and daysUntilExpired <= 90:
      print(companyName +  "'s contract will expire in " + str(int(daysUntilExpired)) + " days")
      body = body + "<br>" + companyName +  "'s contract will expire in " + str(int(daysUntilExpired)) + " days <br>"
  
  thisSheet = thisSheet + 1
  print("Sheet: " + str(thisSheet))
  try:
    sheet = wb.sheet_by_index(thisSheet)
  except:
    break
  

# send mail
msg.attach(MIMEText(body, 'html'))
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
print("mail Sent...")