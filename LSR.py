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
sheet = wb.sheet_by_index(0)

endOfList = 22
today = int(sheet.cell_value(0, 0))

for x in range(1, endOfList):
  companyName = sheet.cell_value(x, 1)
  NumberOfLicenses = sheet.cell_value(x, 2)
  lsapExp = sheet.cell_value(x, 3)
  
  daysUntilExpired = lsapExp - today

  if daysUntilExpired < 0:
    print(companyName +  "'s contract expired " + str((int(daysUntilExpired) *-1)) + " days ago")
    body = body + "<br>" + companyName +  "'s contract expired " + str((int(daysUntilExpired) *-1)) + " days ago <br>"
  elif daysUntilExpired > 0 and daysUntilExpired <= 90:
    print(companyName +  "'s contract will expire in " + str(int(daysUntilExpired)) + " days")
    body = body + "<br>" + companyName +  "'s contract will expire in " + str(int(daysUntilExpired)) + " days <br>"
  

# send mail
msg.attach(MIMEText(body, 'html'))
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)