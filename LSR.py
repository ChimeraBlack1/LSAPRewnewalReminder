import math
import xlrd
import xlwt
from datetime import date

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
  elif daysUntilExpired > 0 and daysUntilExpired <= 90:
    print(companyName +  "'s contract will expire in " + str(int(daysUntilExpired)) + " days")
  # else:
  #   print("not expiring in the next 90 days")

# print(str(companyName))
# print(str(NumberOfLicenses))
# print(int(lsapExp))