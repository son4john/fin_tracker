import locale
import xlsxwriter
import fileIO
from myDateTime import currentDate
from myDateTime import RETIREDATE
import sheets

# Changes to the order of the google sheets file will affect the code

data = sheets.Sheet1()

cash = float(data[0][1]) + float(data[1][1]) + float(data[5][1])
debt = float(data[2][1]) + float(data[6][1]) + float(data[7][1])  + float(data[8][1]) + float(data[9][1])
invst = float(data[3][1]) + float(data[4][1])


#Create Output File
workbook = xlsxwriter.Workbook("FREPORT_" + currentDate() + ".xlsx")
worksheet = workbook.add_worksheet('Summary')

#Formats
dollar_format = workbook.add_format({'num_format': '$#,##0.00'})
title_format = workbook.add_format({'bold': True, 'font_size': 18})
subtitle_format = workbook.add_format({'italic': True, 'font_size': 13})
worksheet.set_column(0,0,30)
worksheet.set_column(1,1,20)

#Titles
worksheet.write('A1', 'FINANCIAL REPORT', title_format)
worksheet.write('A4', 'CASH', subtitle_format)
worksheet.write('B4', cash, dollar_format)
worksheet.write('A5', 'DEBT', subtitle_format)
worksheet.write('B5', debt, dollar_format)
worksheet.write('A6', 'INVESTMENTS', subtitle_format)
worksheet.write('B6', invst, dollar_format)




workbook.close()


