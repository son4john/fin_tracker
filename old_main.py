import locale
import xlsxwriter
import fileIO
from myDateTime import currentDate
from myDateTime import RETIREDATE
import sheets

# Changes to the order of the google sheets file will affect the code

data = sheets.Sheet1()

#CASH
if (data[0][0] == "BOAS" and data[1][0] == "BOAD" and data[6][0]== "DSCD"):
    cash = float(data[0][1]) + float(data[1][1]) + float(data[6][1])
else:
    print("Error In Cash Setup")
    cash = 0.0
#DEBT
if (data[4][0] == "BOAC" and data[7][0] == "DSCC" and data[10][0] == "REMT" and data[11][0] == "WFMT" and data[8][0] == "BBYC"):
    debt = float(data[4][1]) + float(data[7][1]) + float(data[10][1])  + float(data[11][1]) + float(data[8][1])
else:
    print("Error In Debt Setup")
    debt = 0.0
#INVESTMENTS
if (data[3][0] == "EDGE" and data[5][0] == "RHOD" and data[12][0] == "CBCT" and data[13][0] == "UPCT" and data[14][0] == "NOKA" and data[15][0] == "BAYL" and data[16][0] == "TIAA"):
    invst = float(data[3][1]) + float(data[5][1]) + float(data[12][1]) + float(data[13][1]) + float(data[14][1]) + float(data[15][1]) + float(data[16][1])
else:
    print("Error In Investment Setup")
    invst = 0.0
#SPENDING
RentPayment = 1700.0
montlhlySpending = float(data[2][1])

#Create Output File
workbook = xlsxwriter.Workbook("FREPORT_" + currentDate() + ".xlsx")
worksheet = workbook.add_worksheet('Summary')

#Formats
dollar_format = workbook.add_format({'num_format': '$#,##0.00'})
title_format = workbook.add_format({'bold': True, 'font_size': 18})
subtitle_format = workbook.add_format({'italic': True, 'font_size': 13})
note_format = workbook.add_format({'font_size': 12})
percent_format = workbook.add_format({'num_format': '0.00%'})
worksheet.set_column(0,0,30)
worksheet.set_column(1,1,20)
worksheet.set_column(3,3,30)
worksheet.set_column(4,4,20)

#Titles
worksheet.write('A1', 'Overview', title_format)
worksheet.write('A3', 'CASH', subtitle_format)
worksheet.write('B3', cash, dollar_format)
worksheet.write('A4', 'DEBT', subtitle_format)
worksheet.write('B4', debt, dollar_format)
worksheet.write('A5', 'INVESTMENTS', subtitle_format)
worksheet.write('B5', invst, dollar_format)
worksheet.write('D5', 'GOAL', subtitle_format)
worksheet.write('E5', invst - float(data[18][1]), dollar_format)

worksheet.write('A7', 'Other', title_format)
worksheet.write('A9', 'Family Spending', subtitle_format)
worksheet.write('B9', montlhlySpending + 8000, dollar_format)
worksheet.write('C9', 'Note: A 8000 dollar budget a negative balance means i went over.')
worksheet.write('A10', 'Tenant Balance', subtitle_format)
worksheet.write('B10', RentPayment, dollar_format)
worksheet.write('C10', 'Note: The Amount of Money owed for the rental home.')

worksheet.write('A12', 'Debt', title_format)
worksheet.write('A14', 'Net Debt', subtitle_format)
worksheet.write('B14', debt + cash, dollar_format)
worksheet.write('C14', 'Note: Cash plus DEBT')

worksheet.write('A15', 'Total Value', subtitle_format)
worksheet.write('B15', debt + cash + invst, dollar_format)
worksheet.write('C15', 'Note: Cash plus Debt plus Invst')

worksheet.write('A16', 'Debt to Asset', subtitle_format)
worksheet.write('B16', (debt*-1) / (cash + invst),percent_format)
worksheet.write('C16', 'Note: Debt to cash plus assets ratio')



workbook.close()


