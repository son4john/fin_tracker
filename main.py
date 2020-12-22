import locale
import xlsxwriter
from myDateTime import currentDate
from myDateTime import RETIREDATE

input_file = open("data.csv", "r")
workbook = xlsxwriter.Workbook("FREPORT_" + currentDate() + ".xlsx")

data = input_file.read()
accounts = []
investment = 0
debt = 0
income = 0
expense = 0

for rows in data.split("\n"):
    account_info = []
    for col in rows.split(","):
        account_info.append(col)
    accounts.append(account_info)

worksheet = workbook.add_worksheet('Summary')

dollar_format = workbook.add_format({'num_format': '$#,##0.00'})
title_format = workbook.add_format({'bold': True, 'font_size': 18})
subtitle_format = workbook.add_format({'italic': True, 'font_size': 13})
worksheet.set_column('B:B', 20)

#Titles
worksheet.write('A1', 'Financial Report', title_format)
worksheet.write('C4', 'Current', subtitle_format)
worksheet.write('D4', 'Change', subtitle_format)
worksheet.write('G5', 'Total Current Debt', subtitle_format)
worksheet.write('G6', 'Yearly Projected Income', subtitle_format)


worksheet.write('E2', 500000, dollar_format)
worksheet.write('D1', currentDate(), title_format)
worksheet.write('E1', RETIREDATE, title_format)




#Calculate Investments, Debt, Income, Expenses
for rows in accounts:
    if rows[0] == "Merril Edge Investments":
        investment += float(rows[1])
    if rows[0] == "TIAA 403K":
        investment += float(rows[1])
    if rows[0] == "Baylor 401K":
        investment += float(rows[1])
    if rows[0] == "Nokia 401K":
        investment += float(rows[1])
    if rows[0] == "Robinhood":
        investment += float(rows[1])
    if rows[0] == "Bestbuy Credit":
        debt += float(rows[1])
    if rows[0] == "Samsung Credit":
        debt += float(rows[1])
    if rows[0] == "Granbury Loan":
        debt += float(rows[1])
    if rows[0] == "Auburn Hills Loan":
        debt += float(rows[1])
    if rows[0] == "Bank of America Credit":
        debt += float(rows[1])
    if rows[0] == "Discover Credit":
        debt += float(rows[1])        
    if rows[0] == "Baylor Income":
        income += float(rows[1])
    if rows[0] == "VectorNav Income":
        income += float(rows[1])
    if rows[0] == "Granbury Income":
        income += float(rows[1])
income *= 12                
    
worksheet.write('D2', investment, dollar_format)
worksheet.write('H5', debt, dollar_format)
worksheet.write('H6', income, dollar_format)        

rowNumber = 4 
for rows in accounts:
    columnNumber = 1
    worksheet.write(rowNumber ,columnNumber, rows[0])
    try:
        worksheet.write(rowNumber, columnNumber + 1, float(rows[1]), dollar_format)
        
    except (IndexError, ValueError):
        worksheet.write(rowNumber, columnNumber + 1, -1)
    try:
        worksheet.write(rowNumber ,columnNumber + 2, float(rows[2]), dollar_format)
    except (IndexError, ValueError):
        worksheet.write(rowNumber ,columnNumber + 2,-1)
    rowNumber += 1



workbook.close()
input_file.close()


