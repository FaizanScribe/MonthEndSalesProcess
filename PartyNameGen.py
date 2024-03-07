import xlwings as xw

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\2. Feb24  - Copy.xlsm"

wb = xw.Book(fileName)
ws1 = wb.sheets[0]
ws2 = wb.sheets('SheetY')
noOFSheets = int(ws1.range("AR1").value)
print(noOFSheets)

DName = ()
SerialNO = ()

for i in range(6, noOFSheets-13):
    wsTemp = wb.sheets[i]
    DName += (wsTemp.name,)
    SerialNO += (i,)


ws2['M3'].options(transpose=True).value = DName

ws2['A3'].options(transpose=True).value = DName
ws2['B3'].options(transpose=True).value = SerialNO

# PARTY NAME GENERATOR FOR APPLYING DISCOUNTS AND SALES REPORTS


