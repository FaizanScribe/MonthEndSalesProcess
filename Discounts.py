# PLACE THE DISCOUNTS TO THEIR RESPECTIVE SHEETS
# FIRST THE DISCOUNT PERCENTAGES ARE PLACED IN THE SHEETS
# SECOND THE PRODUCTWISE DISCOUNTS ARE PLACED IN THE SHEETS

import xlwings as xw

import LastEmptyRow

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\2. Feb24 .xlsm"

wb = xw.Book(fileName)
ws1 = wb.sheets[0]
ws2 = wb.sheets('SheetY')
noOFSheets = int(ws1.range("AR1").value)
print(noOFSheets)
DPer = ()
ultra = ()

# APPLY DISCOUNTS AND PERCENTAGES IN ALL SHEETS FROM SHEET Y

for i in range(3, noOFSheets-13):
    r1 = "C" + str(i)
    DPer += (ws2.range(r1).value,)

for i in range(6, noOFSheets-13):
    wsTemp = wb.sheets[i]
    wsTemp.range("I26").value = DPer[i-6]

LRofSheetY = LastEmptyRow.Find_Last_Row(fileName, "SheetY", "E")

for i in range(3,LRofSheetY):
    SName = ws2.range("E" + str(i)).value
    TempWS = wb.sheets(SName)
    TempWS.range('I30').value = ws2.range("J" + str(i)).value

#this works all OK


