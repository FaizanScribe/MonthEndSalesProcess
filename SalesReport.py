import xlwings as xw
# GENERATES SALES REPORT
fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\2. Feb24  - Copy.xlsm"

wb = xw.Book(fileName)
ws1 = wb.sheets[0]
ws2 = wb.sheets('SheetY')
noOFSheets = int(ws1.range("AR1").value)
print(noOFSheets)
Vrx = ()
Fit = ()
Bl = ()
Noor = ()
Gt = ()

for i in range(6, noOFSheets-13):
    wsTemp = wb.sheets[i]
    Vrx += (wsTemp.range("D6").value,)
    Fit += (wsTemp.range("D7").value,)
    Bl += (wsTemp.range("D9").value,)
    Noor += (wsTemp.range("D9").value,)
    Gt += (wsTemp.range("D20").value,)

ws2['N3'].options(transpose=True).value = Vrx
ws2['O3'].options(transpose=True).value = Fit
ws2['P3'].options(transpose=True).value = Bl
ws2['Q3'].options(transpose=True).value = Noor
ws2['R3'].options(transpose=True).value = Gt