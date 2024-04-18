import xlwings as xw
import pandas as pd

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\April24.xlsm"
wb = xw.Book(fileName)
pslipSheet = wb.sheets("SheetX")

df = pd.read_excel(fileName, sheet_name="SheetX", usecols='AP:AW')

table = pd.pivot_table(df, values=['QTY', 'AMOUNT'], index=['TALLY PARTY', 'TALLY PRODUCT'], aggfunc="sum")
pslipSheet.range("AY1").value = table
