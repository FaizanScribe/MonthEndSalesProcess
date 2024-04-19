import xlwings as xw
import pandas as pd

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\April24.xlsm"
wb = xw.Book(fileName)
pslipSheet = wb.sheets("Pivots2")
df = pd.read_excel(fileName, sheet_name="PSLIP", usecols='AI:AP')
BLdf1 = pd.read_excel(fileName, sheet_name="BL FILTER", usecols='AJ:AO')
inv1list = ['VRX', 'C']
inv2list = ['B']
inv3list = ['A']
inv4list = ['S']
inv5list = ['F']

conditionVrxCH = df['BRAND'].isin(inv1list)
conditionBL = df['BRAND'].isin(inv2list)
conditionAbdullah = df['BRAND'].isin(inv3list)
conditionSylvet = df['BRAND'].isin(inv4list)
conditionFrame = df['BRAND'].isin(inv5list)

VrxCHdf = df[conditionVrxCH]
BLdf = df[conditionBL]
Abdullahdf = df[conditionAbdullah]
Sylvetdf = df[conditionSylvet]
Framedf = df[conditionFrame]

tableVC = pd.pivot_table(VrxCHdf, values=['QTY', 'AMOUNT'], index=['TALLY PARTY', 'TALLY PRODUCT'], aggfunc="sum")
pslipSheet.range("A1").value = tableVC

tableBL = pd.pivot_table(BLdf1, values=['QTY INPCS', 'PRICE'], index=['TALLY PARTY', 'TALLY PRODUCT'], aggfunc="sum")
pslipSheet.range("F1").value = tableBL

tableAbd = pd.pivot_table(Abdullahdf, values=['QTY', 'AMOUNT'], index=['TALLY PARTY', 'TALLY PRODUCT'], aggfunc="sum")
pslipSheet.range("K1").value = tableAbd

tableSyl = pd.pivot_table(Sylvetdf, values=['QTY', 'AMOUNT'], index=['TALLY PARTY', 'TALLY PRODUCT'], aggfunc="sum")
pslipSheet.range("N1").value = tableSyl

tableFrame = pd.pivot_table(Framedf, values=['QTY', 'AMOUNT'], index=['TALLY PARTY', 'TALLY PRODUCT'], aggfunc="sum")
pslipSheet.range("Q1").value = tableFrame

# THIS THING WORKS