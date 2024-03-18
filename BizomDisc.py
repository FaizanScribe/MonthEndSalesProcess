import xlwings as xw
import pandas as pd
import LastEmptyRow


# bizom discounts. How to go about?
# get the list of dealers in bizom
# make a tuple out of that
# now let's begin

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\March 24.xlsm"
wb = xw.Book(fileName)
SheetX = wb.sheets('SheetX')
customers = ()
# products = ()
pricediff = 0
pricediffT = ()

LRCustomer = LastEmptyRow.Find_Last_Row(fileName, SheetName=SheetX, ColumnLetter='AH')

for i in range(4, LRCustomer):
    # if SheetX.range('AI' + str(i)).value == 'y':
    customers += (SheetX.range('AH' + str(i)).value, )

LRProd = LastEmptyRow.Find_Last_Row(fileName, SheetName=SheetX, ColumnLetter='AK')

# for i in range(4, LRProd):
#     products += (SheetX.range('AK' + str(i)).value,)

df = pd.read_excel(fileName, sheet_name='PSLIP', usecols='E:K')

for cust in customers:
    for i in range(4, LRProd):
        p = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetX.range('AK' + str(i)).value), 'RE QTY'].sum()
        pVal = p * SheetX.range('AN' + str(i)).value
        pricediff += pVal

        p2 = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == SheetX.range('AK' + str(i)).value), 'LE QTY'].sum()
        pVal2 = p2 * SheetX.range('AN' + str(i)).value
        pricediff += pVal2
    pricediffT += (pricediff,)
    pricediff = 0

SheetX['AG4'].options(transpose=True).value = pricediffT
