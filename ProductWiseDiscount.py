# trials
# this is productwise discount of month end data.
# here the items are U 1.50, 1.56, 1.67 AND NDEL 1.56 BM S+

import pandas as pd
import xlwings as xw
import openpyxl
import pyarrow
import LastEmptyRow
fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\2. Feb24  - Copy.xlsm"
wb = xw.Book(fileName)
sheetY = wb.sheets("SheetY")
PSLIP = wb.sheets("PSLIP")
df = pd.read_excel(fileName, sheet_name="PSLIP", usecols= 'E:K')

Customer = ()
u15Disc = ()
u150Amount = ()

u156Disc = ()
u156Amount = ()

u167Disc = ()
u167Amount = ()

nDelDisc = ()
nDelAmount = ()

RowNo = LastEmptyRow.Find_Last_Row(fileName, "SheetY", "D")

# loop will run from D3 to D RowNo

for i in range(3,RowNo):

    if sheetY.range("D" + str(i)).value == "y":
        Customer += (sheetY.range("A" + str(i)).value,)


# print(Customer)

sheetY['E3'].options(transpose=True).value = Customer

# qualified customers list is ready

# find last row of PSLIP

LRPslip = LastEmptyRow.Find_Last_Row(fileName, "PSLIP", "E")
# print(LRPslip)
# loop will run from 2 to LRPslip

TotalCustomer = len(Customer)

print(TotalCustomer)
for cust in Customer:
    u15QtyR = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('F2').value), 'RE QTY'].sum()
    u150RAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('F2').value), 'RE PRICE'].sum()
    u15QtyL = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('F2').value), 'LE QTY'].sum()
    u150LAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('F2').value), 'LE PRICE'].sum()
    u15Disc += (sheetY.range("F1").value * (u15QtyR + u15QtyL),)
    u150Amount += ((u150RAmount + u150LAmount),)

    u156QtyR = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('G2').value), 'RE QTY'].sum()
    u156RAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('G2').value), 'RE PRICE'].sum()
    u156QtyL = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('G2').value), 'LE QTY'].sum()
    u156LAmount = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('G2').value), 'LE PRICE'].sum()
    u156Disc += (sheetY.range("G1").value * (u156QtyR + u156QtyL),)
    u156Amount += ((u156RAmount + u156LAmount),)

    u167QtyR = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('H2').value), 'RE QTY'].sum()
    u167RAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('H2').value), 'RE PRICE'].sum()
    u167QtyL = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('H2').value), 'LE QTY'].sum()
    u167LAmount = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('H2').value), 'LE PRICE'].sum()
    u167Disc += (sheetY.range("H1").value * (u167QtyR + u167QtyL),)
    u167Amount += ((u167RAmount + u167LAmount),)

    nDelQtyR = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('I2').value), 'RE QTY'].sum()
    nDelRAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == sheetY.range('I2').value), 'RE PRICE'].sum()
    nDelQtyL = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('I2').value), 'LE QTY'].sum()
    nDelLAmount = df.loc[(df['CUSTOMER'] == cust) & (df['LE PRODUCT'] == sheetY.range('I2').value), 'LE PRICE'].sum()
    nDelDisc += (sheetY.range("I1").value * (nDelQtyR + nDelQtyL),)
    nDelAmount += ((nDelRAmount + nDelLAmount),)

sheetY['F3'].options(transpose=True).value = u15Disc
sheetY['G3'].options(transpose=True).value = u156Disc
sheetY['H3'].options(transpose=True).value = u167Disc
sheetY['I3'].options(transpose=True).value = nDelDisc
sheetY['J3'].options(transpose=True).value = tuple(map(lambda a, b, c, d: a+b+c+d, u150Amount, u156Amount, u167Amount,nDelAmount ))

print("THIS IS A TEST")
