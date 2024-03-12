# crizal new products @ 20% discount
# crizal old products @ 25% discount

import xlwings as xw
import pandas as pd
import LastEmptyRow
import AllCustTuple

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\March 24.xlsm"
wb = xw.Book(fileName)
PSLIP = wb.sheets('PSLIP')
SheetY = wb.sheets('SheetY')

# create a list of discount applicable products : done
# put that list in a tuple: done
# make a df of PSLIP data region: done
# find the last row of PSLIP: done
# loop to the last row for Re product and Le product
# check if RE PRODUCT or LE PRODUCT " is in" the applicable product TUPLE
# if yes, then check if the RE Product and LE product contain "23"
# if yes, the disc % is 20 else disc % is 25
# apply discount calculation on Disc = (Re_Price/1.12)*Disc
# place this for ALL Dealers

# NewDiscProductsCrizal = ()
# OldDiscProductsCrizal = ()

CrizalAmount = ()
CrizalDiscount = ()
CrizalAmt = 0.00
CrizalDisc = 0.00

Customers = AllCustTuple.CutomersList(fileName)
SheetY['V3'].options(transpose=True).value = Customers

# LRCrizalNew = LastEmptyRow.Find_Last_Row(fileName, 'SheetY', 'Z')
# LRCrizalOld = LastEmptyRow.Find_Last_Row(fileName, 'SheetY', 'AA')

# for i in range(3, LRCrizalNew):
#     NewDiscProductsCrizal += (SheetY.range("T" + str(i)).value,)
#
# for i in range(3, LRCrizalOld):
#     OldDiscProductsCrizal += (SheetY.range("T" + str(i)).value,)


df = pd.read_excel(fileName, sheet_name="PSLIP", usecols= 'E:G,I,K')


for cust in Customers:

    BCT156CrzRockREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T3').value), 'RE PRICE'].sum()
    BCT156CrzRockLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T3').value), 'LE PRICE'].sum()
    CrizalAmt += (BCT156CrzRockREAmount + BCT156CrzRockLEAmount)
    CrizalDisc += ((BCT156CrzRockREAmount + BCT156CrzRockLEAmount)/1.12)*SheetY.range('U3').value

    BCT16CrzRockREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T4').value), 'RE PRICE'].sum()
    BCT16CrzRockLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T4').value), 'LE PRICE'].sum()
    CrizalAmt += (BCT16CrzRockREAmount + BCT16CrzRockLEAmount)
    CrizalDisc += ((BCT16CrzRockREAmount + BCT16CrzRockLEAmount)/1.12)*SheetY.range('U4').value


    BCT156CrzEProREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T5').value), 'RE PRICE'].sum()
    BCT156CrzEProLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T5').value), 'LE PRICE'].sum()
    CrizalAmt += (BCT156CrzEProREAmount + BCT156CrzEProLEAmount)
    CrizalDisc += ((BCT156CrzEProREAmount + BCT156CrzEProLEAmount)/1.12)*SheetY.range('U5').value

    BCT15CrzEProREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T6').value), 'RE PRICE'].sum()
    BCT15CrzEProLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T6').value), 'LE PRICE'].sum()
    CrizalAmt += (BCT15CrzEProREAmount + BCT15CrzEProLEAmount)
    CrizalDisc += ((BCT15CrzEProREAmount + BCT15CrzEProLEAmount)/1.12)*SheetY.range('U6').value

    BCT160CrzEProREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T7').value), 'RE PRICE'].sum()
    BCT160CrzEProLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T7').value), 'LE PRICE'].sum()
    CrizalAmt += (BCT160CrzEProREAmount + BCT160CrzEProLEAmount)
    CrizalDisc += ((BCT160CrzEProREAmount + BCT160CrzEProLEAmount)/1.12)*SheetY.range('U7').value

    BCT160CrzRockREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T8').value), 'RE PRICE'].sum()
    BCT160CrzRockLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T8').value), 'LE PRICE'].sum()
    CrizalAmt += (BCT160CrzRockREAmount + BCT160CrzRockLEAmount)
    CrizalDisc += ((BCT160CrzRockREAmount + BCT160CrzRockLEAmount)/1.12)*SheetY.range('U8').value

    AWBCTRockREAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T9').value), 'RE PRICE'].sum()
    AWBCTRockLEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T9').value), 'LE PRICE'].sum()
    CrizalAmt += (AWBCTRockREAmount + AWBCTRockLEAmount)
    CrizalDisc += ((AWBCTRockREAmount + AWBCTRockLEAmount)/1.12)*SheetY.range('U9').value

    Prod1REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T10').value), 'RE PRICE'].sum()
    Prod1LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T10').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod1REAmount + Prod1LEAmount)
    CrizalDisc += ((Prod1REAmount + Prod1LEAmount)/1.12)*SheetY.range('U10').value

    Prod2REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T11').value), 'RE PRICE'].sum()
    Prod2LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T11').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod2REAmount + Prod2LEAmount)
    CrizalDisc += ((Prod2REAmount + Prod2LEAmount)/1.12)*SheetY.range('U11').value

    Prod3REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T12').value), 'RE PRICE'].sum()
    Prod3LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T12').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod3REAmount + Prod3LEAmount)
    CrizalDisc += ((Prod3REAmount + Prod3LEAmount)/1.12)*SheetY.range('U12').value

    Prod4REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T13').value), 'RE PRICE'].sum()
    Prod4LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T13').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod4REAmount + Prod4LEAmount)
    CrizalDisc += ((Prod4REAmount + Prod4LEAmount)/1.12)*SheetY.range('U13').value

    Prod5REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T14').value), 'RE PRICE'].sum()
    Prod5LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T14').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod5REAmount + Prod5LEAmount)
    CrizalDisc += ((Prod5REAmount + Prod5LEAmount)/1.12)*SheetY.range('U14').value

    Prod6REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T15').value), 'RE PRICE'].sum()
    Prod6LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T15').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod6REAmount + Prod6LEAmount)
    CrizalDisc += ((Prod6REAmount + Prod6LEAmount)/1.12)*SheetY.range('U15').value

    Prod7REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T16').value), 'RE PRICE'].sum()
    Prod7LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T16').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod7REAmount + Prod7LEAmount)
    CrizalDisc += ((Prod7REAmount + Prod7LEAmount)/1.12)*SheetY.range('U16').value

    Prod8REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T17').value), 'RE PRICE'].sum()
    Prod8LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T17').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod8REAmount + Prod8LEAmount)
    CrizalDisc += ((Prod8REAmount + Prod8LEAmount)/1.12)*SheetY.range('U17').value

    Prod9REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T18').value), 'RE PRICE'].sum()
    Prod9LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T18').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod9REAmount + Prod9LEAmount)
    CrizalDisc += ((Prod9REAmount + Prod9LEAmount)/1.12)*SheetY.range('U18').value

    Prod10REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T19').value), 'RE PRICE'].sum()
    Prod10LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T19').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod10REAmount + Prod10LEAmount)
    CrizalDisc += ((Prod10REAmount + Prod10LEAmount)/1.12)*SheetY.range('U19').value

    Prod11REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T20').value), 'RE PRICE'].sum()
    Prod11LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T20').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod11REAmount + Prod11LEAmount)
    CrizalDisc += ((Prod11REAmount + Prod11LEAmount)/1.12)*SheetY.range('U20').value

    Prod12REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T21').value), 'RE PRICE'].sum()
    Prod12LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T21').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod12REAmount + Prod12LEAmount)
    CrizalDisc += ((Prod12REAmount + Prod12LEAmount)/1.12)*SheetY.range('U21').value

    Prod13REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T22').value), 'RE PRICE'].sum()
    Prod13LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T22').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod13REAmount + Prod13LEAmount)
    CrizalDisc += ((Prod13REAmount + Prod13LEAmount)/1.12)*SheetY.range('U22').value

    Prod14REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T23').value), 'RE PRICE'].sum()
    Prod14LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T23').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod14REAmount + Prod14LEAmount)
    CrizalDisc += ((Prod14REAmount + Prod14LEAmount)/1.12)*SheetY.range('U23').value

    Prod15REAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T24').value), 'RE PRICE'].sum()
    Prod15LEAmount = df.loc[(df['CUSTOMER'] == cust) & (df['RE PRODUCT'] == SheetY.range('T24').value), 'LE PRICE'].sum()
    CrizalAmt += (Prod15REAmount + Prod15LEAmount)
    CrizalDisc += ((Prod15REAmount + Prod15LEAmount)/1.12)*SheetY.range('U24').value


    CrizalAmount += (CrizalAmt,)
    CrizalDiscount += (CrizalDisc,)
    CrizalAmt = 0.00
    CrizalDisc = 0.00

SheetY['W3'].options(transpose=True).value = CrizalAmount
SheetY['X3'].options(transpose=True).value = CrizalDiscount
