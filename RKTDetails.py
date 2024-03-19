import xlwings as xw
import pandas as pd

fileName = "C:\\Users\\USER\\Desktop\\BizomDiscounting\\March 24.xlsm"
wb = xw.Book(fileName)
PSLIP = wb.sheets('PSLIP')
SheetY = wb.sheets('SheetY')
Prisha = wb.sheets('PRISHA OPTICALS')
Ehm = wb.sheets('EYE HOSPITAL MAHAGAMA')
vn = wb.sheets('VISHNU NETRALAYA')
df = pd.read_excel(fileName, sheet_name='SheetY', usecols='CB:CS')

PrishaVRX = round(Prisha.range("D6").value/1.12,2)
vnVRX = round(vn.range("D6").value/1.12,2)
EhmCH = round(Ehm.range("D9").value, 2)

VrxReValRktH = df.loc[(df['CUSTOMER'] == 'RKT HINOO') & (df['RE BRAND'] == 'VRX'), 'RE PRICE'].sum()
VrxLeValRktH = df.loc[(df['CUSTOMER'] == 'RKT HINOO') & (df['LE BRAND'] == 'VRX'), 'LE PRICE'].sum()
VrxRKTH = round((VrxReValRktH + VrxLeValRktH)/1.12,2)

VrxReValRktC = df.loc[(df['CUSTOMER'] == 'RKT  CHIROUNDI') & (df['RE BRAND'] == 'VRX'), 'RE PRICE'].sum()
VrxLeValRktC = df.loc[(df['CUSTOMER'] == 'RKT  CHIROUNDI') & (df['LE BRAND'] == 'VRX'), 'LE PRICE'].sum()
VrxRKTC = round((VrxReValRktC + VrxLeValRktC)/1.12, 2)

VrxReValRktB = df.loc[(df['CUSTOMER'] == 'RKT BARIATU') & (df['RE BRAND'] == 'VRX'), 'RE PRICE'].sum()
VrxLeValRktB = df.loc[(df['CUSTOMER'] == 'RKT BARIATU') & (df['LE BRAND'] == 'VRX'), 'LE PRICE'].sum()
VrxRKTB = round((VrxReValRktB + VrxLeValRktB)/1.12,2)

print("Prisha VRX ;     ", PrishaVRX)
print("Vishnu Netralaya VRX :      ", vnVRX)
print("Mahagama CH :      ", EhmCH)
print("RKT H VRX     :", VrxRKTH)
print("RKT C VRX     :", VrxRKTC)
print("RKT B VRX     :", VrxRKTB)


