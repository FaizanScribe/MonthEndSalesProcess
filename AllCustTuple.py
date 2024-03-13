def CutomersList(fileName, cellLocation):


    import xlwings as xw


    wb = xw.Book(fileName)
    ws1 = wb.sheets[0]
    noOFSheets = int(ws1.range(cellLocation).value)
    DName = ()

    for i in range(6, noOFSheets-13):
        wsTemp = wb.sheets[i]
        DName += (wsTemp.name,)

    return DName


