import xlrd
import xlwt
def loadExcel(filename):
    f = 0
    try:
        f = xlrd.open_workbook(filename)
    except:
        print("Cannot open the file " + filename)
        return 0
    return f
def pickup(f, sheetId=0, lis=[], requestWholeTable=False):
    sheet = f.sheet_by_index(sheetId)
    rows = sheet.nrows
    cols = sheet.ncols
    data = []
    nlis = []
    if not requestWholeTable:
        for i in range(len(lis)):
            for j in range(cols):
                if lis[i] == sheet.cell(0, j).value:
                    nlis.append(j)
    else:
        for i in range(cols):
            nlis.append(i)
    for i in range(rows):
        tmp = []
        row = sheet.row_values(i)
        for j in nlis:
            tmp.append(row[j])
        data.append(tmp)
    return data


def getData(filename, sheetId=0, lis=[], requireHeader=False, requestWholeTable=False):
    f = loadExcel(filename)
    data = pickup(f, sheetId, lis, requestWholeTable)
    if not requireHeader:
        del (data[0])
    return data


def filterData(f, colId=0, val=[]):
    data = []
    for x in f:
        if x[colId] in val:
            data.append(x)
    return data


def printToFile(filename, data):
    f = xlwt.Workbook(encoding="gbk")
    sheet1 = f.add_sheet("sheet1", cell_overwrite_ok=True)
    txtStyle = xlwt.XFStyle()
    txtStyle.num_format_str = '@'
    for i in range(len(data)):
        for j in range(len(data[i])):
            sheet1.write(i, j, data[i][j], style=txtStyle)
    f.save(filename + '.xls')