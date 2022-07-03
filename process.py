import xlrd
from openpyxl import load_workbook
import argparse
from config import get_config

def xlWriteLine(sheet, content, irow):
    sheet.insert_rows(irow+1)
    for icol, icell in enumerate(content):
        # colchar = convertColumnLetter(icol)
        # sheet[f'{colchar}{irow+1}'] = icell.value
        sheet.cell(irow+1, icol+1).value = icell.value

parser = argparse.ArgumentParser('xlsx reader', add_help=False)
parser.add_argument('--cfg', default="default.yaml", type=str)
args = parser.parse_args()

cfg = get_config(args.cfg)

wtFile = load_workbook(cfg["Template"])
wtPointer = {}
sheetDict = {}
for isheet in cfg["Sheets"]:
    sheetDict[isheet["name"]] = {
        "prefix_lines": isheet["prefix_lines"],
        "posfix_string": isheet["posfix_string"],
    }
    wtPointer[isheet["name"]] = isheet["prefix_lines"]

for fileName in cfg["Files"]:
    ifile = xlrd.open_workbook(fileName)
    fileSheetNames = ifile.sheet_names()
    for sheetInfo in cfg["Sheets"]:
        sheetName = sheetInfo["name"]
        if sheetName not in fileSheetNames:
            print(f"{fileName} 文件中没有表格 {sheetName}")
        else:
            isheet = ifile.sheet_by_name(sheetName)
            wtSheet = wtFile.get_sheet_by_name(sheetName)
            rdpt = sheetDict[sheetName]["prefix_lines"]
            for idx, irow in enumerate(isheet.get_rows()):
                if idx < rdpt:
                    continue
                # 1 对应 string
                if irow[0].ctype == 1 and irow[0].value == sheetDict[sheetName]["posfix_string"]:
                    break
                xlWriteLine(wtSheet, irow, wtPointer[sheetName])
                wtPointer[sheetName] += 1

wtFile.save(cfg["Output"])