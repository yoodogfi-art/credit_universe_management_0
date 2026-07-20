import json
import openpyxl

fnguide = openpyxl.load_workbook("fnguide.xlsx")
ws = fnguide.active   # 또는 fnguide["시트명"]


for cell in ws["A"]:
    if cell.value is not None:
        cell.value = str(cell.value).zfill(5)

fnguide.save("fnguide.xlsx")