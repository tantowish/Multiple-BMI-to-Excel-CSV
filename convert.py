import openpyxl
import csv
import pandas as pd

def convert(file):
    excel = openpyxl.load_workbook(file+".xlsx")
    
    sheet = excel.active
    
    col = csv.writer(open(file+".csv", 'w', newline=""))

    for r in sheet.rows:
        col.writerow([cell.value for cell in r])