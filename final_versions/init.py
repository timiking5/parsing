import openpyxl
from datetime import date
import os
# import sys


if __name__ == '__main__':
    filename = date.today().strftime("%d.%m.%y") + ".xlsx"
    wb = openpyxl.Workbook()
    list_names = ["mkb_curr", "vtb",
                  "vtb_prem", "sovkom",
                  "gazprom", "sber",
                  "sber_curr", "alfa",
                  "alfa_curr", "rshb",
                  "rshb_prem", "psb"]
    wb.worksheets[0].title = "mkb"
    for list_name in list_names:
        wb.create_sheet(list_name)
    wb.save(filename=filename)
