import openpyxl

workbook = openpyxl.Workbook()

ws = workbook.active

ws['A1'] = 'wangkly'

workbook.save('new.xlsx')

