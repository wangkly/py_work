from openpyxl import load_workbook
from openpyxl import Workbook
from decimal import *

def writeFile(targetWb,data,factor=0):
    nanjing =  targetWb['南京']
    for k,v in data.items():
        code = k[0:1] #第一个字符 excel列
        num = k[1:] #剩余字符 excel 行
        #计算出新key
        newkey = str(code) + str(int(num) + factor * 7)
        print(newkey,v)
        nanjing[newkey] = Decimal(str(v)) / Decimal('10000')
    return
