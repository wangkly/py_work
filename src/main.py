import string
from decimal import *
getcontext().prec = 28
from compute import compute
from compute_jifu import computeJF
from openpyxl import Workbook
from openpyxl import load_workbook
from write_file import writeFile

wb = load_workbook('/Users/wangkly/py_work/xuTest/src/辅助余额表.xlsx')
template = load_workbook('/Users/wangkly/py_work/xuTest/src/template.xlsx')

def entry(workbook,name):
    worksheet = workbook[str(name)]
    #保费收入（正的就是负的，负的就是正的）
    d1 = compute(worksheet,['PC10'])
    writeFile(template,d1,0)
    d2 = compute(worksheet,['PC20'])
    writeFile(template,d2,1)
    d3 = compute(worksheet,['PC30','PC40'])
    writeFile(template,d3,2)

    #计算理赔给付 ，各类给付，退保金额
    d4 = computeJF(worksheet,['PC10'])
    writeFile(template,d4,0)
    d5 = computeJF(worksheet,['PC20'])
    writeFile(template,d5,1)
    d6 = computeJF(worksheet,['PC30','PC40'])
    writeFile(template,d6,2)


    #计算合计
    nanjing = template['南京'] #ws
    columns = list(string.ascii_uppercase[2:16])
    rows = [12,19,26]
    for column in columns:
        aTotal = Decimal('0')
        for row in rows:
            akey = str(column)+str(row)
            print(akey,nanjing[akey])
            aTotal +=  Decimal(str( nanjing[akey].value ))  #C6，C7,C8,C9,C10,C11 
        nanjing[str(column)+'29'] = aTotal

    template.save('test.xlsx')

    # for k,v in d6.items():
    #     print(k,v)
    print('结束')
    return

# entry(wb,'10月')
