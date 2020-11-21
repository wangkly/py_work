from compute import compute
from compute_jifu import computeJF
from openpyxl import Workbook
from openpyxl import load_workbook
from write_file import writeFile

wb = load_workbook('/Users/wangkly/py_work/xuTest/src/辅助余额表.xlsx')
template = load_workbook('/Users/wangkly/py_work/xuTest/src/template.xlsx')

def entry(workbook,name):
    worksheet = workbook[str(name)]
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

    template.save('test.xlsx')

    # for k,v in d6.items():
    #     print(k,v)
    print('结束')
    return

entry(wb,'10月')
