from compute import compute
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
    template.save('test.xlsx')

    # for k,v in d3.items():
    #     print(k,v)
    print('结束')
    return

entry(wb,'10月')
