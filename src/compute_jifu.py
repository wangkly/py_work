from decimal import *
getcontext().prec = 28
from openpyxl import Workbook
from openpyxl import load_workbook
from gongsi import gongsi_filter
from options_filter import outer_filter
from functools import reduce
from rowsum import rowSum
from qimo import qimo

#计算理赔给付 ，其他各类给付 ，退保金额
def computeJF(worksheet,lirunType,companyCode = '1900'):
    #返回结果 ,dict
    resultDict={}

    #第0行，标题行
    row0 = [item.value for item in list(worksheet.rows)[0]]
    gonsi_index = row0.index('公司代码')

    #公司代码 筛选 1900
    gongsiFilterd=list(filter(gongsi_filter(gonsi_index,companyCode),list(worksheet.rows)))

    #筛科目
    kemuIndex = row0.index('科目')

    #理赔给付筛选 赔付支出-赔款支出:6511010000
    lipeiOptions = ['6511010000']
    lipeiFilterd = list(filter(outer_filter(lipeiOptions,kemuIndex),gongsiFilterd))
    
    #其他给付 赔付支出-年金给付：6511040000 ，赔付支出-死伤医疗给付：6511020000 ，赔付支出-满期给付：6511030000
    qitaOptions =['6511040000','6511020000','6511030000']
    qitaFilterd = list(filter(outer_filter(qitaOptions,kemuIndex),gongsiFilterd))

    #退保金 退保金-退保亏损
    tuibaoOptions =['6531010000','6531030000']
    tuibaoFilterd = list(filter(outer_filter(tuibaoOptions,kemuIndex),gongsiFilterd))


    #利润中心描述筛选 
    lirunIndex = row0.index('利润中心')
    lirunOptins=lirunType 

    #理赔给付 （个险、团险、银保对应筛选后的数据），用作下面筛选短期健康险，意外险等
    lipeiBaseData = list(filter(outer_filter(lirunOptins,lirunIndex),lipeiFilterd))

    #其他给付 
    qitaBaseData = list(filter(outer_filter(lirunOptins,lirunIndex),qitaFilterd))

    #退保金额
    tuibaoBaseData = list(filter(outer_filter(lirunOptins,lirunIndex),tuibaoFilterd))

    #险种大类描述 筛选
    xianzhongIndex = row0.index('险种大类')

    ########################   短期健康险   #############################
    #names=['短期费补医疗健康','短期普疾健康','短期定返医疗健康','短期重疾健康']
    xianzhongOptions=['18','17','19','16']

    #短期健康险 理赔给付 本年 、 小计
    lipeiduanqi = list(filter(outer_filter(xianzhongOptions,xianzhongIndex),lipeiBaseData))
    #理赔 短期健康 本期
    K6 = sum(list(map(rowSum,lipeiduanqi)))
    resultDict['K6'] = K6
    #理赔 短期健康 小计
    L6 = sum(list(map(qimo,lipeiduanqi)))
    resultDict['L6'] = L6

    #短期健康 其他给付 本年 、小计
    qitaduanqi = list(filter(outer_filter(xianzhongOptions,xianzhongIndex),qitaBaseData))
    #理赔 短期健康 本期
    M6 = sum(list(map(rowSum,qitaduanqi)))
    resultDict['M6'] = M6
    #理赔 短期健康 小计
    N6 = sum(list(map(qimo,qitaduanqi)))
    resultDict['N6'] = N6

    #短期健康 退保金额 本年 、小计
    tuibaoduanqi = list(filter(outer_filter(xianzhongOptions,xianzhongIndex),tuibaoBaseData))
    #理赔 短期健康 本期
    O6 = sum(list(map(rowSum,tuibaoduanqi)))
    resultDict['O6'] = O6
    #理赔 短期健康 小计
    P6 = sum(list(map(qimo,tuibaoduanqi)))
    resultDict['P6'] = P6



    ##########################   意外伤害险 ############################
    yiwaiOptins=['9']

    #意外险 理赔给付 本年 、 小计
    lipeiyiwai = list(filter(outer_filter(yiwaiOptins,xianzhongIndex),lipeiBaseData))
    #理赔 意外 本期
    lipeiYWCollect = map(rowSum,lipeiyiwai)
    K7 = sum(list(lipeiYWCollect))
    resultDict['K7'] = K7
    #理赔 意外 小计
    L7 = sum(list(map(qimo,lipeiyiwai)))
    resultDict['L7'] = L7

    #意外险 其他给付 本年 、 小计
    qitayiwai = list(filter(outer_filter(yiwaiOptins,xianzhongIndex),qitaBaseData))
    #理赔 意外 本期
    qitaYWCollect = map(rowSum,qitayiwai)
    M7 = sum(list(qitaYWCollect))
    resultDict['M7'] = M7
    #理赔 意外 小计
    N7 = sum(list(map(qimo,qitayiwai)))
    resultDict['N7'] = N7

    #意外险 退保金额 本年 、 小计
    tuibaoyiwai = list(filter(outer_filter(yiwaiOptins,xianzhongIndex),tuibaoBaseData))
    #理赔 意外 本期
    O7 = sum(list(map(rowSum,tuibaoyiwai)))
    resultDict['O7'] = O7
    #理赔 意外 小计
    P7 = sum(list(map(qimo,tuibaoyiwai)))
    resultDict['P7'] = P7


    ######################   一般寿险  ###################
    #普通定期寿险：1，普通两全寿险：2，普通年金寿险：23，普通养老年金寿险：25，普通终生寿险：3，长期定返医疗健康：13，长期普疾健康：11,长期重疾健康：10
    putongOptins=['1','2','23','25','3','13','11','10']

    #一般寿险 理赔给付 本年 、 小计
    lipeiyiban = list(filter(outer_filter(putongOptins,xianzhongIndex),lipeiBaseData))
    #理赔 一般寿险 本期
    lipeiYBCollect = map(rowSum,lipeiyiban)
    K8 = sum(list(lipeiYBCollect))
    resultDict['K8'] = K8
    #理赔 意外 小计
    L8 = sum(list(map(qimo,lipeiyiban)))
    resultDict['L8'] = L8


    #一般寿险 其他给付 本年 、 小计
    qitayiban = list(filter(outer_filter(putongOptins,xianzhongIndex),qitaBaseData))
    #理赔 一般寿险 本期
    M8 = sum(list(map(rowSum,qitayiban)))
    resultDict['M8'] = M8
    #理赔 意外 小计
    N8 = sum(list(map(qimo,qitayiban)))
    resultDict['N8'] = N8

    #一般寿险 退保金额 本年 、 小计
    tuibaoyiban = list(filter(outer_filter(putongOptins,xianzhongIndex),tuibaoBaseData))
    #理赔 一般寿险 本期
    O8 = sum(list(map(rowSum,tuibaoyiban)))
    resultDict['O8'] = O8
    #理赔 意外 小计
    P8 = sum(list(map(qimo,tuibaoyiban)))
    resultDict['P8'] = P8


    ##################################   分红类保险（银保才有）###################
    # 分红两全寿险:4,分红年金寿险:6 
    fenhongOptions=['4','6']

#****************** 分红类  理赔给付 本年 、 小计 **************#
    lipeifenhong = list(filter(outer_filter(fenhongOptions,xianzhongIndex),lipeiBaseData))
    #理赔 分红类  本期
    lipeiFHCollect = map(rowSum,lipeifenhong)
    K9 = sum(list(lipeiFHCollect))
    resultDict['K9'] = K9
    #理赔 意外 小计
    L9 = sum(list(map(qimo,lipeifenhong)))
    resultDict['L9'] = L9

#****************  分红 其他给付 *************#
    qitafenhong = list(filter(outer_filter(fenhongOptions,xianzhongIndex),qitaBaseData))
    #理赔 分红类  本期
    M9 = sum(list(map(rowSum,qitafenhong)))
    resultDict['M9'] = M9
    #理赔 意外 小计
    N9 = sum(list(map(qimo,qitafenhong)))
    resultDict['N9'] = N9


#*************  分红 退保金额 **************#
    tuibaofenhong = list(filter(outer_filter(fenhongOptions,xianzhongIndex),tuibaoBaseData))
    #理赔 分红类  本期
    O9 = sum(list(map(rowSum,tuibaofenhong)))
    resultDict['O9'] = O9
    #理赔 意外 小计
    P9 = sum(list(map(qimo,tuibaofenhong)))
    resultDict['P9'] = P9
    

    ############################  万能寿险  ######################
    wnOptions=['7']

#*************** 万能寿险  理赔给付 本年 、 小计********#
    lipeiWN = list(filter(outer_filter(wnOptions,xianzhongIndex),lipeiBaseData))
    #理赔 万能寿险  本期
    lipeiWNCollect = map(rowSum,lipeiWN)
    K10 = sum(list(lipeiWNCollect))
    resultDict['K10'] = K10
    #理赔 意外 小计
    L10 = sum(list(map(qimo,lipeiWN)))
    resultDict['L10'] = L10
    
#***********#  万能寿险 其他给付 *************#
    qitaWN = list(filter(outer_filter(wnOptions,xianzhongIndex),qitaBaseData))
    #理赔 万能寿险  本期
    M10 = sum(list(map(rowSum,qitaWN)))
    resultDict['M10'] = M10
    #理赔 意外 小计
    N10 = sum(list(map(qimo,qitaWN)))
    resultDict['N10'] = N10

#***********# 万能寿险 退保金额 *************#
    tuibaoWN = list(filter(outer_filter(wnOptions,xianzhongIndex),tuibaoBaseData))
    #理赔 万能寿险  本期
    O10 = sum(list(map(rowSum,tuibaoWN)))
    resultDict['O10'] = O10
    #理赔 意外 小计
    P10 = sum(list(map(qimo,tuibaoWN)))
    resultDict['P10'] = P10

    #计算各种小计
    columns = ['K','L','M','N','O','P']
    rows = [6,7,8,9,10,11]
    for column in columns:
        aTotal = Decimal('0')
        for row in rows:
            akey = str(column)+str(row)
            aTotal +=  Decimal(str( resultDict.get(akey,0) ))  #C6，C7,C8,C9,C10,C11 
        resultDict[str(column)+'12'] = aTotal


    return resultDict

