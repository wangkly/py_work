from openpyxl import Workbook
from openpyxl import load_workbook
from gongsi import gongsi_filter
from options_filter import outer_filter
from functools import reduce
from rowsum import rowSum
from qimo import qimo


#定义一个通用的函数,返回一个dict,通过传入的参数分别计算 个险，团险，银保的数据
#worksheet 要操作的辅助余额表，lirunType利润中心（个险，团险，银保）
def compute(worksheet,lirunType,companyCode = '1900'):
    #返回结果 ,dict
    resultDict={}

    #第0行，标题行
    row0 = [item.value for item in list(worksheet.rows)[0]]
    gonsi_index = row0.index('公司代码')

    #公司代码 筛选 1900
    gongsiFilterd=list(filter(gongsi_filter(gonsi_index,companyCode),list(worksheet.rows)))

    #筛科目
    kemuIndex = row0.index('科目')
    # names=['保费收入－首年首期','保费收入－首年续期','保费收入-保全','保费收入-团体短期意外险保费收入']
    kemuOptins =['6031010001','6031010002','6031030000','6031050000']
    kemuFilterd = list(filter(outer_filter(kemuOptins,kemuIndex),gongsiFilterd))

    #理赔给付筛选 赔付支出-赔款支出:6511010000
    lipeiOptions = ['6511010000']
    lipeiFilterd = list(filter(outer_filter(lipeiOptions,kemuIndex),gongsiFilterd))

    #利润中心描述筛选 
    lirunIndex = row0.index('利润中心')
    lirunOptins=lirunType 
    lirunFilterd = list(filter(outer_filter(lirunOptins,lirunIndex),kemuFilterd))

    #理赔给付 （个险、团险、银保对应筛选后的数据），用作下面筛选短期健康险，意外险等
    lipeiBaseData = list(filter(outer_filter(lirunOptins,lirunIndex),lipeiFilterd))

    #险种大类描述 筛选
    xianzhongIndex = row0.index('险种大类')

    ########################   短期健康险   #############################
    #names=['短期费补医疗健康','短期普疾健康','短期定返医疗健康','短期重疾健康']
    xianzhongOptions=['18','17','19','16']
    xianzhongFilterd =  list(filter(outer_filter(xianzhongOptions,xianzhongIndex),lirunFilterd))

    #先map取每一行期间借方，期间贷方之和list
    rowCollect = map(rowSum,xianzhongFilterd)
    total = sum(list(rowCollect))
    # print('短期健康险',total)
    resultDict['C6'] = total 

    #本年保费收入 G6
    G6 =  sum(list(map(qimo,xianzhongFilterd)))
    resultDict['G6'] = G6 

    #短期健康险 期缴
    #缴费方式index
    jiaofeiIndex = row0.index('缴费方式')
    qijiaoOptions=['2']
    qijiao  = list(filter(outer_filter(qijiaoOptions,jiaofeiIndex),xianzhongFilterd))
    qijiaoTotal = sum(list(map(rowSum,qijiao)))
    # print('短期健康-期缴',qijiaoTotal)
    resultDict['D6'] = qijiaoTotal

    #本年保费收入 首年 期缴
    H6 = sum(list(map(qimo,qijiao)))
    # print('H6',H6)
    resultDict['H6'] = H6

    #短期健康险-续期 科目：保费收入续期
    kemuOptions2=['6031020000']
    xuqiFilterd = list(filter(outer_filter(kemuOptions2,kemuIndex),gongsiFilterd))
    lirunFilterd2 = list(filter(outer_filter(lirunOptins,lirunIndex),xuqiFilterd))
    xianzhongFilterd2 =  list(filter(outer_filter(xianzhongOptions,xianzhongIndex),lirunFilterd2))
    xuqiTotal = sum(list(map(rowSum,xianzhongFilterd2)))
    # print('短期健康-续期',xuqiTotal)
    resultDict['E6'] = xuqiTotal

    #本年保费收入 续期
    I6=sum(list(map(qimo,xianzhongFilterd2)))
    # print('I6',I6)
    resultDict['I6']=I6

    #短期健康险 理赔给付 本年 、 小计
    lipeiduanqi = list(filter(outer_filter(xianzhongOptions,xianzhongIndex),lipeiBaseData))
    print('短期理赔给付',len(lipeiduanqi))
    #理赔 短期健康 本期
    lipeiCollect = map(rowSum,lipeiduanqi)
    K6 = sum(list(lipeiCollect))
    resultDict['K6'] = K6
    #理赔 短期健康 小计
    L6 = sum(list(map(qimo,lipeiduanqi)))
    resultDict['L6'] = L6

    ##########################   意外伤害险 ############################
    yiwaiOptins=['9']
    yiwaiFilterd = list(filter(outer_filter(yiwaiOptins,xianzhongIndex),lirunFilterd))
    yiwaiTotal = sum(list(map(rowSum,yiwaiFilterd)))
    resultDict['C7'] = yiwaiTotal

    #本年保费收入 G7
    G7=sum(list(map(qimo,yiwaiFilterd)))
    # print('G7',G7)
    resultDict['G7'] = G7

    #意外险 期缴
    yiwaiQijiao = list(filter(outer_filter(qijiaoOptions,jiaofeiIndex),yiwaiFilterd))
    D7 = sum(list(map(rowSum,yiwaiQijiao)))
    # print('D7',D7)
    resultDict['D7'] = D7

    #本年保费收入 H7
    H7= sum(list(map(qimo,yiwaiQijiao)))
    # print('H7',H7)
    resultDict['H7'] = H7

    #意外险 续期 本月
    yiwaiXQFilterd =  list(filter(outer_filter(yiwaiOptins,xianzhongIndex),lirunFilterd2))
    yiwaiXQ =  sum(list(map(rowSum,yiwaiXQFilterd)))
    # print('E7',yiwaiXQ)
    resultDict['E7'] = yiwaiXQ

    #意外 续期 本年
    yiwaiXQY = sum(list(map(qimo,yiwaiXQFilterd)))
    # print('I7',yiwaiXQY)
    resultDict['I7'] = yiwaiXQY

    #意外险 理赔给付 本年 、 小计
    lipeiyiwai = list(filter(outer_filter(yiwaiOptins,xianzhongIndex),lipeiBaseData))
    print('意外理赔给付',len(lipeiyiwai))
    #理赔 意外 本期
    lipeiYWCollect = map(rowSum,lipeiyiwai)
    K7 = sum(list(lipeiYWCollect))
    resultDict['K7'] = K7
    #理赔 意外 小计
    L7 = sum(list(map(qimo,lipeiyiwai)))
    resultDict['L7'] = L7




    ######################   一般寿险  ###################
    #普通定期寿险：1，普通两全寿险：2，普通年金寿险：23，普通养老年金寿险：25，普通终生寿险：3，长期定返医疗健康：13，长期普疾健康：11,长期重疾健康：10
    putongOptins=['1','2','23','25','3','13','11','10']
    yibanFilterd = list(filter(outer_filter(putongOptins,xianzhongIndex),lirunFilterd))

    yibanTotal = sum(list(map(rowSum,yibanFilterd)))
    # print('一般寿险',yibanTotal)
    resultDict['C8'] = yibanTotal

    #一般寿险 本年 G8
    yibanY =  sum(list(map(qimo,yibanFilterd)))
    # print('G8',yibanY)
    resultDict['G8'] = yibanY

    #一般寿险 期缴
    yibanQJ = list(filter(outer_filter(qijiaoOptions,jiaofeiIndex),yibanFilterd))
    D8 =  sum(list(map(rowSum,yibanQJ)))
    # print('D8',D8)
    resultDict['D8'] = D8

    #一般寿险 期缴 本年
    H8= sum(list(map(qimo,yibanQJ)))
    # print('H8',H8)
    resultDict['H8'] = H8

    #一般寿险 续期
    yibanXQFilterd = list(filter(outer_filter(putongOptins,xianzhongIndex),lirunFilterd2))
    yibanXQ = sum(list(map(rowSum,yibanXQFilterd)))
    # print('E8',yibanXQ)
    resultDict['E8'] = yibanXQ

    #一般寿险 续期 本年
    yibanXQY = sum(list(map(qimo,yibanXQFilterd)))
    # print('I8',yibanXQY)
    resultDict['I8'] = yibanXQY



    #一般寿险 理赔给付 本年 、 小计
    lipeiyiban = list(filter(outer_filter(putongOptins,xianzhongIndex),lipeiBaseData))
    print('一般理赔给付',len(lipeiyiban))
    #理赔 一般寿险 本期
    lipeiYBCollect = map(rowSum,lipeiyiban)
    K8 = sum(list(lipeiYBCollect))
    resultDict['K8'] = K8
    #理赔 意外 小计
    L8 = sum(list(map(qimo,lipeiyiban)))
    resultDict['L8'] = L8


    ##################################   分红类保险（银保才有）###################
    # 分红两全寿险:4,分红年金寿险:6 
    fenhongOptions=['4','6']
    fenhongFilterd = list(filter(outer_filter(fenhongOptions,xianzhongIndex),lirunFilterd))
    #分红类 本月 首年
    fenhongTotal = sum(list(map(rowSum,fenhongFilterd)))
    resultDict['C9'] = fenhongTotal

    #本年保费收入 G9
    G9=sum(list(map(qimo,fenhongFilterd)))
    resultDict['G9'] = G9

    #期缴 本月
    fhQijiao  = list(filter(outer_filter(qijiaoOptions,jiaofeiIndex),fenhongFilterd))
    fhQijiaoTotal = sum(list(map(rowSum,fhQijiao)))
    resultDict['D9'] = fhQijiaoTotal
    #期缴 本年
    H9 = sum(list(map(qimo,fhQijiao)))
    resultDict['H9'] = H9

    #续期 本月
    fhXQFilterd =  list(filter(outer_filter(fenhongOptions,xianzhongIndex),lirunFilterd2))
    fhXQ =  sum(list(map(rowSum,fhXQFilterd)))
    resultDict['E9'] = fhXQ

    #续期 本年
    fhXQY = sum(list(map(qimo,fhXQFilterd)))
    resultDict['I9'] = fhXQY

    #分红类  理赔给付 本年 、 小计
    lipeifenhong = list(filter(outer_filter(fenhongOptions,xianzhongIndex),lipeiBaseData))
    print('分红类 理赔给付',len(lipeifenhong))
    #理赔 分红类  本期
    lipeiFHCollect = map(rowSum,lipeifenhong)
    K9 = sum(list(lipeiFHCollect))
    resultDict['K9'] = K9
    #理赔 意外 小计
    L9 = sum(list(map(qimo,lipeifenhong)))
    resultDict['L9'] = L9

    ############################  万能寿险  ######################
    wnOptions=['7']
    wnFilterd = list(filter(outer_filter(wnOptions,xianzhongIndex),lirunFilterd))
    #万能寿险 本月 首年
    wnTotal = sum(list(map(rowSum,wnFilterd)))
    resultDict['C10'] = wnTotal

    #本年保费收入 G9
    G10=sum(list(map(qimo,wnFilterd)))
    resultDict['G10'] = G10

    #期缴 本月
    wnQijiao  = list(filter(outer_filter(qijiaoOptions,jiaofeiIndex),wnFilterd))
    wnQijiaoTotal = sum(list(map(rowSum,wnQijiao)))
    resultDict['D10'] = wnQijiaoTotal
    #期缴 本年
    H10 = sum(list(map(qimo,wnQijiao)))
    resultDict['H10'] = H10

    #续期 本月
    wnXQFilterd =  list(filter(outer_filter(wnOptions,xianzhongIndex),lirunFilterd2))
    wnXQ =  sum(list(map(rowSum,wnXQFilterd)))
    resultDict['E10'] = wnXQ
    #续期 本年
    wnXQY = sum(list(map(qimo,wnXQFilterd)))
    resultDict['I10'] = wnXQY

    #万能寿险  理赔给付 本年 、 小计
    lipeiWN = list(filter(outer_filter(wnOptions,xianzhongIndex),lipeiBaseData))
    print('万能寿险 理赔给付',len(lipeiWN))
    #理赔 万能寿险  本期
    lipeiWNCollect = map(rowSum,lipeiWN)
    K10 = sum(list(lipeiWNCollect))
    resultDict['K10'] = K10
    #理赔 意外 小计
    L10 = sum(list(map(qimo,lipeiWN)))
    resultDict['L10'] = L10


    return resultDict

