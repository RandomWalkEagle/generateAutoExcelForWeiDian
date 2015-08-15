# -*- coding: utf-8 -*- 
#python runxlrd.py --help

import  xdrlib ,sys
import xlrd
import sys
#import getopt
from pyExcelerator import *


class OrderInfo(object):
    """客户信息收集"""
    def __init__(self, accountName, accountNumber, accountAddress, accountOrdNumber):
        self.accountName =  accountName
        self.accountNumber = accountNumber
        self.accountAddress = accountAddress
        self.accountOrdNumber = accountOrdNumber
        self.goodsName_Number = {}
        self.goodsName_Price= {}

    def appendGoodsInfo(self, goodsName, goodsNumber, goodsPrice):
        number = goodsNumber.strip()
        price  = goodsPrice.strip()

        if self.goodsName_Number.has_key(goodsName):
            self.goodsName_Number[goodsName] += int(number)
        else:
            self.goodsName_Number[goodsName] = int(number)

        if self.goodsName_Price.has_key(goodsName):
            self.goodsName_Price[goodsName] = float(price)
        else:            
            self.goodsName_Price[goodsName] = float(price)

def handleExcel(fileName, outputName, sheetIndex) :
    data  = xlrd.open_workbook(fileName);
    table = data.sheets()[sheetIndex]
    nrows = table.nrows
    ncols = table.ncols
    print nrows, ncols

    #获取订单描述列（18列，对应S）
    print "1、开始统计订单商品数量 ---》》》》》"
    mapCount = {}    
    for i in  range(1, nrows) :
        #print i, table.cell_value(i,18);
        strGoodsName = table.cell_value(i,18).split(']')
        for info in strGoodsName :
            nameAndCount = info.split(u'[数量:')
            if len(nameAndCount) == 2 :
                goodsName = nameAndCount[0]
                name = goodsName.strip()
                goodsCount = nameAndCount[1].split(',')[0]
                count = str(goodsCount).strip()
                
                if name in mapCount :
                    mapCount[name] += int(count)
                else :
                    #print goodsName, goodsCount
                    mapCount[name] = int(count)

    goodsNamecount = 0
    goodsCount = 0
    sortedMapCount = sorted(mapCount.iteritems(), key=lambda d:d[1], reverse = True)

#    print outputName
    w = Workbook()  
    ws = w.add_sheet(u'订单商品数量统计')
    ws.write(0,0,u'商品名称')
    ws.write(0,1,u'商品数量')

    row = 1
    #访问List列表
    for key in sortedMapCount:
        #print type(key)
        #print '商品名=%s, 商品数量=%s' %(key[0].encode('utf-8'), key[1])
        ws.write(row, 0, key[0])
        ws.write(row, 1, key[1])
        goodsNamecount += 1
        goodsCount += key[1]
        row += 1
    ws.write(row, 0, u'商品总数: %d' %(goodsNamecount))
    ws.write(row, 1, u'商品数量总数: %d' %(goodsCount))
    w.save(outputName)
    print '统计完成，商品总数: %d, 商品数量总数: %d' %(goodsNamecount, goodsCount)




    print "2、开始根据地址合并订单 ---》》》》》"
    orderInfos = {}
    for i in  range(1, nrows) :
        #获取收件人电话
        strNumber = table.cell_value(i,7)
        #获取收件人名称
        strName = table.cell_value(i,6)
        #获取客户地址
        strAddress = table.cell_value(i,14)
        #获取下订单的人手机号
        strOrdNumber = table.cell_value(i, 29)
        orderInfo = None
        if orderInfos.has_key(strNumber):
            orderInfo = orderInfos[strNumber]
        else:
            orderInfo = OrderInfo(strName, strNumber, strAddress, strOrdNumber)
            orderInfos[strNumber] = orderInfo

        #获取商品名称清单
        strGoodsName = table.cell_value(i,18).split(']')
        #print type(strGoodsName), len(strGoodsName)
        #获取商品数量清单
        strGoodsCount = table.cell_value(i,9).split(',')
        #print type(strGoodsCount), strGoodsCount
        #获取每个商品的价格 
        strGoodsPrice = table.cell_value(i,10).split(',')
        #print type(strGoodsPrice), strGoodsPrice
        
        tick = 0
        for strC in strGoodsName :
            strD = strC.split(u'[数量:')
            if len(strD) == 2 :
                goodsName = strD[0]
                goodsName.strip()
                orderInfo.appendGoodsInfo(goodsName, strGoodsCount[tick], strGoodsPrice[tick])
                tick += 1
        

    ws = w.add_sheet(u'客户地址与合并订单');
    #访问字典
    #ws.write(0, 0, u'下单人')
    #ws.write(0, 1, u'收件人')
    #ws.write(0, 2, u'手机号码')
    #ws.write(0, 3, u'地址')
    nrows = 0
    count = 0
    totalPrice = 0
    sortedOrderInfos = sorted(orderInfos.iteritems(), key=lambda (k, v):int(v.accountOrdNumber.split('-')[1]), reverse = True)
    lastItem = None
    orderCount = 0
    for item in sortedOrderInfos:
        key = item[0]
        value = item[1]
        if lastItem == None:
            lastItem = item[1]

        accountOrdNumber = value.accountOrdNumber.split('-')[1]
        #print int(accountOrdNumber), type(int(accountOrdNumber))        
        if int(lastItem.accountOrdNumber.split('-')[1]) != int(accountOrdNumber):
            ws.write(nrows, 0, u'>>>>>>>>')
            ws.write(nrows, 1, u'>>>>>>>>')
            ws.write(nrows, 2, u'>>>>>>>>')
            ws.write(nrows, 3, u'下单人订单数总计：%d' %(orderCount))
            ws.write(nrows, 5, u'下单人订单总金额：')
            nrows += 1
            ws.write(nrows, 0, u'>>>>>>>>')
            ws.write(nrows, 1, u'>>>>>>>>')
            ws.write(nrows, 2, u'>>>>>>>>')
            ws.write(nrows, 3, u'>>>>>>>>')
            ws.write(nrows, 4, u'>>>>>>>>')
            ws.write(nrows, 5, u'>>>>>>>>')
            ws.write(nrows, 6, u'>>>>>>>>')
 
            nrows += 2
            lastItem = value
            orderCount = 1
        else:
            orderCount += 1

        ws.write(nrows, 0, u'下单人')
        ws.write(nrows, 1, u'收件人')
        ws.write(nrows, 2, u'手机号码')
        ws.write(nrows, 3, u'地址')
        nrows += 1

        if orderInfos.has_key(accountOrdNumber):
            ws.write(nrows, 0, orderInfos[accountOrdNumber].accountName+value.accountOrdNumber)
        else :
            ws.write(nrows, 0, value.accountOrdNumber)
        ws.write(nrows, 1, value.accountName)
        ws.write(nrows, 2, value.accountNumber)
        ws.write(nrows, 3, value.accountAddress)
        nrows += 1

        ws.write(nrows, 1, u'商品名')
        ws.write(nrows, 2, u'商品数')
        ws.write(nrows, 3, u'商品单价')
        ws.write(nrows, 4, u'商品总价')
        ws.write(nrows, 5, u'商品实际总价')
        ws.write(nrows, 6, u'快递价')
        nrows += 1

        subTotalPrice = 0
        for key,subvalue in value.goodsName_Number.iteritems():
            ws.write(nrows, 1, key)
            ws.write(nrows, 2, subvalue)
            ws.write(nrows, 3, value.goodsName_Price[key])
            ws.write(nrows, 4, subvalue * value.goodsName_Price[key])
            nrows += 1
            
            count += subvalue
            subTotalPrice += subvalue * value.goodsName_Price[key]
        totalPrice += subTotalPrice
        ws.write(nrows, 4, u'总价: %d' %(subTotalPrice))

        #填入商品信息
        nrows += 2

    ws.write(nrows, 0, u'>>>>>>>>')
    ws.write(nrows, 1, u'>>>>>>>>')
    ws.write(nrows, 2, u'>>>>>>>>')
    ws.write(nrows, 3, u'下单人订单数总计：%d' %(orderCount))
    ws.write(nrows, 5, u'下单人订单总金额：')
    nrows += 1
    ws.write(nrows, 0, u'>>>>>>>>')
    ws.write(nrows, 1, u'>>>>>>>>')
    ws.write(nrows, 2, u'>>>>>>>>')
    ws.write(nrows, 3, u'>>>>>>>>')
    ws.write(nrows, 4, u'>>>>>>>>')
    ws.write(nrows, 5, u'>>>>>>>>')
    ws.write(nrows, 6, u'>>>>>>>>')
    nrows += 1

    ws.write(nrows, 2, u'商品数量总数: %d' %(count))
    ws.write(nrows, 4, u'商品总价: %d' %(totalPrice))
    print '合并完成，商品数量总数: %d, 商品总价: %d' %(count, totalPrice)
    w.save(outputName)
        

def main(argv):
    '''
    try:
        opts, args = getopt.getopt(argv[1:])
    except getopt.GetoptError, err:
        print str(err)
        sys.exit()

    fileName = "input.xls"
    '''
    if len(sys.argv) < 2:
        print "请输入xls文件名： python main.py xls文件名"
        return
    fileName = sys.argv[1]
    outputName = fileName.split(".")[0] + "_统计.xls"
    handleExcel(fileName, outputName, 0);

if __name__=="__main__":
    main(sys.argv)
    