#encoding:utf-8
import xlrd
import xlwt
import os.path
import Product

def opExcel(fileName):
    workbook = xlrd.open_workbook(fileName)
    sheet1 = workbook.sheet_by_index(0)
    index = 1
    productList = []
    product = Product.Product(index,sheet1.row_values(1)[3],{},[])
    for i in range(1,sheet1.nrows):
        if(sheet1.cell(i,0).ctype == 0):
            content = Product.Content(sheet1.row_values(i)[1],sheet1.row_values(i)[2],sheet1.row_values(i)[4],sheet1.row_values(i)[5])
            product.contentList[sheet1.row_values(i)[1]] = content
            product.nameList.append(sheet1.row_values(i)[1])
        else:
            index = index +1
            if len(product.contentList)!=0:
                productList.append(product)
                product = Product.Product(index,sheet1.row_values(i)[3],{},[])
            content = Product.Content(sheet1.row_values(i)[1],sheet1.row_values(i)[2],sheet1.row_values(i)[4],sheet1.row_values(i)[5])
            product.contentList[sheet1.row_values(i)[1]] = content
            product.nameList.append(sheet1.row_values(i)[1])

    productList.append(product)
    return productList

def writeFile(productList1,productList2,fileName):
    wbk = xlwt.Workbook(encoding='utf-8')
    sheet_w = wbk.add_sheet('sheet1', cell_overwrite_ok=True)
    sheet_w.col(3).width = 5000
    # tittle_style = xlwt.easyxf('font: height 300, name SimSun, colour_index red, bold on; align: wrap on, vert centre, horiz center;')
    # sheet_w.write_merge(0,2,0,8,u'中文',tittle_style)
    sheet_w.write(0,0,u'序号')
    sheet_w.write(0,1,u'标准中文名称')
    sheet_w.write(0,2,u'INCI名')
    sheet_w.write(0,3,u'原料含量(%)')
    sheet_w.write(0,4,u'原料中成份含量(%)')
    sheet_w.write(0,5,u'实际成份含量（%）')
    sheet_w.write(0,6,u'序号')
    sheet_w.write(0,7,u'标准中文名称')
    sheet_w.write(0,8,u'INCI名')
    sheet_w.write(0,9,u'原料含量(%)')
    sheet_w.write(0,10,u'原料中成份含量(%)')
    sheet_w.write(0,11,u'实际成份含量（%）')
    line = 1
    index = 1
    for product in productList1:
        sheet_w.write_merge(line,line+len(product.contentList)-1,0,0,index)
        sheet_w.write_merge(line,line+len(product.contentList)-1,3,3,product.percentage)
        for name in product.nameList:
            content = product.contentList[name]
            sheet_w.write(line,1,content.name_cn)
            sheet_w.write(line,2,content.inci)
            sheet_w.write(line,4,content.percentage)
            sheet_w.write(line,5,content.perInAll)
            line += 1
        index+=1
    line = 1
    index = 1
    for product in productList2:
        sheet_w.write_merge(line,line+len(product.contentList)-1,6,6,index)
        sheet_w.write_merge(line,line+len(product.contentList)-1,9,9,product.percentage)
        for name in product.nameList:
            content = product.contentList[name]
            sheet_w.write(line,7,content.name_cn)
            sheet_w.write(line,8,content.inci)
            sheet_w.write(line,10,content.percentage)
            sheet_w.write(line,11,content.perInAll)
            line += 1
        index+=1
    wbk.save(fileName)

def sort(productList1,productList2):
    resultList1 = []
    resultList2 = []
    resultList3 = []
    for product1 in productList1:
        for product2 in productList2:
            nameSet1 = set(product1.nameList)
            nameSet2 = set(product2.nameList)
            if nameSet1 == nameSet2 :
                productList2.remove(product2)
                product2.nameList = product1.nameList
                resultList1.append(product1)
                resultList2.append(product2)
                break;
        else:
            resultList3.append(product1)
    resultList1 = resultList1 + resultList3
    resultList2 = resultList2 + productList2
    map = {}
    map[0] = resultList1
    map[1] = resultList2
    return map


def main(fileName1,fileName2):
    productList1 = opExcel(fileName1)
    productList2 = opExcel(fileName2)
    map = sort(productList1,productList2)
    # print len(map[0])
    # print len(map[1])
    resultFileName = os.path.join(os.path.dirname(fileName1),'result.xls')
    writeFile(map[0],map[1],resultFileName)
    return resultFileName
    # for product in map[0]:
    #     print pr
    # print '&&&'
    # for product in map[1]:
    #     print product.percentage