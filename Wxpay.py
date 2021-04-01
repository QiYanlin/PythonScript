import csv
import xlsxwriter
from decimal import Decimal
from operator import itemgetter

excelData = []
pivotData = {}
excel = xlsxwriter.Workbook('微信new.xls')


def read():
    print('正在读取微信回款数据')
    i = 0
    with open('微信.csv', encoding='UTF-8-sig') as wechat:
        f_csv = csv.reader(wechat)
        for row in f_csv:
            if i == 0:
                row.insert(0, '交易日期')
                excelData.append(row)
            elif row[6][0:3] == '`GA':
                for x in range(0, len(row)):
                    row[x] = row[x][1:]
                row.insert(0, row[0][0:10])
                row[1] = row[1][11:]
                excelData.append(row)
                if row[0] not in pivotData:
                    pivotData[row[0]] = [row[0], '0', '0', '0']
                pivotData[row[0]][1] = str(Decimal(pivotData[row[0]][1]) + Decimal(row[13]))
                pivotData[row[0]][2] = str(Decimal(pivotData[row[0]][2]) + Decimal(row[17]))
                pivotData[row[0]][3] = str(Decimal(pivotData[row[0]][3]) + Decimal(row[23]))
            i += 1
    print('GA数据:' + str(len(excelData) - 1) + '条')
    del wechat


def sheet1():
    print('正在生成GA数据sheet')
    ga = excel.add_worksheet('微信回款')
    for i in range(0, len(excelData)):
        for x in range(0, len(excelData[i])):
            if i != 0 and (x == 13 or x == 17 or x == 23):
                ga.write_number(i, x, float(excelData[i][x]))
            else:
                ga.write(i, x, excelData[i][x])


def sheet2():
    print('正在生成数据透视sheet')
    sheet = excel.add_worksheet('微信数据透视')
    title = [['交易日期', '应结订单金额', '退款金额', '手续费', '收入']]
    total = [['总计', '0', '0', '0', '0']]
    data = list(pivotData.values())
    data.sort(key=itemgetter(0))
    data = title + data + total
    lens = len(data)
    for i in range(0, lens):
        if 0 < i < lens - 1:
            data[i].append(str(Decimal(data[i][1]) - Decimal(data[i][2])))
            for k in range(1, len(data[i])):
                data[lens - 1][k] = str(Decimal(data[lens - 1][k]) + Decimal(data[i][k]))
        for x in range(0, len(data[i])):
            if i != 0 and x != 0:
                sheet.write_number(i, x, float(data[i][x]))
            else:
                sheet.write(i, x, data[i][x])
    excel.close()
    print('微信回款数据处理成功')


if __name__ == '__main__':
    read()
    sheet1()
    sheet2()
