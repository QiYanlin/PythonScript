import re
import xlsxwriter
from decimal import Decimal
from operator import itemgetter

excelData = []
pivotData = {}
excel = xlsxwriter.Workbook('支付宝new.xls')


def read():
    print('正在读取支付宝回款数据')
    data = re.findall('<Data ss:Type="String">([\s\W\w\S]*?)</Data>', open('支付宝.xls', 'rb').read().decode('utf-8'))
    del data[0], data[0]
    lens = len(data)
    print('总共' + str(lens) + '条数据')
    column = 22
    for row in range(0, int(lens / column)):
        arr = data[row * column:(row + 1) * column]
        if row == 0:
            arr.insert(1, '入账日期')
            excelData.append(arr)
        elif arr[4][0:2] == 'GA':
            arr.insert(1, arr[1][0:10])
            arr[2] = arr[2][11:]
            arr[7] = 0 if arr[7].isspace() else arr[7]
            arr[8] = 0 if arr[8].isspace() else arr[8]
            excelData.append(arr)
            if arr[1] not in pivotData:
                pivotData[arr[1]] = [arr[1], '0', '0', '0', '0', '0', '0', '0', '0']
            str1 = arr[6][0:4]
            if str1 == '收费':
                pivotData[arr[1]][1] = str(Decimal(pivotData[arr[1]][1]) + Decimal(arr[7]))
                pivotData[arr[1]][2] = str(Decimal(pivotData[arr[1]][2]) + Decimal(arr[8]))
            elif str1 == '退费':
                pivotData[arr[1]][3] = str(Decimal(pivotData[arr[1]][3]) + Decimal(arr[7]))
                pivotData[arr[1]][4] = str(Decimal(pivotData[arr[1]][4]) + Decimal(arr[8]))
            elif str1 == '退款（交':
                pivotData[arr[1]][5] = str(Decimal(pivotData[arr[1]][5]) + Decimal(arr[7]))
                pivotData[arr[1]][6] = str(Decimal(pivotData[arr[1]][6]) + Decimal(arr[8]))
            elif str1 == '在线支付':
                pivotData[arr[1]][7] = str(Decimal(pivotData[arr[1]][7]) + Decimal(arr[7]))
                pivotData[arr[1]][8] = str(Decimal(pivotData[arr[1]][8]) + Decimal(arr[8]))
    print('GA数据:' + str(len(excelData) - 1) + '条')
    del data


def sheet1():
    print('正在生成GA数据sheet')
    ga = excel.add_worksheet('支付宝回款')
    income = outgo = '0'
    for i in range(0, len(excelData)):
        for x in range(0, len(excelData[i])):
            if (x == 7 or x == 8) and i > 0:
                if x == 7:
                    income = str(Decimal(income) + Decimal(excelData[i][x]))
                if x == 8:
                    outgo = str(Decimal(outgo) + Decimal(excelData[i][x]))
                ga.write_number(i, x, float(excelData[i][x]))
            else:
                ga.write(i, x, excelData[i][x])
    ga.write_number(i + 1, 7, float(income))
    ga.write_number(i + 1, 8, float(outgo))


def sheet2():
    print('正在生成数据透视sheet')
    sheet = excel.add_worksheet('支付宝数据透视')
    title = [['入账时间', '收费 求和项:收入（+元）', '收费 求和项:支出（-元）', '退费 求和项:收入（+元）', '退费 求和项:支出（-元）', '退款（交易退款） 求和项:收入（+元）',
              '退款（交易退款） 求和项:支出（-元）', '在线支付 求和项:收入（+元）', '在线支付 求和项:支出（-元）', '求和项:收入（+元）汇总', '求和项:支出（-元）汇总', '手续费', '收入']]
    total = [['总计', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0']]
    data = list(pivotData.values())
    data.sort(key=itemgetter(0))
    data = title + data + total
    lens = len(data)
    for i in range(0, lens):
        if 0 < i < lens - 1:
            data[i].append(str(Decimal(data[i][1]) + Decimal(data[i][3]) + Decimal(data[i][5]) + Decimal(data[i][7])))
            data[i].append(str(Decimal(data[i][2]) + Decimal(data[i][4]) + Decimal(data[i][6]) + Decimal(data[i][8])))
            data[i].append(str(Decimal(data[i][2]) - Decimal(data[i][3])))
            data[i].append(str(Decimal(data[i][7]) - Decimal(data[i][6])))
            for k in range(1, len(data[i])):
                data[lens - 1][k] = str(Decimal(data[lens - 1][k]) + Decimal(data[i][k]))
        for x in range(0, len(data[i])):
            if i != 0 and x != 0:
                sheet.write_number(i, x, float(data[i][x]))
            else:
                sheet.write(i, x, data[i][x])
    excel.close()
    print('支付宝回款数据处理成功')


if __name__ == '__main__':
    read()
    sheet1()
    sheet2()
