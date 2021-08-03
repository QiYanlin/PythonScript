import re
import csv
import sys
import numpy
import datetime
import xlsxwriter
from decimal import Decimal
from dateutil.relativedelta import relativedelta

excelData = []
excel = xlsxwriter.Workbook('单据new.xls')
time = datetime.datetime.today()


def check_date():
    print('正在检查 时间 格式')
    flag = 0
    try:
        global time
        time = datetime.datetime.strptime(sys.argv[1], '%Y.%m')
    except ValueError:
        flag = 1
        print('时间 ' + sys.argv[1] + ' 格式错误！')
    if flag == 0:
        print('时间 格式正确')
    return flag


def check_num():
    print('正在检查 序号 格式')
    flag = 0
    try:
        int(sys.argv[2])
    except ValueError:
        flag = 1
        print('序号 ' + sys.argv[2] + ' 格式错误！')
    if flag == 0:
        print('序号 格式正确')
    return flag


def check_data():
    print('正在检查 消费事由 格式')
    check_flag = i = 0
    with open('单据.csv', encoding='UTF-8-sig') as check:
        f_csv = csv.reader(check)
        for row in f_csv:
            if i > 1 and '租赁费' in row[7]:
                data_flag = 0
                date_arr = sub_str(row[9]).split('-')
                if len(date_arr) != 2:
                    check_flag = data_flag = 1
                else:
                    try:
                        start = datetime.datetime.strptime(date_arr[0], '%Y.%m.%d')
                        end = datetime.datetime.strptime(date_arr[1], '%Y.%m.%d')
                        if start > end:
                            check_flag = data_flag = 1
                    except ValueError:
                        check_flag = data_flag = 1
                if data_flag == 1:
                    print('单号 ' + row[1] + ' 格式不正确！')
            i += 1
    if check_flag == 0:
        print('消费事由 格式正确')
    return check_flag


def read():
    print('正在读取OA报销单数据')
    temp = []
    i = 0
    with open('单据.csv', encoding='UTF-8-sig') as oa:
        f_csv = csv.reader(oa)
        for row in f_csv:
            if i > 1:
                date_arr = sub_str(row[6], 1).split('-')
                if len(date_arr) > 1 and len(date_arr[1]) == 1:
                    date_arr[1] = '0' + date_arr[1]
                if len(date_arr) == 1 and len(date_arr[0]) == 1:
                    date_arr[0] = '0' + date_arr[0]
                store_arr = get_store_arr(row[0])
                data = [row[1], row[2], store_arr[0], '-'.join(date_arr).split(' ')[0], row[9], '付' + row[0] + row[9],
                        row[7], row[8], datetime.datetime.strptime(row[5], '%Y-%m-%d').strftime('%Y-%m-%d'), '0', '0',
                        '', '', '', row[4], '', row[3], store_arr[1]]
                attention_str = ''
                if '租赁费' in row[7]:
                    date_arr = sub_str(row[9]).split('-')
                    start = datetime.datetime.strptime(date_arr[0], '%Y.%m.%d')
                    end = datetime.datetime.strptime(date_arr[1], '%Y.%m.%d')
                    if (end - start).days < 25:
                        month = datetime.datetime.strptime(re.sub(r'(\d+$)', '01', date_arr[0]), '%Y.%m.%d')
                        if month >= time + relativedelta(months=+1):
                            attention_str = '注意'
                            data[10] = data[7]
                        else:
                            data[9] = data[7]
                    else:
                        if get_month(time, start) > 1:
                            data[10] = data[7]
                        elif get_month(time, end) > 1:
                            data[10] = str((Decimal(get_month(time, end) - 1) / Decimal(get_month(start, end))
                                            * Decimal(data[7])).quantize(Decimal('0.00')))
                            data[9] = str(Decimal(row[8]) - Decimal(data[10]))
                        else:
                            data[9] = data[7]

                    if data[10] != '0':
                        data[11] = str(time.year + (time.month == 12)) + '-' + str(time.month == 12 or time.month + 1)
                        if get_month(time, start) > 1:
                            start_month = start.month
                            if start.day > 15:
                                start_month += 1
                            data[11] = str(start.year + (start_month == 13)) + '-' + str(
                                start_month == 13 or start_month)
                        data[11] = datetime.datetime.strptime(data[11], '%Y-%m').strftime('%Y/%m')
                        data[12] = str(end.year - (end.month == 1)) + '-' + str(12 if end.month == 1 else end.month - 1)
                        data[12] = datetime.datetime.strptime(data[12], '%Y-%m').strftime('%Y/%m')
                        if end.day > 15:
                            data[12] = end.strftime('%Y/%m')

                        data[13] = get_month(datetime.datetime.strptime(data[11], '%Y/%m'),
                                             datetime.datetime.strptime(data[12] + '/16', '%Y/%m/%d'))
                    data[7] = data[9]
                data.insert(9, attention_str)
                temp.append(data)
            i += 1
    num_data = numpy.array(temp)
    index = numpy.lexsort([num_data[:, 0], num_data[:, 3], num_data[:, 1], num_data[:, 8]])
    last_data = []
    last_data_2 = []
    row_data = []
    num1 = 0
    num2 = int(sys.argv[2])
    expense = ''
    for i in index:
        if last_data and last_data[0] != temp[i][0]:
            append_row(last_data, last_data_2, row_data, expense, num1, num2)
            row_data = []
            last_data_2 = []
            num1 = 0
            num2 += 1
        last_data = temp[i][:]
        expense = last_data[6]
        if last_data[4] not in row_data:
            row_data.append(last_data[4])
        if last_data[11] != '0':
            last_data_2.append(temp[i][:])
        if '租赁费' in temp[i][6] and (temp[i][7] == '0' or temp[i][7] == '0.00'):
            continue
        num1 += 1
        temp[i][18] = temp[i][17] = ''
        temp[i].insert(0, num1)
        temp[i].insert(1, num2)
        del (temp[i][12], temp[i][12])
        temp[i][14] = temp[i][12] = temp[i][13] = ''
        excelData.append(add_km_data(temp[i]))
    append_row(last_data, last_data_2, row_data, expense, num1, num2)


def sheet1():
    print('正在生成OA数据sheet')
    ga = excel.add_worksheet(time.strftime('%Y-%m'))
    title = ['序号1', '序号2', 'OA号', '申请人', '门店', '附件序号', '费用明细(分期请注明全额)', '费用明细(分期请注明全额)', '费用名称', '科目编码', '科目方向', '科目金额',
             '支付日期', '注意', '开始月份', '结束月份', '摊销月份', '收款单位名称', '接收时间', '备注', '客商编码', ':客商', '渠道', '', '租赁费', '物业费', '电费',
             '广告费', '服务费', '服务进项 - 总和费', '进项 - 拆分1AD2', '2AD2-AE2', '3', '4', '5', '6', '7', '8', '检验AD2-SUM(AE2:AL2)']
    excelData.insert(0, title)
    for i in range(0, len(excelData)):
        for x in range(0, len(excelData[i])):
            if i != 0 and (x == 0 or x == 1 or x == 11):
                ga.write_number(i, x, float(excelData[i][x]))
            else:
                ga.write(i, x, excelData[i][x])
    excel.close()
    print('OA数据处理成功')


def sub_str(old, month_flag=0):
    new = ''
    flag = 0
    for ch in old:
        if month_flag == 1 and (ch == '月'):
            ch = '-'
        if re.match(r'\d|\.|-', ch) is None:
            flag = 1
        if flag == 0 and ch != '':
            new += ch
    return new


def get_store_arr(store_name):
    ccb_str = '银行存款'
    if '平台支付' in store_name:
        ccb_str = '银行存款-建行'
        store_name = re.sub(r'平台支付|（|）|，|！|\(|\)|,|!', '', store_name)
    return [store_name, ccb_str]


def add_km_data(data_arr):
    km_list = [
        {
            'ykbc': '费用支出/租赁费/门店租赁费',
            'kmbm': '66011101',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '费用支出/物业管理费',
            'kmbm': '660112',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '费用支出/广告宣传费/业务宣传费',
            'kmbm': '66011802',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '资本性支出/押金、保证金',
            'kmbm': '122109',
            'fzhs1': ':客商',
            'fzhs2': ''
        },
        {
            'ykbc': '费用支出/服务费/运营服务费',
            'kmbm': '66012005',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '费用支出/能耗费',
            'kmbm': '66011301',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '其他应付款-其他',
            'kmbm': '224199',
            'fzhs1': '',
            'fzhs2': ''
        },
        {
            'ykbc': '费用支出/能耗费/电费',
            'kmbm': '66011301',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '银行存款',
            'kmbm': '1002',
            'fzhs1': '755939861610902:银行账户',
            'fzhs2': ''
        },
        {
            'ykbc': '银行存款-建行',
            'kmbm': '1002',
            'fzhs1': '1205 0161 5300 0000 2279:银行账户',
            'fzhs2': ''
        },
        {
            'ykbc': '费用支出／办公费／其他',
            'kmbm': '66010399',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '费用支出／清洁卫生费',
            'kmbm': '660105',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '费用支出／邮电费／网络通讯费',
            'kmbm': '66011003',
            'fzhs1': '100056:部门档案',
            'fzhs2': '00005:渠道'
        },
        {
            'ykbc': '预付',
            'kmbm': '112309',
            'fzhs1': ':客商',
            'fzhs2': ''
        },
        {
            'ykbc': '预付账款-单位往来',
            'kmbm': '112303',
            'fzhs1': ':客商',
            'fzhs2': ''
        }
    ]
    kmbm = ''
    fzhs1 = ''
    fzhs2 = ''
    for km in km_list:
        if km['ykbc'] == data_arr[8]:
            kmbm = km['kmbm']
            fzhs1 = km['fzhs1']
            fzhs2 = km['fzhs2']
    data_arr.insert(9, kmbm)
    kmfx = '贷' if kmbm == '1002' else '借'
    data_arr.insert(10, kmfx)
    data_arr.insert(20, '')
    fzhs1 = '04:部门档案' if fzhs1 == '100056:部门档案' and '扭蛋' in data_arr[4] else fzhs1
    data_arr.insert(21, fzhs1)
    data_arr.insert(22, fzhs2)
    return data_arr


def get_month(start, end):
    start_month = start.month
    end_month = end.month
    if start.day > 15:
        start_month += 1
    if end.day < 16:
        end_month -= 1
    return end.year * 12 + end_month - (start.year * 12 + start_month) + 1


def append_row(last_data, last_data_2, row_data, expense, num1, num2):
    if last_data_2:
        for l2 in last_data_2:
            num1 += 1
            excelData.append(add_km_data([num1, num2, l2[0], l2[1], l2[2], l2[3], l2[4], l2[5], '预付', l2[11], l2[8], '',
                                          l2[12], l2[13], l2[14], l2[15], l2[16], '']))
    num1 += 1
    excelData.append(add_km_data([num1, num2, last_data[0], last_data[1], last_data[2], last_data[3],
                                  ','.join(row_data), '付' + last_data[2] + ','.join(row_data), last_data[18],
                                  last_data[17], last_data[8], '', '', '', '', last_data[15],
                                  '保证金' if num1 == 2 and '押金' in expense else '未回', '']))


if __name__ == '__main__':
    if check_date() == 0 and check_num() == 0 and check_data() == 0:
        read()
        sheet1()
