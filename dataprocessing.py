import xlrd

data1 = xlrd.open_workbook('2020年10月.xlsx')
table1 = data1.sheets()[0]
data2 = xlrd.open_workbook('2020年11月.xlsx')
table2 = data2.sheets()[0]
data3 = xlrd.open_workbook('2020年12月.xlsx')
table3 = data3.sheets()[0]

nrows = table1.nrows  # 获取行数
ncols = table1.ncols  # 获取列数


# 提取表格中单元格中的打卡数据，并存储在列表a中
def getcontent1():
    a = []
    b = []
    c = []

    for i in range(4, 16):
        b = []
        for j in range(6, 37):
            # print('-'*10)
            c = table1.cell(i, j).value
            b.append(c)
            # print('-' * 10)
            # print(b)

        a.append(b)
    return a


def getcontent2():
    a = []
    b = []
    c = []

    for i in range(4, 16):
        b = []
        for j in range(6, 36):
            # print('-'*10)
            c = table2.cell(i, j).value
            b.append(c)
            # print('-' * 10)
            # print(b)

        a.append(b)
    return a


def getcontent3():
    a = []
    b = []
    c = []

    for i in range(4, 16):
        b = []
        for j in range(6, 37):
            # print('-'*10)
            c = table3.cell(i, j).value
            b.append(c)
            # print('-' * 10)
            # print(b)

        a.append(b)
    return a


# w = getcontent()
#
# print(w[0][1])


#   处理每个单元格中的数据，去掉换行，拆分时和分，并存储在列表a中
def dealdata(string):
    c = string.replace('  \n', ":")
    c = c.split(':')

    flag = 0
    a = []
    b = []
    i = 0
    while i < len(c):
        b.append(c[i])
        if i % 2 != 0:
            a.append(b)
            b = []
        i = i + 1

    return a


# print(dealdata(w[0][1]))


#  整理数据，提取出每个时段中最早到达时间和最晚离开时间
def simplify(aray):
    j = 0
    a = []
    l = len(aray) - 1

    # 判断是否存在18:20这个特殊的时间点，如果存在的话，看他和哪个时间段闭合
    flag = 0
    flag1 = 0
    flag2 = 0
    for i in range(len(aray)):
        if int(aray[i][0]) * 60 + int(aray[i][1]) == 1100:
            flag = 1
        if 780 <= int(aray[i][0]) * 60 + int(aray[i][1]) <= 945:
            flag1 = 1
        if 1200 <= int(aray[i][0]) * 60 + int(aray[i][1]) <= 1350:
            flag2 = 1

    while j < l:
        # print('-'*10)
        # print(aray)
        # print(l)
        # print(j)
        if 390 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 645 and 390 <= int(aray[j + 1][0]) * 60 + int(
                aray[j + 1][1]) <= 645:
            aray.pop(j + 1)
            j = 0
            l = l - 1
            continue
        if 660 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 750 and 660 <= int(aray[j + 1][0]) * 60 + int(
                aray[j + 1][1]) <= 750:
            aray.pop(j)
            j = 0
            l = l - 1
            continue
        if 780 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 945 and 780 <= int(aray[j + 1][0]) * 60 + int(
                aray[j + 1][1]) <= 945:
            aray.pop(j + 1)
            j = 0
            l = l - 1
            continue
        if 1020 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 1100 and 1020 <= int(aray[j + 1][0]) * 60 + int(
                aray[j + 1][1]) <= 1100:
            aray.pop(j)
            j = 0
            l = l - 1
            continue
        if 1100 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 1170 and 1100 <= int(aray[j + 1][0]) * 60 + int(
                aray[j + 1][1]) <= 1170:
            aray.pop(j + 1)
            j = 0
            l = l - 1
            continue
        if 1200 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 1350 and 1200 <= int(aray[j + 1][0]) * 60 + int(
                aray[j + 1][1]) <= 1350:
            aray.pop(j)
            j = 0
            l = l - 1
            continue

        j = j + 1

    return aray, flag, flag1, flag2


# 计算每个人每天的工作时间
def calday(b):
    # a = getcontent()
    # b = dealdata(a[0][0])
    aray, flag, flag1, flag2 = simplify(b)
    time1 = 0
    time2 = 0
    time3 = 0
    time4 = 0
    time5 = 0
    time6 = 0
    for j in range(len(aray)):
        if 390 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 645:
            time1 = int(aray[j][0]) * 60 + int(aray[j][1])
        if 660 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 750:
            time2 = int(aray[j][0]) * 60 + int(aray[j][1])
        if 780 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 945:
            time3 = int(aray[j][0]) * 60 + int(aray[j][1])
        if flag == 1 and flag1 == 1:
            time4 = 1100
        if flag == 0 and 1020 <= int(aray[j][0]) * 60 + int(aray[j][1]) < 1100:
            time4 = int(aray[j][0]) * 60 + int(aray[j][1])
        if flag == 1 and flag2 == 1:
            time5 = 1100
        if flag == 0 and 1100 < int(aray[j][0]) * 60 + int(aray[j][1]) < 1170:
            time5 = int(aray[j][0]) * 60 + int(aray[j][1])
        if 1200 <= int(aray[j][0]) * 60 + int(aray[j][1]) <= 1350:
            time6 = int(aray[j][0]) * 60 + int(aray[j][1])
    if time1 > 0 and time2 > 0:
        sum1 = time2 - time1
    else:
        sum1 = 0
    if time3 > 0 and time4 > 0:
        sum2 = time4 - time3
    else:
        sum2 = 0
    if time5 > 0 and time6 > 0:
        sum3 = time6 - time5
    else:
        sum3 = 0

    sum = sum1 + sum2 + sum3
    # print("time1:%d" % time1)
    # print("time2:%d" % time2)
    # print("time3:%d" % time3)
    # print("time4:%d" % time4)
    # print("time5:%d" % time5)
    # print("time6:%d" % time6)
    # print(sum)
    return sum


# 打印每个人每天的工作时间
def calmon(fun):
    list = []
    # a = getcontent1()
    a = fun
    for i in range(len(a)):
        # print("-" * 20)
        sum = 0
        for j in range(len(a[i])):
            b = dealdata(a[i][j])
            c = calday(b)
            sum = sum + c
        list.append(sum)
    # print(list)
    return list


def min2hour(list):
    newlist = []
    for i in range(len(list)):
        a = ''
        c = list[i]%60
        b = (list[i]-c)//60
        a = str(b)+"小时"+str(c)+"分钟"
        newlist.append(a)
    return newlist


list1 = calmon(getcontent1())
slist1 = min2hour(list1)
list2 = calmon(getcontent2())
slist2 = min2hour(list2)
list3 = calmon(getcontent3())
slist3 = min2hour(list3)
# print(list1)
# print(list2)
# print(list3)
