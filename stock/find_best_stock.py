import sys
import os
from my_package.excel.excel_rw import *

# 程序功能：
# 本程序用于从excel表中筛选出符合要求的基金。
# 筛选规则是：
# 假设我们有一个基金列表，这些基金都是至少成立3年的。我们在这些基金中，把
# 前3年业绩总体排名前1/10的选出来，再把前2年总体排名前1/10选出来，再把前1年
# 总体排名前1/10选出来，看看有没有基金在这3个排名中都出现，如果都出现，就是
# 满足筛选要求的，立即停止筛选，输出基金名称，否则继续把前1年排名放宽到前1/5,
# 然后再看有没有基金出现在新的这3个排名表中，如果出现，就是满足筛选要求的，立
# 即停止筛选, 以此类推. 伪代码如下:
# for n4 in [3年前1/10, 3年前1/5, 3年前1/3, 3年前1/2]:
#    for n3 in [2年前1/10, 2年前1/5, 2年前1/3, 2年前1/2]:
#        for n2 in [1年前1/10, 1年前1/5, 1年前1/3, 1年前1/2]:
#           for n1 in [半年前1/10, 半年前1/5, 半年前1/3, 半年前1/2]:
#              if (n3 & n2 & n1 & n4) != None:
#                 break;
#
# usage:
# find_best_stock.py 股票型_3y.xls
# 其中，股票型_3y.xls 晨星网上按3年业绩排序的股票基金列表，注意只复制下来有3年业绩的，然后粘贴到excel表格中.
# 如果要筛选其他类型，只需要在晨星网上选其他类型基金来排序即可.
# 执行后,屏幕输出筛选结果.

# 如果要把5年回报率也考虑进去,先修改return_rate_col，然后执行命令：
# usage find_best_stock.py 股票型_5y.xls
# 其中，股票型_5y.xls 晨星网上按3年业绩排序的股票基金列表，注意只复制下来有5年业绩的，然后粘贴到excel表格中


def read_col(file: str, col: int) -> [str]:
    reader = ExcelReader()
    reader.open(file, 0)
    reader.seek(3)
    return reader.read_col(col)


def output_funds_pool(funds_pool):
    for e1 in funds_pool:
        for e2 in e1:
            print(e2)
            print("--------------------------------------------")
        print("================================================")


def find_best(intersection, it):
    try:
        l = next(it)
    except StopIteration:
        print(intersection)
        exit()
    for f in l:
        intersection_new = intersection & set(f)
        if len(intersection_new) == 0:
            continue
        else:
            find_best(intersection_new, it)

file = sys.argv[2]
year = sys.argv[1]
if year == '3':        
    return_rate_col = (10, 9, 8) # 依次是'3年回报', '2年回报', '1年回报', '半年回报','3个月回报','1个月回报' 所在列数,从0开始
elif year == '5':
    return_rate_col = (11, 10, 9, 8) # 依次是'5年回报','3年回报', '2年回报', '1年回报', '半年回报','3个月回报','1个月回报' 所在列数
else:
    sys.exit(0)	
# 某年回报率前1/10, 1/5, 1/3, 1/2
# topn = [10, 5, 3, 2]
topn = [4]
# names = read_col(file, title['基金名称'])
funds_pool = []
# 读取'基金名称'列
# names = ['诺安上证新兴产业ETF','国联安上证商品ETF联接',....]
names = read_col(file, 2)
for rt in return_rate_col:
    # returns_rate = [11.7, 23.9, ....]
    returns_rate = read_col(file, rt)
    # print(returns_rate)
    # funds = {'诺安上证新兴产业ETF':11.7, '国联安上证商品ETF联接':23.9}
    funds = dict(zip(names, returns_rate))
    # print(funds.items())
    funds_sorted = sorted(funds.items(), key=lambda item:item[1], reverse=True)
    # print(funds_sorted)
    funds_name_sorted = []
    for i in range(len(funds_sorted)):
        # funds_name_sorted =
        funds_name_sorted.append(funds_sorted[i][0])
    # print(funds_name_sorted)
    funds_top = []
    for n in topn:
        end = int(len(funds_name_sorted)/n)
        # funds_top = [[n年回报率前1/10]，[n年回报率前1/5]，[n年回报率前1/3]，[n年回报率前1/2]]
        funds_top.append(funds_name_sorted[0:end])
    # funds_pool =
    # [[[3年回报率前1/10]，[3年回报率前1/5]，[3年回报率前1/3]，[3年回报率前1/2]],
    # [[2年回报率前1/10]，[2年回报率前1/5]，[2年回报率前1/3]，[2年回报率前1/2]],
    # [[1年回报率前1/10]，[1年回报率前1/5]，[1年回报率前1/3]，[1年回报率前1/2]],
    # [[0.5年回报率前1/10]，[0.5年回报率前1/5]，[0.5年回报率前1/3]，[0.5年回报率前1/2]]]
    funds_pool.append(funds_top)

    it = iter(funds_pool)
    l = next(it)
for three in l:
    find_best(set(three), it)

