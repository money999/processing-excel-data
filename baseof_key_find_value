import re, openpyxl
'''
在杂乱的excel中找到key，根据偏移量获得key对应的value，重复的key对应的value求和（这一步其实是数据表透视）
'''

# 字典按key排序
def sorted_dict(adict):
    keys = list(adict.keys())
    keys.sort()
    return dict(zip(keys, map(adict.get, keys)))

# 获得全表数据
def get_sheet_all_data(sheet):
    ws = sheet
    return ws['A1':'%s%d' % (openpyxl.utils.get_column_letter(ws.max_column), ws.max_row)]

# 根据正则式挑选key
# 目标字典，正则解析，工作表，正则匹配后的value相对key的水平偏移量
def statistic_data(dic, reg, sheet, offset):
    for i in range(1, sheet.max_row+1):
        for j in range(1, sheet.max_column+1):
            if sheet.cell(row=i, column=j).value:
                res = re.search(reg, str(sheet.cell(row=i, column=j).value))
                if res:
                    if dic.get(res.group()):
                        dic[res.group()].append(sheet.cell(row=i, column=j+offset).value)
                    else:
                        dic[res.group()] = [sheet.cell(row=i, column=j+offset).value]


#col表示列项下标从1开始！！！
def all_data_poly(dic, polydic, col):
    maxcol = 6 #需要多少列项
    for i in polydic:
        if dic.get(i) == None:
            dic[i] = [x * 0 for x in range(6)]
        dic[i][col + 1] = polydic[i]
    return dic

if __name__ == '__main__':
    # wb = openpyxl.Workbook()
    # ws = wb.create_sheet(0)
    # ws.title = 'hello'
    # ws['A6'] = 553
    # ws.cell(1,1,789)
    # wb.save('lll.xlsx')

    wb = openpyxl.load_workbook('2017年公务生产车V1.0.xlsx')
    print(wb.sheetnames)
    reg = re.compile(r'闽\w{6}', re.I)
    dic = {};

    ws = wb.get_sheet_by_name(wb.sheetnames[3])
    statistic_data(dic, reg, ws, 1)
    ws = wb.get_sheet_by_name(wb.sheetnames[5])
    statistic_data(dic, reg, ws, 2)
    print(dic)

    dic = sorted_dict(dic)
    print(dic)

    for i in dic:
        dic[i] = sum(dic[i])

    print(dic)
    # for i in dic:
    #     print(i[0], sum(i[1]))
    # for i in dic:
    #     print(i)

    # for tup in ws['A1':'C6']:
    #     for i in tup:
    #         print(i.value, end=' ')

    #print(wb.get_sheet_names())
