# -*- coding: UTF-8 -*-

from __future__ import division
import xlrd
import os
import math
from xlwt import Workbook, Formula
import xlrd
import copy
import pandas
import time
import types
#
def is_chinese(uchar):
    """判断一个unicode是否是汉字"""
    if str(uchar) >= '/u4e00' and str(uchar) <= '/u9fa5':
        return True
    else:
        return False


def is_num(unum):
    try:
        unum + 1
    except TypeError:
        return 0
    else:
        return 1

# 不 带颜色的读取
def load_file(content):
    # 打开文件
    global workbook, file_excel
    file_excel = str(content)
    file = (file_excel + '.xls')  # 文件名及中文合理性
    if not os.path.exists(file):  # 判断文件是否存在
        file = (file_excel + '.xlsx')
        if not os.path.exists(file):
            print("文件不存在")
    workbook = xlrd.open_workbook(file)
    # print('suicce')

def load_file_with_twolist(file_name):

    file_list =[]
    load_file(file_name)

    Sheetname = workbook.sheet_names()
    for name in range(len(Sheetname)):

        table = workbook.sheets()[name]
        ttype = table.name
        nrows = table.nrows
        for n in range(nrows):
            # 获取每行内容
            a = table.row_values(n)
            mid = []
            for i in range(len(a)):

                if is_chinese(a[i]):
                    a[i].encode('utf-8')
                elif is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:
                        a[i] = int(a[i])

                mid.append(a[i])
            file_list.append(mid)
    return file_list




 # {'4001':{1:{name:cc1001,spec:30*50,sort:1,time:2018-1-2}}{2:{}}}
def mix_compose_mesg(self_made_compose_list):
    num = 0
    self_made_compose_dist = {}
    for x in self_made_compose_list :
        if isinstance(x[1],str) and '代' not in x[1]   and '代' not in x[2] or isinstance(x[1], int):
            if x[0] not in self_made_compose_dist:
                self_made_compose_dist.setdefault(x[0], {})
                self_made_compose_dist[x[0]].setdefault(num, {})
                self_made_compose_dist[x[0]][num].setdefault('name', x[1])
                self_made_compose_dist[x[0]][num].setdefault('spec', x[2])
                self_made_compose_dist[x[0]][num].setdefault('sort', 1)
                self_made_compose_dist[x[0]][num].setdefault('time', x[3])
                num = num +1
            else:
                key_list = self_made_compose_dist[x[0]].keys()
                for y in key_list:
                    num = 1
                    flag = 0
                    if x[1] in self_made_compose_dist[x[0]][y].values() and x[2]  in self_made_compose_dist[x[0]][y].values():
                        benci_shul = x[3]
                        if benci_shul == '':
                            benci_shul = 0
                        # print (self_made_compose_dist[x[0]][y]['time'],benci_shul,x[0])
                        if self_made_compose_dist[x[0]][y]['time'] > benci_shul:
                            self_made_compose_dist[x[0]][y]['time'] = benci_shul
                            self_made_compose_dist[x[0]][y]['sort'] += 1
                        elif self_made_compose_dist[x[0]][y]['time'] == benci_shul:
                            pass
                        else:
                            self_made_compose_dist[x[0]][y]['sort'] += 1
                        flag = 1
                        break
                if flag == 0:
                    self_made_compose_dist[x[0]].setdefault(num, {})
                    self_made_compose_dist[x[0]][num].setdefault('name', x[1])
                    self_made_compose_dist[x[0]][num].setdefault('spec', x[2])
                    self_made_compose_dist[x[0]][num].setdefault('sort', 1)
                    self_made_compose_dist[x[0]][num].setdefault('time', x[3])


    return self_made_compose_dist




def output_dict(self_made_compose_dist):
    book = Workbook()
    sheet2 = book.add_sheet(u'自制件')

    i = 0
    line = 0
    for key, value in self_made_compose_dist.items():
        for s, d in value.items():
            sheet2.write(i, line, key)
            sheet2.write(i, line+1, d['name'])
            sheet2.write(i, line+2, d['spec'])
            sheet2.write(i, line+3, d['sort'])



            if isinstance(d['time'],str):
                timetime=''
            else:
                cccc = pandas.to_datetime(d['time'] - 25569, unit='d')
                timetime = str(pandas.Period(cccc, freq='D'))
            sheet2.write(i, line+4,timetime)
            line = line + 5
        line = 0
        i = i + 1


    book.save('3.xls')  # 存储excel
    book = xlrd.open_workbook('3.xls')
    print('----------------------------------------------------------------------------------------')
    print('----------------------------------------------------------------------------------------')
    print(u'计算完成')

    print('----------------------------------------------------------------------------------------')

    print('----------------------------------------------------------------------------------------')

    time.sleep(10)




if __name__ == "__main__":
    self_made_compose_list = load_file_with_twolist('used_maopi')
    self_made_compose_dist = mix_compose_mesg(self_made_compose_list)
    output_dict(self_made_compose_dist)