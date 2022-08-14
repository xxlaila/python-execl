#!/usr/bin/env python
# -*- coding:   utf-8 -*-
"""
@Time     :   2022/7/31 14:50
@Author   :   xxlaila
@File     :   import_data_mysql.py
"""
import pymysql
from openpyxl.reader.excel import load_workbook
from builtins import int
import uuid
from datetime import datetime

def ReadExeclFiles(path):
    ### openpyxl版本
    # 读取excel文件
    workbook = load_workbook(path)
    # 获得所有工作表的名字
    sheets = workbook.get_sheet_names()
    # 获得第一张表
    worksheet = workbook.get_sheet_by_name(sheets[0])

    return worksheet

def import_database(cur, data):

    # 将表中每一行数据读到 sqlstr 数组中
    for row in data.rows:

        sqlstr = []

        for cell in row:
            sqlstr.append(cell.value)

        uid = str(uuid.uuid4())
        suid = ''.join(uid.split('-'))
        now = datetime.now()
        ###
        # 将每行数据存到数据库中
        ## 数据中心
        valuestr = str(sqlstr[0])
        old_data = cur.execute("select * from scphci_datacenter where name='%s' " % valuestr)
        if int(old_data) == 0:
            cur.execute("insert into scphci_datacenter(id, name, address, phone, domain, comment, date_created, date_updated, created_by) values(%s, %s, 'null', 'null', 'null', '新增',%s, %s, 'admin')", (suid, valuestr, now, now))
        else:
            cur.execute("update scphci_datacenter set address='%s', date_updated='%s' where name='%s'" % (str(sqlstr[1]), now, valuestr))


# 输出数据库中内容
def readTable(cursor):
    # 选择全部
    cursor.execute("select * from scphci_datacenter")
    # 获得返回值，返回多条记录，若没有结果则返回()
    results = cursor.fetchall()

    for i in range(0, results.__len__()):
        for j in range(0, 4):
            print(results[i][j], end='\t')

        print('\n')


if __name__ == '__main__':
    conn = pymysql.connect(host='1.1.1.1', user='root', passwd='123456', port=3306, charset='utf8')
    # 创建游标链接
    cur = conn.cursor()
    # 新建一个database
    # 选择 students 这个数据库
    cur.execute("use sre")
    execl_data = ReadExeclFiles("./files/List.xlsx")
    # 将 excel 中的数据导入 数据库中
    try:
        import_database(cur, execl_data)
        readTable(cur)
    except Exception as e:
        # 发生错误时回滚
        conn.rollback()
        print(str(e))
    else:
        # 关闭游标链接
        cur.close()
        conn.commit() # 事务提交
        print('事务处理成功')
    # 关闭数据库服务器连接，释放内存
    conn.close()
