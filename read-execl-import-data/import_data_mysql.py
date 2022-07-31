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

def importExcelToMysql(cur, path):
    ### openpyxl版本
    # 读取excel文件
    workbook = load_workbook(path)
    # 获得所有工作表的名字
    sheets = workbook.get_sheet_names()
    # 获得第一张表
    worksheet = workbook.get_sheet_by_name(sheets[0])

    # 将表中每一行数据读到 sqlstr 数组中
    for row in worksheet.rows:

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
        cur.execute("insert into scphci_datacenter(id, name, address, phone, domain, comment, date_created, date_updated, created_by) values(%s, %s, 'null', 'null', 'null', '新增',%s, %s, 'admin')", (suid, valuestr, now, now))


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

    # 将 excel 中的数据导入 数据库中
    try:
        importExcelToMysql(cur, "./files/List.xlsx")
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
