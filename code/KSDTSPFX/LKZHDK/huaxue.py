import numpy as np
import math
import pandas as  pd
import pymysql
import os
import matplotlib.pyplot  as plt
import decimal
import cx_Oracle
import matplotlib.ticker as ticker
import openpyxl

class DTFX:
    def __init__(self):
        self.db = cx_Oracle.connect('gkeva2020/ksy#2020#reta@10.0.200.103/ksydb01std')
        self.cursor = self.db.cursor()

    def __del__(self):
        self.cursor.close()
        self.db.close()

    def set_list_precision(self,L):
        for i in range(len(L)):
            if isinstance(L[i], float) or isinstance(L[i],decimal.Decimal):
                L[i] = format(L[i],'.2f')

    # 市级报告 总体概括 表格
    def ZTGK_CITY_TABLE(self, dsh):
        sql = ""
        sql = "select mc from c_ds where DS_H = " + dsh
        self.cursor.execute(sql)
        ds_mc = self.cursor.fetchone()[0]

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + ds_mc
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析总体概括(化学).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = "select count(jmx.ksh) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh " \
              "where jmx.tzh=6 and jmx.kmh = 005 and jbxx.ds_h=" + dsh
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]
        print(num)

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jbxx.xb_h=1 and jmx.kmh = 005 and jmx.tzh=6 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jbxx.xb_h=1 and jmx.kmh = 005 and jmx.tzh=6"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jbxx.xb_h=2 and jmx.kmh = 005 " \
              r"and jmx.tzh=6 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where  jbxx.xb_h=2 and jmx.kmh = 005 and jmx.tzh=6"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=1 or jbxx.kslb_h=3) " \
              r" and jmx.kmh = 005 and jmx.tzh=6 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where " \
              r"(jbxx.kslb_h=1 or jbxx.kslb_h=3) and jmx.tzh=6 and jmx.kmh = 005"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=2 or jbxx.kslb_h=4)" \
              r" and jmx.kmh = 005 and jmx.tzh=6 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where " \
              r"(jbxx.kslb_h=2 or jbxx.kslb_h=4) and jmx.kmh = 005 and jmx.tzh=6"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=1 or jbxx.kslb_h=2) " \
              r"and jmx.tzh=6 and jmx.kmh = 005 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where " \
              r"(jbxx.kslb_h=1 or jbxx.kslb_h=2) and jmx.kmh = 005 and jmx.tzh=6"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=3 or jbxx.kslb_h=4) and jmx.kmh = 005 " \
              r"and jmx.tzh=6 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=3 or jbxx.kslb_h=4) and jmx.kmh = 005 and jmx.tzh=6"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jmx.tzh=6 and jmx.kmh = 005 and jbxx.ds_h=" + dsh

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        sql = r"select avg(jmx.zf)from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join GKEVA2020.jbxx on" \
              r" jbxx.ksh=jmx.ksh where jmx.tzh=6 and jmx.kmh = 005"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(writer, sheet_name="各类别考生成绩比较(化学)", index=None)

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from GKEVA2020.c_xq where xq_h like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        # 全省
        sql = "select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.tzh=6 and kmh = 005"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 110)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 全市
        sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"where jmx.tzh=6 and kmh = 005 and  KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 110)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.tzh=6 and kmh = 005 and  KSH LIKE '" + str(xqh[0]) + r"%'"
            self.cursor.execute(sql)
            print(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
            result.append(result[1] / 110)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区考生成绩比较(化学)", index=None)
        writer.save()

    # 市级报告 总体概括 图
    def ZTGK_CITY_IMG(self, dsh):
        sql = ""
        sql = "select mc from c_ds where DS_H=" + dsh
        self.cursor.execute(sql)
        ds_mc = self.cursor.fetchone()[0]

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + ds_mc
        if not os.path.exists(path):
            os.makedirs(path)

        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        # 全省
        sql = "select count(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.tzh=6 and kmh = 005"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select zf,count(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.tzh=6 and kmh = 005  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())

        province = [None] * 101
        for item in items:
            province[int(item[0])] =item[1] / num * 100
        x = list(range(101))
        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市
        sql = r"select count(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.tzh=6 and kmh = 005 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select zf,count(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.tzh=6 " \
              "and kmh = 005  and ksh like '" + dsh + r"%'GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 101

        for item in items:
            city[int(item[0])] =item[1] / num * 100

        x = list(range(101))

        plt.plot(x, city, color='springgreen', marker='.', label='全市')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\地市及全省考生单科成绩分布(化学).png', dpi=1200)
        plt.close()

    # 省级报告 原始分概括 表
    def YSFGK_PROVICNE_TABLE(self):

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析原始分概括(化学).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(jmx.ksh) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh " \
              "where jmx.tzh=6 and jmx.kmh = 005"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jbxx.xb_h=1 and jmx.kmh = 005 and jmx.tzh=6 "

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100))  # 差异系数

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jbxx.xb_h=2 and jmx.kmh = 005 " \
              r"and jmx.tzh=6 "

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100))  # 差异系数

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=1 or jbxx.kslb_h=3) " \
              r" and jmx.kmh = 005 and jmx.tzh=6"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100))  # 差异系数

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=2 or jbxx.kslb_h=4)" \
              r" and jmx.kmh = 005 and jmx.tzh=6 "

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100))  # 差异系数

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=1 or jbxx.kslb_h=2) " \
              r"and jmx.tzh=6 and jmx.kmh = 005 "

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100))  # 差异系数

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where (jbxx.kslb_h=3 or jbxx.kslb_h=4) and jmx.kmh = 005 " \
              r"and jmx.tzh=6 "

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100))  # 差异系数

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.jbxx on jbxx.ksh=jmx.ksh where jmx.tzh=6 and jmx.kmh = 005 "

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(writer, sheet_name="各类别考生成绩比较(化学)", index=None)
        writer.save()

    # 市级报告 单题分析 表
    def DTFX_CITY_TABLE(self,dsh):
        sql = ""
        sql = "select mc from c_ds where DS_H = " + dsh
        self.cursor.execute(sql)
        ds_mc = self.cursor.fetchone()[0]

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + ds_mc
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(化学).xlsx")

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=005 " \
              r"and jmx.tzh=6 and jmx.ksh like '"+dsh+r"%'"
        self.cursor.execute(sql)
        total =  self.cursor.fetchone()[0]

        low = int(total/3)
        high = int(total/1.5)

        df = pd.DataFrame(data=None,columns=["题号","分值","本市平均分","全省平均分","本市得分率","高分组得分率","中间组得分率","低分组得分率"])

        idxs = list(range(7, 14))
        # for idx in idxs:
        #     row = []
        #     if idx<10:
        #         row.append("0"+str(idx))
        #     else:
        #         row.append(str(idx))
        #
        #     num = 6.00
        #     row.append(num)
        #
        #     sql = r"SELECT avg(kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx where ksh like '"+dsh+r"%' and kmh = 005 and idx = "+str(idx)
        #     self.cursor.execute(sql)
        #     row.append(self.cursor.fetchone()[0])
        #
        #     sql = r"SELECT avg(kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx where  kmh = 005 and idx = " + str(idx)
        #     self.cursor.execute(sql)
        #     row.append(self.cursor.fetchone()[0])
        #
        #     row.append(row[2]/num)
        #
        #     sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
        #           r"right join (select a.*,rownum rn from (select jmx.ksh from " \
        #           r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '"+dsh+r"%' and jmx.kmh=005 " \
        #           r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
        #           r"where b.rn between 1 and "+str(low)+r" and amx.kmh=005 and amx.idx="+str(idx)
        #     self.cursor.execute(sql)
        #     row.append(self.cursor.fetchone()[0]/low/num)
        #
        #     sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
        #           r"right join (select a.*,rownum rn from (select jmx.ksh from " \
        #           r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '" + dsh + r"%' and jmx.kmh=005 " \
        #           r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
        #           r"where b.rn between "+str(low+1)+r" and " + str(high) + r" and amx.kmh=005 and amx.idx="+str(idx)
        #     self.cursor.execute(sql)
        #     row.append(self.cursor.fetchone()[0] / (high-low)/num)
        #
        #     sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
        #           r"right join (select a.*,rownum rn from (select jmx.ksh from " \
        #           r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '" + dsh + r"%' and jmx.kmh=005 " \
        #           r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
        #           r"where b.rn between "+str(high+1)+r" and " + str(total) + " and amx.kmh=005 and amx.idx="+str(idx)
        #     self.cursor.execute(sql)
        #     row.append(self.cursor.fetchone()[0] / (total-high) / num)
        #
        #     self.set_list_precision(row)
        #     df.loc[len(df)] = row

        dths = [26,27,28,35,36]
        for dth in dths:
            row = []
            row.append(str(dth))
            if dth == 26 or dth ==28:
                num = 14.00
            elif dth in [27,35,36]:
                num = 15.00
            else:
                num = 10.00
            row.append(num)

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=005 and jmx.tzh="+str(dth)+" and ksh like '"+dsh+r"%'"
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=005 and jmx.tzh="+str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            row.append(row[2]/num)

            sql = r"select sum(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.kmh=005 and jmx.ksh like '"+dsh+"%' and jmx.tzh="+str(dth)+" ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN 1 and "+str(low)+r" and c.kmh=005 and c.tzh = "+str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone[0]/low/num)

            sql = r"select sum(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.kmh=005 and jmx.ksh like '" + dsh + "%' and jmx.tzh="+str(dth)+" ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN "+str(low+1)+" and " + str(high) + r" and c.kmh=005 and c.tzh = " +str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone[0] / (high-low) / num)

            sql = r"select sum(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.kmh=005 and jmx.ksh like '" + dsh + "%' and jmx.tzh="+str(dth)+" ORDER BY jmx.zf desc) a) b " \
                r"on c.ksh=b.ksh where b.rn BETWEEN "+str(high+1)+" and " + str(total) + r" and c.kmh=005 and c.tzh = " +str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone[0] / (total-high) / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer,sheet_name="地市考生单题作答情况(化学)",index=None)
        writer.save()

    # 市级报告 单题分析 画图
    # def DTFX_CITY_IMG(self,dsh):
    #     sql = "select mc from c_ds where DS_H = " + dsh
    #     self.cursor.execute(sql)
    #     ds_mc = self.cursor.fetchone()[0]
    #
    #     pwd = os.getcwd()
    #     father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
    #     path = father_path + r"\考生答题分析"
    #
    #     if not os.path.exists(path):
    #         os.makedirs(path)
    #     path = path + "\\" + ds_mc
    #     if not os.path.exists(path):
    #         os.makedirs(path)
    #
    #     sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=005 and jmx.ksh like '"+dsh+r"%'"
    #     self.cursor.execute(sql)
    #     total = self.cursor.fetchone()[0]
    #     ph_num = int(total * 0.27)
    #
    #     idxs = list(range(7, 14))
    #     dths = [26,27,28,35,36]
    #     txt = idxs + dths
    #
    #     x = [] # 难度
    #     y = [] # 区分度
    #
    #     for idx in idxs:
    #         num = 6.0
    #         sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where amx.ksh like '"+dsh+"%' and kmh = 005 and idx = " + str(idx)
    #         self.cursor.execute(sql)
    #         difficulty = self.cursor.fetchone()[0] / total / num #难度
    #
    #         # 前27%得分率
    #         sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
    #               r"(select ksh,yw from (select ksh,yw,rownum rn from " \
    #               r"(select ksh,yw from kscj where ksh like '"+dsh+"%' ORDER BY yw desc) a ) b " \
    #               r"where b.rn BETWEEN 1 and "+str(ph_num)+") c on amx.ksh = c.ksh where amx.kmh = 005 and amx.idx = "+str(idx)
    #         print(sql)
    #         self.cursor.execute(sql)
    #         ph = self.cursor.fetchone()[0] / ph_num / num
    #
    #         # 后27%得分率
    #         sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
    #               r"(select ksh,yw from (select ksh,yw,rownum rn from " \
    #               r"(select ksh,yw from kscj where ksh like '" + dsh + "%' ORDER BY yw desc) a ) b " \
    #               r"where b.rn BETWEEN "+str(total-ph_num)+r" and " + str(total) + r") c on amx.ksh = c.ksh where amx.kmh = 005 and amx.idx = " + str(idx)
    #         print(sql)
    #         self.cursor.execute(sql)
    #         pl = self.cursor.fetchone()[0] / (total-ph_num) / num
    #
    #         x.append(difficulty)
    #         y.append(ph-pl)
    #
    #     for dth in dths:
    #         if dth == 26 or dth == 28:
    #             num = 14.00
    #         elif dth in [27, 35, 36]:
    #             num = 15.00
    #         else:
    #             num = 10.00
    #
    #         sql = r"select sum(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where  jmx.kmh=005 and jmx.ksh like '"+dsh+r"%' and jmx.tzh="+str(dth)
    #         self.cursor.execute(sql)
    #         print(sql)
    #         difficulty = self.cursor.fetchone()[0] / total / num # 难度
    #         x.append(difficulty)
    #
    #         sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
    #               r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where " \
    #               r"jmx.kmh=005 and jmx.tzh = 6 and jmx.ksh like '"+dsh+r"%') b on a.ksh=b.ksh where a.kmh=005 and a.tzh="+str(dth)
    #         self.cursor.execute(sql)
    #         result = np.array(self.cursor.fetchall(), dtype="float64")
    #
    #         xt_score = np.array(result[:, 0], dtype="float64")
    #         zf_score = np.array(result[:, 1], dtype="float64")
    #
    #         n = len(xt_score)
    #
    #         D_a = n * np.sum(xt_score * zf_score)
    #         D_b = np.sum(zf_score) * np.sum(xt_score)
    #         D_c = n * np.sum(xt_score**2) - np.sum(xt_score)**2
    #         D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score)**2
    #
    #
    #         qfd = (D_a-D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
    #         y.append(qfd)
    #         print(x)
    #         print(y)
    #
    #
    #     print(x,y)
    #     plt.rcParams['figure.figsize'] = (15.0,6.0)
    #     plt.scatter(x,y)
    #     plt.xlim((0, 1))
    #     plt.ylim((0, 1))
    #     plt.xlabel("难度")
    #     plt.ylabel("区分度")
    #     ax = plt.gca()
    #     ax.spines['right'].set_color('none')
    #     ax.spines['top'].set_color('none')
    #     ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
    #     ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
    #     th = []
    #     for i in range(len(x)):
    #         th.append(txt[i])
    #         plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] , y[i] + 0.008),arrowprops=dict(arrowstyle='-'))
    #     plt.savefig(path + '\\各题难度-区分度分布散点图(化学).png', dpi=1200)
    #     plt.close()