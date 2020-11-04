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

    def set_list_precision(self, L):
        for i in range(len(L)):
            if isinstance(L[i], float) or isinstance(L[i], decimal.Decimal):
                L[i] = format(L[i], '.2f')

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析总体概括(政治).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]
        print(num)

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and xb_h=1) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and xb_h=1) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and xb_h=2) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and xb_h=2) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=1 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and (jbxx.kslb_h=1 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=2 or jbxx.kslb_h=4)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and (jbxx.kslb_h=2 or jbxx.kslb_h=4)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=1 or jbxx.kslb_h=2)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and (jbxx.kslb_h=1 or jbxx.kslb_h=2)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=4 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and (jbxx.kslb_h=4 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where jbxx.ds_h=" + dsh + " ) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

            
        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(writer, sheet_name="各类别考生成绩比较(政治)", index=None)

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from GKEVA2020.c_xq where xq_h like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        # 全省
        sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=6 and kscj.zh!=0 and jmx.kmh = 006 "
        print(sql)
        self.cursor.execute(sql)

        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 100)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 全市
        sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=6 and kscj.zh!=0  and jmx.kmh = 006 and  jmx.KSH LIKE '" + dsh + "%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 100)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            print(xqh)
            sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
                  r"jmx right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=6" \
                  r" and kscj.zh!=0  and jmx.kmh = 006 and  jmx.KSH LIKE '" + xqh[0] + r"%'"
            self.cursor.execute(sql)
            print(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
            result.append(result[1] / 100)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区考生成绩比较(政治)", index=None)
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
        sql = "select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and kmh = 006"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and kmh = 006  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())

        province = [None] * 101
        for item in items:
            print(item[1] / num * 100)
            print(item[0])
            province[int(item[0])] =item[1] / num * 100
        x = list(range(101))
        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市
        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and kmh = 006 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 " \
              "and kmh = 006  and ksh like '" + dsh + r"%'GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 101

        for item in items:
            city[int(item[0])] =item[1] / num * 100

        x = list(range(101))

        plt.plot(x, city, color='springgreen', marker='.', label='全市')
        plt.xlim((0, 100))
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\地市及全省考生单科成绩分布(政治).png', dpi=1200)
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

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析原始分概括(政治).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 ) b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]
        print(num)

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and xb_h=1) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and xb_h=2) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0  and (jbxx.kslb_h=1 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and  (jbxx.kslb_h=2 or jbxx.kslb_h=4)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0  and (jbxx.kslb_h=1 or jbxx.kslb_h=2)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0  and (jbxx.kslb_h=4 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r") b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(writer, sheet_name="各类别考生成绩比较(政治)", index=None)
        writer.save()

    # 省级报告 总体概括 图
    def YSFGK_PROVINCE_IMG(self):

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        # 全省
        sql = "select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and kmh = 006"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and kmh = 006  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())

        province = [None] * 101
        for item in items:
            province[int(item[0])] = item[1] / num * 100
        x = list(range(101))
        plt.plot(x, province, color='springgreen', marker='.', label='全省')

        plt.xlim((0, 100))
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\全省考生单科成绩分布(政治).png', dpi=1200)
        plt.close()

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(政治).xlsx")

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 " \
              r"and jmx.tzh=6 and jmx.ksh like '"+dsh+r"%'"
        self.cursor.execute(sql)
        total =  self.cursor.fetchone()[0]

        low = int(total/3)
        high = int(total/1.5)

        df = pd.DataFrame(data=None,columns=["题号","分值","本市平均分","全省平均分","本市得分率","高分组得分率","中间组得分率","低分组得分率"])

        idxs = list(range(12, 24))
        for idx in idxs:
            row = []
            if idx<10:
                row.append("0"+str(idx))
            else:
                row.append(str(idx))

            num = 4.00
            row.append(num)

            sql = r"SELECT avg(kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx where ksh like '"+dsh+r"%' and kmh = 006 and idx = "+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            sql = r"SELECT avg(kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx where  kmh = 006 and idx = " + str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            row.append(row[2]/num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '"+dsh+r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0]/low/num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high-low) /num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between "+str(high+1)+r" and " + str(total) + " and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (total-high) / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        dths = [38,40]

        for dth in dths:
            row = []
            row.append(str(dth))
            if dth == 38:
                num = 14.00
            elif dth == 40:
                num = 26.00
            else:
                num = 10.00
            row.append(num)

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh="+str(dth)+" and ksh like '"+dsh+r"%'"
            print(sql)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh="+str(dth)
            print(sql)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            print(row)
            row.append(row[2]/num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.kmh=006 and jmx.ksh like '"+dsh+"%' and jmx.tzh=6 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN 1 and "+str(low)+r" and c.kmh=006 and c.tzh = "+str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0]/num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.kmh=006 and jmx.ksh like '" + dsh + "%' and jmx.tzh=6 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN "+str(low+1)+" and " + str(high) + r" and c.kmh=006 and c.tzh = " +str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"where jmx.kmh=006 and jmx.ksh like '" + dsh + "%' and jmx.tzh=6 ORDER BY jmx.zf desc) a) b " \
                r"on c.ksh=b.ksh where b.rn BETWEEN "+str(high+1)+" and " + str(total) + r" and c.kmh=006 and c.tzh = " + str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer,sheet_name="地市考生单题作答情况(政治)",index=None)
        writer.save()

    # 市级报告 单题分析 画图
    def DTFX_CITY_IMG(self, dsh):
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

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=6 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = list(range(12, 24))
        dths = [38,40]
        txt = idxs + dths

        x = []  # 难度
        y = []  # 区分度

        for idx in idxs:
            num = 4.0
            sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where amx.ksh like '" + dsh + "%' and kmh = 006 and idx = " + str(
                idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度

            # 前27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=6 and " \
                  r"jmx.ksh like '" + dsh + r"%' ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN 1 and " + str(ph_num) + r" and amx.idx = " + str(idx) + " and amx.kmh=006"
            print(sql)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=6 and " \
                  r"jmx.ksh like '" + dsh + r"%' ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN " + str(total - ph_num) + r" and " + str(total) + " and amx.idx = " + str(idx) + " and amx.kmh=006"
            print(sql)
            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total - ph_num) / num

            x.append(difficulty)
            y.append(ph - pl)

        print(x, y)

        for dth in dths:
            if dth == 38:
                num = 14.00
            elif dth == 40:
                num = 26.00
            else:
                num = 10.00

            sql = r"select avg(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where  jmx.kmh=006 and jmx.ksh like '" + dsh + r"%' and jmx.tzh=" + str(dth)
            self.cursor.execute(sql)
            print(sql)
            difficulty = self.cursor.fetchone()[0] / num  # 难度
            x.append(difficulty)

            sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
                  r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where " \
                  r"jmx.kmh=006 and jmx.tzh=6 and jmx.ksh like '" + dsh + r"%') b on a.ksh=b.ksh where a.kmh=006 and a.tzh=" + str(dth)
            self.cursor.execute(sql)
            result = np.array(self.cursor.fetchall(), dtype="float64")

            xt_score = np.array(result[:, 0], dtype="float64")
            zf_score = np.array(result[:, 1], dtype="float64")

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

            qfd = (D_a - D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
            y.append(qfd)

        print(x, y)
        plt.rcParams['figure.figsize'] = (15.0, 6.0)
        plt.scatter(x, y)
        plt.xlim((0, 1))
        plt.ylim((0, 1))
        plt.xlabel("难度")
        plt.ylabel("区分度")
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        th = []
        for i in range(len(x)):
            th.append(txt[i])
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i], y[i] + 0.008), arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(政治).png', dpi=1200)
        plt.close()

    # 市级报告附录 原始分概括
    def YSFFX_CITY_TABLE(self,dsh):

        sql = ""
        sql = "select mc from c_ds where DS_H = " + dsh
        self.cursor.execute(sql)
        ds_mc = self.cursor.fetchone()[0]

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析(附录)"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + ds_mc
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题水平分析原始分概括(政治).xlsx")

        city_num = [0] * 101
        province_num = [0] * 101

        province_total = 0
        city_total = 0

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"where jmx.kmh=006 and jmx.tzh=6 and zf!=0 GROUP BY jmx.zf ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()
        for item in items:
            province_num[int(item[0])] = item[1]
            province_total += item[1]  # 人数

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=6 " \
              r"and jmx.ksh like '"+dsh+r"%' and zf!=0 GROUP BY jmx.zf ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()
        for item in items:
            city_num[int(item[0])] = item[1]
            city_total += item[1]  # 人数

        df = pd.DataFrame(data=None,
                          columns=['一分段', '人数(本市)', '百分比(本市)', '累计百分比(本市)', '人数(全省)', '百分比(全省)', '累计百分比(全省)'])

        i = 100
        acc_city = 0
        acc_province = 0
        while i > 1:
            if city_num[i] > 0:
                acc_city += city_num[i]  # 累计百分比
                acc_province += province_num[i]  # 累计百分比
                row = []
                row.append(i)
                row.append(city_num[i])  # 本市人数
                row.append((city_num[i] / city_total) * 100)  # 本市百分比
                row.append((acc_city / city_total) * 100)  # 本市累计百分比

                row.append(province_num[i])
                row.append((province_num[i] / province_total) * 100)  # 全省百分比
                row.append((acc_province / province_total) * 100)  # 全省累计百分比
                self.set_list_precision(row)
                df.loc[len(df)] = row

            i = i - 1

        df.to_excel(excel_writer=writer, sheet_name='地市及全省考生一分段概括(政治)', index=None)

        writer.save()
        
    # 市级报告附录 单题分析
    def DTFX_CITY_APPENDIX(self, dsh):

        sql = "select mc from c_ds where DS_H = " + dsh
        self.cursor.execute(sql)
        ds_mc = self.cursor.fetchone()[0]

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析(附录)"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + ds_mc
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(政治).xlsx")

        rows = []
        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX where ksh like '"+dsh+r"%' and tzh=6 and kmh=006"
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(12, 24)

        for idx in idxs:

            a_h = 0
            b_h = 0
            c_h = 0
            d_h = 0

            a_m = 0
            b_m = 0
            c_m = 0
            d_m = 0

            a_l = 0
            b_l = 0
            c_l = 0
            d_l = 0

            a_t = 0
            b_t = 0
            c_t = 0
            d_t = 0

            row = []
            # 高分组
            sql = r"select amx.da,count(amx.da) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select a.*,rownum rn from (select ksh,zf from TYMHPT.T_GKPJ2020_TKSTZCJMX where" \
                  r" ksh like '"+dsh+r"%' and tzh=6 and kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)+" GROUP BY amx.da"
            print(sql)
            self.cursor.execute(sql)
            items = []
            items = self.cursor.fetchall()
            for item in items:
                if item[0] == 'A':
                    a_h = item[1]
                    a_t += a_h
                if item[0] == 'B':
                    b_h = item[1]
                    b_t += b_h
                if item[0] == 'C':
                    c_h = item[1]
                    c_t += c_h
                if item[0] == 'D':
                    d_h = item[1]
                    d_t += d_h

            # 中间组
            sql = r"select amx.da,count(amx.da) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select a.*,rownum rn from (select ksh,zf from TYMHPT.T_GKPJ2020_TKSTZCJMX where" \
                  r" ksh like '" + dsh + r"%' and tzh=6 and kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
            print(sql)
            self.cursor.execute(sql)
            items = []
            items = self.cursor.fetchall()
            for item in items:
                if item[0] == 'A':
                    a_m = item[1]
                    a_t += a_m
                if item[0] == 'B':
                    b_m = item[1]
                    b_t += b_m
                if item[0] == 'C':
                    c_m = item[1]
                    c_t += c_m
                if item[0] == 'D':
                    d_m = item[1]
                    d_t += d_m

            # 低分组
            sql = r"select amx.da,count(amx.da) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select a.*,rownum rn from (select ksh,zf from TYMHPT.T_GKPJ2020_TKSTZCJMX where" \
                  r" ksh like '" + dsh + r"%' and tzh=6 and kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN " + str(high+1) + r" and " + str(total) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
            print(sql)
            self.cursor.execute(sql)
            items = []
            items = self.cursor.fetchall()
            for item in items:
                if item[0] == 'A':
                    a_l = item[1]
                    a_t += a_l
                if item[0] == 'B':
                    b_l = item[1]
                    b_t += b_l
                if item[0] == 'C':
                    c_l = item[1]
                    c_t += c_l
                if item[0] == 'D':
                    d_l = item[1]
                    d_t += d_l

            row.append((a_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选A
            row.append((a_h / low) * 100)  # 高分组选A
            row.append((a_m / (high - low)) * 100)  # 中间组选A
            row.append((a_l / (total - high)) * 100)  # 低分组选A

            row.append((b_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选B
            row.append((b_h / low) * 100)  # 高分组选B
            row.append((b_m / (high - low)) * 100)  # 中间组选B
            row.append((b_l / (total - high)) * 100)  # 低分组选B

            row.append((c_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选C
            row.append((c_h / low) * 100)  # 高分组选C
            row.append((c_m / (high - low)) * 100)  # 中间组选C
            row.append((c_l / (total - high)) * 100)  # 低分组选C

            row.append((d_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选D
            row.append((d_h / low) * 100)  # 高分组选D
            row.append((d_m / (high - low)) * 100)  # 中间组选D
            row.append((d_l / (total - high)) * 100)  # 低分组选D
            row.insert(0, str(idx))
            self.set_list_precision(row)
            rows.append(row)

        df = pd.DataFrame(data=None, columns=["题号", "全部(A)", "高分组(A)", "中间组(A)", "低分组(A)",
                                              "全部(B)", "高分组(B)", "中间组(B)", "低分组(B)",
                                              "全部(C)", "高分组(C)", "中间组(C)", "低分组(C)",
                                              "全部(D)", "高分组(D)", "中间组(D)", "低分组(D)"])

        for i in range(len(rows)):
            
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(政治)")

        writer.save()

    # 省级报告 各市考生成绩比较
    def GSQKFX_PROVINCE(self):
        sql = "select ds_h,mc from c_ds"
        self.cursor.execute(sql)
        dss = self.cursor.fetchall()

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "各市情况分析(政治).xlsx")

        df = pd.DataFrame(data=None, columns=["地市代码", "地市全称", "人数", "比率", "平均分", "标准差", "差异系数(%)"])

        row = []
        # 全省
        sql = r"select count(zf),avg(zf),STDDEV_SAMP(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX where kmh=006 and tzh=6 "
        self.cursor.execute(sql)
        row = list(self.cursor.fetchone())
        total = row[0]
        row.append(float(row[2]) / float(row[1]))
        row.insert(1, row[0] / total * 100)
        row.insert(0, "全省")
        row.insert(0, "00")
        self.set_list_precision(row)
        df.loc[len(df)] = row

        for ds in dss:
            sql = r"select count(zf),avg(zf),STDDEV_SAMP(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
                  r"where kmh=006 and tzh=6 and ksh like '" + ds[0] + r"%'"
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.append(float(row[2]) / float(row[1]))
            row.insert(1, row[0] / total * 100)
            row.insert(0, ds[1])
            row.insert(0, ds[0])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市考生成绩比较(政治)")
        writer.save()

    # 省级报告 单题分析 画图
    def DTFX_PROVINCE(self):
        sql = ""

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "考生单题分析(政治).xlsx")

        df = pd.DataFrame(data=None, columns=["题号", "分值", "平均分", "标准差", "难度", "区分度"])

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=6 "
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = list(range(12, 24))
        dths = [38, 40]
        txt = idxs + dths

        x = [] # 难度
        y = [] # 区分度

        for idx in idxs:
            num = 6.0
            sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where kmh = 006 and idx = " + str(idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num #难度

            # 前27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=6 " \
                  r" ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN 1 and "+ str(ph_num) +r" and amx.idx = "+str(idx)+" and amx.kmh=006"
            print(sql)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=6  " \
                  r" ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN " + str(total-ph_num) + r" and "+str(total)+" and amx.idx = " + str(idx) + " and amx.kmh=006"
            print(sql)
            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total-ph_num) / num

            x.append(difficulty)
            y.append(ph-pl)

            row = []
            sql = r"select avg(kgval),STDDEV_SAMP(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX where kmh=006 and idx="+str(idx)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0,str(num))
            row.insert(0,str(idx))
            row.append(difficulty)
            row.append(ph-pl)
            self.set_list_precision(row)
            df.loc[len(df)] = row


        for dth in dths:
            if dth == 38:
                num = 14.00
            elif dth == 40:
                num = 26.00
            else:
                num = 10.00

            sql = r"select avg(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006  and jmx.tzh="+str(dth)
            self.cursor.execute(sql)
            print(sql)
            difficulty = self.cursor.fetchone()[0] / num # 难度
            x.append(difficulty)

            sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
                  r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where " \
                  r"jmx.kmh=006 and jmx.tzh = 6) b on a.ksh=b.ksh where a.kmh=006 and a.tzh="+str(dth)
            self.cursor.execute(sql)
            result = np.array(self.cursor.fetchall(), dtype="float64")

            xt_score = np.array(result[:, 0], dtype="float64")
            zf_score = np.array(result[:, 1], dtype="float64")

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score**2) - np.sum(xt_score)**2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score)**2

            qfd = (D_a-D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
            y.append(qfd)

            row = []
            sql = r"select avg(zf),STDDEV_SAMP(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX where kmh=006 and tzh="+str(dth)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0, str(num))
            row.insert(0, str(dth))
            row.append(difficulty)
            row.append(qfd)
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer,index=None,sheet_name="考生单题作答情况(政治)")
        writer.save()

        plt.rcParams['figure.figsize'] = (15.0,6.0)
        plt.scatter(x,y)
        plt.xlim((0, 1))
        plt.ylim((0, 1))
        plt.xlabel("难度")
        plt.ylabel("区分度")
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        th = []
        for i in range(len(x)):
            th.append(txt[i])
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] , y[i] + 0.008),arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(政治).png', dpi=1200)
        plt.close()

    # 省级报告(附录) 原始分概括
    def YSFGK_PROVINCE_APPENDIX(self):
        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析(附录)"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "原始分概括(政治).xlsx")

        sql = "select count(*) from kscj where sx!=0 and kl=1"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = r"select zf,count(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX where zf!=0 and kmh=006 and tzh=6 group by zf order by zf desc"
        self.cursor.execute(sql)
        results = self.cursor.fetchall()

        num = 0

        for result in results:
            row = []
            row.append(result[0])
            row.append(result[1])
            row.append((result[1] / total) * 100)
            num += result[1]
            row.append((num / total) * 100)
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="全省考生一分段(政治)")

        writer.save()


    # 省级报告附录 单题分析
    def DTFX_PROVINCE_APPENDIX(self):

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析(附录)"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "考生答题分析单题分析(政治).xlsx")

        rows = []
        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX where  tzh=6 and kmh=006"
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(12, 24)

        for idx in idxs:

            a_h = 0
            b_h = 0
            c_h = 0
            d_h = 0

            a_m = 0
            b_m = 0
            c_m = 0
            d_m = 0

            a_l = 0
            b_l = 0
            c_l = 0
            d_l = 0

            a_t = 0
            b_t = 0
            c_t = 0
            d_t = 0

            row = []
            # 高分组
            sql = r"select amx.da,count(amx.da) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select a.*,rownum rn from (select ksh,zf from TYMHPT.T_GKPJ2020_TKSTZCJMX where" \
                  r"  tzh=6 and kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)+" GROUP BY amx.da"
            print(sql)
            self.cursor.execute(sql)
            items = []
            items = self.cursor.fetchall()
            for item in items:
                if item[0] == 'A':
                    a_h = item[1]
                    a_t += a_h
                if item[0] == 'B':
                    b_h = item[1]
                    b_t += b_h
                if item[0] == 'C':
                    c_h = item[1]
                    c_t += c_h
                if item[0] == 'D':
                    d_h = item[1]
                    d_t += d_h

            # 中间组
            sql = r"select amx.da,count(amx.da) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select a.*,rownum rn from (select ksh,zf from TYMHPT.T_GKPJ2020_TKSTZCJMX where" \
                  r"  tzh=6 and kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
            print(sql)
            self.cursor.execute(sql)
            items = []
            items = self.cursor.fetchall()
            for item in items:
                if item[0] == 'A':
                    a_m = item[1]
                    a_t += a_m
                if item[0] == 'B':
                    b_m = item[1]
                    b_t += b_m
                if item[0] == 'C':
                    c_m = item[1]
                    c_t += c_m
                if item[0] == 'D':
                    d_m = item[1]
                    d_t += d_m

            # 低分组
            sql = r"select amx.da,count(amx.da) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select a.*,rownum rn from (select ksh,zf from TYMHPT.T_GKPJ2020_TKSTZCJMX where" \
                  r"   tzh=6 and kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN " + str(high+1) + r" and " + str(total) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
            print(sql)
            self.cursor.execute(sql)
            items = []
            items = self.cursor.fetchall()
            for item in items:
                if item[0] == 'A':
                    a_l = item[1]
                    a_t += a_l
                if item[0] == 'B':
                    b_l = item[1]
                    b_t += b_l
                if item[0] == 'C':
                    c_l = item[1]
                    c_t += c_l
                if item[0] == 'D':
                    d_l = item[1]
                    d_t += d_l

            row.append((a_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选A
            row.append((a_h / low) * 100)  # 高分组选A
            row.append((a_m / (high - low)) * 100)  # 中间组选A
            row.append((a_l / (total - high)) * 100)  # 低分组选A

            row.append((b_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选B
            row.append((b_h / low) * 100)  # 高分组选B
            row.append((b_m / (high - low)) * 100)  # 中间组选B
            row.append((b_l / (total - high)) * 100)  # 低分组选B

            row.append((c_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选C
            row.append((c_h / low) * 100)  # 高分组选C
            row.append((c_m / (high - low)) * 100)  # 中间组选C
            row.append((c_l / (total - high)) * 100)  # 低分组选C

            row.append((d_t / (a_t + b_t + c_t + d_t)) * 100)  # 全部选D
            row.append((d_h / low) * 100)  # 高分组选D
            row.append((d_m / (high - low)) * 100)  # 中间组选D
            row.append((d_l / (total - high)) * 100)  # 低分组选D
            row.insert(0, str(idx))
            self.set_list_precision(row)
            rows.append(row)

        df = pd.DataFrame(data=None, columns=["题号", "全部(A)", "高分组(A)", "中间组(A)", "低分组(A)",
                                              "全部(B)", "高分组(B)", "中间组(B)", "低分组(B)",
                                              "全部(C)", "高分组(C)", "中间组(C)", "低分组(C)",
                                              "全部(D)", "高分组(D)", "中间组(D)", "低分组(D)"])

        for i in range(len(rows)):
            
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(政治)")

        writer.save()