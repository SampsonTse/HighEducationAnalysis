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

        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

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
        
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        # 全省
        sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=6 and kscj.zh!=0 and jmx.kmh = 006 and jmx.zf!=0"
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
              r"right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=6 and kscj.zh!=0 " \
              r"and jmx.zf!=0 and jmx.kmh = 006 and  jmx.KSH LIKE '" + dsh + "%'"
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
            
            sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
                  r"jmx right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=6" \
                  r" and kscj.zh!=0 and jmx.zf!=0 and jmx.kmh = 006 and  jmx.KSH LIKE '" + xqh[0] + r"%'"
            self.cursor.execute(sql)
            
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
        sql = "select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and jmx.kmh = 006"
        
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and jmx.kmh = 006  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())

        province = [None] * 101
        for item in items:
            province[int(item[0])] = item[1] / num * 100
        x = list(range(101))
        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市
        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and jmx.kmh = 006 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 " \
              "and jmx.kmh = 006  and jmx.ksh like '" + dsh + r"%'GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 101

        for item in items:
            city[int(item[0])] = item[1] / num * 100

        x = list(range(101))

        plt.plot(x, city, color='springgreen', marker='.', label='全市')
        plt.xlim((0, 100))
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\地市及全省考生单科成绩分布(政治).png', dpi=600)
        plt.close()

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
        sql = "select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx left join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and jmx.kmh = 006"
        
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx left join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=6 and jmx.kmh = 006  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
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
        plt.savefig(path + '\\全省考生单科成绩分布(政治).png', dpi=600)
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

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and  xb_h=1) b on jmx.ksh=b.ksh " \
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
              r"where kscj.zh!=0 and  xb_h=2) b on jmx.ksh=b.ksh " \
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
              r"where kscj.zh!=0 and  (jbxx.kslb_h=1 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
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
              r"where kscj.zh!=0 and (jbxx.kslb_h=1 or jbxx.kslb_h=2)) b on jmx.ksh=b.ksh " \
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

        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 ) b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"

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

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
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

            sql = r"SELECT avg(kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=amx.ksh where amx.ksh like '"+dsh+r"%' and amx.kmh = 006 and idx = "+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            sql = r"SELECT avg(kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=amx.ksh where  amx.kmh = 006 and idx = " + str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            row.append(row[2]/num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '"+dsh+r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0]/low/num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high-low) /num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
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
                num = 24.00
            else:
                num = 10.00
            row.append(num)

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh="+str(dth)+" and ksh like '"+dsh+r"%'"
            
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh="+str(dth)
            
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            
            row.append(row[2]/num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.ksh like '"+dsh+"%' and jmx.tzh=6 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN 1 and "+str(low)+r" and c.kmh=006 and c.tzh = "+str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0]/num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.ksh like '" + dsh + "%' and jmx.tzh=6 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN "+str(low+1)+" and " + str(high) + r" and c.kmh=006 and c.tzh = " +str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.ksh like '" + dsh + "%' and jmx.tzh=6 ORDER BY jmx.zf desc) a) b " \
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

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
              r" where jmx.kmh=006 and jmx.tzh=6 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = list(range(12, 24))
        dths = [38,40]
        dths2 = [39]
        txt = idxs + dths +dths2

        x = []  # 难度
        y = []  # 区分度

        for idx in idxs:
            num = 4.00
            sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where amx.ksh like '" + dsh + "%' and kmh = 006 and idx = " + str(
                idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度

            # 前27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.zf,jmx.ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh = 006 and jmx.tzh=6 and " \
                  r"jmx.ksh like '" + dsh + r"%' ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN 1 and " + str(ph_num) + r" and amx.idx = " + str(idx) + " and amx.kmh=006"
            
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.zf,jmx.ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh = 006 and jmx.tzh=6 and " \
                  r"jmx.ksh like '" + dsh + r"%' ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN " + str(total - ph_num) + r" and " + str(total) + " and amx.idx = " + str(idx) + " and amx.kmh=006"

            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total - ph_num) / num

            x.append(difficulty)
            y.append(ph - pl)

        for dth in dths:
            if dth == 38:
                num = 14.00
            elif dth == 40:
                num = 24.00
            else:
                num = 10.00

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
                  r" where  jmx.kmh=006 and jmx.ksh like '" + dsh + r"%' and jmx.tzh=" + str(dth)
            
            self.cursor.execute(sql)
            
            difficulty = self.cursor.fetchone()[0] / num  # 难度
            x.append(difficulty)

            sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
                  r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where " \
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

        for xth in dths2:
            num = 12.00
            sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where sxt.ksh like '" + dsh + "%' and kmh=006 and xth=" + str(xth)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度
            x.append(difficulty)

            sql = r"select zh,b.sum from kscj right join " \
                  r"(select a.*,rownum rn from (select sum(xtval) sum,sxt.ksh from " \
                  r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where kmh = 006 and xth=" + str(xth) + r" and sxt.ksh " \
                  r"like '" + dsh + r"%' GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn "
            self.cursor.execute(sql)
            result = np.array(self.cursor.fetchall(), dtype="float64")
            zf_score = np.array(result[:, 0], dtype="float64")
            xt_score = np.array(result[:, 1], dtype="float64")

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

            qfd = (D_a - D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
            y.append(qfd)
        
        plt.rcParams['figure.figsize'] = (15.0, 4.00)
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
        plt.savefig(path + '\\各题难度-区分度分布散点图(政治).png', dpi=600)
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
              r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.tzh=6 and jmx.zf!=0 GROUP BY jmx.zf ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()
        for item in items:
            province_num[int(item[0])] = item[1]
            province_total += item[1]  # 人数

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.tzh=6 " \
              r"and jmx.ksh like '"+dsh+r"%' and jmx.zf!=0 GROUP BY jmx.zf ORDER BY jmx.zf desc"
        
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
        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
              r" where jmx.ksh like '"+dsh+r"%' and jmx.tzh=6 and jmx.kmh=006"
        
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
                  r"(select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where" \
                  r" jmx.ksh like '"+dsh+r"%' and jmx.tzh=6 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)+" GROUP BY amx.da"
            
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
                  r"(select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where" \
                  r" jmx.ksh like '" + dsh + r"%' and jmx.tzh=6 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
            
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
                  r"(select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.tzh=6 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN " + str(high+1) + r" and " + str(total) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
            
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
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        row = list(self.cursor.fetchone())
        total = row[0]
        row.append(float(row[2]) / float(row[1]) * 100)
        row.insert(1, row[0] / total * 100)
        row.insert(0, "全省")
        row.insert(0, "00")
        self.set_list_precision(row)
        df.loc[len(df)] = row

        for ds in dss:
            sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
                  r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + ds[0] + r") b on j" \
                   r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.append(float(row[2]) / float(row[1]) * 100)
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

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join " \
              r"gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.tzh=6 "
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = list(range(12, 24))
        dths = [38, 40]
        dths2 = [39]
        txt = idxs + dths + dths2

        x = [] # 难度
        y = [] # 区分度

        for idx in idxs:
            num = 4.00
            sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where kmh = 006 and idx = " + str(idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num #难度

            # 前27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=6 " \
                  r" ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN 1 and "+ str(ph_num) +r" and amx.idx = "+str(idx)+" and amx.kmh=006"

            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=6  " \
                  r" ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN " + str(total-ph_num) + r" and "+str(total)+" and amx.idx = " + str(idx) + " and amx.kmh=006"

            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total-ph_num) / num

            x.append(difficulty)
            y.append(ph-pl)

            row = []
            sql = r"SELECT avg(kgval),stddev_samp(amx.kgval) FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=amx.ksh where  amx.kmh = 006 and idx = " + str(idx)
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
                num = 24.00
            else:
                num = 10.00

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006  and jmx.tzh="+str(dth)
            
            self.cursor.execute(sql)
            
            difficulty = self.cursor.fetchone()[0] / num # 难度
            x.append(difficulty)

            sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
                  r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where " \
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
            sql = r"select avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx  where jmx.kmh=006 and jmx.tzh=" + str(dth)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0, str(num))
            row.insert(0, str(dth))
            row.append(difficulty)
            row.append(qfd)
            self.set_list_precision(row)
            df.loc[len(df)] = row

        for xth in dths2:
            row = []

            row.append(str(xth))
            num = 12.00
            row.append(str(num))

            sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where   kmh=006 and xth=" + str(xth)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度
            x.append(difficulty)

            sql = r"select zh,b.sum from kscj right join " \
                  r"(select a.*,rownum rn from (select sum(xtval) sum,sxt.ksh from " \
                  r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where kmh = 006 and xth=" + str(xth) + r"  GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn "
            self.cursor.execute(sql)
            result = np.array(self.cursor.fetchall(), dtype="float64")
            zf_score = np.array(result[:, 0], dtype="float64")
            xt_score = np.array(result[:, 1], dtype="float64")

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

            qfd = (D_a - D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
            y.append(qfd)



            row = []
            sql = r"SELECT avg(xtval),stddev_samp(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where  sxt.kmh = 006 and xth = " + str(xth)
            self.cursor.execute(sql)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0, str(num))
            row.insert(0, str(xth))
            row.append(difficulty)
            row.append(qfd)
            self.set_list_precision(row)

            df.loc[len(df)] = row

        df.to_excel(writer,index=None,sheet_name="考生单题作答情况(政治)")
        writer.save()

        writer2 = pd.ExcelWriter(path+r'\\'+'全省难度区分度（政治）.xlsx')
        df = pd.DataFrame(data=None,columns=['题号','难度','区分度'])

        plt.rcParams['figure.figsize'] = (15.0,4.00)
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
            row = [txt[i],x[i],y[i]]
            self.set_list_precision(row)
            df.loc[len(df)] = row
            th.append(txt[i])
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] , y[i] + 0.008),arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(政治).png', dpi=600)
        plt.close()
        df.to_excel(writer2,sheet_name="难度-区分度",index=None)
        writer2.save()
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

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
              r" where jmx.zf!=0 and jmx.kmh=006 and jmx.tzh=6 group by jmx.zf order by jmx.zf desc"
        
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
        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join " \
              r"gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where  jmx.tzh=6 and jmx.kmh=006"
        
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
                  r"(select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where" \
                  r"  jmx.tzh=6 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)+" GROUP BY amx.da"
            
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
                  r"(select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where" \
                  r"  jmx.tzh=6 and jmx.kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
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
                  r"(select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where" \
                  r"  jmx.tzh=6 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
                  r"where b.rn BETWEEN " + str(high+1) + r" and " + str(total) + r" and amx.kmh=006 and amx.idx=" + str(idx) + " GROUP BY amx.da"
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

    def MF_LF_CITY_TABLE(self,dsh):
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析零分率满分率(政治).xlsx")
        df = pd.DataFrame(data=None, columns=['题号', '零分人数', '零分率', '满分人数', '满分率'])

        rows = []

        idxs = list(range(12, 24))
        dths = [38, 40]
        dths2 = [39]
        txt = idxs + dths + dths2

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        for idx in idxs:
            sql = r"SELECT count(case when amx.kgval=4 then 1 else null end) num2 " \
                  r"FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=amx.ksh where kscj.zh!=0 and amx.ksh like '"+dsh+"%' and amx.kmh = 006 and idx="+str(idx)

            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0,total-row[0])

            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for dth in dths:
            if dth == 38:
                num = 14
            elif dth == 40:
                num = 24
            else:
                num = 10

            sql = r"select  count(case when jmx.zf=0 then 1 else null end) num1," \
                  r"count(case when jmx.zf="+str(num)+r" then 1 else null end) num2 " \
                  r"from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join GKEVA2020.kscj kscj on" \
                  r" kscj.ksh=jmx.ksh where jmx.kmh=006 and kscj.zh!=0 and jmx.tzh="+str(dth)+r" and jmx.ksh like '"+dsh+r"%'"
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for dth in dths2:
            num = 12
            sql = r"SELECT count(case when sxt.xtval=0 then 1 else null end) num2," \
                  r"count(case when sxt.xtval="+str(num)+r" then 1 else null end) num3 FROM " \
                  r"GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt  right join gkeva2020.kscj kscj" \
                  r" on kscj.ksh=sxt.ksh where kscj.zh!=0 and sxt.ksh like '"+dsh+r"%' and sxt.kmh = 006 and sxt.xth="+str(dth)

            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for i in range(len(rows)):
            rows[i].insert(0,txt[i])
            df.loc[len(df)] = rows[i]

        df.to_excel(writer, sheet_name="各市单题零分率满分率(政治)", index=None)
        writer.save()

    def MF_LF_PRO_TABLE(self):


        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "全省" + "考生答题分析单题分析零分率满分率(政治).xlsx")
        df = pd.DataFrame(data=None, columns=['题号', '零分人数', '零分率', '满分人数', '满分率'])

        rows = []

        idxs = list(range(12, 24))
        dths = [38, 40]
        dths2 = [39]
        txt = idxs + dths + dths2

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 ) b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=6 and jmx.zf!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        for idx in idxs:
            sql = r"SELECT count(case when amx.kgval=4 then 1 else null end) num2 " \
                  r"FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=amx.ksh where kscj.zh!=0 and  amx.kmh = 006 and idx=" + str(
                idx)

            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0, total - row[0])

            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for dth in dths:
            if dth == 38:
                num = 14
            elif dth == 40:
                num = 24
            else:
                num = 10

            sql = r"select  count(case when jmx.zf=0 then 1 else null end) num1," \
                  r"count(case when jmx.zf=" + str(num) + r" then 1 else null end) num2 " \
                  r"from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join GKEVA2020.kscj kscj on" \
                  r" kscj.ksh=jmx.ksh where jmx.kmh=006 and kscj.zh!=0 and jmx.tzh=" + str(dth)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for dth in dths2:
            num = 12
            sql = r"SELECT count(case when sxt.xtval=0 then 1 else null end) num2," \
                  r"count(case when sxt.xtval=" + str(num) + r" then 1 else null end) num3 FROM " \
                  r"GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt  right join gkeva2020.kscj kscj" \
                  r" on kscj.ksh=sxt.ksh where kscj.zh!=0 and  sxt.kmh = 006 and sxt.xth=" + str(dth)

            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for i in range(len(rows)):
            rows[i].insert(0, txt[i])
            df.loc[len(df)] = rows[i]

        df.to_excel(writer, sheet_name="各市单题零分率满分率(政治)", index=None)
        writer.save()

    def DTFX_CITY_TABLE_39(self,dsh):
        ql = ""
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

        df = pd.read_excel(path + '\\' + ds_mc + "考生答题分析单题分析(政治).xlsx",sheet_name=0)
        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(政治).xlsx")

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
              r"and jmx.tzh=6 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        low = int(total / 3)
        high = int(total / 1.5)

        xths = [39]
        for xth in xths:
            row = []

            row.append(str(xth))
            num = 12.00
            row.append(str(num))

            sql = r"SELECT avg(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
                  r"sxt.ksh like '" + dsh + r"%' and sxt.kmh = 006 and xth = " + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])

            sql = r"SELECT avg(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where  sxt.kmh = 006 and xth = " + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0])
            row.append(row[2] / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between 1 and " + str(low) + r" and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / low / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between " + str(low + 1) + r" and " + str(high) + r" and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high - low) / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between " + str(high + 1) + r" and " + str(total) + " and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (total - high) / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, sheet_name="地市考生单题作答情况(政治)", index=None)
        writer.save()

    # 市级报告 各区县占比
    def GQXZB_CITY_TABLE(self,dsh):
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "各区县各分数段分布情况(政治).xlsx")

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from GKEVA2020.c_xq where xq_h like '" + dsh + r"%'"

        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
              r"and jmx.tzh=6 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        mf = 100

        low = int(total / 3)
        high = int(total / 1.5)

        df = pd.DataFrame(data=None,columns=["区县号","区县名","高分组占比","高分组得分率","中间组占比","中间组得分率","低分组占比","低分组的得分率"])
        for xqh in xqhs:
            row = [xqh[0],xqh[1]]
            sql = "select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.ksh like '"+xqh[0]+r"%'"
            self.cursor.execute(sql)
            if self.cursor.fetchone()[0] == 0:
                continue
            
            sql = r"select count(zf),avg(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on jmx.ksh = b.ksh " \
                  r"where b.rn between 1 and " + str(low) + r" and jmx.ksh like '"+xqh[0]+"%' and jmx.tzh=6"
            self.cursor.execute(sql)
            result = list(self.cursor.fetchone())
            result[0] = result[0] / low * 100
            if result[1] != None:
                result[1] = result[1] / mf
            else:
                result[1] = "/"
            row = row +result

            sql = r"select count(zf),avg(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on jmx.ksh = b.ksh " \
                  r"where b.rn between " + str(low+1) + r" and " + str(high) + r" and jmx.ksh like '" + xqh[0] + "%' and jmx.tzh=6"
            self.cursor.execute(sql)
            result = list(self.cursor.fetchone())

            result[0] = result[0] / (high-low) * 100
            if result[1] != None:
                result[1] = result[1] / mf
            else:
                result[1] = "/"
            row = row + result

            sql = r"select count(zf),avg(zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=6 order by jmx.zf desc) a) b on jmx.ksh = b.ksh " \
                  r"where b.rn between " + str(high + 1) + r" and " + str(total) + r" and jmx.ksh like '" + xqh[0] + "%' and jmx.tzh=6"
            self.cursor.execute(sql)
            result = list(self.cursor.fetchone())

            result[0] = result[0] / (total - high) * 100
            if result[1] != None:
                result[1] = result[1] / mf
            else:
                result[1] = "/"
            row = row + result
            self.set_list_precision(row)
            df.loc[len(df)] = row

        
        df.to_excel(writer,sheet_name="各县区分组分布",index=None)
        writer.save()

    def JGFX_CITY_TABLE(self, dsh):
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析结构分析(政治).xlsx")

        df1 = pd.read_excel(path + "\\" + ds_mc + "考生答题分析单题分析(政治).xlsx", sheet_name=0)

        txts = df1['题号'].tolist()
        mean_province = df1['全省平均分'].tolist()
        mean_city = df1['本市平均分'].tolist()
        mean_high = df1['高分组得分率'].tolist()
        mean_mid = df1['中间组得分率'].tolist()
        mean_low = df1['低分组得分率'].tolist()


        row = []

        df2 = pd.DataFrame(data=None, columns=['题型', '题号', '分值','平均分','标准差','差异系数','全省平均分','全市得分率','全省得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ["选择题题(必做)", "12-23", "48.00"]
        num = 48.00
        mean_c=0
        mean_p=0
        mean_h=0
        mean_m=0
        mean_l = 0
        for i in range(12):
            mean_c += mean_city[i]
            mean_p += mean_province[i]
            mean_h += mean_high[i]
            mean_l += mean_low[i]
            mean_m += mean_mid[i]
        row.append(mean_c)
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12,13,14,15,16,17,18,19,20,21,22,23) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1]/row[-2]*100)
        row.append(mean_p)
        row.append(mean_c/num)
        row.append(mean_p/num)
        row.append(mean_h/12)
        row.append(mean_m/12)
        row.append(mean_l/12)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        row = ["非选择题题(必做)", "38-40", "52.00"]
        num = 52.00
        mean_c = 0
        mean_p = 0
        mean_h = 0
        mean_m = 0
        mean_l = 0
        for i in range(12,15):
            mean_c += mean_city[i]
            mean_p += mean_province[i]
            mean_h += mean_high[i]
            mean_l += mean_low[i]
            mean_m += mean_mid[i]
        row.append(mean_c)
        sql = r"select STDDEV_SAMP(a.score+b.score) from " \
              r"(select sxt.ksh,sum(sxt.xtval) score from GKEVA2020.T_GKPJ2020_TSJBNKSXT " \
              r"sxt right join GKEVA2020.kscj kscj on kscj.ksh=sxt.ksh where sxt.kmh=006 " \
              r"and sxt.xth in (39) GROUP BY sxt.ksh) a left join (select jmx.ksh,sum(jmx.zf) " \
              r"score from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where  jmx.kmh=006 and jmx.tzh" \
              r" in (38,40) GROUP BY jmx.ksh) b on a.ksh=b.ksh where a.ksh like '"+dsh+r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_p)
        row.append(mean_c / num)
        row.append(mean_p / num)
        row.append(mean_h / 3)
        row.append(mean_m / 3)
        row.append(mean_l / 3)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer,sheet_name="题型",index=None)
        
        df2 = pd.DataFrame(data=None, columns=['知识板块', '题号', '分值','平均分','标准差','差异系数','全省平均分','全市得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ["生产与消费；企业的经营与发展(必做)", "12", "4.00"]
        num = 4.00
        row.append(mean_city[0])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[0])
        row.append(row[3] / num)
        row.append(mean_high[0])
        row.append(mean_mid[0])
        row.append(mean_low[0])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["国家财政(必做)", "13", "4.00"]
        num = 4.00
        row.append(mean_city[1])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (13) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[1])
        row.append(row[3] / num)
        row.append(mean_high[1])
        row.append(mean_mid[1])
        row.append(mean_low[1])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["结算与信用工具(必做)", "14", "4.00"]
        num = 4.00
        row.append(mean_city[2])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (14) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[2])
        row.append(row[3] / num)
        row.append(mean_high[2])
        row.append(mean_mid[2])
        row.append(mean_low[2])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["分配制度(必做)", "15", "4.00"]
        num = 4.00
        row.append(mean_city[3])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (15) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[3])
        row.append(row[3] / num)
        row.append(mean_high[3])
        row.append(mean_mid[3])
        row.append(mean_low[3])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["行政机关：政府(必做)", "16", "4.00"]
        num = 4.00
        row.append(mean_city[4])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (16) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[4])
        row.append(row[3] / num)
        row.append(mean_high[4])
        row.append(mean_mid[4])
        row.append(mean_low[4])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["人民代表大会制度(必做)", "17", "4.00"]
        num = 4.00
        row.append(mean_city[5])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (17) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[5])
        row.append(row[3] / num)
        row.append(mean_high[5])
        row.append(mean_mid[5])
        row.append(mean_low[5])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["决定国际关系的因素；时代主题(必做)", "18", "4.00"]
        num = 4.00
        row.append(mean_city[6])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[6])
        row.append(row[3] / num)
        row.append(mean_high[6])
        row.append(mean_mid[6])
        row.append(mean_low[6])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["文化创新；文化对人的影响(必做)", "19", "4.00"]
        num = 4.00
        row.append(mean_city[7])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (19) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[7])
        row.append(row[3] / num)
        row.append(mean_high[7])
        row.append(mean_mid[7])
        row.append(mean_low[7])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["文化与社会；文化的多样性与文化传播(必做)", "20", "4.00"]
        num = 4.00
        row.append(mean_city[8])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (20) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[8])
        row.append(row[3] / num)
        row.append(mean_high[8])
        row.append(mean_mid[8])
        row.append(mean_low[8])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["唯物辩证法的联系观；唯物辩证法的发展观(必做)", "21", "4.00"]
        num = 4.00
        row.append(mean_city[9])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (21) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[9])
        row.append(row[3] / num)
        row.append(mean_high[9])
        row.append(mean_mid[9])
        row.append(mean_low[9])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["价值观、人生观（必做）", "22", "4.00"]
        num = 4.00
        row.append(mean_city[10])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (22) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[10])
        row.append(row[3] / num)
        row.append(mean_high[10])
        row.append(mean_mid[10])
        row.append(mean_low[10])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["社会存在与社会意识(必做)", "23", "4.00"]
        num = 4.00
        row.append(mean_city[11])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (23) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[11])
        row.append(row[3] / num)
        row.append(mean_high[11])
        row.append(mean_mid[11])
        row.append(mean_low[11])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["消费对生产的反作用(必做)", "38", "14.00"]
        num = 14.00
        row.append(mean_city[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=38 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[12])
        row.append(row[3] / num)
        row.append(mean_high[12])
        row.append(mean_mid[12])
        row.append(mean_low[12])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        row = ["公民政治参与的途径和方式(必做)", "39", "12.00"]
        num = 12.00
        row.append(mean_city[14])
        sql = r"SELECT stddev_samp(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r"sxt.ksh like '" + dsh + r"%' and sxt.kmh = 006 and xth = 39"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[14])
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["辩证唯物注意认识论原理；文化生活(必做)", "40", "26.00"]
        num = 26.00
        row.append(mean_city[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=40 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[13])
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer,sheet_name="知识板块",index=None)

        df2 = pd.DataFrame(data=None, columns=['考核能力', '题号', '分值','平均分','标准差','差异系数','全省平均分','全市得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ["从试题的文字表述中获取回答问题的有关信息", "12、14、16、19、20、21、22", "28.00"]
        num = 28.00
        row.append(mean_city[0]+mean_city[2]+mean_city[4]+mean_city[7]+mean_city[8]+mean_city[9]+mean_city[10])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[0]+mean_province[2]+mean_province[4]+mean_province[7]+mean_province[8]+mean_province[9]+mean_province[10])
        row.append(row[3] / num)
        row.append((mean_high[0]+mean_high[2]+mean_high[4]+mean_high[7]+mean_high[8]+mean_high[9]+mean_high[10])/7)
        row.append((mean_mid[0]+mean_mid[2]+mean_mid[4]+mean_mid[7]+mean_mid[8]+mean_mid[9]+mean_mid[10])/7)
        row.append((mean_low[0]+mean_low[2]+mean_low[4]+mean_low[7]+mean_low[8]+mean_low[9]+mean_low[10])/7)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["调动和运用自主学习过程中获得的重大时事和相关信息", "13、17、23", "12.00"]
        num = 12.00
        row.append(mean_city[1] + mean_city[5] + mean_city[11] )
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (13,17,23) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[1] + mean_province[5] + mean_province[11])
        row.append(row[3] / num)
        row.append((mean_high[1] + mean_high[5] + mean_high[11]) / 3)
        row.append((mean_mid[1] + mean_mid[5] + mean_mid[11]) / 3)
        row.append((mean_low[1] + mean_low[5] + mean_low[11]) / 3)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["调动和运用自主学习过程中获得的重大时事和相关信息", "15", "4.00"]
        num = 4.00
        row.append(mean_city[3])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (15) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[3])
        row.append(row[3] / num)
        row.append(mean_high[3])
        row.append(mean_mid[3])
        row.append(mean_low[3])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["综合检索和选用自己的“知识库”中的有关知识和技能", "18", "4.00"]
        num = 4.00
        row.append(mean_city[6])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (18) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[6])
        row.append(row[3] / num)
        row.append(mean_high[6])
        row.append(mean_mid[6])
        row.append(mean_low[6])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["从试题的图表等形式中获取回答问题的有关信息;准确描述试题所设计的学科的基本概念、观点和原理", "38", "14.00"]
        num = 14.00
        row.append(mean_city[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=38 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[12])
        row.append(row[3] / num)
        row.append(mean_high[12])
        row.append(mean_mid[12])
        row.append(mean_low[12])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["准确描述试题所设计的学科的基本概念、观点和原理", "39", "12.00"]
        num = 12.00
        row.append(mean_city[14])
        sql = r"SELECT stddev_samp(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r"sxt.ksh like '" + dsh + r"%' and sxt.kmh = 006 and xth = 39"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[14])
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["准确描述试题所设计的学科的基本概念、观点和原理;全面阐释或评价理论问题和现实问题", "40", "26.00"]
        num = 26.00
        row.append(mean_city[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=40 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(mean_province[13])
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer,sheet_name="考核能力",index=False)
        writer.save()

    def JGFX_PRO_TABLE(self):
        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "全省" + "考生答题分析结构分析(政治).xlsx")

        df1 = pd.read_excel(path + "\\" + "考生单题分析(政治).xlsx", sheet_name=0)

        txts = df1['题号'].tolist()
        mean_province = df1['平均分'].tolist()

        row = []

        df2 = pd.DataFrame(data=None,
                           columns=['题型', '题号', '分值', '平均分', '标准差', '难度'])

        row = ["选择题题(必做)", "12-23", "48.00"]
        num = 48.00
        mean_p = 0
        for i in range(12):
            mean_p += mean_province[i]
        row.append(mean_p)
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12,13,14,15,16,17,18,19,20,21,22,23) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(mean_p / num)
        df2.loc[len(df2)] = row

        row = ["非选择题题(必做)", "38-40", "52.00"]
        num = 52.00
        mean_p = 0
        for i in range(12, 15):
            mean_p += mean_province[i]
        row.append(mean_p)
        sql = r"select STDDEV_SAMP(a.score+b.score) from " \
              r"(select sxt.ksh,sum(sxt.xtval) score from GKEVA2020.T_GKPJ2020_TSJBNKSXT " \
              r"sxt right join GKEVA2020.kscj kscj on kscj.ksh=sxt.ksh where sxt.kmh=006 " \
              r"and sxt.xth in (39) GROUP BY sxt.ksh) a left join (select jmx.ksh,sum(jmx.zf) " \
              r"score from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where  jmx.kmh=006 and jmx.tzh" \
              r" in (38,40) GROUP BY jmx.ksh) b on a.ksh=b.ksh "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(mean_p / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer, sheet_name="题型", index=None)

        df2 = pd.DataFrame(data=None,
                           columns=['知识板块', '题号', '分值', '平均分', '标准差', '难度'])

        row = ["生产与消费；企业的经营与发展(必做)", "12", "4.00"]
        num = 4.00
        row.append(mean_province[0])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["国家财政(必做)", "13", "4.00"]
        num = 4.00
        row.append(mean_province[1])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (13) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["结算与信用工具(必做)", "14", "4.00"]
        num = 4.00
        row.append(mean_province[2])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (14) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["分配制度(必做)", "15", "4.00"]
        num = 4.00
        row.append(mean_province[3])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (15) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["行政机关：政府(必做)", "16", "4.00"]
        num = 4.00
        row.append(mean_province[4])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (16) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["人民代表大会制度(必做)", "17", "4.00"]
        num = 4.00
        row.append(mean_province[5])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (17) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["决定国际关系的因素；时代主题(必做)", "18", "4.00"]
        num = 4.00
        row.append(mean_province[6])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["文化创新；文化对人的影响(必做)", "19", "4.00"]
        num = 4.00
        row.append(mean_province[7])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (19) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["文化与社会；文化的多样性与文化传播(必做)", "20", "4.00"]
        num = 4.00
        row.append(mean_province[8])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (20) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["唯物辩证法的联系观；唯物辩证法的发展观(必做)", "21", "4.00"]
        num = 4.00
        row.append(mean_province[9])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (21) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["价值观、人生观（必做）", "22", "4.00"]
        num = 4.00
        row.append(mean_province[10])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (22) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["社会存在与社会意识(必做)", "23", "4.00"]
        num = 4.00
        row.append(mean_province[11])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (23) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["消费对生产的反作用(必做)", "38", "14.00"]
        num = 14.00
        row.append(mean_province[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=38 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["公民政治参与的途径和方式(必做)", "39", "12.00"]
        num = 12.00
        row.append(mean_province[14])
        sql = r"SELECT stddev_samp(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r" sxt.kmh = 006 and xth = 39"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["辩证唯物注意认识论原理；文化生活(必做)", "40", "26.00"]
        num = 26.00
        row.append(mean_province[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=40 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer, sheet_name="知识板块", index=None)

        df2 = pd.DataFrame(data=None,
                           columns=['考核能力', '题号', '分值', '平均分', '标准差', '难度'])

        row = ["从试题的文字表述中获取回答问题的有关信息", "12、14、16、19、20、21、22", "28.00"]
        num = 28.00
        row.append(
            mean_province[0] + mean_province[2] + mean_province[4] + mean_province[7] + mean_province[8] + mean_province[9] + mean_province[10])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (12) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["调动和运用自主学习过程中获得的重大时事和相关信息", "13、17、23", "12.00"]
        num = 12.00
        row.append(mean_province[1] + mean_province[5] + mean_province[11])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (13,17,23) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["调动和运用自主学习过程中获得的重大时事和相关信息", "15", "4.00"]
        num = 4.00
        row.append(mean_province[3])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (15) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["综合检索和选用自己的“知识库”中的有关知识和技能", "18", "4.00"]
        num = 4.00
        row.append(mean_province[6])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (18) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["从试题的图表等形式中获取回答问题的有关信息;准确描述试题所设计的学科的基本概念、观点和原理", "38", "14.00"]
        num = 14.00
        row.append(mean_province[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=38 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["准确描述试题所设计的学科的基本概念、观点和原理", "39", "12.00"]
        num = 12.00
        row.append(mean_province[14])
        sql = r"SELECT stddev_samp(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r" sxt.kmh = 006 and xth = 39"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ["准确描述试题所设计的学科的基本概念、观点和原理;全面阐释或评价理论问题和现实问题", "40", "26.00"]
        num = 26.00
        row.append(mean_province[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=40 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[3] / num)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer, sheet_name="考核能力", index=False)
        writer.save()

    def MF_LF_CITY_TABLE_39(self,dsh):
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析零分率满分率(政治).xlsx")
        df = pd.read_excel(path + '\\' + ds_mc + "考生答题分析单题分析零分率满分率(政治).xlsx",sheet_name=0)

        rows = []

        dths2 = [39]
        txt = dths2

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=3 and jmx.zf!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        for dth in dths2:
            num = 12
            sql = r"SELECT count(case when sxt.xtval=0 then 1 else null end) num2," \
                  r"count(case when sxt.xtval="+str(num)+r" then 1 else null end) num3 FROM " \
                  r"GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt  right join gkeva2020.kscj kscj" \
                  r" on kscj.ksh=sxt.ksh where kscj.zh!=0 and sxt.ksh like '"+dsh+r"%' and sxt.kmh = 006 and sxt.xth="+str(dth)

            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for i in range(len(rows)):
            rows[i].insert(0,txt[i])
            df.loc[len(df)] = rows[i]

        df.to_excel(writer, sheet_name="各市单题零分率满分率(政治)", index=None)
        writer.save()

    def FSX_TXT(self):

        sql = "select * from gkeva2020.c_ds"
        self.cursor.execute(sql)
        dss = self.cursor.fetchall()

        for ds in dss:
            dsh = ds[0]
            print(ds[1])

            sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
                  r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
                  r"and jmx.tzh=6 and jmx.ksh like '" + dsh + r"%'"
            self.cursor.execute(sql)
            total = self.cursor.fetchone()[0]

            low = int(total / 3)
            high = int(total / 1.5)

            sql = r"select b.* from (select a.*,rownum rn from" \
                  r" (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and " \
                  r"jmx.ksh like '" + dsh + "%' and jmx.tzh=6 and jmx.zf!=0 ORDER BY jmx.zf desc) a) b where  b.rn ="+ str(low)
            self.cursor.execute(sql)
            print(self.cursor.fetchone()[1])

            sql = r"select b.* from (select a.*,rownum rn from" \
                  r" (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and " \
                  r"jmx.ksh like '" + dsh + "%' and jmx.tzh=6 and jmx.zf!=0 ORDER BY jmx.zf desc) a) b where  b.rn =" + str(high)
            self.cursor.execute(sql)
            print(self.cursor.fetchone()[1])


    def FSX_TXT(self):


        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
              r"and jmx.tzh=6 "
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        low = int(total / 3)
        high = int(total / 1.5)

        sql = r"select b.* from (select a.*,rownum rn from" \
              r" (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and " \
              r" jmx.tzh=6 and jmx.zf!=0 ORDER BY jmx.zf desc) a) b where  b.rn ="+ str(low)
        self.cursor.execute(sql)
        print(self.cursor.fetchone()[1])

        sql = r"select b.* from (select a.*,rownum rn from" \
              r" (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and " \
              r" and jmx.tzh=6 and jmx.zf!=0 ORDER BY jmx.zf desc) a) b where  b.rn =" + str(high)
        self.cursor.execute(sql)
        print(self.cursor.fetchone()[1])


    def nd_qfd_39(self):

        sql = "select * from gkeva2020.c_ds"
        self.cursor.execute(sql)
        cities = self.cursor.fetchall()

        x = []
        y = []

        xth = 39
        for city in cities:

            if city[0] == "01":
                continue

            excel_path = r"C:\Users\SampsonTse\Desktop\各地市政治难度区分度\\"+city[1]+r"考生答题分析单题区分度难度(政治).xlsx"
            writer = pd.ExcelWriter(excel_path)
            df = pd.read_excel(excel_path)

            sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
                  r" where jmx.kmh=006 and jmx.tzh=6 and jmx.ksh like '" + city[0] + r"%'"
            self.cursor.execute(sql)
            total = self.cursor.fetchone()[0]

            row = [39]

            num = 12.00
            sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where sxt.ksh like '" + city[0] + "%' and kmh=006 and xth=" + str(xth)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度
            row.append(difficulty)

            sql = r"select zh,b.sum from kscj right join " \
                  r"(select a.*,rownum rn from (select sum(xtval) sum,sxt.ksh from " \
                  r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where kmh = 006 and xth=" + str(xth) + r" and sxt.ksh " \
                  r"like '" + city[0] + r"%' GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn "
            self.cursor.execute(sql)
            result = np.array(self.cursor.fetchall(), dtype="float64")
            zf_score = np.array(result[:, 0], dtype="float64")
            xt_score = np.array(result[:, 1], dtype="float64")

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

            qfd = (D_a - D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
            row.append(qfd)
            self.set_list_precision(row)
            df.loc[len(df)-1] = row
            df.to_excel(writer,sheet_name="难度-区分度",index=None)
            writer.save()





