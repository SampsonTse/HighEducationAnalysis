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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析总体概括(历史).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and xb_h=1) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and xb_h=2) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=1 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=2 or jbxx.kslb_h=4)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=1 or jbxx.kslb_h=2)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and jbxx.ds_h=" + dsh + " and (jbxx.kslb_h=4 or jbxx.kslb_h=3)) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计

        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

            
        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(writer, sheet_name="各类别考生成绩比较(历史)", index=None)

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from GKEVA2020.c_xq where xq_h like '" + dsh + r"%'"
        
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        # 全省
        sql = r"select count(jmx.zf),avg(jmx.zf),STDDEV_SAMP(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r"right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=9 and kscj.zh!=0 and jmx.kmh = 006 and jmx.zf!=0"
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
              r"right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=9 and kscj.zh!=0 " \
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
                  r"jmx right join GKEVA2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.tzh=9" \
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

        df.to_excel(excel_writer=writer, sheet_name="各县区考生成绩比较(历史)", index=None)
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
        sql = "select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=9 and jmx.kmh = 006"
        
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=9 and jmx.kmh = 006  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())

        province = [None] * 101
        for item in items:
            province[int(item[0])] = item[1] / num * 100
        x = list(range(101))
        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市
        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=9 and jmx.kmh = 006 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=9 " \
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
        plt.savefig(path + '\\地市及全省考生单科成绩分布(历史).png', dpi=600)
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
        sql = "select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx left join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=9 and jmx.kmh = 006"
        
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx left join gkeva2020.kscj kscj on kscj.ksh = jmx.ksh where jmx.tzh=9 and jmx.kmh = 006  GROUP BY (jmx.zf) ORDER BY jmx.zf desc"
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
        plt.savefig(path + '\\全省考生单科成绩分布(历史).png', dpi=600)
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

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析原始分概括(历史).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 ) b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 计算维度为男
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0 and  xb_h=1) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

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
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(writer, sheet_name="各类别考生成绩比较(历史)", index=None)
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(历史).xlsx")

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
              r"and jmx.tzh=9 and jmx.ksh like '"+dsh+r"%'"
        self.cursor.execute(sql)
        total =  self.cursor.fetchone()[0]

        low = int(total/3)
        high = int(total/1.5)

        df = pd.DataFrame(data=None,columns=["题号","分值","本市平均分","全省平均分","本市得分率","高分组得分率","中间组得分率","低分组得分率"])

        idxs = list(range(24, 36))
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
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between 1 and "+str(low)+r" and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0]/low/num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between "+str(low+1)+r" and " + str(high) + r" and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high-low) /num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between "+str(high+1)+r" and " + str(total) + " and amx.kmh=006 and amx.idx="+str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (total-high) / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        dths = [41,45,46,47]

        for dth in dths:
            row = []
            row.append(str(dth))
            if dth == 41:
                num = 25.00
            elif dth in [45, 46, 47]:
                num = 15.00
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
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.ksh like '"+dsh+"%' and jmx.tzh=9 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN 1 and "+str(low)+r" and c.kmh=006 and c.tzh = "+str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0]/num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.ksh like '" + dsh + "%' and jmx.tzh=9 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN "+str(low+1)+" and " + str(high) + r" and c.kmh=006 and c.tzh = " +str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.ksh like '" + dsh + "%' and jmx.tzh=9 ORDER BY jmx.zf desc) a) b " \
                r"on c.ksh=b.ksh where b.rn BETWEEN "+str(high+1)+" and " + str(total) + r" and c.kmh=006 and c.tzh = " + str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer,sheet_name="地市考生单题作答情况(历史)",index=None)
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
              r" where jmx.kmh=006 and jmx.tzh=9 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = list(range(24, 36))
        dths = [41,45,46,47]
        dths2 = [42]
        txt = idxs + dths + dths2

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
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh = 006 and jmx.tzh=9 and " \
                  r"jmx.ksh like '" + dsh + r"%' ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN 1 and " + str(ph_num) + r" and amx.idx = " + str(idx) + " and amx.kmh=006"

            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.zf,jmx.ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh = 006 and jmx.tzh=9 and " \
                  r"jmx.ksh like '" + dsh + r"%' ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN " + str(total - ph_num) + r" and " + str(total) + " and amx.idx = " + str(idx) + " and amx.kmh=006"

            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total - ph_num) / num

            x.append(difficulty)
            y.append(ph - pl)

        for dth in dths:
            if dth == 41:
                num = 25.00
            elif dth in [45,46,47]:
                num = 15.00

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
                  r" where  jmx.kmh=006 and jmx.ksh like '" + dsh + r"%' and jmx.tzh=" + str(dth)

            self.cursor.execute(sql)

            difficulty = self.cursor.fetchone()[0] / num  # 难度
            x.append(difficulty)

            sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
                  r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where " \
                  r"jmx.kmh=006 and jmx.tzh=9 and jmx.ksh like '" + dsh + r"%') b on a.ksh=b.ksh where a.kmh=006 and a.tzh=" + str(dth)
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
            print(sql)
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
        plt.savefig(path + '\\各题难度-区分度分布散点图(历史).png', dpi=600)
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题水平分析原始分概括(历史).xlsx")

        city_num = [0] * 101
        province_num = [0] * 101

        province_total = 0
        city_total = 0

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
              r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.tzh=9 and jmx.zf!=0 GROUP BY jmx.zf ORDER BY jmx.zf desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()
        for item in items:
            province_num[int(item[0])] = item[1]
            province_total += item[1]  # 人数

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.tzh=9 " \
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

        df.to_excel(excel_writer=writer, sheet_name='地市及全省考生一分段概括(历史)', index=None)

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(历史).xlsx")

        rows = []
        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
              r" where jmx.ksh like '"+dsh+r"%' and jmx.tzh=9 and jmx.kmh=006"
        
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(24, 36)

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
                  r" jmx.ksh like '"+dsh+r"%' and jmx.tzh=9 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
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
                  r" jmx.ksh like '" + dsh + r"%' and jmx.tzh=9 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
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
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.tzh=9 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
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

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(历史)")

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

        writer = pd.ExcelWriter(path + '\\' + "各市情况分析(历史).xlsx")

        df = pd.DataFrame(data=None, columns=["地市代码", "地市全称", "人数", "比率", "平均分", "标准差", "差异系数(%)"])

        row = []
        # 全省
        sql = r"select count(jmx.zf),avg(jmx.zf),stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX " \
              r"jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh " \
              r"where kscj.zh!=0) b on jmx.ksh=b.ksh " \
              r"where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
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
                  r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.append(float(row[2]) / float(row[1]) * 100)
            row.insert(1, row[0] / total * 100)
            row.insert(0, ds[1])
            row.insert(0, ds[0])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市考生成绩比较(历史)")
        writer.save()

    # 省级报告 单题分析
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

        writer = pd.ExcelWriter(path + '\\' + "考生单题分析(历史).xlsx")

        df = pd.DataFrame(data=None, columns=["题号", "分值", "平均分", "标准差", "难度", "区分度","高分组得分率", "中间组得分率", "低分组得分率"])

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join " \
              r"gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and jmx.tzh=9 "
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        low = int(total / 3)
        high = int(total / 1.5)

        idxs = list(range(24, 36))
        dths = [41,45,46,47]
        dths2 = [42]
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
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=9 " \
                  r" ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh w" \
                  r"here b.rn BETWEEN 1 and "+ str(ph_num) +r" and amx.idx = "+str(idx)+" and amx.kmh=006"

            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select zf,ksh from T" \
                  r"YMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh = 006 and jmx.tzh=9  " \
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

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where  jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between 1 and " + str(low) + r" and amx.kmh=006 and amx.idx=" + str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / low / num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between " + str(low + 1) + r" and " + str(high) + r" and amx.kmh=006 and amx.idx=" + str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high - low) / num)

            sql = r"select sum(kgval) from GKEVA2020.T_GKPJ2020_TKSKGDAMX amx " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on amx.ksh = b.ksh " \
                  r"where b.rn between " + str(high + 1) + r" and " + str(total) + " and amx.kmh=006 and amx.idx=" + str(idx)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (total - high) / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        for dth in dths:
            if dth == 41:
                num = 25.00
            elif dth in [45, 46, 47]:
                num = 15.00

            sql = r"select avg(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx  where jmx.kmh=006  and jmx.tzh="+str(dth)
            
            self.cursor.execute(sql)
            
            difficulty = self.cursor.fetchone()[0] / num # 难度
            x.append(difficulty)

            sql = r"select a.zf,b.zf,b.ksh from TYMHPT.T_GKPJ2020_TKSTZCJMX a right join " \
                  r"(select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where " \
                  r"jmx.kmh=006 and jmx.tzh=9) b on a.ksh=b.ksh where a.kmh=006 and a.tzh="+str(dth)
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

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006  and jmx.tzh=9 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN 1 and " + str(low) + r" and c.kmh=006 and c.tzh = " + str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 and  jmx.tzh=9 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN " + str(low + 1) + " and " + str(
                high) + r" and c.kmh=006 and c.tzh = " + str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

            sql = r"select avg(c.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX c right join" \
                  r" (select a.*,rownum rn from (select jmx.ksh,jmx.zf from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx " \
                  r" right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006  and jmx.tzh=9 ORDER BY jmx.zf desc) a) b " \
                  r"on c.ksh=b.ksh where b.rn BETWEEN " + str(high + 1) + " and " + str(
                total) + r" and c.kmh=006 and c.tzh = " + str(dth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / num)

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

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where  jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between 1 and " + str(low) + r" and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / low / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between " + str(low + 1) + r" and " + str(high) + r" and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high - low) / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where  jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between " + str(high + 1) + r" and " + str(total) + " and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (total - high) / num)

            self.set_list_precision(row)

            df.loc[len(df)] = row

        df.to_excel(writer,index=None,sheet_name="考生单题作答情况(历史)")
        writer.save()

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
            th.append(txt[i])
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] , y[i] + 0.008),arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(历史).png', dpi=600)
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

        writer = pd.ExcelWriter(path + '\\' + "原始分概括(历史).xlsx")

        sql = "select count(*) from kscj where sx!=0 and kl=1"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = r"select jmx.zf,count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh" \
              r" where jmx.zf!=0 and jmx.kmh=006 and jmx.tzh=9 group by jmx.zf order by jmx.zf desc"
        
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

        df.to_excel(writer, index=None, sheet_name="全省考生一分段(历史)")

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

        writer = pd.ExcelWriter(path + '\\' + "考生答题分析单题分析(历史).xlsx")

        rows = []
        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join " \
              r"gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where  jmx.tzh=9 and jmx.kmh=006"
        
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(24, 36)

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
                  r"  jmx.tzh=9 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
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
                  r"  jmx.tzh=9 and jmx.kmh=006 ORDER BY zf desc) a) b on amx.ksh=b.ksh " \
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
                  r"  jmx.tzh=9 and jmx.kmh=006 ORDER BY jmx.zf desc) a) b on amx.ksh=b.ksh " \
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

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(历史)")

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析零分率满分率(历史).xlsx")
        df = pd.DataFrame(data=None, columns=['题号', '零分人数', '零分率', '满分人数', '满分率'])

        rows = []

        idxs = list(range(24, 36))
        dths = [41, 45, 46, 47]
        dths2 = [42]
        txt = idxs + dths + dths2

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
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
            if dth == 41:
                num = 25
            elif dth in [45, 46, 47]:
                num = 15

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
            print(sql)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for i in range(len(rows)):
            rows[i].insert(0,txt[i])
            df.loc[len(df)] = rows[i]

        df.to_excel(writer, sheet_name="各市单题零分率满分率(历史)", index=None)
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

        writer = pd.ExcelWriter(path + '\\' + "全省" + "考生答题分析单题分析零分率满分率(历史).xlsx")
        df = pd.DataFrame(data=None, columns=['题号', '零分人数', '零分率', '满分人数', '满分率'])

        rows = []

        idxs = list(range(24, 36))
        dths = [41, 45, 46, 47]
        dths2 = [42]
        txt = idxs + dths + dths2

        sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
              r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 ) b on j" \
              r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh=9 and jmx.zf!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        for idx in idxs:
            sql = r"SELECT count(case when amx.kgval=4 then 1 else null end) num2 " \
                  r"FROM GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=amx.ksh where kscj.zh!=0 and  amx.kmh = 006 and idx=" + str(idx)

            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(0, total - row[0])

            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for dth in dths:
            if dth == 41:
                num = 25
            elif dth in [45, 46, 47]:
                num = 15

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
            print(sql)
            self.cursor.execute(sql)
            row = list(self.cursor.fetchone())
            row.insert(1, row[0] / total)
            row.append(row[2] / total)
            self.set_list_precision(row)
            rows.append(row)

        for i in range(len(rows)):
            rows[i].insert(0, txt[i])
            df.loc[len(df)] = rows[i]

        df.to_excel(writer, sheet_name="各市单题零分率满分率(历史)", index=None)
        writer.save()

    def DTFX_CITY_TABLE_42(self,dsh):
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

        df = pd.read_excel(path + '\\' + ds_mc + "考生答题分析单题分析(历史).xlsx",sheet_name=0)
        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(历史).xlsx")

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
              r"and jmx.tzh=9 and jmx.ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        low = int(total / 3)
        high = int(total / 1.5)

        xths = [42]
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
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between 1 and " + str(low) + r" and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / low / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between " + str(low + 1) + r" and " + str(high) + r" and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (high - low) / num)

            sql = r"select sum(xtval) from GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
                  r"right join (select a.*,rownum rn from (select jmx.ksh from " \
                  r"TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join gkeva2020.kscj kscj " \
                  r"on kscj.ksh=jmx.ksh where jmx.ksh like '" + dsh + r"%' and jmx.kmh=006 " \
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on sxt.ksh = b.ksh " \
                  r"where b.rn between " + str(high + 1) + r" and " + str(total) + " and sxt.kmh=006 and sxt.xth=" + str(xth)
            self.cursor.execute(sql)
            row.append(self.cursor.fetchone()[0] / (total - high) / num)

            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, sheet_name="地市考生单题作答情况(历史)", index=None)
        writer.save()

    def XZT_COUNT_CITY(self,dsh):

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

        fig,ax = plt.subplots(figsize=(12,5),dpi=80)
        X = ["45","46","47"]

        Y = []
        for x in X:
            sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
                  r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 and jbxx.ds_h=" + dsh + r") b on j" \
                  r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh="+x+r" and jmx.zf!=0"
            print(sql)
            self.cursor.execute(sql)
            Y.append(self.cursor.fetchone()[0])

        for (x,y) in zip(range(3),Y):
            print(x,y)
            ax.text(x,y+3, y, va='center', fontsize=16)

        plt.bar(X, Y)
        plt.savefig(path + '\\各选做题人数(历史).png', dpi=600)

    def XZT_COUNT_PROVINCE(self):

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        fig,ax = plt.subplots(figsize=(12,5),dpi=80)
        X = ["45","46","47"]

        Y = []
        for x in X:
            sql = r"select count(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right join (select kscj.ksh from " \
                  r"GKEVA2020.kscj kscj left join GKEVA2020.jbxx jbxx on jbxx.ksh=kscj.ksh where kscj.zh!=0 ) b on j" \
                  r"mx.ksh=b.ksh where jmx.kmh = 006 and jmx.tzh="+x+r" and jmx.zf!=0"
            print(sql)
            self.cursor.execute(sql)
            Y.append(self.cursor.fetchone()[0])

        for (x,y) in zip(range(3),Y):
            print(x,y)
            ax.text(x,y+3, y, va='center', fontsize=16)

        plt.bar(X, Y)
        plt.savefig(path + '\\各选做题人数(历史).png', dpi=600)

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析结构分析(历史).xlsx")

        df1 = pd.read_excel(path + "\\" + ds_mc + "考生答题分析单题分析(历史).xlsx", sheet_name=0)

        txts = df1['题号'].tolist()
        mean_city = df1['本市平均分'].tolist()
        mean_high = df1['高分组得分率'].tolist()
        mean_mid = df1['中间组得分率'].tolist()
        mean_low = df1['低分组得分率'].tolist()

        row = []

        df2 = pd.DataFrame(data=None, columns=['题型', '题号', '分值', '平均分','标准差','差异系数','得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ['单选题(必做)', '24-35', '48.00']
        num = 48.00
        mean_c = 0
        mean_h = 0
        mean_m = 0
        mean_l = 0
        for i in range(12):
            mean_c = mean_c + mean_city[i]
            mean_h = mean_h + mean_high[i]
            mean_m = mean_m + mean_mid[i]
            mean_l = mean_l + mean_low[i]
        row.append(mean_c)
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (24,25,26,27,28,29,30,31,32,33,34,35) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_h/12)
        row.append(mean_m/12)
        row.append(mean_l/12)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['必考题(41)', '41', '25.00']
        num = 25.00
        row.append(mean_city[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=41 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[12])
        row.append(mean_mid[12])
        row.append(mean_low[12])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['必考题(42)', '42', '12.00']
        num = 12.00
        row.append(mean_city[16])
        sql = r"SELECT avg(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r"sxt.ksh like '" + dsh + r"%' and sxt.kmh = 006 and xth = 42"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[16])
        row.append(mean_mid[16])
        row.append(mean_low[16])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        row = ['选考题(45)', '45', '15.00']
        num = 15.00
        row.append(mean_city[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=45 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['选考题(46)', '46', '15.00']
        num = 15.00
        row.append(mean_city[14])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=46 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['选考题(47)', '47', '15.00']
        num = 15.00
        row.append(mean_city[15])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=47 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[15])
        row.append(mean_mid[15])
        row.append(mean_low[15])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer,sheet_name="题型",index=None)

        df2 = pd.DataFrame(data=None, columns=['知识板块', '题号', '分值', '平均分','标准差','差异系数','得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ['古代中国的政治制度(必做)','24','4.00']
        num = 4.00
        row.append(mean_city[0])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
             r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
             r" amx.idx in (24) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[0])
        row.append(mean_mid[0])
        row.append(mean_low[0])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代中国的科学技术与文学艺术(必做)', '25', '4.00']
        num = 4.00
        row.append(mean_city[1])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (25) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[1])
        row.append(mean_mid[1])
        row.append(mean_low[1])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代中国的经济(必做)', '26', '4.00']
        num = 4.00
        row.append(mean_city[2])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (26) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[2])
        row.append(mean_mid[2])
        row.append(mean_low[2])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['中国传统文化主流思想的演变(必做)', '27', '4.00']
        num = 4.00
        row.append(mean_city[3])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (27) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[3])
        row.append(mean_mid[3])
        row.append(mean_low[3])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['经济结构的变化与资本主义的曲折发展：思想解放的潮流(必做)', '28', '4.00']
        num = 4.00
        row.append(mean_city[4])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (28) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[4])
        row.append(mean_mid[4])
        row.append(mean_low[4])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['经济结构的变化与资本主义的曲折发展(必做)', '29', '4.00']
        num = 4.00
        row.append(mean_city[5])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (29) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[5])
        row.append(mean_mid[5])
        row.append(mean_low[5])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['近代中国的民主革命(必做)', '30', '4.00']
        num = 4.00
        row.append(mean_city[6])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (30) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[6])
        row.append(mean_mid[6])
        row.append(mean_low[6])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['中国特色社会主义建设的道路(必做)', '31', '4.00']
        num = 4.00
        row.append(mean_city[7])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (31) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[7])
        row.append(mean_mid[7])
        row.append(mean_low[7])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代希腊:罗马的政治制度(必做)', '32', '4.00']
        num = 4.00
        row.append(mean_city[8])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (32) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[8])
        row.append(mean_mid[8])
        row.append(mean_low[8])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['西方人文精神的发展(必做)', '33', '4.00']
        num = 4.00
        row.append(mean_city[9])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (33) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[9])
        row.append(mean_mid[9])
        row.append(mean_low[9])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['科学社会主义理论的诞生和国际工人运动(必做)', '34', '4.00']
        num = 4.00
        row.append(mean_city[10])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (34) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[10])
        row.append(mean_mid[10])
        row.append(mean_low[10])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['第二次世界大战后的经济全球化趋势(必做)', '35', '4.00']
        num = 4.00
        row.append(mean_city[11])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (35) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[11])
        row.append(mean_mid[11])
        row.append(mean_low[11])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['现代中国的对外关系：第二次世界大战后世界政治格局的演变；第二次世界大战后经济的全球化趋势(必做)', '41', '25.00']
        num = 25.00
        row.append(mean_city[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=38 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[12])
        row.append(mean_mid[12])
        row.append(mean_low[12])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代中国的政治制度；古代中国的经济发展；中国传统文化主流思想的演变；古代中国的科学技术与文学艺术等(必做)', '42', '12.00']
        num = 12.00
        row.append(mean_city[16])
        sql = r"SELECT avg(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r"sxt.ksh like '" + dsh + r"%' and sxt.kmh = 006 and xth = 42"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[16])
        row.append(mean_mid[16])
        row.append(mean_low[16])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['历史上的重大变革(选做1)', '45', '15.00']
        num = 15.00
        row.append(mean_city[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=45 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['20世纪的战争与和平(选做2)', '46', '15.00']
        num = 15.00
        row.append(mean_city[14])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=46 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['中外历史任务平手(选做3)', '47', '15.00']
        num = 15.00
        row.append(mean_city[15])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=47 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[15])
        row.append(mean_mid[15])
        row.append(mean_low[15])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        df2.to_excel(writer,sheet_name="知识板块",index=None)

        df2 = pd.DataFrame(data=None, columns=['考核能力', '题号', '分值', '平均分','标准差','差异系数','得分率','高分组得分率','中间组得分率','低分组得分率'])
        row = ['获取和解读信息：调用和运用知识(必做)', '24-28,31-35', '40.00']
        num = 40.00
        mean_c = 0
        mean_h = 0
        mean_m = 0
        mean_l = 0
        for i in [0,1,2,3,4,7,8,9,10,11]:
            mean_c = mean_c + mean_city[i]
            mean_h = mean_h + mean_high[i]
            mean_m = mean_m + mean_mid[i]
            mean_l = mean_l + mean_low[i]
        row.append(mean_c)
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (24,25,26,27,28,31,32,33,34,35) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_h/10)
        row.append(mean_m/10)
        row.append(mean_l/10)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息:调动和运用知识；描述和阐释事物(必做)', '29-30', '8.00']
        num = 8.00
        row.append(mean_city[5] + mean_city[6] )
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (29,30) and amx.ksh like '" + dsh + r"%'and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append((mean_high[5] + mean_high[6]) / 2)
        row.append((mean_mid[5] + mean_mid[6])/ 2)
        row.append((mean_low[5] + mean_low[6] ) / 2)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物；论证和探讨问题(必做)', '41-42', '37.00']
        num = 37.00
        row.append(mean_city[12]+mean_city[16])
        sql = "select STDDEV_SAMP(a.score+b.score) from " \
              "(select sxt.ksh,sum(sxt.xtval) score from GKEVA2020.T_GKPJ2020_TSJBNKSXT" \
              " sxt right join GKEVA2020.kscj kscj on kscj.ksh=sxt.ksh where  sxt.kmh=006" \
              " and sxt.xth in (42) GROUP BY sxt.ksh) a left join (select jmx.ksh,sum(jmx.zf)" \
              " score from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where  " \
              "jmx.kmh=006 and jmx.tzh in (41) GROUP BY jmx.ksh) b on a.ksh=b.ksh"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append((mean_high[12]+mean_high[16])/2)
        row.append((mean_mid[12]+mean_mid[16])/2)
        row.append((mean_low[12]+mean_low[16])/2)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物(选做1)', '45', '15.00']
        num = 15.00
        row.append(mean_city[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=45 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物(选做2)', '46', '15.00']
        num = 15.00
        row.append(mean_city[14])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=46 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物(选做3)', '47', '15.00']
        num = 15.00
        row.append(mean_city[15])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=47 and ksh like '" + dsh + r"%'"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[15])
        row.append(mean_mid[15])
        row.append(mean_low[15])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        df2.to_excel(writer,sheet_name='考核能力',index=None)
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "各区县各分数段分布情况(历史).xlsx")

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from GKEVA2020.c_xq where xq_h like '" + dsh + r"%'"

        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        sql = r"select count(*) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx right " \
              r"join gkeva2020.kscj kscj on kscj.ksh=jmx.ksh where jmx.kmh=006 " \
              r"and jmx.tzh=9 and jmx.ksh like '" + dsh + r"%'"
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
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on jmx.ksh = b.ksh " \
                  r"where b.rn between 1 and " + str(low) + r" and jmx.ksh like '"+xqh[0]+"%' and jmx.tzh=9"
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
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on jmx.ksh = b.ksh " \
                  r"where b.rn between " + str(low+1) + r" and " + str(high) + r" and jmx.ksh like '" + xqh[0] + "%' and jmx.tzh=9"
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
                  r"and jmx.tzh=9 order by jmx.zf desc) a) b on jmx.ksh = b.ksh " \
                  r"where b.rn between " + str(high + 1) + r" and " + str(total) + r" and jmx.ksh like '" + xqh[0] + "%' and jmx.tzh=9"
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

    def JGFX_PRO_TABLE(self):
        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        writer = pd.ExcelWriter(path + '\\' + "全省" + "考生答题分析结构分析(历史).xlsx")

        df1 = pd.read_excel(path + "\\" + "考生单题分析(历史).xlsx", sheet_name=0)
        print(df1)
        means = df1['平均分'].tolist()
        mean_high = df1['高分组得分率'].tolist()
        mean_mid = df1['中间组得分率'].tolist()
        mean_low = df1['低分组得分率'].tolist()


        row = []

        df2 = pd.DataFrame(data=None, columns=['题型', '题号', '分值', '平均分', '标准差','差异系数','得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ['单选题(必做)', '24-35', '48.00']
        num = 48.00
        mean_c = 0
        mean_h = 0
        mean_m = 0
        mean_l = 0
        for i in range(12):
            mean_c = mean_c + means[i]
            mean_h = mean_h + mean_high[i]
            mean_m = mean_m + mean_mid[i]
            mean_l = mean_l + mean_low[i]
        row.append(mean_c)
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (24,25,26,27,28,29,30,31,32,33,34,35) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_h / 12)
        row.append(mean_m / 12)
        row.append(mean_l / 12)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['必考题(41)', '41', '25.00']
        num = 25.00
        row.append(means[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=41 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[12])
        row.append(mean_mid[12])
        row.append(mean_low[12])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['必考题(42)', '42', '12.00']
        num = 12.00
        row.append(means[16])
        sql = r"SELECT avg(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r"sxt.kmh = 006 and xth = 42"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[16])
        row.append(mean_mid[16])
        row.append(mean_low[16])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        row = ['选考题(45)', '45', '15.00']
        num = 15.00
        row.append(means[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=45"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['选考题(46)', '46', '15.00']
        num = 15.00
        row.append(means[14])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=46 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['选考题(47)', '47', '15.00']
        num = 15.00
        row.append(means[15])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=47 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[15])
        row.append(mean_mid[15])
        row.append(mean_low[15])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        df2.to_excel(writer,sheet_name="题型",index=None)

        df2 = pd.DataFrame(data=None, columns=['知识板块', '题号', '分值', '平均分', '标准差','差异系数','得分率','高分组得分率','中间组得分率','低分组得分率'])

        row = ['古代中国的政治制度(必做)','24','4.00']
        num = 4.00
        row.append(means[0])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (24) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[0])
        row.append(mean_mid[0])
        row.append(mean_low[0])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代中国的科学技术与文学艺术(必做)', '25', '4.00']
        num = 4.00
        row.append(means[1])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (25) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[1])
        row.append(mean_mid[1])
        row.append(mean_low[1])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代中国的经济(必做)', '26', '4.00']
        num = 4.00
        row.append(means[2])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (26) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[2])
        row.append(mean_mid[2])
        row.append(mean_low[2])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['中国传统文化主流思想的演变(必做)', '27', '4.00']
        num = 4.00
        row.append(means[3])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (27) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[3])
        row.append(mean_mid[3])
        row.append(mean_low[3])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['经济结构的变化与资本主义的曲折发展：思想解放的潮流(必做)', '28', '4.00']
        num = 4.00
        row.append(means[4])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (28) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[4])
        row.append(mean_mid[4])
        row.append(mean_low[4])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['经济结构的变化与资本主义的曲折发展(必做)', '29', '4.00']
        num = 4.00
        row.append(means[5])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (29) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[5])
        row.append(mean_mid[5])
        row.append(mean_low[5])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['近代中国的民主革命(必做)', '30', '4.00']
        num = 4.00
        row.append(means[6])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (30) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[6])
        row.append(mean_mid[6])
        row.append(mean_low[6])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['中国特色社会主义建设的道路(必做)', '31', '4.00']
        num = 4.00
        row.append(means[7])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (31) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[7])
        row.append(mean_mid[7])
        row.append(mean_low[7])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代希腊:罗马的政治制度(必做)', '32', '4.00']
        num = 4.00
        row.append(means[8])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (32) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[8])
        row.append(mean_mid[8])
        row.append(mean_low[8])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['西方人文精神的发展(必做)', '33', '4.00']
        num = 4.00
        row.append(means[9])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (33) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[9])
        row.append(mean_mid[9])
        row.append(mean_low[9])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['科学社会主义理论的诞生和国际工人运动(必做)', '34', '4.00']
        num = 4.00
        row.append(means[10])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (34) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[10])
        row.append(mean_mid[10])
        row.append(mean_low[10])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['第二次世界大战后的经济全球化趋势(必做)', '35', '4.00']
        num = 4.00
        row.append(means[11])
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (35) and amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[11])
        row.append(mean_mid[11])
        row.append(mean_low[11])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['现代中国的对外关系：第二次世界大战后世界政治格局的演变；第二次世界大战后经济的全球化趋势(必做)', '41', '25.00']
        num = 25.00
        row.append(means[12])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=38 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[12])
        row.append(mean_mid[12])
        row.append(mean_low[12])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['古代中国的政治制度；古代中国的经济发展；中国传统文化主流思想的演变；古代中国的科学技术与文学艺术等(必做)', '42', '12.00']
        num = 12.00
        row.append(means[16])
        sql = r"SELECT avg(xtval) FROM GKEVA2020.T_GKPJ2020_TSJBNKSXT sxt " \
              r"right join gkeva2020.kscj kscj on kscj.ksh=sxt.ksh where " \
              r" sxt.kmh = 006 and xth = 42"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[16])
        row.append(mean_mid[16])
        row.append(mean_low[16])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['历史上的重大变革(选做1)', '45', '15.00']
        num = 15.00
        row.append(means[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=45 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['20世纪的战争与和平(选做2)', '46', '15.00']
        num = 15.00
        row.append(means[14])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=46 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['中外历史任务平手(选做3)', '47', '15.00']
        num = 15.00
        row.append(means[15])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=47 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[15])
        row.append(mean_mid[15])
        row.append(mean_low[15])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        df2.to_excel(writer,sheet_name="知识板块",index=None)

        df2 = pd.DataFrame(data=None,
                           columns=['知识板块', '题号', '分值', '平均分', '标准差', '差异系数', '得分率', '高分组得分率', '中间组得分率', '低分组得分率'])

        row = ['获取和解读信息：调用和运用知识(必做)', '24-28,31-35', '40.00']
        num = 40.00
        num = 40.00
        mean_c = 0
        mean_h = 0
        mean_m = 0
        mean_l = 0
        for i in [0, 1, 2, 3, 4, 7, 8, 9, 10, 11]:
            mean_c = mean_c + means[i]
            mean_h = mean_h + mean_high[i]
            mean_m = mean_m + mean_mid[i]
            mean_l = mean_l + mean_low[i]
        row.append(mean_c)
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (24,25,26,27,28,31,32,33,34,35) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_h / 10)
        row.append(mean_m / 10)
        row.append(mean_l / 10)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息:调动和运用知识；描述和阐释事物(必做)', '29-30', '8.00']
        num = 8.00
        row.append(means[5] + means[6] )
        sql = r"select stddev_samp(a.score) from (SELECT sum(amx.kgval) score from " \
              r"GKEVA2020.T_GKPJ2020_TKSKGDAMX amx right join GKEVA2020.kscj kscj on amx.ksh=kscj.ksh where" \
              r" amx.idx in (29,30) and  amx.kmh=006 GROUP BY amx.ksh) a"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append((mean_high[5] + mean_high[6]) / 2)
        row.append((mean_mid[5] + mean_mid[6]) / 2)
        row.append((mean_low[5] + mean_low[6]) / 2)
        self.set_list_precision(row)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物；论证和探讨问题(必做)', '41-42', '37.00']
        num = 37.00
        row.append(means[12] + means[16])
        sql = "select STDDEV_SAMP(a.score+b.score) from " \
              "(select sxt.ksh,sum(sxt.xtval) score from GKEVA2020.T_GKPJ2020_TSJBNKSXT" \
              " sxt right join GKEVA2020.kscj kscj on kscj.ksh=sxt.ksh where  sxt.kmh=006" \
              " and sxt.xth in (42) GROUP BY sxt.ksh) a left join (select jmx.ksh,sum(jmx.zf)" \
              " score from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where  " \
              "jmx.kmh=006 and jmx.tzh in (41) GROUP BY jmx.ksh) b on a.ksh=b.ksh"
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append((mean_high[12] + mean_high[16]) / 2)
        row.append((mean_mid[12] + mean_mid[16]) / 2)
        row.append((mean_low[12] + mean_low[16]) / 2)
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物(选做1)', '45', '15.00']
        num = 15.00
        row.append(means[13])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=45 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[13])
        row.append(mean_mid[13])
        row.append(mean_low[13])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物(选做2)', '46', '15.00']
        num = 15.00
        row.append(means[14])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=46 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[14])
        row.append(mean_mid[14])
        row.append(mean_low[14])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row

        row = ['获取和解读信息；调动和运用知识；描述和阐释事物(选做3)', '47', '15.00']
        num = 15.00
        row.append(means[15])
        sql = r"select stddev_samp(jmx.zf) from TYMHPT.T_GKPJ2020_TKSTZCJMX jmx where jmx.kmh=006 and jmx.tzh=47 "
        self.cursor.execute(sql)
        row.append(self.cursor.fetchone()[0])
        row.append(row[-1] / row[-2] * 100)
        row.append(row[3] / num)
        row.append(mean_high[15])
        row.append(mean_mid[15])
        row.append(mean_low[15])
        self.set_list_precision(row)
        df2.loc[len(df2)] = row


        df2.to_excel(writer,sheet_name='考核能力',index=None)
        writer.save()

    def MF_LF_CITY_TABLE_42(self,dsh):
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析零分率满分率(历史).xlsx")
        df = pd.read_excel(path + '\\' + ds_mc + "考生答题分析单题分析零分率满分率(历史).xlsx",sheet_name=0)

        rows = []

        dths2 = [42]
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

        df.to_excel(writer, sheet_name="各市单题零分率满分率(历史)", index=None)
        writer.save()










