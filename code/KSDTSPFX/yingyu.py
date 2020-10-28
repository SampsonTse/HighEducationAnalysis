import numpy as np
import pandas as  pd
import pymysql
import os
import matplotlib.pyplot  as plt
import decimal
import cx_Oracle
import matplotlib.ticker as ticker
import math
import openpyxl

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
np.set_printoptions(precision=2)


# 英语考生答题水平分析
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

    # 市级报告 总体概括 制表
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析总体概括(英语).xlsx")

        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r'select count(a.wy) from kscj  a right join JBXX b on a.KSH = b.KSH ' \
              r'WHERE b.DS_H=' + dsh + ' and a.wy!=0'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj  a right join JBXX  b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and b.XB_H = 1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1])) * 100) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH where b.XB_H = 1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and b.XB_H = 2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH where b.XB_H = 2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX  b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH where a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别考生成绩比较(英语)", excel_writer=writer, index=None)

        # 文科
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r'select count(a.wy) from kscj   a right join JBXX   b on a.KSH = b.KSH ' \
              r'WHERE b.DS_H=' + dsh + r' and a.kl=2 and a.wy!=0'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and b.XB_H = 1 and a.kl=2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 1 and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.wy)  num,AVG(A.wy)  mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H=" + dsh + r" and b.XB_H = 2 and a.kl=2 and a.wy!=0"
        print(sql)
        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 2 and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + " and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and a.kl=2 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH and a.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别文科考生成绩比较(英语)", excel_writer=writer, index=None)

        # 理科
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r'select count(a.wy) from kscj  a right join JBXX  b on a.KSH = b.KSH WHERE b.DS_H=' + dsh + r' and a.kl=1'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H=" + dsh + r" and b.XB_H = 1 and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        self.cursor.execute(sql)
        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where b.XB_H = 1 and a.kl=1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and b.XB_H = 2 and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 2 and a.kl=1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=1 and a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.wy)   num,AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H=" + dsh + r" and a.kl=1 and a.wy!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.wy)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH and a.kl=1 where a.wy!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别理科考生成绩比较(英语)", excel_writer=writer, index=None)

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from c_xq where  xq_h like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A where  a.wy!=0 "
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A " \
              r"where a.wy!=0 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = "select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE  a.wy!=0 and B.XQ_H = " + xqh[0]
            self.cursor.execute(sql)
            print(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
            result.append(result[1] / 150)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区考生成绩比较(英语)", index=None)

        # 各区县理考生成绩比较

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A where A.kl=1"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A " \
              r"where A.kl=1 and a.wy!=0 and A.KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = "select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE A.kl=1 and a.wy!=0 and B.XQ_H = " + xqh[0]
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
            result.append(result[1] / 150)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区理科考生成绩比较(英语)", index=None)

        # 各区县文科考生成绩比较

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A where A.kl=2 and a.wy!=0"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A " \
              r"where A.kl=2 and a.wy!=0 and A.KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = "select count(wy),AVG(A.wy)   mean,STDDEV_SAMP(A.wy)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE A.kl=2 and a.wy!=0 and B.XQ_H = " + xqh[0]
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
            result.append(result[1] / 150)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区文科考生成绩比较(英语)", index=None)

        writer.save()

    # 市级报告 总体概括 画图
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

        # 全省
        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(wy) FROM kscj where wy!=0"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = r"SELECT wy,COUNT(wy) FROM kscj WHERE wy !=0 GROUP BY wy"
        print(sql)
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(151))

        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市
        sql = "SELECT COUNT(wy) FROM kscj where KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = r"SELECT wy,COUNT(wy) FROM kscj WHERE wy != 0 and KSH LIKE '" + dsh + r"%' GROUP BY  wy"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 151

        for item in items:
            city[item[0]] = round(item[1] / num * 100, 2)

        x = list(range(151))

        plt.plot(x, city, color='springgreen', marker='.', label='全市')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center',bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\地市及全省考生单科成绩分布(英语).png', dpi=1200)
        plt.close()
        

        # 全省文科
        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(wy) FROM kscj where kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "SELECT wy,COUNT(wy) FROM kscj WHERE wy != 0 and kl=2 GROUP BY  wy "
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(151))

        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市文科
        sql = "SELECT COUNT(wy) FROM kscj where kl=2 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全市人数

        sql = r"SELECT wy,COUNT(wy) FROM kscj WHERE wy != 0 and kl=2 and KSH LIKE '" + dsh + r"%' GROUP BY  wy"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 151

        for item in items:
            city[item[0]] = round(item[1] / num * 100, 2)

        x = list(range(151))

        plt.plot(x, city, color='springgreen', marker='.', label='全市')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center',bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\地市及全省文科考生单科成绩分布(英语).png', dpi=1200)
        plt.close()
        

        # 全省理科
        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(wy) FROM kscj where kl=1"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "SELECT wy,COUNT(wy) FROM kscj WHERE wy != 0 and kl=1 GROUP BY  wy "
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(151))

        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市理科
        plt.rcParams['figure.figsize'] = (15.0, 6)
        sql = "SELECT COUNT(wy) FROM kscj where kl=1 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全市人数

        sql = r"SELECT wy,COUNT(wy) FROM kscj WHERE wy != 0 and kl=1 and KSH LIKE '" + dsh + r"%' GROUP BY  wy"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 151

        for item in items:
            city[item[0]] = round(item[1] / num * 100, 2)

        x = list(range(151))

        plt.plot(x, city, color='springgreen', marker='.', label='全市')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center',bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\地市及全省理科考生单科成绩分布(英语).png', dpi=1200)
        plt.close()

    # 省级报告 原始分概括
    def YSFGK_PROVINCE_TABLE(self):

        sql = ""

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        xbs = [1, 2]
        hjs = [["1", "3"], ["2", "4"]]
        ywjs = [["1", "2"], ["3", "4"]]

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析原始分概括(英语).xlsx")

        # 全省考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(*) from kscj  a right join JBXX  b on a.ksh = b.ksh"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh where b.xb_h=" + str(xb)
            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if xb == 1:
                results.insert(0, '男')
            else:
                results.insert(0, '女')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 户籍
        for hj in hjs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh where b.kslb_h = " + str(
                hj[0]) + " or b.kslb_h = " + hj[1]

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in hj:
                results.insert(0, '城镇')
            else:
                results.insert(0, '农村')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 应往届
        for ywj in ywjs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh where b.kslb_h = " + ywj[0] + " or b.kslb_h = " + \
                  ywj[1]

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in ywj:
                results.insert(0, '应届')
            else:
                results.insert(0, '往届')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        sql = "select count(a.wy)  num,AVG(a.wy)  mean,STDDEV_SAMP(a.wy)  std " \
              "from kscj a right join JBXX  b on a.ksh = b.ksh"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别考生成绩比较(英语)", index=None)

        # 全省文科考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(a.wy)  num " \
              "from kscj  a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=2 and b.xb_h=" + str(xb)
            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if xb == 1:
                results.insert(0, '男')
            else:
                results.insert(0, '女')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 户籍
        for hj in hjs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=2 and (b.kslb_h = " + hj[0] + " or b.kslb_h = " + hj[1] + ")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in hj:
                results.insert(0, '城镇')
            else:
                results.insert(0, '农村')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 应往届
        for ywj in ywjs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=2 and (b.kslb_h = " + ywj[0] + " or b.kslb_h = " + ywj[1] + ")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in ywj:
                results.insert(0, '应届')
            else:
                results.insert(0, '往届')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
              "from kscj   a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别文科考生成绩比较(英语)", index=None)

        # 全省理科考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(*) from kscj   a right join JBXX   b on a.ksh = b.ksh where a.kl=1"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=1 and b.xb_h=" + str(xb)
            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if xb == 1:
                results.insert(0, '男')
            else:
                results.insert(0, '女')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 户籍
        for hj in hjs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=1 and (b.kslb_h = " + hj[0] + " or b.kslb_h = " + hj[1] + ")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in hj:
                results.insert(0, '城镇')
            else:
                results.insert(0, '农村')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 应往届
        for ywj in ywjs:
            sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=1 and (b.kslb_h = " + ywj[0] + " or b.kslb_h = " + ywj[1] + ")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in ywj:
                results.insert(0, '应届')
            else:
                results.insert(0, '往届')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        sql = "select count(a.wy)   num,AVG(a.wy)   mean,STDDEV_SAMP(a.wy)   std " \
              "from kscj   a right join JBXX  b on a.ksh = b.ksh where a.kl=1"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别理科考生成绩比较(英语)", index=None)

        writer.save()

    # 省级报告 原始分概括 画图
    def YSFGK_PROVINCE_IMG(self):

        sql = ""

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\考生答题分析"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + "全省"
        if not os.path.exists(path):
            os.makedirs(path)

        plt.rcParams['figure.figsize'] = (15.0, 6)
        plt.xlim((0, 150))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "select count(*) from kscj where wy!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        score = [None] * 151
        sql = "select wy,count(wy) from kscj where wy!=0 group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            score[item[0]] = item[1] / total

        x = list(range(151))

        plt.plot(x, score, color='springgreen', marker='.', label='全省')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(25))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\' + '全省考生单科成绩分布(英语).png', dpi=1200)
        plt.close()

        plt.rcParams['figure.figsize'] = (15.0, 6)
        plt.xlim((0, 150))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "select count(*) from kscj where wy!=0 and kl = 2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        score = [None] * 151
        sql = "select wy,count(wy) from kscj where wy!=0 and kl = 2 group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            score[item[0]] = item[1] / total

        x = list(range(151))

        plt.plot(x, score, color='springgreen', marker='.', label='全省')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(25))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\' + '全省文科考生单科成绩分布(英语).png', dpi=1200)
        plt.close()

        plt.rcParams['figure.figsize'] = (15.0, 6)
        plt.xlim((0, 150))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "select count(*) from kscj where wy!=0 and kl=1"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        score = [None] * 151
        sql = "select wy,count(wy) from kscj where wy!=0 and kl=1 group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            score[item[0]] = item[1] / total

        x = list(range(151))

        plt.plot(x, score, color='springgreen', marker='.', label='全省')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(25))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center', bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\' + '全省理科考生单科成绩分布(英语).png', dpi=1200)
        plt.close()

    # 市级报告 单题分析
    def DTFX_CITY_TABLE(self,dsh):

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(英语).xlsx")

        sql = r"select count(*) from kscj where ksh like '"+dsh+r"%'  "
        self.cursor.execute(sql)
        num_ks = self.cursor.fetchone()[0]

        sql = r"select count(*) from kscj "
        self.cursor.execute(sql)
        num_t = self.cursor.fetchone()[0]

        low = int(num_ks/3)
        high = int(num_ks/1.5)

        df = pd.DataFrame(data=None,columns=['题号','分值','本市平均分','全省平均分','本市得分率','高分组得分率','中间组得分率','低分组得分率'])

        kgths = list(range(1,41))
        zgths = list(range(61,82))

        for kgth in kgths:

            if kgth in range(1,21):
                num = 2.0
            else:
                num = 1.5

            row = []
            row.append(str(kgth+20))
            row.append(num)

            total = 0

            # 全省平均分
            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX a right join jbxx b " \
                  "on a.ksh = b.ksh where a.idx=" + str(kgth) + " and kmh=101"
            self.cursor.execute(sql)
            avg_province = self.cursor.fetchone()[0] / num_t

            # 本市计算高分组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ RIGHT JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE wykmh=101 and jbxx.ds_h="+dsh+" ORDER BY KSCJ.wy DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = "+str(kgth)+" and c.kmh=101 and b.rn BETWEEN 1 and "+str(low)
            self.cursor.execute(sql)
            sum_h = float(self.cursor.fetchone()[0])
            total = total + sum_h
            dfl_h = sum_h/ low / num

            # 本市计算中间组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ RIGHT JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE wykmh=101 and jbxx.ds_h=" + dsh + " ORDER BY KSCJ.wy DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=101 and b.rn BETWEEN "+str(low+1)+" and " + str(high)
            self.cursor.execute(sql)
            sum_m = float(self.cursor.fetchone()[0])
            total = total + sum_m
            dfl_m = sum_m / (high - low) / num

            # 本市计算低分组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ RIGHT JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE wykmh=101 and jbxx.ds_h=" + dsh + " ORDER BY KSCJ.wy DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=101 and b.rn BETWEEN "+str(high+1)+" and " + str(num_ks)
            self.cursor.execute(sql)
            sum_l = float(self.cursor.fetchone()[0])
            total = total + sum_l
            dfl_l = sum_l / (num_ks - high) / num

            row.append(total/num_ks) # 全市平均分
            row.append(avg_province) # 全省平均分
            row.append(total/num_ks/5) # 全市得分率
            row.append(dfl_h) #高分组
            row.append(dfl_m) #中间组
            row.append(dfl_l) #低分组

            self.set_list_precision(row)
            print(row)
            df.loc[len(df)] = row

        for zgth in zgths:

            row = []
            num = 0
            row.append(str(zgth))
            if zgth in range(61,71):
                num = 1.5
            elif zgth in range(71,81):
                num = 1.00
            elif zgth == 81:
                num = 25.00
            row.append(num)

            total = 0

            # 全省平均分
            sql = "select sum(b.sum) from " \
                  "(select sum(a.xtval) as sum,a.xth,a.ksh from T_GKPJ2020_TSJBNKSXT a " \
                  "right join jbxx on jbxx.ksh=a.ksh where a.kmh = 101 and a.xth = "+str(zgth)+" GROUP BY a.ksh,a.xth) b"
            print(sql)
            self.cursor.execute(sql)
            avg_province = self.cursor.fetchone()[0] / num_t

            # 高分组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.wy from kscj where ksh like \'"+dsh+"%\'  ORDER BY KSCJ.wy desc) a ) b " \
                  "where b.rn BETWEEN 1 and "+str(low)+") c on sxt.ksh = c.ksh where sxt.kmh=101 and sxt.xth="+str(zgth)+" GROUP BY sxt.ksh) d"
            print(sql)
            self.cursor.execute(sql)
            sum_h = float(self.cursor.fetchone()[0])
            total = total + sum_h
            dfl_h = sum_h / low / num

            # 中间组组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.wy from kscj where ksh like \'"+dsh+"%\' ORDER BY KSCJ.wy desc) a ) b " \
                  "where b.rn BETWEEN "+str(low+1)+" and " + str(high) + ") c on sxt.ksh = c.ksh where sxt.kmh=101 and sxt.xth=" + str(zgth) + " GROUP BY sxt.ksh) d"
            print(sql)
            self.cursor.execute(sql)
            sum_m = float(self.cursor.fetchone()[0])
            total = total + sum_m
            dfl_m = sum_m / (high - low) / num

            # 低分组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.wy from kscj where ksh like \'"+dsh+"%\'  ORDER BY KSCJ.wy desc) a ) b " \
                  "where b.rn BETWEEN " + str(high + 1) + " and " + str(num_ks) + ") c on sxt.ksh = c.ksh where sxt.kmh=101 and sxt.xth=" + str(zgth) + " GROUP BY sxt.ksh) d"
            print(sql)
            self.cursor.execute(sql)
            sum_l = float(self.cursor.fetchone()[0])
            total = total + sum_l
            dfl_l = sum_l / (num_ks - high) / num

            row.append(total/num_ks)  # 全市平均分
            row.append(avg_province)  # 全省平均分
            row.append(total / num_ks / num)  # 全市得分率
            row.append(dfl_h)  # 高分组
            row.append(dfl_m)  # 中间组
            row.append(dfl_l)  # 低分组

            self.set_list_precision(row)
            df.loc[len(df)] = row
            print(row)

        df.to_excel(excel_writer=writer,sheet_name="地市考生单题分析情况(英语)",index=False)
        writer.save()

    # 市级报告 单题分析 画图
    def DTFX_CITY_IMG(self,dsh):
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

        sql = r"select count(ksh) from (SELECT DISTINCT ksh from kscj where ksh like '01%' ) a"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = range(1,40)
        xths = range(61,82)

        x = []  # 难度
        y = []  # 区分度

        for idx in idxs:
            if idx in range(1,21):
                num = 2.0
            else:
                num = 1.5
            sql = r"select sum(kgval),sxt.ksh FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where amx.ksh like '"+dsh+"%' and kmh = 101 and idx = " + str(idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度

            # 前27%得分率
            sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select ksh,wy from (select ksh,wy,rownum rn from " \
                  r"(select ksh,wy from kscj where ksh like '" + dsh + "%' ORDER BY wy desc) a ) b " \
                  r"where b.rn BETWEEN 1 and " + str(ph_num) + ") c on amx.ksh = c.ksh where amx.kmh = 101 and amx.idx = " + str(idx)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select ksh,wy from (select ksh,wy,rownum rn from " \
                  r"(select ksh,wy from kscj where ksh like '" + dsh + "%' ORDER BY wy desc) a ) b " \
                  r"where b.rn BETWEEN " + str(total - ph_num) + r" and " + str(total) + r") c on amx.ksh = c.ksh where amx.kmh = 101 and amx.idx = " + str(idx)
            print(sql)
            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total - ph_num) / num

            x.append(difficulty)
            y.append(ph - pl)

        for xth in xths:
            if xth in range(61, 71):
                num = 1.5
            elif xth in range(71, 81):
                num = 1.00
            elif xth == 81:
                num = 25.00

            sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where sxt.ksh like '" + dsh + "%' and kmh=101 and xth=" + str(xth)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度
            x.append(difficulty)


            sql = r"select wy,b.sum from kscj right join " \
                  r"(select a.*,rownum rn from (select sum(xtval) sum,sxt.ksh from " \
                  r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where kmh = 101 and xth=" + str(xth) + r" and sxt.ksh " \
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

        plt.rcParams['figure.figsize'] = (15.0, 6)
        plt.scatter(x, y)
        plt.xlim((0, 1))
        plt.ylim((0, 1))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        for i in range(len(x)):
            plt.annotate(str(i+20), xy=(x[i], y[i]), xytext=(x[i] + 0.008, y[i] + 0.008),
                         arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(英语).png', dpi=1200)
        plt.close()

    # 市级报告附录 原始分分析
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题水平分析原始分概括(英语).xlsx")

        city_num = [0]*151
        province_num = [0]*151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,columns=['一分段','人数(本市)','百分比(本市)','累计百分比(本市)','人数(全省)','百分比(全省)','累计百分比(全省)'])


        # 地市
        sql = r"select wy,count(wy) from kscj where wy!=0 and ksh like '"+dsh+r"%' group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  #人数

        # 全省
        sql = r"select wy,count(wy) from kscj where wy!=0 group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            province_num[item[0]] = item[1]
            province_total += item[1] #人数

        i = 150
        acc_city = 0
        acc_province = 0
        while i>1:
            if city_num[i] > 0:
                acc_city += city_num[i] # 累计百分比
                acc_province += province_num[i] # 累计百分比
                row = []
                row.append(i)
                row.append(city_num[i]) # 本市人数
                row.append((city_num[i]/city_total)*100) # 本市百分比
                row.append((acc_city/city_total)*100) # 本市累计百分比

                row.append(province_num[i])
                row.append((province_num[i] / province_total)*100)  # 全省百分比
                row.append((acc_province / province_total)*100)  # 全省累计百分比
                self.set_list_precision(row)
                df.loc[len(df)] = row
            i = i - 1

        df.to_excel(excel_writer=writer,sheet_name='地市及全省考生一分段概括(英语)',index=None)


        # 文科生
        city_num = [0] * 151
        province_num = [0] * 151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,columns=['一分段', '人数(本市)', '百分比(本市)', '累计百分比(本市)', '人数(全省)', '百分比(全省)', '累计百分比(全省)'])

        # 地市
        sql = r"select wy,count(wy) from kscj where kl = 2 and wy!=0 and ksh like '" + dsh + r"%' group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  # 人数

        # 全省
        sql = r"select wy,count(wy) from kscj where kl=2 and wy!=0 group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            province_num[item[0]] = item[1]
            province_total += item[1]  # 人数

        i = 150
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
            i = i-1

        df.to_excel(excel_writer=writer, sheet_name='地市及全省文科考生一分段概括(英语)',index=None)

        # 理科生
        city_num = [0] * 151
        province_num = [0] * 151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,
                          columns=['一分段', '人数(本市)', '百分比(本市)', '累计百分比(本市)', '人数(全省)', '百分比(全省)', '累计百分比(全省)'])

        # 地市
        sql = r"select wy,count(wy) from kscj where kl = 1 and wy!=0 and ksh like '" + dsh + r"%' group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  # 人数

        # 全省
        sql = r"select wy,count(wy) from kscj where kl=1 and wy!=0 group by wy order by wy desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            province_num[item[0]] = item[1]
            province_total += item[1]  # 人数

        i = 150
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

        df.to_excel(excel_writer=writer, sheet_name='地市及全省理科考生一分段概括(英语)',index=None)

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

        writer = pd.ExcelWriter(path + '\\'  + "考生答题分析单题分析(英语).xlsx")

        rows = []
        sql = r"select count(*) from kscj "
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(1, 41)

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
            sql = r"select DA,count(DA) as num from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select * from (select a.*,rownum rn from (select ksh,wy from kscj " \
                  r" ORDER BY wy desc) a ) b" \
                  r" where b.rn BETWEEN 1 and " + str(low) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=101 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
            sql = r"select DA,count(DA) as num from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select * from (select a.*,rownum rn from (select ksh,wy from kscj " \
                  r" ORDER BY wy desc) a ) b" \
                  r" where b.rn BETWEEN " + str(low + 1) + " and " + str(high) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=101 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
            sql = r"select DA,count(DA) as num from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select * from (select a.*,rownum rn from (select ksh,wy from kscj " \
                  r" ORDER BY wy desc) a ) b" \
                  r" where b.rn BETWEEN " + str(high + 1) + " and " + str(total) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=101 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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

            row.append((a_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选A
            row.append((a_h / low) * 100)  # 高分组选A
            row.append((a_m / (high - low)) * 100)  # 中间组选A
            row.append((a_l / (total - high)) * 100)  # 低分组选A

            row.append((b_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选B
            row.append((b_h / low) * 100)  # 高分组选B
            row.append((b_m / (high - low)) * 100)  # 中间组选B
            row.append((b_l / (total - high)) * 100)  # 低分组选B

            row.append((c_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选C
            row.append((c_h / low) * 100)  # 高分组选C
            row.append((c_m / (high - low)) * 100)  # 中间组选C
            row.append((c_l / (total - high)) * 100)  # 低分组选C

            row.append((d_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选D
            row.append((d_h / low) * 100)  # 高分组选D
            row.append((d_m / (high - low)) * 100)  # 中间组选D
            row.append((d_l / (total - high)) * 100)  # 低分组选D

            self.set_list_precision(row)
            rows.append(row)

        df = pd.DataFrame(data=None, columns=["题号", "全部(A)", "高分组(A)", "中间组(A)", "低分组(A)",
                                              "全部(B)", "高分组(B)", "中间组(B)", "低分组(B)",
                                              "全部(C)", "高分组(C)", "中间组(C)", "低分组(C)",
                                              "全部(D)", "高分组(D)", "中间组(D)", "低分组(D)"])

        for i in range(len(rows)):
            rows[i].insert(0, i + 21)
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(英语)")
        writer.save()

    # 市级报告附录 单题分析
    def DTFX_CITY_APPENDIX(self,dsh):

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(英语).xlsx")

        rows = []
        sql = r"select count(*) from kscj where ksh like '"+dsh+r"%'"
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(1,41)

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
            sql = r"select DA,count(DA) as num from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select * from (select a.*,rownum rn from (select ksh,wy from kscj " \
                  r"where ksh like '"+dsh+r"%' ORDER BY wy desc) a ) b" \
                  r" where b.rn BETWEEN 1 and "+str(low)+r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=101 and amx.idx="+str(idx)+r" GROUP BY amx.da"
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
            sql = r"select DA,count(DA) as num from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select * from (select a.*,rownum rn from (select ksh,wy from kscj " \
                  r"where ksh like '" + dsh + r"%' ORDER BY wy desc) a ) b" \
                  r" where b.rn BETWEEN "+str(low+1)+" and " + str(high) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=101 and amx.idx=" +str(idx) + r" GROUP BY amx.da"
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
            sql = r"select DA,count(DA) as num from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select * from (select a.*,rownum rn from (select ksh,wy from kscj " \
                  r"where ksh like '" + dsh + r"%' ORDER BY wy desc) a ) b" \
                  r" where b.rn BETWEEN " + str(high+1) + " and " + str(total) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=101 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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

            row.append((a_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选A
            row.append((a_h / low) * 100)  # 高分组选A
            row.append((a_m / (high - low)) * 100)  # 中间组选A
            row.append((a_l / (total - high)) * 100)  # 低分组选A

            row.append((b_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选B
            row.append((b_h / low) * 100)  # 高分组选B
            row.append((b_m / (high - low)) * 100)  # 中间组选B
            row.append((b_l / (total - high)) * 100)  # 低分组选B

            row.append((c_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选C
            row.append((c_h / low) * 100)  # 高分组选C
            row.append((c_m / (high - low)) * 100)  # 中间组选C
            row.append((c_l / (total - high)) * 100)  # 低分组选C

            row.append((d_t / (a_t+b_t+c_t+d_t)) * 100)  # 全部选D
            row.append((d_h / low) * 100)  # 高分组选D
            row.append((d_m / (high - low)) * 100)  # 中间组选D
            row.append((d_l / (total - high)) * 100)  # 低分组选D

            self.set_list_precision(row)
            rows.append(row)

        df = pd.DataFrame(data=None,columns=["题号","全部(A)","高分组(A)","中间组(A)","低分组(A)",
                                             "全部(B)","高分组(B)","中间组(B)","低分组(B)",
                                             "全部(C)","高分组(C)","中间组(C)","低分组(C)",
                                             "全部(D)","高分组(D)","中间组(D)","低分组(D)"])

        for i in range(len(rows)):
            rows[i].insert(0,i+21)
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer,index=None,sheet_name="地市不同层次考生选择题受选率统计(英语)")
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

        writer = pd.ExcelWriter(path + '\\' + "各市情况分析(英语).xlsx")

        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])

        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where yw!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(yw),stddev_samp(wy) from kscj where wy!=0"
        self.cursor.execute(sql)
        item = self.cursor.fetchone()
        row.append(item[0])
        row.append((item[0] / total) * 100)
        row.append(item[1])
        row.append(item[2])
        row.append(item[2] / item[1])
        self.set_list_precision(row)
        df.loc[len(df)] = row

        for ds in dss:
            row = []
            row.append(ds[0])
            row.append(ds[1])


            sql = r"select count(*) as num,avg(wy),stddev_samp(yw) from kscj where wy!=0 and ksh like '" + ds[
                0] + r"%'"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市考生成绩比较(英语)")

        # 文科
        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])
        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where wy!=0 and kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(wy),stddev_samp(wy) from kscj where wy!=0 and kl=2"
        self.cursor.execute(sql)
        item = self.cursor.fetchone()

        row.append(item[0])
        row.append((item[0] / total) * 100)
        row.append(item[1])
        row.append(item[2])
        row.append(item[2] / item[1])
        self.set_list_precision(row)
        df.loc[len(df)] = row

        for ds in dss:
            row = []
            row.append(ds[0])
            row.append(ds[1])


            sql = r"select count(*) as num,avg(wy),stddev_samp(wy) from kscj where wy!=0 and ksh like '" + ds[0] + r"%' and kl=2"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市文科考生成绩比较(英语)")

        # 理科
        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])
        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where wy!=0 and kl=1"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(wy),stddev_samp(wy) from kscj where yw!=0 and kl=1"
        self.cursor.execute(sql)
        item = self.cursor.fetchone()
        row.append(item[0])
        row.append((item[0] / total) * 100)
        row.append(item[1])
        row.append(item[2])
        row.append(item[2] / item[1])
        self.set_list_precision(row)
        df.loc[len(df)] = row

        for ds in dss:
            row = []
            row.append(ds[0])
            row.append(ds[1])


            sql = r"select count(*) as num,avg(wy),stddev_samp(wy) from kscj where wy!=0 and ksh like '" + ds[0] + r"%' and kl=1"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市理科考生成绩比较(英语)")

        writer.save()

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

        writer = pd.ExcelWriter(path + '\\' + "原始分概括(英语).xlsx")

        sql = "select count(*) from kscj where wy!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = "select wy,count(wy) from kscj where wy!=0 group by (wy) order by wy desc"
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

        df.to_excel(writer, index=None, sheet_name="全省考生一分段(英语)")

        sql = "select count(*) from kscj where wy!=0 and kl=1"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = "select wy,count(wy) from kscj where yw!=0 and kl=1  group by (wy) order by wy desc"
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

        df.to_excel(writer, index=None, sheet_name="全省理科考生一分段(英语)")

        sql = "select count(*) from kscj where wy!=0 and kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = "select wy,count(wy) from kscj where wy!=0 and kl=2  group by (wy) order by wy desc"
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

        df.to_excel(writer, index=None, sheet_name="全省文科考生一分段(英语)")
        writer.save()

    # 省级报告 单题分析(图、表)
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

        writer = pd.ExcelWriter(path + '\\' + "考生单题分析(英语).xlsx")
        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])

        idxs = list(range(1,41))
        xths = list(range(61,82))

        x = []
        y = []

        sql = "select count(yw) from kscj"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total*0.27)

        rows = []

        for idx in idxs:
            row = []

            if idx in range(1, 21):
                num = 2.0
            else:
                num = 1.5
            row.append(str(idx+20))
            row.append(num)

            sql = "select avg(kgval),STDDEV_SAMP(kgval) from T_GKPJ2020_TKSKGDAMX amx right join kscj" \
                  " on kscj.ksh=amx.ksh where amx.kmh=101 and amx.idx = "+str(idx)
            print(sql)
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            mean = result[0]
            std = result[1]
            diffculty = mean/num

            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx " \
                  "right join (select b.* from (select a.*,rownum rn from " \
                  "(select ksh,yw from kscj order by yw desc) a) b where rn BETWEEN 1 and "+str(ph_num)+") c " \
                  "on c.ksh = amx.ksh where kmh = 101 and idx = "+str(idx)
            print(sql)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx " \
                  "right join (select b.* from (select a.*,rownum rn from " \
                  "(select ksh,yw from kscj order by yw desc) a) b where rn BETWEEN "+str(total-ph_num)+" and " + str(total) + ") c on " \
                  "c.ksh = amx.ksh where kmh = 101 and idx = "+str(idx)
            print(sql)
            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / ph_num / num

            qfd = ph - pl

            row.append(mean)
            row.append(std)
            row.append(diffculty)
            row.append(qfd)
            print(row)
            self.set_list_precision(row)
            rows.append(row)

            x.append(diffculty)
            y.append(qfd)

        for xth in xths:
            row = []
            print(xth)
            if xth in range(61, 71):
                num = 1.5
            elif xth in range(71, 81):
                num = 1.00
            elif xth == 81:
                num = 25.00
            row.append(str(xth))
            row.append(num)

            sql = "select avg(xtval),STDDEV_SAMP(xtval) from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join kscj on sxt.ksh = kscj.ksh where kmh = 101 and xth =" + str(xth)
            print(sql)
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            mean = result[0]
            std = result[1]
            diffculty = mean/num

            sql = "select yw,b.sum from kscj right join " \
                  "(select a.*,rownum rn from (select sum(xtval)  sum,sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join kscj on kscj.ksh = sxt.ksh where kmh = 101 and xth="+str(xth)+" GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn"

            self.cursor.execute(sql)
            result = np.array(self.cursor.fetchall(),dtype="float64")

            zf_score = np.array(result[:, 0],dtype="float64")
            xt_score = np.array(result[:,1],dtype="float64")

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

            qfd = (D_a - D_b) / (math.sqrt(D_c) * math.sqrt(D_d))

            row.append(mean)
            row.append(std)
            row.append(diffculty)
            row.append(qfd)

            self.set_list_precision(row)
            rows.append(row)
            print(row)
            x.append(diffculty)
            y.append(qfd)

        for i in range(len(rows)):
            df.loc[len(df)] = rows[i]

        df.to_excel(writer,index=None,sheet_name="考生单题作答情况(英语)")
        writer.save()

        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        plt.xlim((0, 1))
        plt.ylim((0, 1))
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        plt.scatter(x, y)

        for i in range(len(x)):
            plt.annotate(rows[i][0], xy=(x[i], y[i]), xytext=(x[i] + 0.008, y[i] + 0.008),
                         arrowprops=dict(arrowstyle='->',connectionstyle="arc3,rad = .2"))
        plt.savefig(path + '\\各题难度-区分度分布散点图(英语).png', dpi=1200)
        plt.close()