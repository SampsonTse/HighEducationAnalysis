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


plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
np.set_printoptions(precision=2)

# 语文考生答题水平分析
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

    # 市级报告 总体概括 制表
    def ZTGK_CITY_TABLE(self,dsh):

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析总体概括(语文).xlsx")


        df = pd.DataFrame(data=None,columns=['维度','人数','比率(%)','平均分','标准差','差异系数','平均分(全省)'])

        sql = r'select count(a.YW) from kscj  a right join JBXX b on a.KSH = b.KSH ' \
              r'WHERE b.DS_H='+dsh+' and a.yw!=0'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0] # 总人数


        # 计算维度为男
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj  a right join JBXX  b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and b.XB_H = 1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append(((float(result[2]) / float(result[1]))*100)*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH where b.XB_H = 1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1,(result[0]/num)*100)
        result.insert(0,'男')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and b.XB_H = 2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH where b.XB_H = 2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '女')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '城镇')


        self.set_list_precision(result)
        df.loc[len(df)] = result


        # 计算维度为农村
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '农村')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '应届')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '往届')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX  b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH where a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '总计')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别考生成绩比较(语文)",excel_writer=writer,index=None)



        # 文科
        df = pd.DataFrame(data=None,columns=['维度','人数','比率(%)','平均分','标准差','差异系数','平均分(全省)'])

        sql = r'select count(a.YW) from kscj   a right join JBXX   b on a.KSH = b.KSH ' \
              r'WHERE b.DS_H='+dsh+r' and a.kl=2 and a.yw!=0'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and b.XB_H = 1 and a.kl=2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 1 and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '男')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.YW)  num,AVG(A.YW)  mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H="+dsh+r" and b.XB_H = 2 and a.kl=2 and a.yw!=0"
        print(sql)
        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 2 and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '女')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '城镇')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '农村')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+" and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '应届')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '往届')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and a.kl=2 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH and a.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '总计')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别文科考生成绩比较(语文)", excel_writer=writer,index=None)

        # 理科
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r'select count(a.YW) from kscj  a right join JBXX  b on a.KSH = b.KSH WHERE b.DS_H='+dsh+r' and a.kl=1'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H="+dsh+r" and b.XB_H = 1 and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        self.cursor.execute(sql)
        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where b.XB_H = 1 and a.kl=1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '男')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and b.XB_H = 2 and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 2 and a.kl=1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '女')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '城镇')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '农村')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '应届')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H="+dsh+r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=1 and a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '往届')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.YW)   num,AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H="+dsh+r" and a.kl=1 and a.yw!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数

        sql = r"select AVG(A.YW)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH and a.kl=1 where a.yw!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num)*100)
        result.insert(0, '总计')


        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别理科考生成绩比较(语文)", excel_writer=writer,index=None)

        # 各区县考生成绩比较
        sql = r"select xq_h,mc from c_xq where  xq_h like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A where  a.yw!=0 "
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A " \
              r"where a.yw!=0 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = "select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE  a.yw!=0 and B.XQ_H = " + xqh[0]
            self.cursor.execute(sql)
            print(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
            result.append(result[1] / 150)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区考生成绩比较(语文)",index=None)

        # 各区县理考生成绩比较

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A where A.kl=1"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A " \
              r"where A.kl=1 and a.yw!=0 and A.KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = "select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE A.kl=1 and a.yw!=0 and B.XQ_H = " + xqh[0]
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
            result.append(result[1] / 150)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区理科考生成绩比较(语文)",index=None)

        # 各区县文科考生成绩比较

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A where A.kl=2 and a.yw!=0"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A " \
              r"where A.kl=2 and a.yw!=0 and A.KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全市')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        for xqh in xqhs:
            result = []
            sql = "select count(YW),AVG(A.YW)   mean,STDDEV_SAMP(A.YW)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE A.kl=2 and a.yw!=0 and B.XQ_H = " + xqh[0]
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            result = list(result)
            if None in result:
                continue
            result.append((float(result[2]) / float(result[1]))*100)  # 差异系数
            result.append(result[1] / 150)
            result.insert(0, xqh[1])
            self.set_list_precision(result)
            df.loc[len(df)] = result

        df.to_excel(excel_writer=writer, sheet_name="各县区文科考生成绩比较(语文)",index=None)

        writer.save()

    # 市级报告 总体概括 画图
    def ZTGK_CITY_IMG(self,dsh):

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
        # plt.figure()
        plt.rcParams['figure.figsize'] = (15.0,6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(YW) FROM kscj where yw!=0"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0] # 全省人数

        sql = r"SELECT YW,COUNT(YW) FROM kscj WHERE YW !=0 GROUP BY YW"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1]/num * 100,2)
        x = list(range(151))

        plt.plot(x,province,color='orange',marker='.',label='全省')

        # 全市
        sql = "SELECT COUNT(YW) FROM kscj where KSH LIKE '"+dsh+r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数


        sql = r"SELECT YW,COUNT(YW) FROM kscj WHERE YW != 0 and KSH LIKE '"+dsh+r"%' GROUP BY  YW"
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
        plt.savefig(path + '\\地市及全省考生单科成绩分布(语文).png', dpi=1200)
        plt.close()



        # 全省文科
        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0,6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(YW) FROM kscj where kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "SELECT YW,COUNT(YW) FROM kscj WHERE YW != 0 and kl=2 GROUP BY  YW "
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(151))

        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市文科
        sql = "SELECT COUNT(YW) FROM kscj where kl=2 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全市人数

        sql = r"SELECT YW,COUNT(YW) FROM kscj WHERE YW != 0 and kl=2 and KSH LIKE '" + dsh + r"%' GROUP BY  YW"
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
        plt.savefig(path + '\\地市及全省文科考生单科成绩分布(语文).png', dpi=1200)
        plt.close()




        # 全省理科
        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(YW) FROM kscj where kl=1"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "SELECT YW,COUNT(YW) FROM kscj WHERE YW != 0 and kl=1 GROUP BY  YW "
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(151))

        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市理科
        plt.rcParams['figure.figsize'] = (15.0, 6)
        sql = "SELECT COUNT(YW) FROM kscj where kl=1 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全市人数

        sql = r"SELECT YW,COUNT(YW) FROM kscj WHERE YW != 0 and kl=1 and KSH LIKE '" + dsh + r"%' GROUP BY  YW"
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
        plt.savefig(path + '\\地市及全省理科考生单科成绩分布(语文).png', dpi=1200)
        plt.close()

    # 省级报告 总体概括
    def ZTGK_PROVINCE_TABLE(self):

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

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析原始分概括(语文).xlsx")

        # 全省考生
        df = pd.DataFrame(data=None,columns=['维度','人数','比率(%)','平均分','标准差','差异系数'])

        sql = "select count(*) from kscj  a right join JBXX  b on a.ksh = b.ksh"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh where b.xb_h="+ str(xb)
            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1,results[0]/num * 100)  #比率
            if xb == 1 :
                results.insert(0,'男')
            else:
                results.insert(0,'女')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 户籍
        for hj in hjs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh where b.kslb_h = "+str(hj[0])+" or b.kslb_h = "+hj[1]

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in hj:
                results.insert(0,'城镇')
            else:
                results.insert(0,'农村')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 应往届
        for ywj in ywjs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh where b.kslb_h = "+ywj[0]+" or b.kslb_h = "+ywj[1]

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in ywj:
                results.insert(0, '应届')
            else:
                results.insert(0, '往届')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        sql = "select count(a.YW)  num,AVG(a.YW)  mean,STDDEV_SAMP(a.YW)  std " \
              "from kscj a right join JBXX  b on a.ksh = b.ksh"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0,'总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer,sheet_name="各类别考生成绩比较(语文)",index=None)



        # 全省文科考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(a.YW)  num " \
              "from kscj  a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=2 and b.xb_h=" + str(xb)
            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if xb == 1:
                results.insert(0, '男')
            else:
                results.insert(0, '女')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 户籍
        for hj in hjs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=2 and (b.kslb_h = "+hj[0]+" or b.kslb_h = "+hj[1]+")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in hj:
                results.insert(0, '城镇')
            else:
                results.insert(0, '农村')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 应往届
        for ywj in ywjs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=2 and (b.kslb_h = "+ywj[0]+" or b.kslb_h = "+ywj[1]+")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in ywj:
                results.insert(0, '应届')
            else:
                results.insert(0, '往届')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
              "from kscj   a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别文科考生成绩比较(语文)", index=None)

        # 全省理科考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(*) from kscj   a right join JBXX   b on a.ksh = b.ksh where a.kl=1"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=1 and b.xb_h=" + str(xb)
            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if xb == 1:
                results.insert(0, '男')
            else:
                results.insert(0, '女')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 户籍
        for hj in hjs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=1 and (b.kslb_h = "+hj[0]+" or b.kslb_h = "+hj[1]+")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in hj:
                results.insert(0, '城镇')
            else:
                results.insert(0, '农村')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        # 应往届
        for ywj in ywjs:
            sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
                  "from kscj   a right join JBXX   b on a.ksh = b.ksh " \
                  "where a.kl=1 and (b.kslb_h = "+ywj[0]+" or b.kslb_h = "+ywj[1]+")"

            self.cursor.execute(sql)
            results = self.cursor.fetchone()
            results = list(results)
            results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
            results.insert(1, results[0] / num * 100)  # 比率
            if "1" in ywj:
                results.insert(0, '应届')
            else:
                results.insert(0, '往届')

            self.set_list_precision(results)
            df.loc[len(df)] = results

        sql = "select count(a.YW)   num,AVG(a.YW)   mean,STDDEV_SAMP(a.YW)   std " \
              "from kscj   a right join JBXX  b on a.ksh = b.ksh where a.kl=1"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1]))*100*100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别理科考生成绩比较(语文)", index=None)

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

        writer = pd.ExcelWriter(path + '\\' + "考生单题分析(语文).xlsx")
        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])

        idxs = [1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13]
        xths = [6, 8, 9, 15, 16, 20, 21, 22]
        txt = ["01", "02", "03", "04", "05", "07", "10", "11", "12", "14", "17", "18", "19", "06", "08", "09", "15",
               "16", "20", "21", "22"]

        x = []
        y = []

        sql = "select count(yw) from kscj"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total*0.27)

        rows = []

        # for idx in idxs:
        #     row = []
        #
        #     num = 3.0
        #     row.append(num)
        #
        #     sql = "select avg(kgval),STDDEV_SAMP(kgval) from T_GKPJ2020_TKSKGDAMX amx right join kscj" \
        #           " on kscj.ksh=amx.ksh where amx.kmh=001 and amx.idx = "+str(idx)
        #     print(sql)
        #     self.cursor.execute(sql)
        #     result = self.cursor.fetchone()
        #     mean = result[0]
        #     std = result[1]
        #     diffculty = mean/num
        #
        #     sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx " \
        #           "right join (select b.* from (select a.*,rownum rn from " \
        #           "(select ksh,yw from kscj order by yw desc) a) b where rn BETWEEN 1 and "+str(ph_num)+") c " \
        #           "on c.ksh = amx.ksh where kmh = 001 and idx = "+str(idx)
        #     print(sql)
        #     self.cursor.execute(sql)
        #     ph = self.cursor.fetchone()[0] / ph_num / num
        #
        #     sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx " \
        #           "right join (select b.* from (select a.*,rownum rn from " \
        #           "(select ksh,yw from kscj order by yw desc) a) b where rn BETWEEN "+str(total-ph_num)+" and " + str(total) + ") c on " \
        #           "c.ksh = amx.ksh where kmh = 001 and idx = "+str(idx)
        #     print(sql)
        #     self.cursor.execute(sql)
        #     pl = self.cursor.fetchone()[0] / ph_num / num
        #
        #     qfd = ph - pl
        #
        #     row.append(mean)
        #     row.append(std)
        #     row.append(diffculty)
        #     row.append(qfd)
        #     print(row)
        #     self.set_list_precision(row)
        #     rows.append(row)
        #
        #     x.append(diffculty)
        #     y.append(qfd)

        for xth in xths:
            row = []
            print(xth)
            if xth in [6, 8, 9, 15, 16, 20]:
                num = 6.0
            elif xth == 21:
                num = 5.0
            elif xth == 22:
                num = 60.00
            row.append(num)

            sql = "select avg(xtval),STDDEV_SAMP(xtval) from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join kscj on sxt.ksh = kscj.ksh where kmh = 001 and dth =" + str(xth)
            print(sql)
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            mean = result[0]
            std = result[1]
            diffculty = mean/num

            sql = r"select a.*,rownum rn from (select sum(xtval),sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt right join " \
                  r"kscj on kscj.ksh = sxt.ksh where kmh = 001 and dth="+str(xth)+r" GROUP BY sxt.ksh) a"
            print(sql)
            self.cursor.execute(sql)
            xt_score = np.array(self.cursor.fetchall(), dtype='float64')
            xt_score = np.delete(xt_score, [1,2], axis=1).flatten()
            sql = r"select yw,kscj.ksh,b.rn from kscj right join (select a.*,rownum rn from " \
                  r"(select sum(xtval),sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt right join kscj on " \
                  r"kscj.ksh = sxt.ksh where kmh = 001 and dth="+str(xth)+" GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn"
            print(sql)
            self.cursor.execute(sql)
            zf_score = np.array(self.cursor.fetchall(), dtype='float64')
            print(zf_score)
            zf_score = np.delete(xt_score, [1, 2], axis=1).flatten()

            print(zf_score)
            print(xt_score)

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


        num = 10.0

        sql = "select avg(xtval),STDDEV_SAMP(xtval) from T_GKPJ2020_TSJBNKSXT sxt " \
              "right join kscj on sxt.ksh = kscj.ksh where kmh = 001 and (dth=13 or dth=23)"

        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        mean = result[0]
        std = result[1]
        diffculty = mean / num

        sql = r"select sum(xtval),sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
              r"where sxt.kmh = 001 and (sxt.dth=13 or sxt.dth=23)  GROUP BY sxt.ksh"
        print(sql)
        self.cursor.execute(sql)
        xt_score = np.array(self.cursor.fetchall(), dtype='float64')
        xt_score = np.delete(xt_score, -1, axis=1).flatten()
        sql = r"select yw from kscj right join " \
              r"(select a.*,rownum rn from (select sum(xtval),sxt.ksh from " \
              r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
              r"where kmh = 001 and (dth=13 or dth=23) GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn "
        print(sql)
        self.cursor.execute(sql)
        zf_score = np.array(self.cursor.fetchall(), dtype='float64').flatten()

        self.cursor.execute(sql)
        zf_score = np.array(self.cursor.fetchall()).flatten()

        n = len(xt_score)

        D_a = n * np.sum(xt_score * zf_score)
        D_b = np.sum(zf_score) * np.sum(xt_score)
        D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
        D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

        qfd = (D_a - D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
        print(qfd)
        x.append(diffculty)
        y.append(qfd)

        row.append(num)
        row.append(mean)
        row.append(std)
        row.append(diffculty)
        row.append(qfd)
        self.set_list_precision(row)
        rows.append(row)
        print(row)


        for i in range(len(rows)):
            rows[i].insert(0,txt[0])
            df.loc[len(df)] = rows[i]

        df.to_excel(writer,index=None,sheet_name="考生单题作答情况(语文)")
        writer.save()

        plt.xlim((0, 1))
        plt.ylim((0, 1))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        plt.scatter(x, y)

        for i in range(len(x)):
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] + 0.008, y[i] + 0.008),
                         arrowprops=dict(arrowstyle='->',connectionstyle="arc3,rad = .2"))
        plt.savefig(path + '\\各题难度-区分度分布散点图(语文).png', dpi=1200)

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(语文).xlsx")

        sql = r"select count(*) from kscj where ksh like '"+dsh+r"%'  "
        self.cursor.execute(sql)
        num_ks = self.cursor.fetchone()[0]

        sql = r"select count(*) from kscj "
        self.cursor.execute(sql)
        num_t = self.cursor.fetchone()[0]

        low = int(num_ks/3)
        high = int(num_ks/1.5)

        df = pd.DataFrame(data=None,columns=['题号','分值','本市平均分','全省平均分','本市得分率','高分组得分率','中间组得分率','低分组得分率'])

        kgths = list(range(1,14))
        zgths = [ 6,8,9, 15, 16, 20, 21, 22]

        for kgth in kgths:
            row = []
            num = 3.0
            if kgth in [1,2,3,4,5]:
                row.append(kgth)
            elif kgth ==6:
                row.append(7)
            elif kgth in [7,8,9]:
                row.append(kgth+3)
            elif kgth ==10:
                row.append(14)
            elif kgth in [11,12,13]:
                row.append(kgth+6)

            row.append(3)

            total = 0

            # 全省平均分
            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX a right join jbxx b " \
                  "on a.ksh = b.ksh where a.idx=" + str(kgth) + " and kmh=001"
            self.cursor.execute(sql)
            avg_province = self.cursor.fetchone()[0] / num_t

            # 本市计算高分组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ left JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE jbxx.ds_h=" + dsh + " ORDER BY KSCJ.yw DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=001 and b.rn BETWEEN 1 and " + str(low)
            print(sql)
            self.cursor.execute(sql)
            sum_h = float(self.cursor.fetchone()[0])
            total = total + sum_h
            dfl_h = sum_h / low / num

            # 本市计算中间组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ left JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE jbxx.ds_h=" + dsh + " ORDER BY KSCJ.yw DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=001 and b.rn BETWEEN " + str(low + 1) + " and " + str(high)
            self.cursor.execute(sql)
            sum_m = float(self.cursor.fetchone()[0])
            total = total + sum_m
            dfl_m = sum_m / (high - low) / num

            # 本市计算低分组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ left JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE jbxx.ds_h=" + dsh + " ORDER BY KSCJ.yw DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=001 and b.rn BETWEEN " + str(high + 1) + " and " + str(num_ks)
            self.cursor.execute(sql)
            sum_l = float(self.cursor.fetchone()[0])
            total = total + sum_l
            dfl_l = sum_l / (num_ks - high) / num

            row.append(total / num_ks)  # 全市平均分
            row.append(avg_province)  # 全省平均分
            row.append(total / num_ks / 3)  # 全市得分率
            row.append(dfl_h)  # 高分组
            row.append(dfl_m)  # 中间组
            row.append(dfl_l)  # 低分组

            self.set_list_precision(row)
            print(row)
            df.loc[len(df)] = row

        for zgth in zgths:

            row = []
            num = 0
            row.append(str(zgth))
            if zgth in [6,8,9,15,16,20]:
                num = 6.0
            elif zgth == 21:
                num = 5.0
            elif zgth == 22:
                num = 60.00
            row.append(num)

            total = 0

            # 全省平均分
            sql = "select sum(b.sum) from " \
                  "(select sum(a.xtval) as sum,a.dth,a.ksh from T_GKPJ2020_TSJBNKSXT a " \
                  "left join jbxx on jbxx.ksh=a.ksh where a.kmh = 001 and a.dth = " + str(
                zgth) + " GROUP BY a.ksh,a.dth) b"
            self.cursor.execute(sql)
            avg_province = self.cursor.fetchone()[0] / num_t

            # 高分组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT SXT " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.yw from kscj where ksh like \'" + dsh + "%\'  ORDER BY KSCJ.yw desc) a ) b " \
                  "where b.rn BETWEEN 1 and " + str(low) + ") c on SXT.ksh = c.ksh " \
                  "where SXT.kmh=001 and SXT.dth=" + str(zgth) + " GROUP BY SXT.ksh) d"
            self.cursor.execute(sql)
            sum_h = float(self.cursor.fetchone()[0])
            total = total + sum_h
            dfl_h = sum_h / low / num

            # 中间组组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT SXT " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.yw from kscj where ksh like \'" + dsh + "%\' ORDER BY KSCJ.yw desc) a ) b " \
                  "where b.rn BETWEEN " + str(low + 1) + " and " + str(high) + ") c on SXT.ksh = c.ksh " \
                  "where SXT.kmh=001 and SXT.dth=" + str(zgth) + " GROUP BY SXT.ksh) d"
            self.cursor.execute(sql)
            sum_m = float(self.cursor.fetchone()[0])
            total = total + sum_m
            dfl_m = sum_m / (high - low) / num

            # 低分组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT SXT " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.yw from kscj where ksh like \'" + dsh + "%\'   ORDER BY KSCJ.yw desc) a ) b " \
                  "where b.rn BETWEEN " + str(high + 1) + " and " + str(num_ks) + ") c on SXT.ksh = c.ksh " \
                  "where SXT.kmh=001 and SXT.dth=" + str(zgth) + " GROUP BY SXT.ksh) d"

            self.cursor.execute(sql)
            sum_l = float(self.cursor.fetchone()[0])
            total = total + sum_l
            dfl_l = sum_l / (num_ks - high) / num

            row.append(total / num_ks)  # 全市平均分
            row.append(avg_province)  # 全省平均分
            row.append(total / num_ks / num)  # 全市得分率
            row.append(dfl_h)  # 高分组
            row.append(dfl_m)  # 中间组
            row.append(dfl_l)  # 低分组

            self.set_list_precision(row)
            df.loc[len(df)] = row
            print(row)

        # 13 23
        row = []
        num = 10.0
        total = 0
        row.append(13)
        row.append(10.0)
        # 全省平均分
        sql = "select sum(b.sum) from " \
              "(select sum(a.xtval) as sum,a.dth,a.ksh from T_GKPJ2020_TSJBNKSXT a " \
              "right join jbxx on jbxx.ksh=a.ksh where a.kmh = 001 and (a.dth=13 or a.dth= 23) GROUP BY a.ksh,a.dth) b"
        self.cursor.execute(sql)
        avg_province = self.cursor.fetchone()[0] / num_t

        # 高分组得分率
        sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT SXT " \
              "right join (select b.* from (SELECT a.*,rownum rn from " \
              "(select kscj.ksh,kscj.yw from kscj where ksh like \'" + dsh + "%\'  ORDER BY KSCJ.yw desc) a ) b " \
              "where b.rn BETWEEN 1 and " + str(low) + ") c on SXT.ksh = c.ksh where SXT.kmh=001 and (SXT.dth=13 or SXT.dth=23)  GROUP BY SXT.ksh) d"
        self.cursor.execute(sql)
        sum_h = float(self.cursor.fetchone()[0])
        total = total + sum_h
        dfl_h = sum_h / low / num

        # 中间组组得分率
        sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT SXT " \
              "right join (select b.* from (SELECT a.*,rownum rn from " \
              "(select kscj.ksh,kscj.yw from kscj where ksh like \'" + dsh + "%\'  ORDER BY KSCJ.yw desc) a ) b " \
              "where b.rn BETWEEN " + str(
            low + 1) + " and " + str(high) + ") c on SXT.ksh = c.ksh where SXT.kmh=001 and (SXT.dth=13 or SXT.dth= 23) GROUP BY SXT.ksh) d"
        self.cursor.execute(sql)
        sum_m = float(self.cursor.fetchone()[0])
        total = total + sum_m
        dfl_m = sum_m / (high - low) / num

        # 低分组得分率
        sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT SXT " \
              "right join (select b.* from (SELECT a.*,rownum rn from " \
              "(select kscj.ksh,kscj.yw from kscj where ksh like \'" + dsh + "%\'  ORDER BY KSCJ.yw desc) a ) b " \
              "where b.rn BETWEEN " + str(high + 1) + " and " + str(num_ks) + ") c on SXT.ksh = c.ksh where SXT.kmh=001 and SXT.dth=13 or SXT.dth= 23 GROUP BY SXT.ksh) d"

        self.cursor.execute(sql)
        sum_l = float(self.cursor.fetchone()[0])
        total = total + sum_l
        dfl_l = sum_l / (num_ks - high) / num

        row.append(total / num_ks)  # 全市平均分
        row.append(avg_province)  # 全省平均分
        row.append(total / num_ks / num)  # 全市得分率
        row.append(dfl_h)  # 高分组
        row.append(dfl_m)  # 中间组
        row.append(dfl_l)  # 低分组

        self.set_list_precision(row)
        df.loc[len(df)] = row


        df.to_excel(excel_writer=writer,sheet_name="地市考生单题分析情况(语文)",index=False)
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

        sql = r"select count(ksh) from (SELECT DISTINCT ksh from kscj where ksh like '"+dsh+r"%') a"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = [1,2,3,4,5,7,8,9,10,11,12,13]
        xths = [6,8,9,15,16,20,21,22]

        x = [] # 难度
        y = [] # 区分度

        for idx in idxs:
            num = 3
            sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where amx.ksh like '"+dsh+"%' and kmh = 001 and idx = " + str(idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num #难度

            # 前27%得分率
            sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select ksh,yw from (select ksh,yw,rownum rn from " \
                  r"(select ksh,yw from kscj where ksh like '"+dsh+"%' ORDER BY yw desc) a ) b " \
                  r"where b.rn BETWEEN 1 and "+str(ph_num)+") c on amx.ksh = c.ksh where amx.kmh = 001 and amx.idx = "+str(idx)
            print(sql)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select ksh,yw from (select ksh,yw,rownum rn from " \
                  r"(select ksh,yw from kscj where ksh like '" + dsh + "%' ORDER BY yw desc) a ) b " \
                  r"where b.rn BETWEEN "+str(total-ph_num)+r" and " + str(total) + r") c on amx.ksh = c.ksh where amx.kmh = 001 and amx.idx = " + str(idx)
            print(sql)
            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total-ph_num) / num

            x.append(difficulty)
            y.append(ph-pl)

        for xth in xths:
            if xth in [6,8,9,15,16]:
                num = 6.0
            elif xth in [21]:
                num = 5.0
            elif xth in [22]:
                num = 60.0

            sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where sxt.ksh like '"+dsh+"%' and kmh=001 and dth="+str(xth)
            self.cursor.execute(sql)
            print(sql)
            difficulty = self.cursor.fetchone()[0] / total / num # 难度
            x.append(difficulty)


            sql = r"select sum(xtval),sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where sxt.kmh = 001 and sxt.dth="+str(xth)+" and sxt.ksh like '"+dsh+r"%' GROUP BY sxt.ksh"
            print(sql)
            self.cursor.execute(sql)
            xt_score = np.array(self.cursor.fetchall(), dtype='float64')
            xt_score = np.delete(xt_score, -1, axis=1).flatten()
            sql = r"select yw from kscj right join " \
                  r"(select a.*,rownum rn from (select sum(xtval),sxt.ksh, from " \
                  r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where kmh = 001 and dth="+str(xth)+r" and sxt.ksh " \
                  r"like '"+dsh+r"%' GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn "
            self.cursor.execute(sql)
            zf_score = np.array(self.cursor.fetchall(),dtype='float64').flatten()

            n = len(xt_score)

            D_a = n * np.sum(xt_score * zf_score)
            D_b = np.sum(zf_score) * np.sum(xt_score)
            D_c = n * np.sum(xt_score**2) - np.sum(xt_score)**2
            D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score)**2


            qfd = (D_a-D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
            y.append(qfd)
            print(y)

        num = 10.0

        sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
              r"where sxt.ksh like '" + dsh + "%' and kmh=001 and (dth=13 or dth=23)"
        self.cursor.execute(sql)
        difficulty = self.cursor.fetchone()[0] / total / num  # 难度
        x.append(difficulty)

        sql = r"select sum(xtval),sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
              r"where sxt.kmh = 001 and (dth=13 or dth=23) and sxt.ksh like '" + dsh + r"%' GROUP BY sxt.ksh"
        self.cursor.execute(sql)
        xt_score = np.array(self.cursor.fetchall(), dtype='float64')
        xt_score = np.delete(xt_score, -1, axis=1).flatten()

        sql = r"select yw from kscj right join " \
              r"(select a.*,rownum rn from (select sum(xtval),sxt.ksh from " \
              r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
              r"where kmh = 001 and (dth=13 or dth=23) and sxt.ksh " \
              r"like '" + dsh + r"%' GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn "
        print(sql)
        self.cursor.execute(sql)
        zf_score = np.array(self.cursor.fetchall(),dtype='float64').flatten()

        n = len(xt_score)

        D_a = n * np.sum(xt_score * zf_score)
        D_b = np.sum(zf_score) * np.sum(xt_score)
        D_c = n * np.sum(xt_score ** 2) - np.sum(xt_score) ** 2
        D_d = n * np.sum(zf_score ** 2) - np.sum(zf_score) ** 2

        qfd = (D_a-D_b) / (math.sqrt(D_c) * math.sqrt(D_d))
        y.append(qfd)

        txt = ["01","02","03","04","05","07","10","11","12","14","17","18","19","06","08","09","15","16","20","21","22","13"]


        print(x,y)

        plt.scatter(x,y)
        plt.xlim((0, 1))
        plt.ylim((0, 1))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        for i in range(len(x)):
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] + 0.008, y[i] + 0.008),arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(语文).png', dpi=1200)
        plt.show()

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题水平分析原始分概括(语文).xlsx")

        city_num = [0]*151
        province_num = [0]*151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,columns=['一分段','人数(本市)','百分比(本市)','累计百分比(本市)','人数(全省)','百分比(全省)','累计百分比(全省)'])


        # 地市
        sql = r"select yw,count(yw) from kscj where yw!=0 and ksh like '"+dsh+r"%' group by yw order by yw desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  #人数

        # 全省
        sql = r"select yw,count(yw) from kscj where yw!=0 group by yw order by yw desc"
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

        df.to_excel(excel_writer=writer,sheet_name='地市及全省考生一分段概括(语文)',index=None)


        # 文科生
        city_num = [0] * 151
        province_num = [0] * 151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,columns=['一分段', '人数(本市)', '百分比(本市)', '累计百分比(本市)', '人数(全省)', '百分比(全省)', '累计百分比(全省)'])

        # 地市
        sql = r"select yw,count(yw) from kscj where kl = 2 and yw!=0 and ksh like '" + dsh + r"%' group by yw order by yw desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  # 人数

        # 全省
        sql = r"select yw,count(yw) from kscj where kl=2 and yw!=0 group by yw order by yw desc"
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

        df.to_excel(excel_writer=writer, sheet_name='地市及全省文科考生一分段概括(语文)',index=None)

        # 理科生
        city_num = [0] * 151
        province_num = [0] * 151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,
                          columns=['一分段', '人数(本市)', '百分比(本市)', '累计百分比(本市)', '人数(全省)', '百分比(全省)', '累计百分比(全省)'])

        # 地市
        sql = r"select yw,count(yw) from kscj where kl = 1 and yw!=0 and ksh like '" + dsh + r"%' group by yw order by yw desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  # 人数

        # 全省
        sql = r"select yw,count(yw) from kscj where kl=1 and yw!=0 group by yw order by yw desc"
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

        df.to_excel(excel_writer=writer, sheet_name='地市及全省理科考生一分段概括(语文)',index=None)

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(语文).xlsx")

        rows = []
        sql = r"select count(*) from kscj where ksh like '"+dsh+r"%'"
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = total / 3
        # 2/3
        high = total / 1.5

        idxs = range(1,14)

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
                  r"(select * from (select a.*,rownum rn from (select ksh,yw from kscj " \
                  r"where ksh like '"+dsh+r"%' ORDER BY yw desc) a ) b" \
                  r" where b.rn BETWEEN 1 and "+str(low)+r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=001 and amx.idx="+str(idx)+r" GROUP BY amx.da"
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
                  r"(select * from (select a.*,rownum rn from (select ksh,yw from kscj " \
                  r"where ksh like '" + dsh + r"%' ORDER BY yw desc) a ) b" \
                  r" where b.rn BETWEEN "+str(low+1)+" and " + str(high) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=001 and amx.idx=" +str(idx) + r" GROUP BY amx.da"
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
                  r"(select * from (select a.*,rownum rn from (select ksh,yw from kscj " \
                  r"where ksh like '" + dsh + r"%' ORDER BY yw desc) a ) b" \
                  r" where b.rn BETWEEN " + str(high+1) + " and " + str(total) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=001 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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

            row.append((a_t/(a_h+a_m+a_l))*100) # 全部选A
            row.append((a_h/low)*100)   # 高分组选A
            row.append((a_m/(high-low))*100) #中间组选A
            row.append((a_l/(total-high))*100) #低分组选A

            row.append((b_t / (b_h+b_m+b_l)) * 100)  # 全部选B
            row.append((b_h / low) * 100)  # 高分组选B
            row.append((b_m / (high - low)) * 100)  # 中间组选B
            row.append((b_l / (total - high)) * 100)  # 低分组选B

            row.append((c_t / (c_h+a_m+c_l)) * 100)  # 全部选C
            row.append((c_h / low) * 100)  # 高分组选C
            row.append((c_m / (high - low)) * 100)  # 中间组选C
            row.append((c_l / (total - high)) * 100)  # 低分组选C

            row.append((d_t / (d_h+d_m+d_l)) * 100)  # 全部选D
            row.append((d_h / low) * 100)  # 高分组选D
            row.append((d_m / (high - low)) * 100)  # 中间组选D
            row.append((d_l / (total - high)) * 100)  # 低分组选D

            self.set_list_precision(row)
            rows.append(row)

        xths = ["01","02","03","04","05","07","10","11","12","14","17","18","19"]

        df = pd.DataFrame(data=None,columns=["题号","全部(A)","高分组(A)","中间组(A)","低分组(A)",
                                             "全部(B)","高分组(B)","中间组(B)","低分组(B)",
                                             "全部(C)","高分组(C)","中间组(C)","低分组(C)",
                                             "全部(D)","高分组(D)","中间组(D)","低分组(D)"])

        for i in range(13):
            rows[i].insert(0,xths[i])
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer,index=None,sheet_name="地市不同层次考生选择题受选率统计(语文)")
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

        writer = pd.ExcelWriter(path + '\\' + "各市情况分析(语文).xlsx")

        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])

        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where yw!=0"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(yw),stddev_samp(yw) from kscj where yw!=0"
        self.cursor.execute(sql)
        item = self.cursor.fetchone()
        row.append(item[0])
        row.append((item[0]/total)*100)
        row.append(item[1])
        row.append(item[2])
        row.append(item[2]/item[1])
        self.set_list_precision(row)
        df.loc[len(df)] = row

        for ds in dss:
            row = []
            row.append(ds[0])
            row.append(ds[1])


            sql = r"select count(*) as num,avg(yw),stddev_samp(yw) from kscj where yw!=0 and ksh like '"+ds[0]+r"%'"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer,index=None,sheet_name="各市考生成绩比较(语文)")

        # 文科
        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])
        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where yw!=0 and kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(yw),stddev_samp(yw) from kscj where yw!=0 and kl=2"
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


            sql = r"select count(*) as num,avg(yw),stddev_samp(yw) from kscj where yw!=0 and ksh like '" + ds[0] + r"%' and kl=2"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市文科考生成绩比较(语文)")

        # 理科
        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])
        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where yw!=0 and kl=1"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(yw),stddev_samp(yw) from kscj where yw!=0 and kl=1"
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


            sql = r"select count(*) as num,avg(yw),stddev_samp(yw) from kscj where yw!=0 and ksh like '" + ds[0] + r"%' and kl=1"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])

            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市理科考生成绩比较(语文)")

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

        writer = pd.ExcelWriter(path + '\\' + "原始分概括(语文).xlsx")

        sql = "select count(*) from kscj where yw!=0"
        self.cursor(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None,columns=['一分段','人数','百分比','累计百分比'])

        sql = "select yw,count(yw) from kscj where yw!=0 group by (yw) order by yw desc"
        self.cursor.execute(sql)
        results = self.cursor.fetchall()

        num = 0

        for result in results:
            row = []
            row.append(result[0])
            row.append(result[1])
            row.append(result[1]/total)
            num += result[1]
            row.append(num/total)
            df.loc[len(df)] = row

        df.to_excel(writer,index=None,sheet_name="全省考生一分段(语文)")

        sql = "select count(*) from kscj where yw!=0 and kl=1"
        self.cursor(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = "select yw,count(yw) from where yw!=0 and kl=1 kscj group by (yw) order by yw desc"
        self.cursor.execute(sql)
        results = self.cursor.fetchall()

        num = 0

        for result in results:
            row = []
            row.append(result[0])
            row.append(result[1])
            row.append(result[1] / total)
            num += result[1]
            row.append(num / total)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="全省理科考生一分段(语文)")

        sql = "select count(*) from kscj where yw!=0 and kl=2"
        self.cursor(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = "select yw,count(yw) from where yw!=0 and kl=2 kscj group by (yw) order by yw desc"
        self.cursor.execute(sql)
        results = self.cursor.fetchall()

        num = 0

        for result in results:
            row = []
            row.append(result[0])
            row.append(result[1])
            row.append(result[1] / total)
            num += result[1]
            row.append(num / total)
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="全省文科考生一分段(语文)")
        writer.save()

