import numpy as np
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

    # 制表
    def ZTKG_CITY_TABLE(self,dsh):

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

    # 画图
    def ZTJG_CITY_IMG(self,dsh):

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
        plt.savefig(path + '\\地市及全省考生单科成绩分布(语文).png', dpi=600)
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
        plt.savefig(path + '\\地市及全省文科考生单科成绩分布(语文).png', dpi=600)
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
        plt.savefig(path + '\\地市及全省理科考生单科成绩分布(语文).png', dpi=600)
        plt.close()

    def ZTKG_PROVINCE_TABLE(self):

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

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析总体概括(语文).xlsx")

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

    def YSFFX_CITY_TABLE(self,dsh):

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





















