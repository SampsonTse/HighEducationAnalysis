import numpy as np
import pandas as  pd
import matplotlib.ticker as ticker
import pymysql
import os
import matplotlib.pyplot  as plt
import math
import decimal
import cx_Oracle
import openpyxl

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
np.set_printoptions(precision=2)


# 文科数学考生答题水平分析
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析总体概括(文科数学).xlsx")


        # 文科
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r'select count(a.sx) from kscj  a right join JBXX  b on a.KSH = b.KSH WHERE b.DS_H=' + dsh + r' and a.kl=2 and a.sx!=0'
        print(sql)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H=" + dsh + r" and b.XB_H = 1 and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        self.cursor.execute(sql)
        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where b.XB_H = 1 and a.kl=2 and a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '男')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and b.XB_H = 2 and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where b.XB_H = 2 and a.kl=2 and a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '女')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2 and a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '城镇')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2 and a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '农村')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH " \
              r"where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2 and a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '应届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH " \
              r"where b.DS_H=" + dsh + r" and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH" \
              r" where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2 and a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '往届')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.sx)   num,AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std " \
              r"from kscj   a right join JBXX   b on a.KSH = b.KSH" \
              r" where b.DS_H=" + dsh + r" and a.kl=2 and a.sx!=0"

        result = []
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数

        sql = r"select AVG(A.sx)   mean from kscj   A right join JBXX   B on A.KSH = B.KSH and a.kl=2 where a.sx!=0"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, (result[0] / num) * 100)
        result.insert(0, '总计')

        self.set_list_precision(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别考生成绩比较(文科数学)", excel_writer=writer, index=None)

        sql = r"select xq_h,mc from c_xq where  xq_h like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        xqhs = list(self.cursor.fetchall())
        xqhs.pop(0)
        # 各区县理考生成绩比较

        df = pd.DataFrame(data=None, columns=['区县', '人数', '平均分', '标准差', '差异系数', '得分率'])

        sql = "select count(sx),AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std FROM kscj   A where A.kl=2"
        self.cursor.execute(sql)
        result = self.cursor.fetchone()
        result = list(result)
        result.append((float(result[2]) / float(result[1])) * 100)  # 差异系数
        result.append(result[1] / 150)
        result.insert(0, '全省')
        self.set_list_precision(result)
        df.loc[len(df)] = result

        sql = r"select count(sx),AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std FROM kscj   A " \
              r"where A.kl=2 and a.sx!=0 and A.KSH LIKE '" + dsh + r"%'"
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
            sql = "select count(sx),AVG(A.sx)   mean,STDDEV_SAMP(A.sx)   std FROM kscj   A " \
                  "right join JBXX   B ON A.KSH = B.KSH WHERE A.kl=2 and a.sx!=0 and B.XQ_H = " + xqh[0]
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

        df.to_excel(excel_writer=writer, sheet_name="各县区考生成绩比较(文科数学)", index=None)

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

        # 全省文科
        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0, 6)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(sx) FROM kscj where kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "SELECT sx,COUNT(sx) FROM kscj WHERE sx != 0 and kl=2 GROUP BY  sx "
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 151

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(151))

        plt.plot(x, province, color='orange', marker='.', label='全省')

        # 全市文科
        sql = "SELECT COUNT(sx) FROM kscj where kl=2 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全市人数

        sql = r"SELECT sx,COUNT(sx) FROM kscj WHERE sx != 0 and kl=2 and KSH LIKE '" + dsh + r"%' GROUP BY  sx"
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
        plt.savefig(path + '\\地市及全省考生单科成绩分布(文科数学).png', dpi=1200)
        plt.close()

    # 省级报告 原始分概括 制表
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

        writer = pd.ExcelWriter(path + '\\' + "全省考生答题分析原始分概括(文科数学).xlsx")

        # 全省考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(*) from kscj  a right join JBXX  b on a.ksh = b.ksh"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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

        sql = "select count(a.sx)  num,AVG(a.sx)  mean,STDDEV_SAMP(a.sx)  std " \
              "from kscj a right join JBXX  b on a.ksh = b.ksh"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别考生成绩比较(文科数学)", index=None)

        # 全省考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(a.sx)  num " \
              "from kscj  a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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

        sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
              "from kscj   a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别考生成绩比较(文科数学)", index=None)

        # 全省考生
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率(%)', '平均分', '标准差', '差异系数'])

        sql = "select count(*) from kscj   a right join JBXX   b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]

        # 性别
        for xb in xbs:
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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
            sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
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

        sql = "select count(a.sx)   num,AVG(a.sx)   mean,STDDEV_SAMP(a.sx)   std " \
              "from kscj   a right join JBXX  b on a.ksh = b.ksh where a.kl=2"
        self.cursor.execute(sql)
        results = self.cursor.fetchone()
        results = list(results)
        results.append((float(results[2]) / float(results[1])) * 100 * 100)  # 差异系数
        results.insert(1, results[0] / num * 100)  # 比率
        results.insert(0, '总计')

        self.set_list_precision(results)
        df.loc[len(df)] = results

        df.to_excel(excel_writer=writer, sheet_name="各类别考生成绩比较(文科数学)", index=None)

        writer.save()

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


        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(文科数学).xlsx")


        sql = r"select count(*) from kscj where ksh like '"+dsh+r"%' and kl = 2 "
        self.cursor.execute(sql)
        num_ks = self.cursor.fetchone()[0]

        sql = r"select count(*) from kscj where kl = 2 "
        self.cursor.execute(sql)
        num_t = self.cursor.fetchone()[0]

        low = int(num_ks/3)
        high = int(num_ks/1.5)


        df = pd.DataFrame(data=None,columns=['题号','分值','本市平均分','全省平均分','本市得分率','高分组得分率','中间组得分率','低分组得分率'])

        kgths = [1,2,3,4,5,6,7,8,9,10,11,12]
        zgths = [13,14,15,16,17,18,19,20,21,22,23]

        for kgth in kgths:

            row = []
            row.append(str(kgth))
            row.append(5)

            total = 0

            # 全省平均分
            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX a right join jbxx b " \
                  "on a.ksh = b.ksh where a.idx=" + str(kgth) + " and kmh=003"
            self.cursor.execute(sql)
            avg_province = self.cursor.fetchone()[0] / num_t

            # 本市计算高分组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ RIGHT JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE KSCJ.kl = 2 and jbxx.ds_h="+dsh+" ORDER BY KSCJ.SX DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = "+str(kgth)+" and c.kmh=003 and b.rn BETWEEN 1 and "+str(low)
            self.cursor.execute(sql)
            sum_h = float(self.cursor.fetchone()[0])
            total = total + sum_h
            dfl_h = sum_h/ low / 5

            # 本市计算中间组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ RIGHT JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE KSCJ.kl = 2 and jbxx.ds_h=" + dsh + " ORDER BY KSCJ.SX DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=003 and b.rn BETWEEN "+str(low+1)+" and " + str(high)
            self.cursor.execute(sql)
            sum_m = float(self.cursor.fetchone()[0])
            total = total + sum_m
            dfl_m = sum_m / (high - low) / 5

            # 本市计算低分组平均分
            sql = "select sum(c.kgval) mean from T_GKPJ2020_TKSKGDAMX c " \
                  "right join (select a.*,rownum rn from " \
                  "(SELECT KSCJ.KSH from KSCJ RIGHT JOIN JBXX ON KSCJ.KSH = JBXX.KSH " \
                  "WHERE KSCJ.kl = 2 and jbxx.ds_h=" + dsh + " ORDER BY KSCJ.SX DESC) a) b " \
                  "on c.ksh = b.ksh where c.idx = " + str(kgth) + " and c.kmh=003 and b.rn BETWEEN "+str(high+1)+" and " + str(num_ks)
            self.cursor.execute(sql)
            sum_l = float(self.cursor.fetchone()[0])
            total = total + sum_l
            dfl_l = sum_l / (num_ks - high) / 5

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
            score_5 = [13,14,15,16]
            score_12 = [17,18,19,20,21]
            score_10 = [22,23]
            row = []
            num = 0
            row.append(str(zgth))
            if zgth in score_5:
                num = 5.00
            elif zgth in score_10:
                num = 10.00
            elif zgth in score_12:
                num = 12.00
            row.append(num)

            total = 0

            # 全省平均分
            sql = "select sum(b.sum) from " \
                  "(select sum(a.xtval) as sum,a.dth,a.ksh from T_GKPJ2020_TSJBNKSXT a " \
                  "right join jbxx on jbxx.ksh=a.ksh where a.kmh = 003 and a.dth = "+str(zgth)+" GROUP BY a.ksh,a.dth) b"
            self.cursor.execute(sql)
            avg_province = self.cursor.fetchone()[0] / num_t


            # 高分组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.sx from kscj where ksh like \'"+dsh+"%\' and kl = 2 ORDER BY KSCJ.sx desc) a ) b " \
                  "where b.rn BETWEEN 1 and "+str(low)+") c on sxt.ksh = c.ksh where sxt.kmh=003 and sxt.dth="+str(zgth)+" GROUP BY sxt.ksh) d"
            self.cursor.execute(sql)
            sum_h = float(self.cursor.fetchone()[0])
            total = total + sum_h
            dfl_h = sum_h / low / num

            # 中间组组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.sx from kscj where ksh like \'"+dsh+"%\' and kl = 2 ORDER BY KSCJ.sx desc) a ) b " \
                  "where b.rn BETWEEN "+str(low+1)+" and " + str(high) + ") c on sxt.ksh = c.ksh where sxt.kmh=003 and sxt.dth=" + str(zgth) + " GROUP BY sxt.ksh) d"
            self.cursor.execute(sql)
            sum_m = float(self.cursor.fetchone()[0])
            total = total + sum_m
            dfl_m = sum_m / (high - low) / num

            # 低分组得分率
            sql = "select sum(d.sum) as avg from (SELECT sum(xtval) as sum from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join (select b.* from (SELECT a.*,rownum rn from " \
                  "(select kscj.ksh,kscj.sx from kscj where ksh like \'"+dsh+"%\'  and kl = 2 ORDER BY KSCJ.sx desc) a ) b " \
                  "where b.rn BETWEEN " + str(high + 1) + " and " + str(num_ks) + ") c on sxt.ksh = c.ksh where sxt.kmh=003 and sxt.dth=" + str(zgth) + " GROUP BY sxt.ksh) d"

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

        df.to_excel(excel_writer=writer,sheet_name="地市考生单题分析情况(文科数学)",index=False)
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

        sql = r"select count(ksh) from (SELECT DISTINCT ksh from kscj where ksh like '"+dsh+r"%' and kl = 2) a"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        idxs = range(1, 13)
        xths = range(13, 24)

        x = []  # 难度
        y = []  # 区分度

        for idx in idxs:
            num = 5.0
            sql = r"select sum(kgval) FROM T_GKPJ2020_TKSKGDAMX amx right join kscj on kscj.ksh = amx.ksh where amx.ksh like '"+dsh+"%' and kmh = 003 and idx = " + str(idx)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度

            # 前27%得分率
            sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select ksh,sx from (select ksh,sx,rownum rn from " \
                  r"(select ksh,sx from kscj where ksh like '" + dsh + "%' ORDER BY sx desc) a ) b " \
                  r"where b.rn BETWEEN 1 and " + str(ph_num) + ") c on amx.ksh = c.ksh where amx.kmh = 003 and amx.idx = " + str(idx)
            print(sql)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            # 后27%得分率
            sql = r"select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx right join " \
                  r"(select ksh,sx from (select ksh,sx,rownum rn from " \
                  r"(select ksh,sx from kscj where ksh like '" + dsh + "%' ORDER BY sx desc) a ) b " \
                  r"where b.rn BETWEEN " + str(total-ph_num) + r" and " + str(total) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh = 003 and amx.idx = " + str(idx)
            print(sql)
            self.cursor.execute(sql)
            pl = self.cursor.fetchone()[0] / (total - ph_num) / num

            x.append(difficulty)
            y.append(ph - pl)

        for xth in xths:
            if xth in [13, 14, 15, 16]:
                num = 5.0
            elif xth in [17, 18, 19, 20, 21]:
                num = 12.0
            elif xth in [22, 23]:
                num = 10.0

            sql = r"select sum(xtval) from T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where sxt.ksh like '" + dsh + "%' and kmh=003 and dth=" + str(xth)
            self.cursor.execute(sql)
            difficulty = self.cursor.fetchone()[0] / total / num  # 难度
            x.append(difficulty)


            sql = r"select sx,b.sum from kscj right join " \
                  r"(select a.*,rownum rn from (select sum(xtval) sum,sxt.ksh from " \
                  r"T_GKPJ2020_TSJBNKSXT sxt right join kscj on kscj.ksh = sxt.ksh " \
                  r"where kmh = 003 and dth=" + str(xth) + r" and sxt.ksh " \
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

        txt = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17",
               "18", "19", "20", "21", "22", "23"]

        plt.rcParams['figure.figsize'] = (15.0,6.0)
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
            th = [txt[i]]
            plt.annotate(txt[i], xy=(x[i], y[i]), xytext=(x[i] , y[i] + 0.008),arrowprops=dict(arrowstyle='-'))
        plt.savefig(path + '\\各题难度-区分度分布散点图(文科数学).png', dpi=1200)
        plt.close()


    # 市级报告附录 原始分分析
    def YSFFX_CITY_TABLE(self, dsh):

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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题水平分析原始分概括(文科数学).xlsx")

        # 文科生
        city_num = [0] * 151
        province_num = [0] * 151

        city_total = 0
        province_total = 0

        df = pd.DataFrame(data=None,
                          columns=['一分段', '人数(本市)', '百分比(本市)', '累计百分比(本市)', '人数(全省)', '百分比(全省)', '累计百分比(全省)'])

        # 地市
        sql = r"select sx,count(sx) from kscj where kl=2 and sx!=0 and ksh like '" + dsh + r"%' group by sx order by sx desc"
        self.cursor.execute(sql)
        items = self.cursor.fetchall()

        for item in items:
            city_num[item[0]] = item[1]
            city_total += item[1]  # 人数

        # 全省
        sql = r"select sx,count(sx) from kscj where kl=2 and sx!=0 group by sx order by sx desc"
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

        df.to_excel(excel_writer=writer, sheet_name='地市及全省考生一分段概括(文科数学)', index=None)

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

        writer = pd.ExcelWriter(path + '\\' + "考生答题分析单题分析(文科数学).xlsx")

        rows = []
        sql = r"select count(*) from kscj where kl=2 "
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(1, 13)

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
                  r"(select * from (select a.*,rownum rn from (select ksh,sx from kscj " \
                  r"where kl=2  ORDER BY sx desc) a ) b" \
                  r" where b.rn BETWEEN 1 and " + str(low) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=003 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
                  r"(select * from (select a.*,rownum rn from (select ksh,sx from kscj " \
                  r"where kl=2  ORDER BY sx desc) a ) b" \
                  r" where b.rn BETWEEN " + str(low + 1) + " and " + str(high) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=003 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
                  r"(select * from (select a.*,rownum rn from (select ksh,sx from kscj " \
                  r"where kl=2  ORDER BY sx desc) a ) b" \
                  r" where b.rn BETWEEN " + str(high + 1) + " and " + str(total) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=003 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
            rows[i].insert(0, i + 1)
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(文科数学)")
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

        writer = pd.ExcelWriter(path + '\\' + ds_mc + "考生答题分析单题分析(文科数学).xlsx")

        rows = []
        sql = r"select count(*) from kscj where kl = 2 and ksh like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        # 1/3
        low = int(total / 3)
        # 2/3
        high = int(total / 1.5)

        idxs = range(1, 13)

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
                  r"(select * from (select a.*,rownum rn from (select ksh,sx from kscj " \
                  r"where kl=2 and ksh like '" + dsh + r"%' ORDER BY sx desc) a ) b" \
                  r" where b.rn BETWEEN 1 and " + str(low) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=003 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
                  r"(select * from (select a.*,rownum rn from (select ksh,sx from kscj " \
                  r"where kl=2 and ksh like '" + dsh + r"%' ORDER BY sx desc) a ) b" \
                  r" where b.rn BETWEEN " + str(low + 1) + " and " + str(high) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=003 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
                  r"(select * from (select a.*,rownum rn from (select ksh,sx from kscj " \
                  r"where kl=2 and ksh like '" + dsh + r"%' ORDER BY sx desc) a ) b" \
                  r" where b.rn BETWEEN " + str(high + 1) + " and " + str(total) + r") c on amx.ksh = c.ksh " \
                  r"where amx.kmh=003 and amx.idx=" + str(idx) + r" GROUP BY amx.da"
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
            rows[i].insert(0, i + 1)
            df.loc[len(df)] = rows[i]

        df.to_excel(excel_writer=writer, index=None, sheet_name="地市不同层次考生选择题受选率统计(文科数学)")
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

        writer = pd.ExcelWriter(path + '\\' + "各市情况分析(文科数学).xlsx")

        df = pd.DataFrame(data=None,columns=["地市代码","地市全称","人数","比率","平均分","标准差","差异系数(%)"])

        # 文科
        row = []
        row.append("00")
        row.append("全省")
        sql = "select count(*) from kscj where sx!=0 and kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        sql = "select count(*) as num,avg(sx),stddev_samp(sx) from kscj where sx!=0 and kl=2"
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


            sql = r"select count(*) as num,avg(sx),stddev_samp(sx) from kscj where sx!=0 and ksh like '" + ds[0] + r"%' and kl=2"
            self.cursor.execute(sql)
            item = self.cursor.fetchone()
            row.append(item[0])
            row.append((item[0] / total) * 100)
            row.append(item[1])
            row.append(item[2])
            row.append(item[2] / item[1])
            self.set_list_precision(row)
            df.loc[len(df)] = row

        df.to_excel(writer, index=None, sheet_name="各市考生成绩比较(文科数学)")
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

        writer = pd.ExcelWriter(path + '\\' + "原始分概括(文科数学).xlsx")

        sql = "select count(*) from kscj where sx!=0 and kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        df = pd.DataFrame(data=None, columns=['一分段', '人数', '百分比', '累计百分比'])

        sql = "select sx,count(sx) from kscj where sx!=0 and kl=2  group by (sx) order by sx desc"
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

        df.to_excel(writer, index=None, sheet_name="全省考生一分段(文科数学)")

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

        writer = pd.ExcelWriter(path + '\\' + "考生单题分析(文科数学).xlsx")
        df = pd.DataFrame(data=None, columns=["题号", "分值", "平均分", "标准差", "难度", "区分度"])

        idxs = list(range(1,13))
        xths = list(range(13,24))

        x = []
        y = []

        sql = "select count(sx) from kscj where kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]
        ph_num = int(total * 0.27)

        rows = []

        for idx in idxs:
            row = []

            if idx<10:
                row.append("0"+str(idx))
            else:
                row.append(str(idx))
            num = 5.0
            row.append(num)

            sql = "select avg(kgval),STDDEV_SAMP(kgval) from T_GKPJ2020_TKSKGDAMX amx right join kscj" \
                  " on kscj.ksh=amx.ksh where amx.kmh=003 and kscj.kl=2 and amx.idx = " + str(idx)
            print(sql)
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            mean = result[0]
            std = result[1]
            diffculty = mean / num

            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx " \
                  "right join (select b.* from (select a.*,rownum rn from " \
                  "(select ksh,sx from kscj where kl=2 order by sx desc) a) b where rn BETWEEN 1 and " + str(ph_num) + ") c " \
                  "on c.ksh = amx.ksh where kmh = 003 and idx = " + str(idx)
            print(sql)
            self.cursor.execute(sql)
            ph = self.cursor.fetchone()[0] / ph_num / num

            sql = "select sum(kgval) from T_GKPJ2020_TKSKGDAMX amx " \
                  "right join (select b.* from (select a.*,rownum rn from " \
                  "(select ksh,sx from kscj where kl=2 order by sx desc) a) b where rn BETWEEN " + str(total - ph_num) + " and " + str(total) + ") c on " \
                  "c.ksh = amx.ksh where kmh = 003 and idx = " + str(idx)
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
            row.append(str(xth))
            score_5 = [13, 14, 15, 16]
            score_12 = [17, 18, 19, 20, 21]
            score_10 = [22, 23]
            num = 0
            if xth in score_5:
                num = 5.0
            elif xth in score_10:
                num = 10.0
            elif xth in score_12:
                num = 12.0
            row.append(num)

            sql = "select avg(xtval),STDDEV_SAMP(xtval) from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join kscj on sxt.ksh = kscj.ksh where kscj.kl=2 and kmh = 003 and dth =" + str(xth)
            print(sql)
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            mean = result[0]
            std = result[1]
            diffculty = mean / num

            sql = "select sx,b.sum from kscj right join " \
                  "(select a.*,rownum rn from (select sum(xtval)  sum,sxt.ksh from T_GKPJ2020_TSJBNKSXT sxt " \
                  "right join kscj on kscj.ksh = sxt.ksh where kscj.kl=2 and kmh = 003 and dth=" + str(xth) + " GROUP BY sxt.ksh) a) b on kscj.ksh = b.ksh ORDER BY b.rn"

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

        df.to_excel(writer, index=None, sheet_name="考生单题作答情况(文科数学)")
        writer.save()

        plt.figure()
        plt.rcParams['figure.figsize'] = (15.0,6.0)
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        plt.xlim((0, 1))
        plt.ylim((0, 1))
        plt.xlabel("难度")
        plt.ylabel("区分度")
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
        plt.scatter(x, y)

        th = []
        for i in range(len(x)):
            th.append(rows[i][0])
            plt.annotate(rows[i][0], xy=(x[i], y[i]), xytext=(x[i] , y[i] + 0.008),
                         arrowprops=dict(arrowstyle='->', connectionstyle="arc3,rad = .2"))
        plt.savefig(path + '\\各题难度-区分度分布散点图(文科数学).png', dpi=1200)
        plt.close()


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

        sql = "select count(*) from kscj where sx!=0 and kl=2"
        self.cursor.execute(sql)
        total = self.cursor.fetchone()[0]

        score = [None] * 151
        sql = "select sx,count(sx) from kscj where sx!=0 and kl=2 group by sx order by sx desc"
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
        plt.savefig(path + '\\' + '全省考生单科成绩分布(文科数学).png', dpi=1200)
        plt.close()