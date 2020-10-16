import numpy as np
import pandas as pd
import pymysql


# 考生答题水平分析
class YW_KSDTSPFX:
    def __init__(self):
        self.db = pymysql.connect('localhost','root','1234','gk2020')
        self.cursor = self.db.cursor()

    def __del__(self):
        self.cursor.close()
        self.db.close()


    def ZTKG_CITY_YW(self,dsh):

        writer = pd.ExcelWriter("广州市考生答题分析总体概括(语文).xlsx")
        np.set_printoptions(precision=2)

        sql = ""

        df = pd.DataFrame(data=None,columns=['维度','人数','比率','平均分','标准差','差异系数','平均分(全省)'])

        sql = r'select count(a.YW) from kscj as a right join jbxx as b on a.KSH = b.KSH WHERE b.DS_H=%s'
        print(sql)
        self.cursor.execute(sql,[dsh])
        num = self.cursor.fetchone()[0] # 总人数


        # 计算维度为男
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and b.XB_H = 1"

        result = []
        self.cursor.execute(sql,[dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where b.XB_H = 1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1,result[0]/num)
        result.insert(0,'男')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and b.XB_H = 2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where b.XB_H = 2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '女')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 1 OR b.KSLB_H = 3)"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 1 OR b.KSLB_H = 3)"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '城镇')

        result = np.array(result)
        df.loc[len(df)] = result


        # 计算维度为农村
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 2 OR b.KSLB_H = 4)"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 2 OR b.KSLB_H = 4)"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '农村')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 1 OR b.KSLB_H = 2)"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 1 OR b.KSLB_H = 2)"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '应届')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 3 OR b.KSLB_H = 4)"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 3 OR b.KSLB_H = 4)"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '往届')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s "

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '总计')

        result = np.array(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别考生成绩比较(语文)",excel_writer=writer)



        # 文科
        df = pd.DataFrame(data=None,columns=['维度','人数','比率','平均分','标准差','差异系数','平均分(全省)'])

        sql = r'select count(a.YW) from kscj as a right join jbxx as b on a.KSH = b.KSH WHERE b.DS_H=%s and a.kl=2'
        print(sql)
        self.cursor.execute(sql, [dsh])
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and b.XB_H = 1 and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where b.XB_H = 1 and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '男')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and b.XB_H = 2 and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where b.XB_H = 2 and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '女')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '城镇')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '农村')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '应届')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '往届')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and a.kl=2"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH and a.kl=2"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '总计')

        result = np.array(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别文科考生成绩比较(语文)", excel_writer=writer)

        # 理科
        df = pd.DataFrame(data=None, columns=['维度', '人数', '比率', '平均分', '标准差', '差异系数', '平均分(全省)'])

        sql = r'select count(a.YW) from kscj as a right join jbxx as b on a.KSH = b.KSH WHERE b.DS_H=%s and a.kl=1'
        print(sql)
        self.cursor.execute(sql, [dsh])
        num = self.cursor.fetchone()[0]  # 总人数

        # 计算维度为男
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and b.XB_H = 1 and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where b.XB_H = 1 and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '男')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为女
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and b.XB_H = 2 and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where b.XB_H = 2 and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '女')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为城镇
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 1 OR b.KSLB_H = 3) and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '城镇')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为农村
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 2 OR b.KSLB_H = 4) and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '农村')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为应届
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 1 OR b.KSLB_H = 2) and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '应届')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为往届
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH where (b.KSLB_H = 3 OR b.KSLB_H = 4) and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '往届')

        result = np.array(result)
        df.loc[len(df)] = result

        # 计算维度为总计
        sql = r"select count(A.YW) as num,AVG(A.YW) as mean,STDDEV_SAMP(A.YW) as std " \
              r"from kscj as a right join jbxx as b on a.KSH = b.KSH where b.DS_H=%s and a.kl=1"

        result = []
        self.cursor.execute(sql, [dsh])
        result = self.cursor.fetchone()
        result = list(result)
        result.append(float(result[2]) / float(result[1]))  # 差异系数

        sql = r"select AVG(A.YW) as mean from kscj as A right join jbxx as B on A.KSH = B.KSH and a.kl=1"
        self.cursor.execute(sql)
        result.append(self.cursor.fetchone()[0])

        result.insert(1, result[0] / num)
        result.insert(0, '总计')

        result = np.array(result)
        df.loc[len(df)] = result

        df.to_excel(sheet_name="各类别文科考生成绩比较(理科)", excel_writer=writer)


        # 各区县考生成绩比较

        sql = r"select xq_h,mc from c_xq where like '"+dsh+r"%'"
        result = list(self.cursor.fetchall())
        result.pop(0)

        df = pd.DataFrame(data=None,columns=['区县','人数','标准差','差异系数','得分率'])

        for item in result:
            pass



        writer.save()

    def test(self,dsh):
        sql = r"select xq_h,mc from c_xq where xq_h like '" + dsh + r"%'"
        print(sql)
        self.cursor.execute(sql)
        result = list(self.cursor.fetchall())
        result.pop(0)
        print(result)





