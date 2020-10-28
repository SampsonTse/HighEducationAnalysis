import numpy as np
import pandas as  pd
import pymysql
import os
import matplotlib.pyplot  as plt
import decimal
import cx_Oracle
import openpyxl
import matplotlib.ticker as ticker

class zffb_img:
    def __init__(self):
        self.db = cx_Oracle.connect('gkeva2020/ksy#2020#reta@10.0.200.103/ksydb01std')
        self.cursor = self.db.cursor()

    def __del__(self):
        self.cursor.close()
        self.db.close()

    def getImg(self,kl,dsh):

        if kl == 1:
            kl_mc = "理科"
        else:
            kl_mc = "文科"

        sql = ""
        sql = "select mc from c_ds where DS_H=" + dsh
        self.cursor.execute(sql)
        ds_mc = self.cursor.fetchone()[0]

        pwd = os.getcwd()
        father_path = os.path.abspath(os.path.dirname(pwd) + os.path.sep + ".")
        path = father_path + r"\成绩概括"

        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + ds_mc
        if not os.path.exists(path):
            os.makedirs(path)
        path = path + "\\" + kl_mc
        if not os.path.exists(path):
            os.makedirs(path)

        plt.rcParams['figure.figsize'] = (15.0, 6)
        plt.xlim((0,750))
        ax = plt.gca()
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')

        sql = "SELECT COUNT(zf0) FROM kscj where kl="+str(kl)
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全省人数

        sql = "SELECT zf0,COUNT(zf0) FROM kscj WHERE zf0 != 0 and kl="+str(kl)+" GROUP BY  zf0 "
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        province = [None] * 751

        for item in items:
            province[item[0]] = round(item[1] / num * 100, 2)
        x = list(range(751))

        plt.plot(x, province, color='springgreen', marker='.', label='全省')


        sql = "SELECT COUNT(zf0) FROM kscj where kl=1 and KSH LIKE '" + dsh + r"%'"
        self.cursor.execute(sql)
        num = self.cursor.fetchone()[0]  # 全市人数

        sql = r"SELECT zf0,COUNT(zf0) FROM kscj WHERE zf0 != 0 and kl="+str(kl)+" and KSH LIKE '" + dsh + r"%' GROUP BY  zf0"
        self.cursor.execute(sql)
        items = list(self.cursor.fetchall())
        city = [None] * 751

        for item in items:
            city[item[0]] = round(item[1] / num * 100, 2)

        x = list(range(751))

        plt.plot(x, city, color='orange', marker='.', label='全市')
        ax.xaxis.set_major_locator(ticker.MultipleLocator(25))
        plt.xlabel('得分')
        plt.ylabel('人数百分比（%）')
        plt.legend(loc='upper center',bbox_to_anchor=(1.05, 1.05))
        plt.savefig(path + '\\'+kl_mc+'总分分布.png', dpi=1200)
        plt.close()