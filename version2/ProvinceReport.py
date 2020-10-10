import pandas as pd
import numpy as np
import pymysql
import math


def get_all_ksids(cursor):
    sql = "select 考生号 from ks_test"
    cursor.execute(sql)
    stuids = np.array(cursor.fetchall()).flatten()
    return stuids

def get_all_wk_ksids(cursor):
    sql = "select 考生号 from kmf_test where 科目号=003"
    cursor.execute(sql)
    stuids = np.array(cursor.fetchall()).flatten()
    return stuids

def get_all_lk_ksids(cursor):
    sql = "select 考生号 from kmf_test where 科目号=002"
    cursor.execute(sql)
    stuids = np.array(cursor.fetchall()).flatten()
    return stuids


def ZFGK(cursor,studentids):
    grades = []
    sql = ""

    for id in studentids:
        sql = "select sum(主观分+客观分) as 总分 from kmf_test where 考生号="+id+" group by 考生号"
        cursor.execute(sql)
        grades.append(cursor.fetchone()[0])

    print(grades)
    grades = np.array(grades)

    # 全省人数
    sql = "select count(*) from ks_test"
    cursor.execute(sql)
    n = cursor.fetchone()[0]

    num = len(grades)
    rate = len(grades) / n * 100
    average = np.mean(grades)
    std = np.std(grades,ddof=1)
    cv = std/average * 100

    return [num,rate,average,std,cv]

"""
    wenli:
        1 理科
        2 文科
"""
def KMYSFGK(cursor,wenli):
    sql = ""
    kmids = []

    df = pd.DataFrame(data=None,columns=['科目','样品数','平均分','标准差','难度','信度'])

    if wenli == 1:
        wl = "002"
        kmids = ["001","002","005","101"]
    elif wenli == 2:
        wl = "003"
        kmids = ["001", "003", "006", "101"]

    for kmid in kmids:
        sql = r"SELECT sum(客观分+主观分) as 单科总分 FROM kmf_test as a " \
              r"RIGHT JOIN (SELECT 考生号 FROM kmf_test WHERE 科目号 = "+ wl +") as b ON a.`考生号` = b.考生号 " \
              r"WHERE a.`科目号`=" + kmid + " GROUP BY(客观分+主观分)"

        cursor.execute(sql)
        grades = np.array(cursor.fetchall()).flatten()

        # 科目总分
        sql = "SELECT (主观分+客观分) as 总分 FROM km WHERE 科目号 = "+kmid
        cursor.execute(sql)
        total = cursor.fetchone()[0]

        num = len(grades)
        mean = np.mean(grades)
        std = np.std(grades,ddof=1)
        difficulty = mean/total
        reliability = get_reliability(cursor,kmid,np.var(grades),wenli)

        result = [kmid,num,round(mean,2),round(std,2),round(difficulty,2),round(reliability,2)]

        df.loc[len(df)] = result

    print(df)

# 计算信度
def get_reliability(cursor,kmid,total_var,wenli):

    sql = ""

    if wenli:
        if wenli == 1:
            wl = "002"
        elif wenli == 2:
            wl = "003"

        sql = r"SELECT STD(a.`分数`) as 标准差  FROM stf_test as a " \
              r"RIGHT JOIN (SELECT 考生号 FROM kmf_test WHERE 科目号 = "+wl+") as b " \
              r"ON a.考生号 = b.考生号 WHERE 科目号 = "+kmid+" GROUP BY a.`小题号` ORDER BY a.`小题号`"

    #计算每题的方差
    cursor.execute(sql)
    st_var = np.array(cursor.fetchall()).flatten() **2
    st_var = np.array(list(filter(None,st_var)))

    n = len(st_var)

    result = (n/(n-1)) * ((total_var-np.sum(st_var))/total_var)
    return result

# 上线概括
def SXGK(cursor,wenli):

    if wenli == 1:
        score_wenli = [160,410,524,"002"]
    elif wenli == 2:
        score_wenli = [160, 430, 536, "003"]

    sql = r"SELECT ELT(INTERVAL(c.总分,%s,%s,%s),'专科','本科','重本') as 分数段,count(c.考生号) as 人数 FROM" \
          r"( SELECT a.考生号,sum(a.客观分+a.主观分) as 总分 FROM kmf_test as a RIGHT JOIN " \
          r"(SELECT 考生号 FROM kmf_test WHERE 科目号 = %s) as b ON a.`考生号` = b.考生号 GROUP BY a.考生号) " \
          r"as c GROUP BY 分数段"

    cursor.execute(sql,score_wenli)
    items = cursor.fetchall()
    print(items)

    sql = "SELECT COUNT(考生号) as num FROM kmf_test WHERE 科目号 = %s"
    cursor.execute(sql,score_wenli[3])
    n = cursor.fetchone()[0]

    result = [0,0,0]

    for item in items:
        if item[0] == "专科":
            result[2] = (result[2]+item[1])/n
        if item[0] == "本科":
            result[1] = (result[1]+item[1])/n
        if item[0] == "重本":
            result[0] = (result[0]+item[1])/n

    return result

# 录取概括
def LQGK(cursor,wenli):
    if wenli == 1:
        score_wenli = [160,410,"002"]
    elif wenli == 2:
        score_wenli = [160,430,"003"]

    sql = r"SELECT ELT(INTERVAL(c.总分,%s,%s),'专科','本科') as 分数段,count(c.考生号) as 人数 FROM" \
          r"( SELECT a.考生号,sum(a.客观分+a.主观分) as 总分 FROM kmf_test as a RIGHT JOIN " \
          r"(SELECT 考生号 FROM kmf_test WHERE 科目号 = %s) as b ON a.`考生号` = b.考生号 GROUP BY a.考生号) " \
          r"as c GROUP BY 分数段"

    cursor.execute(sql, score_wenli)
    items = cursor.fetchall()
    print(items)

    sql = "SELECT COUNT(考生号) as num FROM kmf_test WHERE 科目号 = %s"
    cursor.execute(sql, score_wenli[2])
    n = cursor.fetchone()[0]

    result = [ 0, 0]
    for item in items:
        if item[0] == "专科":
            result[1] = (result[1] + item[1])/n
        if item[0] == "本科":
            result[0] = (result[0] + item[1])/n

    return result


