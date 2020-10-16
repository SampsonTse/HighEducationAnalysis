import pymysql
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math 

plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

#获取所有学生的id
def get_all_ksid(cursor):

    sql = "SELECT 考生号 FROM ks_test"
    cursor.execute(sql)
    results = cursor.fetchall()

    studentids = []

    for result in results:
        studentids.append(result[0])
    return studentids

# 获取文科生以及理科生的ID
# 返回[理科考生号,理科考生号]
def judge_ks_wenli(cursor,studentids):
    
    wk = []
    lk = []
    for id in studentids:
        sql = "select * FROM kmf_test WHERE 科目号=002 and 考生号="+id
        cursor.execute(sql)
        result = cursor.fetchone()
        if result == None:
            wk.append(id)
        else:
            lk.append(id)
            
    return lk,wk


# 将考试科目号分为文科理科的考试号
# 返回[文科考试科目号,理科考试科目号]
def judge_km_wenli(km_ids):
    lk_km_id = []
    wk_km_id = []
    
    
    for km_id in km_ids:
        if km_id == '002' or km_id == '005':
            lk_km_id.append(km_id)
        elif km_id == '003'or km_id == '006':
            wk_km_id.append(km_id)
        else:
            lk_km_id.append(km_id)
            wk_km_id.append(km_id)
    
    
    return lk_km_id,wk_km_id

# 获取城市号的id
# 返回地市号[]
def get_all_cityid(cursor):
    ids = []
    sql = "SELECT * FROM ds"
    cursor.execute(sql)
    results = cursor.fetchall()
    for result in results:
        ids.append(result[0])
    return ids

# 获取科目号
# 返回科目号[]
def get_all_kmid(cursor):
    ids = []
    sql = "SELECT * FROM km"
    cursor.execute(sql)
    results = cursor.fetchall()
    for result in results:
        ids.append(result[0])
    return ids


# 成绩概括->总分概括
# 计算所有文科（理科）总分的(样本)标准差、平均分、差异系数
# 输入文理科所有学号也可以计算总概况
def general_situation(cursor,studentids):
    grades = []
    for id in studentids:
        sql = "SELECT * FROM kmf_test WHERE 考生号="+id
        cursor.execute(sql)
        results = cursor.fetchall()
        total = 0
        for result in results:
            total += result[3] + result[4]
        grades.append(total)
        
    # 全省人数
    sql = "SELECT COUNT(*) FROM ks_test" 
    cursor.execute(sql)
    n = cursor.fetchone()[0] 
    
    rate = len(grades) / n * 100
    average = np.mean(grades)
    std = np.std(grades,ddof = 1)
    cv = std/average * 100
       
    return [len(studentids),round(rate,2),round(std,2),round(average,2),round(cv,2)]


# 获取该科目的信度
def get_wenli_reliability(cursor,studentids,km_id):
    
    total_grades = []    # 学生单科总分

    for id in studentids:
        sql = "SELECT (主观分+客观分) as 总分 FROM kmf_test WHERE 科目号 = " + km_id + " AND 考生号 = " + id
        cursor.execute(sql)
        result = cursor.fetchone()
        total_grades.append(result[0])
    
    total_grades = np.array(total_grades)
    # 试卷总体方差
    total_var = np.var(total_grades)
    
    
    # 获取该科目的小题号，并计算该科目试题数量
    sql = "SELECT 小题号 FROM st WHERE 科目号 = " + km_id
    cursor.execute(sql)
    xth = cursor.fetchall()
    xth = np.array(xth)
    xth = xth.flatten()    
    n = len(xth) 
    
    #获取小题方差
    st_var = []
    
    # 获取每小题的方差
    for st in xth:
        # 该小题所有人的得分
        st_grade = []
        for id in studentids:
            sql = "SELECT 分数 FROM stf_test WHERE 科目号 = " + str(km_id) + " AND 考生号 = " + str(id) + " AND 小题号 = " + str(st)
            cursor.execute(sql)
            a = cursor.fetchone()
            if a != None:
                st_grade.append(a[0])
            else:
                continue

        st_grade = np.array(st_grade)
        st_var.append(np.var(st_grade))


    r = ( n/(n-1)) * ( (total_var - np.nansum(st_var))  /total_var )
    return r
     
# 成绩概括->原始分概括
# 计算文（理）科生单个科目的情况
# return:[科目,人数,平均值,标准差,难度,信度]
def single_km_situation(cursor,studentids,km_id):
    sql = "SELECT * FROM km WHERE 科目号 = " + km_id
    cursor.execute(sql)
    km = cursor.fetchone()
    km_name = km[1]
    km_total = km[2] + km[3]
    
    try:
        km_grades = []

        for id in studentids:
            sql = "SELECT * FROM kmf_test WHERE 科目号=" + km_id + " AND 考生号=" + id
            cursor.execute(sql)
            result = cursor.fetchone()
            km_grades.append(result[3]+result[4])

        # 全省人数
        sql = "SELECT COUNT(*) FROM ks_test" 
        cursor.execute(sql)
        n = cursor.fetchone()[0] 

        average = np.mean(km_grades)
        std = np.std(km_grades,ddof = 1)
        difficulty = average / km_total
        reliability = get_wenli_reliability(cursor,studentids,km_id)

        return [km_name,len(km_grades),round(average,2),round(std,2),round(difficulty,2),round(reliability,2)]

    except TypeError as e:

        return []


# 计算上线率 高分保护线上线率、本科上线率、本科与专科之间
# wl 1->理科 2->文科
# return [高分保护上线率,本科上线率,专科与本科间上线率]
def enrolled_situation(cursor, studentids, wl):
    try:
        # 各分数段录取线
        zb = 0.0
        bk = 0.0
        zk = 0.0

        # 各分数段人数
        num_zb = 0
        num_bk = 0
        num_zk = 0

        #学生总人数
        n = len(studentids)

        if wl == 1:
            zb = 524
            bk = 410
            zk = 160
        elif wl == 2:
            zb = 536
            bk = 430
            zk = 160

        for id in studentids:
            grade = 0
            sql = "SELECT * FROM kmf_test WHERE 考生号 = " + id
            cursor.execute(sql)
            results = cursor.fetchall()
            for result in results:
                grade += result[3] + result[4]

            if grade >= zb :
                num_zb = num_zb + 1
            elif grade >= bk and grade < zb :
                num_bk = num_bk + 1
            elif grade >= zk and grade <bk:
                num_zk = num_zk + 1 

        return [round(num_zb/n* 100,2), round(num_bk/n * 100,2), round(num_zk/n * 100,2)]

    except ZeroDivisionError:
        return []


# 计算上线率 本科上线率、本科与专科之间
# wl 1->理科 2->文科
# return [本科录取率,专科录取率]
def enrolled_situation2(cursor, studentids, wl):
    try:
        # 各分数段录取线
        bk = 0.0
        zk = 0.0

        # 各分数段人数
        num_bk = 0
        num_zk = 0

        #学生总人数
        n = len(studentids)

        if wl == 1:
            bk = 410
            zk = 160
        elif wl == 2:
            bk = 430
            zk = 160

        for id in studentids:
            grade = 0
            sql = "SELECT * FROM kmf_test WHERE 考生号 = " + id
            cursor.execute(sql)
            results = cursor.fetchall()
            for result in results:
                grade += result[3] + result[4]


            if grade >= bk :
                num_bk = num_bk + 1
            elif grade >= zk:
                num_zk = num_zk + 1 


        return [round(num_bk/n * 100,2), round(num_zk/n * 100,2)]
    except ZeroDivisionError:
        return []

# 计算临界率（临界生：指在各录取分数线下10分的考生）
# return:[本科生临界生比率,专科生临界生比率]
def margin_situation(cursor, studentids, wl):
    try:
        # 各分数段录取线
        bk = 0.0
        zk = 0.0

        # 各分数段临界人数
        num_bk = 0
        num_zk = 0

        #学生总人数
        n = len(studentids)

        if wl == 1:
            bk = 410
            zk = 160
        elif wl == 2:
            bk = 430
            zk = 160

        for id in studentids:
            grade = 0
            sql = "SELECT * FROM kmf_test WHERE 考生号 = " + id
            cursor.execute(sql)
            results = cursor.fetchall()
            for result in results:
                grade += result[3] + result[4]


            if grade < bk and grade > (bk - 10) :
                num_bk = num_bk + 1
            elif grade < zk and grade > (zk - 10):
                num_zk = num_zk + 1 

        return [round(num_bk/n * 100, 2), round(num_zk/n * 100, 2)]
    except ZeroDivisionError:
        return []



# 根据地市号计算各城市上线率 高分保护线上线率、本科上线率、本科与专科之间
# return [高分保护上线率,本科上线率,专科与本科间上线率]
def city_enrolled_situation(cursor, cityid, wl):
    
    
    sql = "SELECT * FROM ds WHERE 地市号 = " + str(cityid)
    cursor.execute(sql)
    city = cursor.fetchone()
    
    sql = "SELECT * FROM ks_test WHERE 地市号 = " + str(cityid)
    cursor.execute(sql)
    results = cursor.fetchall()
    
    stu_ids = []
    
    if len(results):
        for result in results :
            stu_ids.append(result[0])
            
        lk_ids,wk_ids = judge_ks_wenli(cursor,stu_ids)
        
        
        if wl == 1 and len(lk_ids):
            result = enrolled_situation(cursor, lk_ids, 1)
            result.insert(0,city[1])
            result.insert(0,str(cityid))
            return result
        
        elif wl == 2 and len(wk_ids):
            result = enrolled_situation(cursor, wk_ids, 2)
            result.insert(0,city[1])
            result.insert(0,str(cityid))
            return result
        else:
            return []
    else:
        return []


# 根据地市号计算各城市上线率:本科上线率、本科与专科之间
# return [本科上线率,本科与专科之间上线率]
def city_enrolled_situation2(cursor, cityid, wl):
    
    sql = "SELECT * FROM ds WHERE 地市号 = " + str(cityid)
    cursor.execute(sql)
    city = cursor.fetchone()
    
    
    sql = "SELECT * FROM ks_test WHERE 地市号 = " + str(cityid)
    cursor.execute(sql)
    results = cursor.fetchall()
    
    stu_ids = []
    
    if len(results):
        for result in results :
            stu_ids.append(result[0])
            
        lk_ids,wk_ids = judge_ks_wenli(cursor,stu_ids)
        
        if wl == 1 and len(lk_ids):
            result = enrolled_situation2(cursor, lk_ids, 1)
            result.insert(0,city[1])
            result.insert(0,str(cityid))
            return result
        elif wl == 2 and len(wk_ids):
            result = enrolled_situation2(cursor, wk_ids, 2)
            result.insert(0,city[1])
            result.insert(0,str(cityid))
            return result
        else:
            return []

    else:
        return []



# 根据地市号计算各城市临界生:本科上线率、本科与专科之间
# return [本科生临界生比率,专科生临界生比率]
def city_margin_situation(cursor, cityid, wl):
    
    sql = "SELECT * FROM ds WHERE 地市号 = " + str(cityid)
    cursor.execute(sql)
    city = cursor.fetchone()
    
    
    sql = "SELECT * FROM ks_test WHERE 地市号 = " + str(cityid)
    cursor.execute(sql)
    results = cursor.fetchall()
    
    stu_ids = []
    
    if len(results):
        for result in results :
            stu_ids.append(result[0])
            
        lk_ids,wk_ids = judge_ks_wenli(cursor,stu_ids)
        
        if wl == 1 and len(lk_ids):
            result = margin_situation(cursor, lk_ids, 1)
            result.insert(0,city[1])
            result.insert(0,str(cityid)) 
            return result
        elif wl == 2 and (wk_ids):
            result = margin_situation(cursor, wk_ids, 2)
            result.insert(0,city[1])
            result.insert(0,str(cityid)) 
            return result 
        else:
            return []
    else:
        return []

# 计算主观题区分度
# return 主观题区分度
def get_zg_discrimination(st,km):

    n = len(st)
    a = n * np.sum(st*km)
    b = np.sum(st) * np.sum(km)
    c = math.sqrt( n * np.sum(st**2) - (np.sum(st))**2 )
    d = math.sqrt( n * np.sum(km**2) - (np.sum(km))**2 )
    
    return ( a -b ) / ( c * d )


# 计算客观题区分度
# return 客观题区分度
def get_kg_discrimination(st,km,stf):


    # 获取排序后的索引
    rank = np.argsort(km)
    rank_low = rank[0:int(len(rank)*0.27)]
    rank_high = rank[len(rank)-int(len(rank)*0.27):]

    grade_high = []
    grade_low = []
      
    
    for i in rank_low:
        grade_low.append(st[i])
    for i in rank_high:
        grade_high.append(st[i])      
        
    grade_high = np.array(grade_high)
    grade_low = np.array(grade_low)

    ave_low = np.mean(grade_low)
    ave_high = np.mean(grade_high)

    result =  (ave_high/stf) - (ave_low/stf)
    return result
    


# 计算试卷的信度,不分文理科
# return 信度
def get_reliability(cursor, kmid):
    
    sql = "SELECT COUNT(*) FROM st WHERE 科目号 = " + kmid
    cursor.execute(sql)
    n = cursor.fetchone()[0]  # 试卷中试题数目
    
    
    total_grades = []  # 每个人的总分
    st_var = []        # 每小题的方差
    total_var = 0      # 总分方差
 
    # 获取所有人的得分并计算方差
    sql = "SELECT * FROM kmf_test WHERE 科目号 = " + kmid
    cursor.execute(sql)
    results = cursor.fetchall()

    for result in results:
        total_grades.append(result[3] + result[4])

    # 获取小题号
    sql = "SELECT 小题号 FROM st WHERE 科目号 = " + kmid
    cursor.execute(sql)
    xth = cursor.fetchall()
    xth = np.array(xth)
    xth = xth.flatten()

    # 计算每个小题的方差
    for i in xth:
        sql = "SELECT 分数 FROM stf_test WHERE 小题号 = "+str(i)+" AND 科目号 ="+kmid
        cursor.execute(sql)
        st_grades = cursor.fetchall()
        if len(st_grades):
            st_grades = np.array(st_grades)
            st_grades = st_grades.flatten()
            st_var.append(np.var(st_grades))
        
    
    total_grades = np.array(total_grades)
    total_var = np.var(total_grades)   # 总体方差

    a = ( n/(n-1)) * ( (total_var - np.sum(st_var))  /total_var )
    
    return a


# 根据科目号得到参与考试的学生的主客观题、全卷的得分情况
# 包括平均分、标准差、难度、区分度、信度
# return [客观题情况，主观题情况，全卷情况]
def zkg_situation(cursor, kmid):
    sql = "SELECT * FROM kmf_test WHERE 科目号 = " + kmid
    cursor.execute(sql)
    results = cursor.fetchall()
    
    n = len(results)
    # 主观题得分、客观题得分、总得分
    kg_total = []
    zg_total = []
    total = [] 
    for result in results:
        kg_total.append(result[3])
        zg_total.append(result[4])
        total.append(result[3] + result[4])

    zg_total = np.array(zg_total)
    kg_total = np.array(kg_total)
    total = np.array(total)
    
    # 客观题满分
    sql = "SELECT 客观分 FROM km WHERE 科目号 = " + kmid
    cursor.execute(sql)
    kg_grade = cursor.fetchone()[0]

    # 主观题满分
    sql = "SELECT 客观分 FROM km WHERE 科目号 = " + kmid
    cursor.execute(sql)
    zg_grade = cursor.fetchone()[0]
              
    # 主观题情况
    sql = "SELECT COUNT(*) FROM st WHERE 是否客观题 = 0 AND `科目号` = 001"
    cursor.execute(sql)
    zg_st_num = cursor.fetchone()[0]
    zg_average = np.mean(zg_total)
    zg_std = np.std(zg_total,ddof = 1)
    zg_qfd = get_zg_discrimination(zg_total,total)
    zg_difficulty = zg_average / zg_grade
    
    
       
    # 客观题情况
    sql = "SELECT COUNT(*) FROM st WHERE 是否客观题 = 1 AND `科目号` = 001"
    cursor.execute(sql)
    kg_st_num = cursor.fetchone()[0]
    kg_average = np.mean(kg_total)
    kg_std = np.std(kg_total,ddof = 1)
    kg_qfd = get_kg_discrimination(kg_total,total,kg_grade)
    kg_difficulty = kg_average / kg_grade
   
    
    # 总体情况    
    total_average = np.mean(total)
    total_std = np.std(total,ddof = 1 )
    reliability = get_reliability(cursor,kmid)
    difficulty = total_average / (zg_grade + kg_grade)
    

    return [
            ["选择题", round(kg_average,2),round(kg_std,2), round(kg_difficulty,2), round(kg_qfd,2), "/"],
            ["非选择题",round(zg_average,2),round(zg_std,2), round(zg_difficulty,2), round(zg_qfd,2), "/"],
            ["全卷",round(total_average),round(total_std,2),round(difficulty,2),"/",round(reliability,2)]
        ]
        
# 答题水平分析-->原始分分析
# 根据考生号，计算各科目的平均分 标准差 差异系数
# return [人数,比率,平均分,标准差,差异系数]
def km_total_grade_analysis(cursor, studentids,km_id):
    sql = "SELECT * FROM km WHERE 科目号 = " + km_id
    cursor.execute(sql)
    km_name = cursor.fetchone()[1]
    
    try:
        km_grades = []

        for id in studentids:
            sql = "SELECT * FROM kmf_test WHERE 科目号=" + km_id + " AND 考生号=" + id
            cursor.execute(sql)
            result = cursor.fetchone()


            km_grades.append(result[3]+result[4])

        sql = "SELECT COUNT(*) FROM ks_test" 
        cursor.execute(sql)
        
        # 全省人数
        n = cursor.fetchone()[0]
        
        rate = len(km_grades) / n * 100
        average = np.mean(km_grades)
        std = np.std(km_grades,ddof = 1)
        cv = std / average * 100

        return [len(km_grades),round(rate,2),round(average,2),round(std,2),round(cv,2)]      
        
    except TypeError:
        return []


# 答题水平分析 -> 各市情况分析
# cityid =0 全省
# return [人数,比率,平均分,标准差,差异系数]
def city_single_km(cursor,kmid,cityid=0, wl=0):
    try:
        sql = ""
        city_name = ""

        if cityid != 0:
            sql = "SELECT 名称 FROM ds WHERE 地市号 = " + str(cityid)
            cursor.execute(sql)
            city_name = cursor.fetchone()[0]
            sql = "SELECT 考生号 FROM ks_test WHERE 地市号 = " + str(cityid)
        elif cityid == 0:
            sql = "SELECT 考生号 FROM ks_test "
            city_name = "全省"

        cursor.execute(sql)
        studentids = cursor.fetchall()
        studentids = np.array(studentids)
        studentids = studentids.flatten()
        
        lk_ids = []
        wk_ids = []
        lk_ids,wk_ids = judge_ks_wenli(cursor,studentids)
        
        lk_ids = np.array(lk_ids)
        wk_ids = np.array(wk_ids)
        
        result = []

        # 不分文理
        if wl == 0:
            result = km_total_grade_analysis(cursor,studentids,kmid)
        elif wl == 1:
            result = km_total_grade_analysis(cursor, lk_ids, kmid)
        elif wl == 2:
            result = km_total_grade_analysis(cursor, wk_ids, kmid)

        result.insert(0,city_name)
        if cityid < 10:
            cityid = "0"+str(cityid)
        else:
            cityid = str(cityid)

        result.insert(0,cityid)
        return result
            
    except:
        return []

# 结构分析
# 计算所给科目号、小题号计算出平均分、标准差、难度
# return [分值,平均分,标准差,难度]
def structural_st_analysis(cursor,kmid,stids):
    
    try:
        sql = "SELECT COUNT(*) FROM stf_test WHERE 科目号 = " + str(kmid) + " AND 小题号 = " + str(stids[0])
        cursor.execute(sql)
        n = cursor.fetchone()[0]
        
        grades = np.zeros(n)
        
        total = 0.0
        
        #计算试题总分
        for stid in stids:
            sql = "SELECT 题分 FROM st WHERE 小题号 = " + str(stid) + " AND 科目号 = " + str(kmid)
            cursor.execute(sql)
            total += cursor.fetchone()[0]
            
            
        # 计算试卷得分
        
        # 得到学生得分
        for stid in stids:
            sql = "SELECT 分数 FROM stf_test WHERE 小题号 = " + str(stid) + " AND 科目号 = " + str(kmid)
            cursor.execute(sql)
            result = cursor.fetchall()
            result = np.array(result)
            result = result.flatten()
            grades += result
        
        average = np.mean(grades)
        std = np.std(grades,ddof=1)
        difficulty = average / total
        return [round(total,2),round(average,2),round(std,2),round(difficulty,2)]
    except:
        return []


# 单题分析
# 计算试卷单题的作答情况
# return [试题总分,平均分,标准差,难度,区分度]
def singlel_st_analysis(cursor,kmid,stids):
    
    try:
        sql = "SELECT COUNT(*) FROM stf_test WHERE 科目号 = " + str(kmid) + " AND 小题号 = " + str(stids[0])
        cursor.execute(sql)
        n = cursor.fetchone()[0]
        
        grades = np.zeros(n)
        
        # 试题分值
        st_total = 0.0
        
        #计算试题总分
        for stid in stids:
            sql = "SELECT 题分 FROM st WHERE 小题号 = " + str(stid) + " AND 科目号 = " + str(kmid)
            cursor.execute(sql)
            st_total += cursor.fetchone()[0]
            
            
        # 计算试卷得分
        total = []
        sql = "SELECT * FROM kmf_test WHERE 科目号 = " + str(kmid)
        cursor.execute(sql)
        results = cursor.fetchall()
        for result in results:
            total.append(result[3] + result[4])
            
        # 学生得分
        for stid in stids:
            sql = "SELECT 分数 FROM stf_test WHERE 小题号 = " + str(stid) + " AND 科目号 = " + str(kmid)
            cursor.execute(sql)
            result = cursor.fetchall()
            result = np.array(result)
            result = result.flatten()
            grades += result
        
        total = np.array(total)
        average = np.mean(grades)
        std = np.std(grades,ddof=1)
        qfd = 0.0
        difficulty = average / st_total
        
        
        sql = "SELECT 是否客观题 FROM st WHERE 小题号 = " + str(stids[0]) + " AND 科目号 = "+str(kmid)
        cursor.execute(sql)
        ifzg = cursor.fetchone()[0]
        
        # 0是客观题
        if ifzg == 1:
            qfd = get_zg_discrimination(grades,total)
        else:
            qfd = get_kg_discrimination(grades,total,st_total)
        
        return [round(st_total,2),round(average,2),round(std,2),round(difficulty,2),round(qfd,2)]
    except:
        return [None,None,None,None,None]


# 获得小题号
# return 小题号
def get_st_ids(cursor,kmid,st):
    if len(str(st)) > 1:
        sql = "SELECT 小题号  FROM st WHERE 科目号 = "+kmid+" AND 小题名称 LIKE \'"+str(st)+"%\'" 
        cursor.execute(sql)
        result = cursor.fetchall()
        result = np.array(result).flatten()
    else:
        sql = "SELECT 小题号  FROM st WHERE 科目号 = "+kmid+" AND 小题名称 = "+str(st)
    cursor.execute(sql)
    result = cursor.fetchall()
    result = np.array(result).flatten()
        
    return result     

def get_single_km_picture(cursor,stu_ids,km_id,path):
    
    sql = "SELECT 主观分,客观分 FROM km WHERE 科目号 = " + km_id
    cursor.execute(sql)
    km = cursor.fetchone()
    total_grade = km[0] + km[1]

    kmf = []
    for stu_id in stu_ids:
        sql = "SELECT 客观分,主观分 FROM kmf_test WHERE 科目号 =" + km_id + " AND 考生号 = " + stu_id
        cursor.execute(sql)
        result = cursor.fetchone()
        kmf.append(result[0]+result[1])

    x = list(set(kmf))
    x.sort()

    y = []
    for i in x:
        y.append(round(kmf.count(i) / len(kmf) * 100, 2))

    plt.figure()
    ax = plt.gca()
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    plt.xlim(0,total_grade)
    plt.plot(x,y)
    plt.savefig(path,dpi=500)

    