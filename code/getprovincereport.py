from operator import index
import provincereport as pr
import pymysql
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math 
import openpyxl
import os


class ProvinceReport(object):

    def __init__(self):
        self.__db = pymysql.connect("localhost","root","1234","higheducation")
        self.__cursor = self.__db.cursor()
        self.__ks_ids = pr.get_all_ksid(self.__cursor)
        self.__km_ids = pr.get_all_kmid(self.__cursor)
        self.__city_ids = pr.get_all_cityid(self.__cursor)


    """
        生成全省概括excel文件
    """
    def get_summary_of_grade(self):
        lk_ks_ids, wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)
        lk_km_ids,wk_km_ids = pr.judge_km_wenli(self.__km_ids)

        output_file = "省级报告/成绩概括.xlsx"
        writer = pd.ExcelWriter(output_file)

        # 生成文科各类别考生总分概括
        df = pd.DataFrame(data=None,columns=["维度","人数","平均数","标准差","平均分","差异系数"])
        result = pr.general_situation(self.__cursor,wk_ks_ids)
        result.insert(0, "总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="文科类各类别考生总分概括",index=False)

        # 生成文科类考生各科目原始分概括
        df = pd.DataFrame(data=None,columns=["科目","样品数","平均数","标准差","难度","信度"])
        for kmid in wk_km_ids:
            result = pr.single_km_situation(self.__cursor,wk_ks_ids,kmid)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="文科类考生各科目原始分概括",index=False)

         # 生成理科各类别考生总分概括
        df = pd.DataFrame(data=None,columns=["维度","人数","平均数","标准差","平均分","差异系数"])
        result = pr.general_situation(self.__cursor,lk_ks_ids)
        result.insert(0, "总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="理科类各类别考生总分概括",index=False)

        # 生成理科类考生各科目原始分概括
        df = pd.DataFrame(data=None,columns=["科目","样品数","平均数","标准差","难度","信度"])
        for kmid in lk_km_ids:
            result = pr.single_km_situation(self.__cursor,lk_ks_ids,kmid)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="理科类考生各科目原始分概括",index=False)

        writer.save()
        writer.close()

        return True

    """
        生成全省文科录取情况
    """
    def get_enrolled_province_wenke(self):
        lk_ks_ids, wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/录取概括_全省概括_文科.xlsx"
        writer = pd.ExcelWriter(output_file)

        # 文科全省各类别考生上线率概括
        df = pd.DataFrame(data=None,columns=["维度","高分保护先上线率(%)","本科上线率(%)","本科与专科线上线率(%)"])
        result = pr.enrolled_situation(self.__cursor,wk_ks_ids,2)
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省文科各类别考生上线率",index=False)

        # 文科全省各类别考生录取率情况
        df = pd.DataFrame(data=None,columns=["维度","本科录取率(%)","专科录取率(%)"])
        result = pr.enrolled_situation2(self.__cursor,wk_ks_ids,2)
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="文科全省各类别考生录取率情况",index=False)

        # 文科全省各类别考生临界生情况
        df = pd.DataFrame(data=None,columns=["维度","本科录取率(%)","专科录取率(%)"])
        result = pr.margin_situation(self.__cursor,wk_ks_ids,2)
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="文科全省各类别考生临界生情况",index=False)

        writer.save()
        writer.close()

        return True

    """
        生成全省理科录取情况
    """
    def get_enrolled_province_like(self):
        lk_ks_ids, wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/录取概括_全省概括_理科.xlsx"
        writer = pd.ExcelWriter(output_file)

        # 理科全省各类别考生上线率概括
        df = pd.DataFrame(data=None,columns=["维度","高分保护先上线率(%)","本科上线率(%)","本科与专科线上线率(%)"])
        result = pr.enrolled_situation(self.__cursor,lk_ks_ids,1)
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省文科各类别考生上线率",index=False)

        # 理科全省各类别考生录取率情况
        df = pd.DataFrame(data=None,columns=["维度","本科录取率(%)","专科录取率(%)"])
        result = pr.enrolled_situation2(self.__cursor,lk_ks_ids,1)
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="文科全省各类别考生录取率情况",index=False)

        # 理科全省各类别考生临界生情况
        df = pd.DataFrame(data=None,columns=["维度","本科录取率(%)","专科录取率(%)"])
        result = pr.margin_situation(self.__cursor,lk_ks_ids,1)
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="文科全省各类别考生临界生情况",index=False)

        writer.save()
        writer.close()

        return True

    """
        各市理科概括
    """
    def get_enrolled_city_like(self):
        lk_ks_ids, wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/录取概括_各市概括_理科.xlsx"
        writer = pd.ExcelWriter(output_file)

        # 各市理科类考生上线率概况
        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","高分保护线上","本科线上","本科、专科线间"])
        result = pr.enrolled_situation(self.__cursor,lk_ks_ids,1)
        result.insert(0,"全省")
        result.insert(0,"00")
        df.loc[len(df)] = result
        for cityid in self.__city_ids:
            result = pr.city_enrolled_situation(self.__cursor,cityid,1)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各市理科类考生上线率概况",index=False)

        # 各市理科类考生考生录取率概况
        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","本科","专科"])
        result = pr.enrolled_situation2(self.__cursor,lk_ks_ids,1)
        result.insert(0,"全省")
        result.insert(0,"00")
        df.loc[len(df)] = result
        for cityid in self.__city_ids:
            result = pr.city_enrolled_situation2(self.__cursor,cityid,1)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各市理科类考生录取率概况",index=False)

        # 各市理科类考生考生临界生概况
        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","本科临界生","专科临界生"])
        result = pr.margin_situation(self.__cursor,lk_ks_ids,1)
        result.insert(0,"全省")
        result.insert(0,"00")
        df.loc[len(df)] = result
        for cityid in self.__city_ids:
            result = pr.city_margin_situation(self.__cursor,cityid,1)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各市理科类考生临界生概况",index=False)

        writer.save()
        writer.close()

        return True


    """
        各市理科概括
    """
    def get_enrolled_city_wenke(self):
        lk_ks_ids, wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/录取概括_各市概括_文科.xlsx"
        writer = pd.ExcelWriter(output_file)

        # 各市文科类考生上线率概况
        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","高分保护线上","本科线上","本科、专科线间"])
        result = pr.enrolled_situation(self.__cursor,wk_ks_ids,2)
        result.insert(0,"全省")
        result.insert(0,"00")
        df.loc[len(df)] = result
        for cityid in self.__city_ids:
            result = pr.city_enrolled_situation(self.__cursor,cityid,2)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各市文科类考生上线率概况",index=False)

        # 各市文科类考生考生录取率概况
        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","本科","专科"])
        result = pr.enrolled_situation2(self.__cursor,wk_ks_ids,2)
        result.insert(0,"全省")
        result.insert(0,"00")
        df.loc[len(df)] = result
        for cityid in self.__city_ids:
            result = pr.city_enrolled_situation2(self.__cursor,cityid,2)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各市文科类考生录取率概况",index=False)

        # 各市文科类考生考生临界生概况
        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","本科临界生","专科临界生"])
        result = pr.margin_situation(self.__cursor,wk_ks_ids,2)
        result.insert(0,"全省")
        result.insert(0,"00")
        df.loc[len(df)] = result
        for cityid in self.__city_ids:
            result = pr.city_margin_situation(self.__cursor,cityid,2)
            if len(result):
                df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各市文科类考生临界生概况",index=False)

        writer.save()
        writer.close()

        return True

    

    """
        考生答题水平分析：语文
    """
    def yuwen_anwsering_situation(self):


        tixing = {
            "选择题(必做)":[1,2,3,4,5,7,10,11,12,14,17,18,19],
            "简答题(必做)":[6,8,9,15],
            "翻译题(必做)":[13],
            "压缩题(必做)":[16,20],
            "作文题(必做)":[22]
        }

        zsbk = {
            "论述类文本阅读(必做)":[1,2,3],
            "实用类文本阅读(必做)":[4,5,6],
            "文学类文本阅读(必做)":[7,8,9],
            "文言文阅读(必做)":[10,11,12,13],
            "古代诗歌阅读(必做)":[14,15],
            "古诗文阅读(必做)":[16],
            "语言文字应用(必做)":[17,18,19,20,21],
            "写作(必做)":[22]
        }

        khnl = {
            "理解(必做)":[1,10,11,13],
            "分析综合(必做)":[2,3,4,5,6,7,8,12],
            "鉴赏评价(必做)":[9,14,15],
            "识记(必做)":[16],
            "表达应用(必做)":[17,18,19,20,21,22]
        }

        
        """
            原始分分析
        """
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        if not os.path.exists("省级报告"):
            os.makedirs("省级报告")
        if not os.path.exists("省级报告/语文考生答题水平分析"):
            os.makedirs("省级报告/语文考生答题水平分析")

        lk_ks_ids,wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/语文考生答题水平分析/原始分概括(语文).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 全省
        result = pr.km_total_grade_analysis(self.__cursor,self.__ks_ids,"001")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别考生比较(语文)",index=False)

        # 文科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,wk_ks_ids,"001")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别文科考生比较(语文)",index=False)

        # 理科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,lk_ks_ids,"001")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别理科考生比较(语文)",index=False)

        writer.save()
        writer.close()


        """
            结构分析
        """
        output_file = "省级报告/语文考生答题水平分析/结构分析(语文).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 主客观分析
        df = pd.DataFrame(data=None,columns=["主客观题","题数","平均分","标准差","难度","区分度","信度"])
        results = pr.zkg_situation(self.__cursor,"001")
        results[0].insert(1,13.00)
        results[1].insert(1,9.00)
        results[2].insert(1,22.00)
        for result in results:
            df.loc[len(df)] = result
        df.to_excel(writer,sheet_name="全省考生客观题得分(语文)",index=False)

        # 全省各题型得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in tixing.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"001",i))
            result = pr.structural_st_analysis(self.__cursor,"001",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各题型得分情况(语文)",index=False)


        # 全省各知识板块得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in zsbk.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"001",i))
            result = pr.structural_st_analysis(self.__cursor,"001",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各知识板块得分情况(语文)",index=False)


        # 全省各考核能力得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in khnl.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"001",i))
            result = pr.structural_st_analysis(self.__cursor,"001",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各考核能力得分情况(语文)",index=False)
        writer.save()
        writer.close()



        """
            单题分析
        """
        output_file = "省级报告/语文考生答题水平分析/单题分析(语文).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 获取所有题号
        st = []
        for value in zsbk.values():
            st.extend(value)

        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])
        for i in st:
            st_ids = pr.get_st_ids(self.__cursor,"001",i)
            result = pr.singlel_st_analysis(self.__cursor,"001",st_ids)
            result.insert(0,i)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="考生单题作答情况(语文)",index=False)
        writer.save()
        writer.close()


        """
            各市情况分析
        """
        output_file = "省级报告/语文考生答题水平分析/各市情况分析(语文).xlsx"
        writer = pd.ExcelWriter(output_file)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"001",0,0)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"001",i,0)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市考生成绩比较(语文)",index=False)


        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"001",0,2)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"001",i,2)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市文科考生成绩比较(语文)",index=False)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"001",0,1)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"001",i,1)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市理科考生成绩比较(语文)",index=False)

        writer.save()
        writer.close()

        pr.get_single_km_picture(self.__cursor,self.__ks_ids,"001","省级报告/语文考生答题水平分析/全省考生单科成绩分布(语文).png")
        pr.get_single_km_picture(self.__cursor,wk_ks_ids,"001","省级报告/语文考生答题水平分析/全省文科考生单科成绩分布(语文).png")
        pr.get_single_km_picture(self.__cursor,lk_ks_ids,"001","省级报告/语文考生答题水平分析/全省理科考生单科成绩分布(语文).png")

        return True
    
    
    """
        考生答题水平分析：文科数学
    """
    def wenkeshuxue_anwsering_situation(self):

        tixing = {
            "选择题(必做)":[1,2,3,4,5,6,7,8,9,10,11,12],
            "填空题(必做)":[13,14,15,16],
            "解答题(必做)":[17,18,19,20,21],
            "解答题(选做1)":[22],
            "解答题(选做2)":[23]
        }

        zsbk = {
            "复数(必做)":[1],
            "集合(必做)":[2],
            "函数(必做)":[3,5],
            "相等关系与不等关系(必做)":[4],
            "统计(必做)":[6],
            "三角函数(必做)":[7,15],
            "平面向量(必做)":[8],
            "程序框图(必做)":[9],
            "解析几何(必做)":[10,12,21],
            "解三角形(必做)":[11],
            "导数(必做)":[13],
            "数列(必做)":[14,18],
            "立体几何(必做)":[16,19],
            "概率与统计(必做)":[17],
            "函数与导数(必做)":[20],
            "坐标系与参数方程(选做1)":[22],
            "不等式选讲(选做2)":[23]
        }

        khnl = {
            "数学运算(必做)":[1,7,8,10,13,14],
            "综合能力(必做)":[2,4,5,11,12,15,16,17,18,19,20],
            "逻辑推理(必做)":[3,6,9],
            "综合能力(选做1)":[22],
            "综合能力(选做2)":[23]
        }

        
        """
            原始分分析
        """
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        if not os.path.exists("省级报告"):
            os.makedirs("省级报告")
        if not os.path.exists("省级报告/文科数学考生答题水平分析"):
            os.makedirs("省级报告/文科数学考生答题水平分析")

        lk_ks_ids,wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/文科数学考生答题水平分析/原始分概括(文科数学).xlsx"
        writer = pd.ExcelWriter(output_file)


        # 文科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,wk_ks_ids,"003")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别考生比较(文科数学)",index=False)

        writer.save()
        writer.close()


        """
            结构分析
        """
        output_file = "省级报告/文科数学考生答题水平分析/结构分析(文科数学).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 主客观分析
        df = pd.DataFrame(data=None,columns=["主客观题","题数","平均分","标准差","难度","区分度","信度"])
        results = pr.zkg_situation(self.__cursor,"003")
        results[0].insert(1,12.00)
        results[1].insert(1,11.00)
        results[2].insert(1,23.00)
        for result in results:
            df.loc[len(df)] = result
        df.to_excel(writer,sheet_name="全省考生客观题得分(文科数学)",index=False)

        # 全省各题型得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in tixing.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"003",i))
            result = pr.structural_st_analysis(self.__cursor,"003",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各题型得分情况(文科数学)",index=False)


        # 全省各知识板块得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in zsbk.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"003",i))
            result = pr.structural_st_analysis(self.__cursor,"003",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各知识板块得分情况(文科数学)",index=False)


        # 全省各考核能力得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in khnl.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"003",i))
            result = pr.structural_st_analysis(self.__cursor,"003",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各考核能力得分情况(文科数学)",index=False)
        writer.save()
        writer.close()



        """
            单题分析
        """
        output_file = "省级报告/文科数学考生答题水平分析/单题分析(文科数学).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 获取所有题号
        st = []
        for value in zsbk.values():
            st.extend(value)

        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])
        for i in st:
            st_ids = pr.get_st_ids(self.__cursor,"003",i)
            result = pr.singlel_st_analysis(self.__cursor,"003",st_ids)
            result.insert(0,i)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="考生单题作答情况(文科数学)",index=False)
        writer.save()
        writer.close()


        """
            各市情况分析
        """
        output_file = "省级报告/文科数学考生答题水平分析/各市情况分析(文科数学).xlsx"
        writer = pd.ExcelWriter(output_file)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"003",0,2)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"003",i,2)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市考生成绩比较(文科数学)",index=False)

        writer.save()
        writer.close()

        pr.get_single_km_picture(self.__cursor,wk_ks_ids,"003","省级报告/文科数学考生答题水平分析/全省考生单科成绩分布(文科数学).png")

        return True
    
    
    """
        考生答题水平分析：理科数学
    """
    def likeshuxue_anwsering_situation(self):


        tixing = {
            "选择题(必做)":[1,2,3,4,5,6,7,8,9,10,11,12],
            "填空题(必做)":[13,14,15,16],
            "解答题(必做)":[17,18,19,20,21],
            "解答题(选做1)":[22],
            "解答题(选做2)":[23]
        }

        zsbk = {
            "集合(必做)":[1],
            "复数(必做)":[2],
            "函数(必做)":[3,5],
            "相等关系与不等关系(必做)":[4],
            "概率(必做)":[6,15],
            "平面向量(必做)":[7],
            "程序框图(必做)":[8],
            "数列(必做)":[9,14],
            "解析几何(必做)":[10,16,19],
            "三角函数(必做)":[11],
            "立体几何(必做)":[12,18],
            "导数(必做)":[13],
            "解三角形(必做)":[17],
            "函数与导数(必做)":[20],
            "概率与统计(必做)":[21],
            "坐标系与参数方程(选做1)":[22],
            "不等式选讲(选做2)":[23]
        }

        khnl = {
            "数学运算(必做)":[1,7,9,13,14],
            "综合能力(必做)":[2,4,5,10,12,16,17,18,19,20,21],
            "逻辑推理(必做)":[3,8,11,15],
            "数学建模(必做)":[6],
            "综合能力(选做1)":[22],
            "综合能力(选做2)":[23]
        }

        
        """
            原始分分析
        """
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        if not os.path.exists("省级报告"):
            os.makedirs("省级报告")
        if not os.path.exists("省级报告/理科数学考生答题水平分析"):
            os.makedirs("省级报告/理科数学考生答题水平分析")

        lk_ks_ids,wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/理科数学考生答题水平分析/原始分概括(理科数学).xlsx"
        writer = pd.ExcelWriter(output_file)


        # 理科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,lk_ks_ids,"002")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别考生比较(理科数学)",index=False)

        writer.save()
        writer.close()


        """
            结构分析
        """
        output_file = "省级报告/理科数学考生答题水平分析/结构分析(理科数学).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 主客观分析
        df = pd.DataFrame(data=None,columns=["主客观题","题数","平均分","标准差","难度","区分度","信度"])
        results = pr.zkg_situation(self.__cursor,"002")
        """
            需要手动更改
        """
        results[0].insert(1,12.00)
        results[1].insert(1,11.00)
        results[2].insert(1,23.00)
        for result in results:
            df.loc[len(df)] = result
        df.to_excel(writer,sheet_name="全省考生客观题得分(理科数学)",index=False)

        # 全省各题型得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in tixing.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"002",i))
            result = pr.structural_st_analysis(self.__cursor,"002",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各题型得分情况(理科数学)",index=False)


        # 全省各知识板块得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in zsbk.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"002",i))
            result = pr.structural_st_analysis(self.__cursor,"002",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各知识板块得分情况(理科数学)",index=False)


        # 全省各考核能力得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in zsbk.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"002",i))
            result = pr.structural_st_analysis(self.__cursor,"002",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各考核能力得分情况(理科数学)",index=False)
        writer.save()
        writer.close()



        """
            单题分析
        """
        output_file = "省级报告/理科数学考生答题水平分析/单题分析(理科数学).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 获取所有题号
        st = []
        for value in zsbk.values():
            st.extend(value)

        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])
        for i in st:
            st_ids = pr.get_st_ids(self.__cursor,"002",i)
            result = pr.singlel_st_analysis(self.__cursor,"002",st_ids)
            result.insert(0,i)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="考生单题作答情况(理科数学)",index=False)
        writer.save()
        writer.close()


        """
            各市情况分析
        """
        output_file = "省级报告/理科数学考生答题水平分析/各市情况分析(理科数学).xlsx"
        writer = pd.ExcelWriter(output_file)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"002",0,1)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"002",i,1)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市考生成绩比较(理科数学)",index=False)

        writer.save()
        writer.close()

        pr.get_single_km_picture(self.__cursor,lk_ks_ids,"002","省级报告/理科数学考生答题水平分析/全省考生单科成绩分布(理科数学).png")

        return True

    

    """
        考生答题水平分析：英语
    """
    def yingyu_anwsering_situation(self):


        tixing = {
            "阅读理解(必做)":[21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40],
            "完形填空(必做)":[41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60],
            "语法填空(必做)":[61,62,63,64,65,66,67,68,69,70],
            "短文改错(必做)":[71,72,73,74,75,76,77,78,79,80],
            "写作题(必做)":[81]
        }

        zsbk = {
            "应用文阅读(人与社会)(必做)":[21,22,23],
            "记叙文阅读(人与自我)(必做)":[24,25,26,27],
            "说明文阅读(人与社会)(必做)":[28,29,30,31],
            "说明性议论文阅读(人与社会)(必做)":[32,33,34,35],
            "说明文阅读(人与自然)(必做)":[36,37,38,39,40],
            "动词(必做)":[41,45,51,55,56,60],
            "名词(必做)":[42,43,46,50,52,53,57,66,76,77],
            "形容词(必做)":[44,47,48,59,68],
            "动词短语搭配(必做)":[49,54],
            "副词(必做)":[58,62,75,79],
            "主从复合句(必做)":[61,72,63],
            "介词(必做)":[64,67,71,78],
            "非谓语动词(必做)":[64,67,71,78],
            "时态(必做)":[65],
            "冠词(必做)":[69,73],
            "句子成分(必做)":[70],
            "连词(必做)":[74],
            "零冠词(必做)":[80],
            "应用文写作(必做)":[81],

        }

        khnl = {
            "提取与理解(必做)":[21,22,23,24,28,29,30,34,71,76,80],
            "理解与推断(必做)":[25,26,27,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,58,59,60,63,65,74,79],
            "理解与判断(必做)":[31,32,36,38,39,73,77],
            "理解与概括(必做)":[33,35,37],
            "总结与归纳(必做)":[37,40],
            "理解与分析(必做)":[61,64,67,70,72,78],
            "辨认与运用(必做)":[62],
            "提取与辨认词性(必做)":[66],
            "比较与分析(必做)":[68],
            "分析运用冠词(必做)":[69],
            "辨认与分析(必做)":[75],
            "分析与综合(必做)":[81],
        }

        
        """
            原始分分析
        """
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        if not os.path.exists("省级报告"):
            os.makedirs("省级报告")
        if not os.path.exists("省级报告/英语考生答题水平分析"):
            os.makedirs("省级报告/英语考生答题水平分析")

        lk_ks_ids,wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/英语考生答题水平分析/原始分概括(英语).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 全省
        result = pr.km_total_grade_analysis(self.__cursor,self.__ks_ids,"101")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别考生比较(英语)",index=False)

        # 文科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,wk_ks_ids,"101")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别文科考生比较(英语)",index=False)

        # 理科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,lk_ks_ids,"101")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别理科考生比较(英语)",index=False)

        writer.save()
        writer.close()


        """
            结构分析
        """
        output_file = "省级报告/英语考生答题水平分析/结构分析(英语).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 主客观分析
        df = pd.DataFrame(data=None,columns=["主客观题","题数","平均分","标准差","难度","区分度","信度"])
        results = pr.zkg_situation(self.__cursor,"101")
        """
            需要手动更改
        """
        results[0].insert(1,40.00)
        results[1].insert(1,21.00)
        results[2].insert(1,61.00)
        for result in results:
            df.loc[len(df)] = result
        df.to_excel(writer,sheet_name="全省考生客观题得分(英语)",index=False)

        # 全省各题型得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in tixing.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"101",i))
            result = pr.structural_st_analysis(self.__cursor,"101",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各题型得分情况(英语)",index=False)


        # 全省各知识板块得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in zsbk.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"101",i))
            result = pr.structural_st_analysis(self.__cursor,"101",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各知识板块得分情况(英语)",index=False)


        # 全省各考核能力得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in khnl.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"101",i))
            result = pr.structural_st_analysis(self.__cursor,"101",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各考核能力得分情况(英语)",index=False)
        writer.save()
        writer.close()



        """
            单题分析
        """
        output_file = "省级报告/英语考生答题水平分析/单题分析(英语).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 获取所有题号
        st = []
        for value in zsbk.values():
            st.extend(value)

        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])
        for i in st:
            st_ids = pr.get_st_ids(self.__cursor,"101",i)
            result = pr.singlel_st_analysis(self.__cursor,"101",st_ids)
            result.insert(0,i)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="考生单题作答情况(英语)",index=False)
        writer.save()
        writer.close()


        """
            各市情况分析
        """
        output_file = "省级报告/英语考生答题水平分析/各市情况分析(英语).xlsx"
        writer = pd.ExcelWriter(output_file)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"101",0,0)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"101",i,0)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市考生成绩比较(英语)",index=False)


        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"101",0,2)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"101",i,2)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市文科考生成绩比较(英语)",index=False)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"101",0,1)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"101",i,1)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市理科考生成绩比较(英语)",index=False)

        writer.save()
        writer.close()

        pr.get_single_km_picture(self.__cursor,self.__ks_ids,"101","省级报告/英语考生答题水平分析/全省考生单科成绩分布(英语).png")
        pr.get_single_km_picture(self.__cursor,wk_ks_ids,"101","省级报告/英语考生答题水平分析/全省文科考生单科成绩分布(英语).png")
        pr.get_single_km_picture(self.__cursor,lk_ks_ids,"101","省级报告/英语考生答题水平分析/全省理科考生单科成绩分布(英语).png")

        return True
    
    
    """
        考生答题水平分析：文科综合
    """
    def wenkezonghe_anwsering_situation(self):

        tixing = {
            "单项选择题(必做)":[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35],
            "非选择题(必做)":[36,37,38,39,40,41,42],
            "选做题_地理(选做1)":[43],
            "选做题_地理(选做2)":[44],
            "选做题_历史(选做1)":[45],
            "选做题_历史(选做2)":[46],
            "选做题_历史(选做3)":[47]
        }

       

        
        """
            原始分分析
        """
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        if not os.path.exists("省级报告"):
            os.makedirs("省级报告")
        if not os.path.exists("省级报告/文科综合考生答题水平分析"):
            os.makedirs("省级报告/文科综合考生答题水平分析")

        lk_ks_ids,wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/文科综合考生答题水平分析/原始分概括(文科综合).xlsx"
        writer = pd.ExcelWriter(output_file)


        # 文科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,wk_ks_ids,"006")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别考生比较(文科综合)",index=False)

        writer.save()
        writer.close()


        """
            结构分析
        """
        output_file = "省级报告/文科综合考生答题水平分析/结构分析(文科综合).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 主客观分析
        df = pd.DataFrame(data=None,columns=["主客观题","题数","平均分","标准差","难度","区分度","信度"])
        results = pr.zkg_situation(self.__cursor,"006")
        """
            手动改
        """
        results[0].insert(1,35.00)
        results[1].insert(1,12.00)
        results[2].insert(1,47.00)
        for result in results:
            df.loc[len(df)] = result
        df.to_excel(writer,sheet_name="全省考生客观题得分(文科综合)",index=False)

        # 全省各题型得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in tixing.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"006",i))
            result = pr.structural_st_analysis(self.__cursor,"006",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各题型得分情况(文科综合)",index=False)

        writer.save()
        writer.close()

        """
            单题分析
        """
        output_file = "省级报告/文科综合考生答题水平分析/单题分析(文科综合).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 获取所有题号
        st = []
        for value in tixing.values():
            st.extend(value)

        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])
        for i in st:
            st_ids = pr.get_st_ids(self.__cursor,"006",i)
            result = pr.singlel_st_analysis(self.__cursor,"006",st_ids)
            result.insert(0,i)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="考生单题作答情况(文科综合)",index=False)
        writer.save()
        writer.close()


        """
            各市情况分析
        """
        output_file = "省级报告/文科综合考生答题水平分析/各市情况分析(文科综合).xlsx"
        writer = pd.ExcelWriter(output_file)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"006",0,2)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"006",i,2)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市考生成绩比较(文科综合)",index=False)

        writer.save()
        writer.close()

        pr.get_single_km_picture(self.__cursor,wk_ks_ids,"006","省级报告/文科综合考生答题水平分析/全省考生单科成绩分布(文科综合).png")
        

        return True



    """
        考生答题水平分析：理科综合
    """
    
    def likezonghe_anwsering_situation(self):

        tixing = {
            "单项选择题(必做)":[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18],
            "多项选择题(必做)":[19,20,21],
            "非选择题(必做)":[22,23,24,25,26,27,28,29,30,31,32],
            "选做题_物理(选做1)":[33],
            "选做题_物理(选做2)":[34],
            "选做题_化学(选做1)":[35],
            "选做题_化学(选做2)":[36],
            "选做题_生物(选做1)":[37],
            "选做题_生物(选做2)":[38]
        }

        
        """
            原始分分析
        """
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        if not os.path.exists("省级报告"):
            os.makedirs("省级报告")
        if not os.path.exists("省级报告/理科综合考生答题水平分析"):
            os.makedirs("省级报告/理科综合考生答题水平分析")

        lk_ks_ids,wk_ks_ids = pr.judge_ks_wenli(self.__cursor,self.__ks_ids)

        output_file = "省级报告/理科综合考生答题水平分析/原始分概括(理科综合).xlsx"
        writer = pd.ExcelWriter(output_file)


        # 理科
        df = pd.DataFrame(data=None,columns=["维度","人数","比率","平均分","标准差","差异系数"])
        result = pr.km_total_grade_analysis(self.__cursor,lk_ks_ids,"005")
        result.insert(0,"总计")
        df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="各类别考生比较(理科综合)",index=False)

        writer.save()
        writer.close()


        """
            结构分析
        """
        output_file = "省级报告/理科综合考生答题水平分析/结构分析(理科综合).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 主客观分析
        df = pd.DataFrame(data=None,columns=["主客观题","题数","平均分","标准差","难度","区分度","信度"])
        results = pr.zkg_situation(self.__cursor,"005")
        """
            手动改
        """
        results[0].insert(1,21.00)
        results[1].insert(1,17.00)
        results[2].insert(1,38.00)
        for result in results:
            df.loc[len(df)] = result
        df.to_excel(writer,sheet_name="全省考生客观题得分(理科综合)",index=False)

        # 全省各题型得分情况
        df = pd.DataFrame(data=None,columns=["题型","题号","分值","平均分","标准差","难度"])
        for key,value in tixing.items():
            st = []
            for i in value:
                st.extend(pr.get_st_ids(self.__cursor,"005",i))
            result = pr.structural_st_analysis(self.__cursor,"005",st)
            result.insert(0,value)
            result.insert(0,key)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="全省考生各题型得分情况(理科综合)",index=False)

        writer.save()
        writer.close()

        """
            单题分析
        """
        output_file = "省级报告/理科综合考生答题水平分析/单题分析(理科综合).xlsx"
        writer = pd.ExcelWriter(output_file)

        # 获取所有题号
        st = []
        for value in tixing.values():
            st.extend(value)

        df = pd.DataFrame(data=None,columns=["题号","分值","平均分","标准差","难度","区分度"])
        for i in st:
            st_ids = pr.get_st_ids(self.__cursor,"005",i)
            result = pr.singlel_st_analysis(self.__cursor,"005",st_ids)
            result.insert(0,i)
            df.loc[len(df)] = result

        df.to_excel(writer,sheet_name="考生单题作答情况(理科综合)",index=False)
        writer.save()
        writer.close()


        """
            各市情况分析
        """
        output_file = "省级报告/理科综合考生答题水平分析/各市情况分析(理科综合).xlsx"
        writer = pd.ExcelWriter(output_file)

        df = pd.DataFrame(data=None,columns=["城市代码","地市名称","人数","比率","平均分","标准差","差异系数"])
        df.loc[len(df)] = pr.city_single_km(self.__cursor,"005",0,1)
        for i in self.__city_ids:
            result = pr.city_single_km(self.__cursor,"005",i,1)
            if result[2] > 0:
                df.loc[len(df)] = result
        
        df.to_excel(writer,sheet_name="各市考生成绩比较(理科综合)",index=False)


        writer.save()
        writer.close()

        pr.get_single_km_picture(self.__cursor,lk_ks_ids,"005","省级报告/理科综合考生答题水平分析/全省考生单科成绩分布(理科综合).png")

        return True

    


    def test(self):
        print(pr.singlel_st_analysis(self.__cursor,"001",[8]))
       


