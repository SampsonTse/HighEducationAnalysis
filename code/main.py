from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe
from CJGK import ZFFB_IMG

import cx_Oracle
import numpy as np

# 市级报告 总体概括 图片
def ztgk_city_img(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTJG_CITY_IMG(dsh)
    lksx.ZTJG_CITY_IMG(dsh)
    wksx.ZTJG_CITY_IMG(dsh)
    yy.ZTJG_CITY_IMG(dsh)
    lkzh.ZTJG_CITY_IMG(dsh)
    wkzh.ZTJG_CITY_IMG(dsh)

# 市级报告 总体概括 表格
def ztgk_city_table(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()


    # yw.ZTKG_CITY_TABLE(dsh)
    lksx.ZTKG_CITY_TABLE(dsh)
    # wksx.ZTKG_CITY_TABLE(dsh)
    # yy.ZTKG_CITY_TABLE(dsh)
    # lkzh.ZTKG_CITY_TABLE(dsh)
    # wkzh.ZTKG_CITY_TABLE(dsh)

# 市级报告 单题分析 表格
def dtfx_city_table(dsh):

    yw = yuwen.DTFX()
    yw.DTFX_CITY_TABLE(dsh)
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()

    # lksx.DTFX_CITY_TABLE(dsh)
    # wksx.DTFX_CITY_TABLE(dsh)
    # yy.DTFX_CITY_TABLE(dsh)

# 省级报告 总体概括 表格
def ztgt_province_table():
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTKG_PROVINCE_TABLE()
    lksx.ZTKG_PROVINCE_TABLE()
    wksx.ZTKG_PROVINCE_TABLE()
    yy.ZTKG_PROVINCE_TABLE()
    lkzh.ZTKG_PROVINCE_TABLE()
    wkzh.ZTKG_PROVINCE_TABLE()

# 市级报告附录 原始分分析
def ysffx_city_table(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.YSFFX_CITY_TABLE(dsh)
    lksx.YSFFX_CITY_TABLE(dsh)
    yy.YSFFX_CITY_TABLE(dsh)
    wksx.YSFFX_CITY_TABLE(dsh)
    lkzh.YSFFX_CITY_TABLE(dsh)
    wkzh.YSFFX_CITY_TABLE(dsh)



if __name__ == '__main__':

    # ztgk_city_table('01')
    # ztgk_city_img('01')
    dtfx_city_table('01')
    # ysffx_city_table('01')

    # zffb = ZFFB_IMG.zffb_img()
    # zffb.getImg(1,'01')
    # zffb.getImg(2,'01')

