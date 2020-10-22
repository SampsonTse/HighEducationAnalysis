from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe

import cx_Oracle
import numpy as np

def GetKSDTSPFX(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTJG_CITY_IMG(dsh)
    yw.ZTKG_CITY_TABLE(dsh)
    yw.ZTKG_PROVINCE_TABLE()

    lksx.ZTJG_CITY_IMG(dsh)
    lksx.ZTKG_CITY_TABLE(dsh)
    lksx.ZTKG_PROVINCE_TABLE()

    wksx.ZTJG_CITY_IMG(dsh)
    wksx.ZTKG_CITY_TABLE(dsh)
    wksx.ZTKG_PROVINCE_TABLE()

    yy.ZTJG_CITY_IMG(dsh)
    yy.ZTKG_CITY_TABLE(dsh)
    yy.ZTKG_PROVINCE_TABLE()

    lkzh.ZTJG_CITY_IMG(dsh)
    lkzh.ZTKG_CITY_TABLE(dsh)
    lkzh.ZTKG_PROVINCE_TABLE()

    wkzh.ZTJG_CITY_IMG(dsh)
    wkzh.ZTKG_CITY_TABLE(dsh)
    wkzh.ZTKG_PROVINCE_TABLE()


if __name__ == '__main__':



    conn = cx_Oracle.connect('gkeva2020/ksy#2020#reta@10.0.200.103/ksydb01std')
    cursor = conn.cursor()
    sql = "select DS_H from C_DS"
    dshs = cursor.execute(sql)
    dshs = np.array(cursor.fetchall()).flatten()

    for dsh in dshs:
        GetKSDTSPFX(dsh)
