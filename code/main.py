from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe
from CJGK import ZFFB_IMG

import cx_Oracle
import numpy as np


def city(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTGK_CITY_TABLE(dsh)
    yw.ZTGK_CITY_IMG(dsh)
    yw.DTFX_CITY_TABLE(dsh)
    yw.DTFX_CITY_APPENDIX(dsh)
    yw.YSFFX_CITY_TABLE(dsh)

    lksx.ZTGK_CITY_TABLE(dsh)
    lksx.ZTGK_CITY_IMG(dsh)
    lksx.DTFX_CITY_TABLE(dsh)
    lksx.DTFX_CITY_APPENDIX(dsh)
    lksx.YSFFX_CITY_TABLE(dsh)

    wksx.ZTGK_CITY_TABLE(dsh)
    wksx.ZTGK_CITY_TABLE(dsh)
    wksx.DTFX_CITY_IMG(dsh)
    wksx.DTFX_CITY_APPENDIX(dsh)
    wksx.YSFFX_CITY_TABLE(dsh)

    yy.ZTGK_CITY_TABLE(dsh)
    yy.ZTGK_CITY_IMG(dsh)
    yy.DTFX_CITY_TABLE(dsh)
    yy.DTFX_CITY_APPENDIX(dsh)
    yy.YSFFX_CITY_TABLE(dsh)


    lkzh.ZTGK_CITY_TABLE(dsh)
    lkzh.ZTGK_CITY_TABLE(dsh)
    lkzh.YSFFX_CITY_TABLE(dsh)


    wkzh.ZTGK_CITY_TABLE(dsh)
    wkzh.ZTGK_CITY_IMG(dsh)
    wkzh.YSFFX_CITY_TABLE(dsh)

def dtfx_city_img(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()

    # yw.DTFX_CITY_IMG(dsh)
    # lksx.DTFX_CITY_IMG(dsh)
    # wksx.DTFX_CITY_IMG(dsh)
    # yy.DTFX_CITY_IMG(dsh)

    yw.DTFX_PROVINCE()

def province():
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likezonghe.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTGK_PROVINCE_TABLE()
    yw.YSFGK_PROVINCE_APPENDIX()

    lksx.ZTGK_PROVINCE_TABLE()
    lksx.YSFGK_PROVINCE_APPENDIX()

    wksx.ZTGK_PROVINCE_TABLE()
    wksx.YSFGK_PROVINCE_APPENDIX()

    yy.ZTGK_PROVINCE_TABLE()
    yy.YSFGK_PROVINCE_APPENDIX()


    lkzh.ZTGK_PROVINCE_TABLE()
    lkzh.YSFGK_PROVINCE_APPENDIX()

    wkzh.ZTGK_PROVINCE_TABLE()
    wkzh.YSFGK_PROVINCE_APPENDIX()






if __name__ == '__main__':

    dtfx_city_img('01')

    province()

    dshs = ["01","02","03","04","05","06","07","08","09","12","13","14","15","16","17","18","19","20","51","52","53"]

    for dsh in dshs:
        city(dsh)





