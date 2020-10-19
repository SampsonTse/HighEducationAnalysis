import getYwData
import getLksxData
import getWksxData
import getWyData
import getWkzhData
import getLksxData
import numpy as np
import matplotlib.pyplot as plt

if __name__ == '__main__':
    YW = getYwData.KSDTSPFX_yw()
    YW.ZTKG_CITY_TABLE('01')
    YW.ZTJG_CITY_IMG('01')

    LKSX = getLksxData.KSDTSPFX_lksx()
    LKSX.ZTKG_CITY_TABLE('01')
    LKSX.ZTJG_CITY_IMG('01')

    WKSX = getWksxData.KSDTSPFX_lksx()
    WKSX.ZTKG_CITY_TABLE('01')
    WKSX.ZTJG_CITY_IMG('01')

    WY = getWyData.KSDTSPFX_WY()
    WY.ZTKG_CITY_TABLE('01')
    WY.ZTJG_CITY_IMG('01')

    LKZH = getLksxData.KSDTSPFX_lkzh()
    LKZH.ZTKG_CITY_TABLE('01')
    LKZH.ZTJG_CITY_IMG('01')

    WKZH = getWkzhData.KSDTSPFX_wkzh()
    WKZH.ZTKG_CITY_TABLE('01')
    WKZH.ZTJG_CITY_IMG('01')