from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe

def GetData(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likeshuxue.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTJG_CITY_IMG(dsh)
    lksx.ZTKG_CITY_TABLE(dsh)
    wksx.ZTJG_CITY_IMG(dsh)
    yy.ZTJG_CITY_IMG(dsh)
    lkzh.ZTJG_CITY_IMG(dsh)
    wkzh.ZTJG_CITY_IMG(dsh)




if __name__ == '__main__':
    dsh = '01'
    GetData(dsh=dsh)