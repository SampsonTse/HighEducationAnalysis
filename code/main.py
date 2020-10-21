from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe

def GetTable_Province(dsh):
    yw = yuwen.DTFX()
    lksx = likeshuxue.DTFX()
    wksx = wenkeshuxue.DTFX()
    yy = yingyu.DTFX()
    lkzh = likeshuxue.DTFX()
    wkzh = wenkezonghe.DTFX()

    yw.ZTKG_PROVINCE_TABLE()
    lksx.ZTKG_PROVINCE_TABLE()
    wksx.ZTKG_PROVINCE_TABLE()
    yy.ZTKG_PROVINCE_TABLE()
    lkzh.ZTKG_PROVINCE_TABLE()
    wkzh.ZTKG_PROVINCE_TABLE()

if __name__ == '__main__':
    dsh = '01'
    GetTable_Province(dsh=dsh)