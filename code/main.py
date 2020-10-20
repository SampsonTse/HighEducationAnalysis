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

    yw.ZTKG_PROVINCE_TABLE()




if __name__ == '__main__':
    dsh = '01'
    GetData(dsh=dsh)