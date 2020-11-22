from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe
from KSDTSPFX.LKZHDK import wuli,shengwu,huaxue
from KSDTSPFX.WKZHDK import lishi,zhengzhi,dili
from KSDTSPFX.WKZHDK import zhengzhi

class city_report:
    def __init__(self,dsh):
        self.dsh = dsh
        self.yw = yuwen.DTFX()
        self.lksx = likeshuxue.DTFX()
        self.wksx = wenkeshuxue.DTFX()
        self.yy = yingyu.DTFX()
        self.lkzh = likezonghe.DTFX()
        self.wkzh = wenkezonghe.DTFX()

        self.wl = wuli.DTFX()
        self.sw = shengwu.DTFX()
        self.hx = huaxue.DTFX()

        self.dl = dili.DTFX()
        self.zz = zhengzhi.DTFX()
        self.ls = lishi.DTFX()

    def __del__(self):
        del self.yw
        del self.wksx
        del self.lksx
        del self.yy
        del self.wkzh
        del self.lkzh

        del self.sw
        del self.wl
        del self.hx
        del self.dl
        del self.zz
        del self.ls

    def test2(self):
        self.lksx.DTFX_CITY_IMG(self.dsh)
        self.lksx.DTFX_CITY_TABLE(self.dsh)
        self.wksx.DTFX_CITY_IMG(self.dsh)
        self.wksx.DTFX_CITY_TABLE(self.dsh)

    def test(self):
        self.yw.GQXZB_CITY_TABLE(self.dsh)
        print("语文")
        self.lksx.GQXZB_CITY_TABLE(self.dsh)
        print("理科数学")
        self.wksx.GQXZB_CITY_TABLE(self.dsh)
        print("文科数学")
        self.yy.GQXZB_CITY_TABLE(self.dsh)
        print("英语")
        self.wl.GQXZB_CITY_TABLE(self.dsh)
        print("物理")
        self.hx.GQXZB_CITY_TABLE(self.dsh)
        print("化学")
        self.sw.GQXZB_CITY_TABLE(self.dsh)
        print("生物")
        self.ls.GQXZB_CITY_TABLE(self.dsh)
        print("历史")
        self.dl.GQXZB_CITY_TABLE(self.dsh)
        print("地理")
        self.zz.GQXZB_CITY_TABLE(self.dsh)
        print("政治")


class city_report_appendix:
    def __init__(self, dsh):
        self.dsh = dsh
        self.yw = yuwen.DTFX()
        self.lksx = likeshuxue.DTFX()
        self.wksx = wenkeshuxue.DTFX()
        self.yy = yingyu.DTFX()
        self.lkzh = likezonghe.DTFX()
        self.wkzh = wenkezonghe.DTFX()

        self.wl = wuli.DTFX()
        self.sw = shengwu.DTFX()
        self.hx = huaxue.DTFX()

        self.dl = dili.DTFX()
        self.zz = zhengzhi.DTFX()
        self.ls = lishi.DTFX()

    def __del__(self):
        del self.yw
        del self.wksx
        del self.lksx
        del self.yy
        del self.wkzh
        del self.lkzh

        del self.sw
        del self.wl
        del self.hx
        del self.dl
        del self.zz
        del self.ls





