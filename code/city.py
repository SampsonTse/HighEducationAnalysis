from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe
from KSDTSPFX.LKZHDK import wuli,shengwu,huaxue
from KSDTSPFX.WKZHDK import lishi,zhengzhi,dili


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

    # 总体概括
    def ztgk(self):

        self.zz.ZTGK_CITY_TABLE(self.dsh)
        self.dl.ZTGK_CITY_TABLE(self.dsh)
        self.ls.ZTGK_CITY_TABLE(self.dsh)
        self.sw.ZTGK_CITY_TABLE(self.dsh)
        self.wl.ZTGK_CITY_TABLE(self.dsh)
        self.hx.ZTGK_CITY_TABLE(self.dsh)

        self.zz.ZTGK_CITY_IMG(self.dsh)
        self.dl.ZTGK_CITY_IMG(self.dsh)
        self.ls.ZTGK_CITY_IMG(self.dsh)
        self.sw.ZTGK_CITY_IMG(self.dsh)
        self.wl.ZTGK_CITY_IMG(self.dsh)
        self.hx.ZTGK_CITY_IMG(self.dsh)

        self.yw.ZTGK_CITY_IMG(self.dsh)
        self.lksx.ZTGK_CITY_IMG(self.dsh)
        self.wksx.ZTGK_CITY_IMG(self.dsh)
        self.yy.ZTGK_CITY_IMG(self.dsh)
        self.lkzh.ZTGK_CITY_IMG(self.dsh)
        self.wkzh.ZTGK_CITY_IMG(self.dsh)

        self.yw.ZTGK_CITY_TABLE(self.dsh)
        self.lksx.ZTGK_CITY_TABLE(self.dsh)
        self.wksx.ZTGK_CITY_TABLE(self.dsh)
        self.yy.ZTGK_CITY_TABLE(self.dsh)
        self.wkzh.ZTGK_CITY_TABLE(self.dsh)
        self.lkzh.ZTGK_CITY_TABLE(self.dsh)

    # 单题分析 综合单科图片没做
    def dtfx(self):
        # self.zz.DTFX_CITY_TABLE(self.dsh)
        # self.dl.DTFX_CITY_TABLE(self.dsh)
        # self.ls.DTFX_CITY_TABLE(self.dsh)
        # self.sw.DTFX_CITY_TABLE(self.dsh)
        # self.wl.DTFX_CITY_TABLE(self.dsh)
        # self.hx.DTFX_CITY_TABLE(self.dsh)

        self.yw.DTFX_CITY_TABLE(self.dsh)
        self.wksx.DTFX_CITY_TABLE(self.dsh)
        self.lksx.DTFX_CITY_TABLE(self.dsh)
        self.yy.DTFX_CITY_TABLE(self.dsh)

        self.yw.DTFX_CITY_IMG(self.dsh)
        self.wksx.DTFX_CITY_IMG(self.dsh)
        self.lksx.DTFX_CITY_IMG(self.dsh)
        self.yy.DTFX_CITY_IMG(self.dsh)

    # 结构分析(暂时做不了)
    def jgfx(self):
        pass

    def test(self):
        # self.hx.DTFX_PROVINCE_APPENDIX()
        # self.sw.DTFX_PROVINCE_APPENDIX()
        # self.wl.DTFX_PROVINCE_APPENDIX()
        # self.ls.DTFX_PROVINCE_APPENDIX()
        # self.zz.DTFX_PROVINCE_APPENDIX()
        self.dl.DTFX_PROVINCE_APPENDIX()


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

    # 原始分概括
    def ysfgk(self):
        self.yw.YSFFX_CITY_TABLE(self.dsh)
        self.lksx.YSFFX_CITY_TABLE(self.dsh)
        self.wksx.YSFFX_CITY_TABLE(self.dsh)
        self.yy.YSFFX_CITY_TABLE(self.dsh)
        self.lkzh.YSFFX_CITY_TABLE(self.dsh)
        self.wkzh.YSFFX_CITY_TABLE(self.dsh)

    # 单题分析
    def dtfx(self):
        self.yw.DTFX_CITY_APPENDIX(self.dsh)
        self.wksx.DTFX_CITY_APPENDIX(self.dsh)
        self.lksx.DTFX_CITY_APPENDIX(self.dsh)
        self.yy.DTFX_CITY_APPENDIX(self.dsh)

    # 结构分析(暂时做不了)
    def jgfx(self):
        pass


