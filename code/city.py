from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe


class city_report:
    def __init__(self,dsh):
        self.dsh = dsh
        self.yw = yuwen.DTFX()
        self.lksx = likeshuxue.DTFX()
        self.wksx = wenkeshuxue.DTFX()
        self.yy = yingyu.DTFX()
        self.lkzh = likezonghe.DTFX()
        self.wkzh = wenkezonghe.DTFX()

    def __del__(self):
        del self.yw
        del self.wksx
        del self.lksx
        del self.yy
        del self.wkzh
        del self.lkzx

    # 总体概括
    def ztgk(self):
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

    # 单题分析(图没有验证)
    def dtfx(self):
        self.yw.DTFX_CITY_TABLE(self.dsh)
        self.kwsx.DTFX_CITY_TABLE(self.dsh)
        self.lksx.DTFX_CITY_TABLE(self.dsh)
        self.yy.DTFX_CITY_TABLE(self.dsh)

        self.yw.DTFX_CITY_IMG(self.dsh)
        self.wksx.DTFX_CITY_IMG(self.dsh)
        self.lksx.DTFX_CITY_IMG(self.dsh)
        self.yy.DTFX_CITY_IMG(self.dsh)

    # 结构分析(暂时做不了)
    def jgfx(self):
        pass


class city_report_appendix:
    def __init__(self,ksh):
        self.ksh = ksh
        self.yw = yuwen.DTFX()
        self.lksx = likeshuxue.DTFX()
        self.wksx = wenkeshuxue.DTFX()
        self.yy = yingyu.DTFX()
        self.lkzh = likezonghe.DTFX()
        self.wkzh = wenkezonghe.DTFX()

    def __del__(self):
        del self.yw
        del self.wksx
        del self.lksx
        del self.yy
        del self.wkzh
        del self.lkzx

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
        self.wlsx.DTFX_CITY_APPENDIX(self.dsh)

    # 结构分析(暂时做不了)
    def jgfx(self):
        pass