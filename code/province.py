from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe
from KSDTSPFX.LKZHDK import wuli,shengwu,huaxue
from KSDTSPFX.WKZHDK import lishi,zhengzhi,dili


class pro_report:
    def __init__(self):
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

    # 原始分概况
    def ysfgk(self):

        self.wl.YSFGK_PROVICNE_TABLE()
        self.sw.YSFGK_PROVICNE_TABLE()
        self.hx.YSFGK_PROVICNE_TABLE()
        self.dl.YSFGK_PROVICNE_TABLE()
        self.zz.YSFGK_PROVICNE_TABLE()
        self.ls.YSFGK_PROVICNE_TABLE()

        self.wl.YSFGK_PROVINCE_IMG()
        self.sw.YSFGK_PROVINCE_IMG()
        self.hx.YSFGK_PROVINCE_IMG()
        self.dl.YSFGK_PROVINCE_IMG()
        self.zz.YSFGK_PROVINCE_IMG()
        self.ls.YSFGK_PROVINCE_IMG()

        self.yw.YSFGK_PROVINCE_IMG()
        self.yw.YSFGK_PROVINCE_TABLE()

        self.lksx.YSFGK_PROVINCE_IMG()
        self.lksx.YSFGK_PROVINCE_TABLE()

        self.wksx.YSFGK_PROVINCE_IMG()
        self.wksx.YSFGK_PROVINCE_TABLE()

        self.yy.YSFGK_PROVINCE_IMG()
        self.yy.YSFGK_PROVINCE_TABLE()

        self.lkzh.YSFGK_PROVINCE_IMG()
        self.lkzh.YSFGK_PROVINCE_TABLE()

        self.wkzh.YSFGK_PROVINCE_IMG()
        self.wkzh.YSFGK_PROVINCE_TABLE()

    # 单题分析
    def dtfx(self):

        self.sw.DTFX_PROVINCE()
        # self.hx.DTFX_PROVINCE()
        # self.wl.DTFX_PROVINCE()
        # self.dl.DTFX_PROVINCE()
        # self.ls.DTFX_PROVINCE()
        # self.zz.DTFX_PROVINCE()
        #
        # self.yw.DTFX_PROVINCE()
        # self.lksx.DTFX_PROVINCE()
        # self.wksx.DTFX_PROVINCE()
        # self.yy.DTFX_PROVINCE()

    # 各市情况分析
    def gkqkfx(self):

        self.sw.GSQKFX_PROVINCE()
        self.hx.GSQKFX_PROVINCE()
        self.wl.GSQKFX_PROVINCE()
        self.dl.GSQKFX_PROVINCE()
        self.ls.GSQKFX_PROVINCE()
        self.zz.GSQKFX_PROVINCE()

        self.yw.GSQKFX_PROVINCE()
        self.lksx.GSQKFX_PROVINCE()
        self.wksx.GSQKFX_PROVINCE()
        self.yy.GSQKFX_PROVINCE()
        self.lkzh.GSQKFX_PROVINCE()
        self.wkzh.GSQKFX_PROVINCE()

    # 结构分析(暂时无法完成)
    def jgfx(self):
        pass


    def test(self):
        self.sw.MF_LF_PROVINCE_TABLE()

class pro_report_appendix:
    def __init__(self):

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

        self.sw.YSFGK_PROVINCE_APPENDIX()
        self.hx.YSFGK_PROVINCE_APPENDIX()
        self.wl.YSFGK_PROVINCE_APPENDIX()
        self.dl.YSFGK_PROVINCE_APPENDIX()
        self.ls.YSFGK_PROVINCE_APPENDIX()
        self.zz.YSFGK_PROVINCE_APPENDIX()

        self.yw.YSFGK_PROVINCE_APPENDIX()
        self.lksx.YSFGK_PROVINCE_APPENDIX()
        self.wksx.YSFGK_PROVINCE_APPENDIX()
        self.yy.YSFGK_PROVINCE_APPENDIX()
        self.lkzh.YSFGK_PROVINCE_APPENDIX()
        self.wkzh.YSFGK_PROVINCE_APPENDIX()

    # 单题分析
    def dtfx(self):
        self.wl.DTFX_PROVINCE_APPENDIX()
        self.sw.DTFX_PROVINCE_APPENDIX()
        # self.hx.DTFX_PROVINCE_APPENDIX()
        #
        # self.dl.DTFX_PROVINCE_APPENDIX()
        # self.ls.DTFX_PROVINCE_APPENDIX()
        # self.zz.DTFX_PROVINCE_APPENDIX()
        #
        # self.yw.DTFX_PROVINCE_APPENDIX()
        # self.lksx.DTFX_PROVINCE_APPENDIX()
        # self.wksx.DTFX_PROVINCE_APPENDIX()
        # self.yy.DTFX_PROVINCE_APPENDIX()

    # 各市情况分析(暂时无法完成)
    def gkqkfx(self):
        pass

    # 结构分析(暂时无法完成)
    def jgfx(self):
        pass


