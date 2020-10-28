from KSDTSPFX import yuwen
from KSDTSPFX import likeshuxue
from KSDTSPFX import wenkeshuxue
from KSDTSPFX import yingyu
from KSDTSPFX import likezonghe
from KSDTSPFX import wenkezonghe


class pro_report:
    def __init__(self):
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
        del self.lkzh

    # 原始分概况
    def ysfgk(self):
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
        self.yw.DTFX_PROVINCE()
        self.lksx.DTFX_PROVINCE()
        self.wksx.DTFX_PROVINCE()
        self.yy.DTFX_PROVINCE()

    # 各市情况分析
    def gkqkfx(self):
        self.yw.GSQKFX_PROVINCE()
        self.lksx.GSQKFX_PROVINCE()
        self.wksx.GSQKFX_PROVINCE()
        self.yy.GSQKFX_PROVINCE()
        self.lkzh.GSQKFX_PROVINCE()
        self.wkzh.GSQKFX_PROVINCE()

    # 结构分析(暂时无法完成)
    def jgfx(self):
        pass

class pro_report_appendix:
    def __init__(self):
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
        self.yw.YSFGK_PROVINCE_APPENDIX()
        self.lksx.YSFGK_PROVINCE_APPENDIX()
        self.wksx.YSFGK_PROVINCE_APPENDIX()
        self.yy.YSFGK_PROVINCE_APPENDIX()
        self.lkzh.YSFGK_PROVINCE_APPENDIX()
        self.wkzh.YSFGK_PROVINCE_APPENDIX()

    # 单题分析
    def dtfx(self):
        self.yw.DTFX_PROVINCE_APPENDIX()
        self.lksx.DTFX_PROVINCE_APPENDIX()
        self.wksx.DTFX_PROVINCE_APPENDIX()
        self.yy.DTFX_PROVINCE_APPENDIX()

    # 各市情况分析(暂时无法完成)
    def gkqkfx(self):
        pass

    # 结构分析(暂时无法完成)
    def jgfx(self):
        pass
