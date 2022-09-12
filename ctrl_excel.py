import sys, os, shutil, glob, time, datetime
import xlwings as xlw
from file_handle import MyFile

class ExcelFile(MyFile):
    def set_wb(self):
        self.wb = xlw.Book(self.fpath)
        return self.wb

    def set_wsht(self, sht_name):
        sht = self.wb.sheets(sht_name)
        return sht

    def set_rng(self, sht_name, rng_val):
        sht = self.set_wsht(sht_name)
        return sht.range(rng_val)

    # def get_header_row(self, sht, basecol='A', baserow = 1):
    #     baserng = basecol + str(baserow)
    #     return sht.Range(baserng).end('up').row

    # def get_lastrow(self, sht, endrow='A1048576'):
    #     return sht.range(endrow).end('up').row
    #
    # def get_lastcolumn(self, sht, endcol='XFD1'):
    #     return sht.range(endcol).end('left').columns
    #
    # def set_range(self, startrow, startcol, endrow, endcol):
    #     return startcol + str(startrow) + ':' + endcol + str(endrow)

    # def read_this_excel(self, sheet=''):
    #     if sheet == '':
    #         self.df = pd.read_excel(self.path)
    #     else:
    #         self.df = pd.read_excel(self.path, sheet_name=sheet)


class CtrlExcel:
    def input_data(self, rng, df):
        rng.options(header=False, index=False).value = df
        return 0

    def overwrite_result(self, rng, df):
        self.del_allval_inrange(rng)
        rng.options(pd.DataFrame, header=False, index=True). value = df

    def del_allval_inrange(self, rng):
        # 対象シート内の範囲を取得
        rng.clear_contents()

    def set_backgroundcolor(self, rng, color_r, color_g, color_b):
        rng.color = color_r, color_g, color_b


