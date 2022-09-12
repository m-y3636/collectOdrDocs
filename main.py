import sys, os, shutil, time, datetime
from set_variables import *
from collect_odr_docs import *
from read_elmes_info import ElmesInfo
from ctrl_excel import ExcelFile
from tkinter import messagebox


class OdrInfo:
    def __init__(self, odr_num):
        self.odr_num = odr_num
        self.odr_name = ''
        self.exm_num = r''
        self.dwg_num = r''
        self.dwg_name = r''
        self.result_dwgsender = None
        self.result_elena = None

    def set_exmnum(self, exm_num):
        self.exm_num = exm_num

    def set_old_odr_num(self, old_odr_num):
        self.old_odr_num = old_odr_num

    def set_rev_odr_nums(self, rev_odr_nums):
        self.rev_odr_nums = rev_odr_nums

    def set_dl_dwg_names(self, dwg_names):
        self.dl_dwg_list = dwg_names

    def get_odrinfo_by_elmes(self):
        old_odr, exm_num = ElmesInfo(self.odr_num).read_info()
        if old_odr: self.old_odr_num = old_odr
        if exm_num: self.exm_num = exm_num


if __name__ == '__main__':
    print('開始：強制終了するには\'Ctrl＋C\'を押してください。')
    # ver確認
    check_ver = CheckVer(tool_fpath, root_dpath)
    check_ver.check_ver()

    os.environ['WDM_SSL_VERIFY'] = '0'

    # 結果表示UIのVBAをインスタンス化
    tool = ExcelFile()
    tool.set_fpath(tool_fpath)

    # オーダー情報をインスタンス化してリストに格納
    odr_data_list = []
    for odr in odr_list:
        odr_data = OdrInfo(odr)
        odr_data.set_dl_dwg_names(dwg_list)
        odr_data_list.append(odr_data)

    # # ELMES情報の読取り(新設時オーダー番号、検討依頼番号追加)
    # for odr_data in odr_data_list:
    #     odr_data.get_odrinfo_by_elmes(odr_data)

    # ★一通りログイン？★
    if 1 in doc_list:
        crawl1 = CrawlSendDwg(tmp_dl_dpath)
        crawl1.get_target_site('http://portal.ina.melco.co.jp/portal/zumen/Main.do')
        crawl1.common_login()
    if 2 in doc_list:
        crawl2 = CrawlElena(tmp_dl_dpath)
        crawl2.get_target_site('https://www.mitsubishi-elevator.net/elena2nd/S01_init')
        crawl2.common_login()
    if 3 in doc_list:
        crawl3 = CrawlEnet(tmp_dl_dpath)
        crawl3.login()
    if 4 in doc_list:
        crawl4 = CrawlOfficeStaff(tmp_dl_dpath)
        crawl4.login()
    if 5 in doc_list:
        crawl5 = CrawlDelisis(tmp_dl_dpath)
        crawl5.login()
    if 6 in doc_list:
        crawl6 = CrawlMetis(tmp_dl_dpath)
        crawl6.login()
    os.environ['WDM_SSL_VERIFY'] = '1'

    # # # Eネット(旧工事番号、改修工事番号取得) # モダニのみ(#4, #5検索に利用)
    # odr_data = []
    # for odr_data in odr_data_list:
    #     old_odr, rev_odrs = crawl1.get_odr_info()
    #     odr_data.set_old_odr_num(old_odr)
    #     odr_data.set_rev_odr_nums(rev_odrs)

    # Eネット検索
    if 1 in doc_list: crawl1.dl_result_docs(odr_data_list)
    # # 図面伝送
    if 2 in doc_list: crawl2.dl_result_docs(odr_data_list)
    # # ELENA-Ⅱ
    if 3 in doc_list: crawl3.dl_application(odr_data_list)
    # # OfficeSTAFF
    if 4 in doc_list: crawl4.dl_old_itemtable(odr_data_list)
    # # DELISIS
    if 5 in doc_list: crawl5.dl_delisis_docs(odr_data_list)
    # # METIS
    if 6 in doc_list: crawl6.dl_metis_docs(odr_data_list)

    if doc_list ==[]:
        print('setting error')
        sys.exit()
    # odr_data_listを返してエクセルに結果表示
    tool_wb = tool.set_wb()
    for i, odr_data in enumerate(odr_data_list):
        rng_val = f'B{str(i + 14)}'
        tool_rng = tool.set_rng('PrintDWG', rng_val)
        if (odr_data.result_dwgsender == False) and (odr_data.result_dwgsender == False):
            tool_rng.color = 255, 204, 204
            err_flg = 1
        else:
            tool_rng.color = 204, 255, 204

    messagebox.showinfo('Result', '処理完了')
