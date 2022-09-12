import sys, os, shutil, glob, time, datetime
import tkinter as tk
from err_handle import *
from crawler import CrawlToTargetSite
from file_handle import OdrFile
from set_variables import debug


class CheckVer:
    def __init__(self, tool_fpath, root_dpath):
        self.tool_fpath = tool_fpath
        self.root_dpath = root_dpath
        self.tool_name = os.path.basename(tool_fpath)

    def check_ver(self):
        latest_tool = glob.glob(self.root_dpath + r'\工事ファイル収集ツール_Ver*.xlsm')[0]
        latest_fname = os.path.basename(latest_tool)
        if self.tool_name != latest_fname:
            tk.messagebox('ツールが最新版ではありません。\n更新してください。')
            sys.exit()


# 図面伝送のクローリング用クラス
class CrawlSendDwg(CrawlToTargetSite):
    @ErrHandle.err_handle(999)
    def __init__(self, dl_dir):
        super(CrawlSendDwg, self).__init__()
        self.site_name = r'図面伝送'
        self.dl_dir = dl_dir

    @ErrHandle.err_handle(999)
    def set_root_url(self, url):
        self.url = url

    @ErrHandle.err_handle(999)
    def get_target_site(self, url):
        self.set_root_url(url)
        self.set_config()
        self.execute_Edge()
        self.get_root_html()

    # 共通認証
    @ErrHandle.err_handle(210)
    def common_login(self):
        # debug用ログイン機能
        if debug == 1:
            self.input_text_by_name('ID', 'LH21475')
            self.input_text_by_name('PASS', '&Yuuchan362422')
            self.onclick_js('logon_button()')
        # 入力を待つ
        for i in range(60):
            time.sleep(5)
            # ログインできたか確認する
            if self.driver.current_url == self.url:
                break

    # 纏め処理
    @ErrHandle.err_handle(200)
    def dl_result_docs(self, odr_data_list):
        # 工事番号分繰り返し
        for odr_data in odr_data_list:
            # 検索
            self.prog_odr = odr_data
            self.search_odr_num()
            # 検索結果表から文字列検索してDL
            self.dl_from_dwglist()
            # frame戻る
            self.driver.switch_to.default_content()
            self.down_iframe(0)
            self.down_frame(0)
            self.click_element_by_xpath('/html/body/table/tbody/tr/td[1]/a')
            self.driver.switch_to.default_content()
            if self.prog_odr.result_dwgsender != False:
                self.prog_odr.result_dwgsender = True
        self.close_browser()
        self.quit_browser()

    @ErrHandle.err_handle(220)
    def search_odr_num(self):
        self.down_iframe(0)
        self.down_frame(1)
        self.down_frame(0)
        # チェックボックス操作
        self.check_on_by_name('newFlag')
        self.check_on_by_name('modaniFlag')
        # テキスト削除
        self.text_clear_by_name('orderNo')
        self.input_text_by_name('orderNo', self.prog_odr.odr_num[0:8])
        self.onclick_js('search()')
        self.driver.switch_to.default_content()
        self.down_iframe(0)
        self.down_frame(1)
        self.down_frame(1)
        self.click_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[1]/a')
        self.driver.switch_to.default_content()

    @ErrHandle.err_handle(230)
    def dl_from_dwglist(self):
        self.down_iframe(0)
        self.down_frame(1)
        self.down_frame(1)
        # 物件名取得
        odr_name = self.get_html_attr('/html/body/form/table[1]/tbody/tr[1]/td[5]', "textContent")[0]
        self.prog_odr.odr_name = odr_name.strip()
        # self.odr_data.odr_name = odr_name
        # 取得対象の据付図分繰り返す
        odr_rows = self.driver.find_elements_by_xpath('/html/body/form/table[3]/tbody/tr')
        del odr_rows[0:3]
        for dwg in self.prog_odr.dl_dwg_list:
            self.prog_odr.dwg_name = dwg
            self.search_dwg_intable(dwg, odr_rows)
            # for odr_row in odr_rows:
            #     dwg_name_elem = odr_row.find_element_by_xpath('td[1]')
            #     dwg_num_elem = odr_row.find_element_by_xpath('td[3]')
            #     click_rev_elem = odr_row.find_element_by_xpath('td[6]')
            #     doc_type_elem = odr_row.find_element_by_xpath('td[5]')
            #     dwg_name = dwg_name_elem.text.strip()
            #     dwg_num = dwg_num_elem.text.strip()
            #     dwg_rev = click_rev_elem.text.strip()
            #     doc_type = doc_type_elem.text.strip()
            #     self.prog_dwg_num = dwg_num
            #     if dwg_name == dwg:
            #         self.click_dl_cell(odr_rows)
            #         doc_name = dwg
            #         for i in range(5):
            #             # クリック
            #             click_rev_elem.click()
            #             # DLファイルをインスタンス化して、情報纏め
            #             odr_file = OdrFile()
            #             odr_file.set_dl_site(self.site_name)
            #             odr_file.set_dwg_data(dwg_num, dwg_rev, dwg_name, odr_num, odr_name, doc_type)
            #             odr_file.mv_to_resultd()
            #             if os.path.exists(odr_file.result_fpath):
            #                 break
            #         if not os.path.exists(odr_file.result_fpath):
            #             raise Exception

    @ErrHandle.err_handle(230)
    def search_dwg_intable(self, dwg, odr_rows):
        self.prog_odr.dwg_name = dwg
        for odr_row in odr_rows:
            self.exe_dl_intabe(odr_row)

    @ErrHandle.err_handle(230)
    def exe_dl_intabe(self, odr_row):
        dwg_name_elem = odr_row.find_element_by_xpath('td[1]')
        dwg_num_elem = odr_row.find_element_by_xpath('td[3]')
        click_rev_elem = odr_row.find_element_by_xpath('td[6]')
        doc_type_elem = odr_row.find_element_by_xpath('td[5]')
        dwg_name = dwg_name_elem.text.strip()
        dwg_num = dwg_num_elem.text.strip()
        dwg_rev = click_rev_elem.text.strip()
        doc_type = doc_type_elem.text.strip()
        if dwg_name == self.prog_odr.dwg_name:
            self.prog_odr.dwg_num = dwg_num
            for i in range(5):
                click_rev_elem.click()
                odr_file = OdrFile()
                odr_file.set_dl_site(self.site_name)
                odr_file.set_dwg_data(dwg_num, dwg_rev, dwg_name, self.prog_odr.odr_num, self.prog_odr.odr_name, doc_type)
                odr_file.mv_to_resultd()
                if os.path.exists(odr_file.result_fpath):
                    break
            if not os.path.exists(odr_file.result_fpath):
                raise Exception


# ELENA-Ⅱのクローリング用クラス
class CrawlElena(CrawlToTargetSite):
    def __init__(self, dl_dir):
        super(CrawlElena, self).__init__()
        self.site_name = r'ELENA-Ⅱ'
        self.dl_dir = dl_dir

    @ErrHandle.err_handle(999)
    def set_root_url(self, url):
        self.url = url

    @ErrHandle.err_handle(999)
    def get_target_site(self, url):
        self.set_root_url(url)
        self.set_config()
        self.execute_Edge()
        self.get_root_html()

    # Darwin_Login
    @ErrHandle.err_handle(310)
    def common_login(self):
        # debug用ログイン機能
        if debug == 1:
            self.input_text_by_name('IDToken1', 'Oya.Hiroyuki@zy.MitsubishiElectric.co.jp')
            self.input_text_by_name('IDToken2', 'hiro33##')
            self.onclick_js("LoginSubmit('ログイン')")
        # 入力を待つ
        for i in range(60):
            time.sleep(5)
            # ログインできたか確認する
            if self.driver.current_url == self.url:
                break

    # 纏め処理
    @ErrHandle.err_handle(300)
    def dl_result_docs(self, odr_data_list):
        # 工事番号分繰り返し
        for odr_data in odr_data_list:
            # 検索
            self.prog_odr = odr_data
            self.search_odr_num()
            # 検索結果表から文字列検索してDL
            self.dl_from_dwglist()
            # self.click_element_by_xpath('/html/body/table/tbody/tr/td[1]/a')
            if self.prog_odr.result_elena != False:
                self.prog_odr.result_elena = True
        self.close_browser()
        self.quit_browser()

    @ErrHandle.err_handle(320)
    def search_odr_num(self):
        # チェックボックス操作
        # self.check_on_by_name('newFlag')
        # self.check_on_by_name('modaniFlag')
        # テキスト削除
        self.text_clear_by_id('melcoNo')
        self.input_text_by_id('melcoNo', self.prog_odr.odr_num[0:8])
        self.click_element_by_id('bSearch')
        self.click_element_by_xpath('/html/body/form[1]/div[2]/div[2]/div/div[3]/table/tbody/tr/td[5]/button[2]')

    @ErrHandle.err_handle(330)
    def dl_from_dwglist(self):
        # ★新規ウィンドウへ移動
        self.move_to_new_tab()
        # target_url = self.driver.current_url
        # 物件名取得
        time.sleep(1)
        odr_name = self.get_html_attr('/html/body/form/table[1]/tbody/tr[4]/td', "textContent")[0]
        self.prog_odr.odr_name = odr_name.replace(r'Project Name：', '').strip()
        # 取得対象の資料分繰り返す
        odr_rows = self.driver.find_elements_by_xpath('/html/body/form/table[2]/tbody/tr')
        self.search_dwg_intable(odr_rows)
        self.driver.close()
        self.move_to_new_tab()

    @ErrHandle.err_handle(330)
    def search_dwg_intable(self, odr_rows):
        del odr_rows[0:3]
        for odr_row in odr_rows:
            self.exe_dl_intabe(odr_row)

    @ErrHandle.err_handle(330)
    def exe_dl_intabe(self, odr_row):
        if odr_row.find_elements_by_xpath('td[4]'):
            dwg_name_elem = odr_row.find_element_by_xpath('td[1]')
            dwg_num_elem = odr_row.find_element_by_xpath('td[4]')
            rev_elem = odr_row.find_element_by_xpath('td[7]')
            doc_type_elem = odr_row.find_element_by_xpath('td[6]')
            dwg_name = dwg_name_elem.text.strip()
            dwg_num = dwg_num_elem.text.strip()
            dwg_rev = rev_elem.text.strip()
            doc_type = doc_type_elem.text.strip()
            # if dwg_name == self.prog_odr.dwg_name:
            self.prog_odr.dwg_num = dwg_num
            for i in range(5):
                rev_elem.click()
                odr_file = OdrFile()
                odr_file.set_dl_site(self.site_name)
                odr_file.set_dwg_data(dwg_num, dwg_rev, dwg_name, self.prog_odr.odr_num, self.prog_odr.odr_name, doc_type)
                odr_file.mv_to_resultd()
                if os.path.exists(odr_file.result_fpath):
                    break
            if not os.path.exists(odr_file.result_fpath):
                raise Exception

class CrawlEnet(CrawlToTargetSite):
    def __init__(self):
        super(CrawlSendDwg, self).__init__()
        self.site_name = r'Eネット'

    def set_odr_dwg(self, odr_list, dwg_list):
        self.odr_list = odr_list
        self.dwg_list = dwg_list

    def set_root_url(self, url):
        self.url = url

    def get_target_site(self, url):
        self.set_root_url(url)
        self.set_config()
        self.execute_Edge()
        self.get_root_html()

    def login(self):
        pass

    def search_odr(self):
        pass

    def dl_application(self):
        self.search_odr()
        pass

    def get_odr_info(self):
        pass
        old_odr = ''
        rev_odrs = []
        return old_odr, rev_odrs


class CrawlMetis(CrawlToTargetSite):
    def __init__(self):
        super(CrawlSendDwg, self).__init__()
        self.site_name = r'図面伝送'

    def set_odr_info(self, odr_data):
        self.odr_num = odr_data.odr_num
        self.dwg_list = odr_data.dwg_list
        self.exm_list = odr_data.exm_list

    def set_root_url(self, url):
        self.url = url

    def get_target_site(self, url):
        self.set_root_url(url)
        self.set_config()
        self.execute_Edge()
        self.get_root_html()

    def search_exm_num(self, exm_num):
        # 自由入力値のため、区切り文字などの扱いに注意
        exm_splited = exm_num.splite('-')
        exm_num1 = exm_splited[1]
        exm_num2 = exm_splited[2]
        exm_num3 = exm_splited[3]
        exm_num4 = exm_splited[4]
        self.onclick_js('search_reset();')
        self.accept_alert()
        self.driver.refresh()
        self.down_iframe(1)
        self.down_frame(0)
        self.input_text_by_name('condition.inqNo2', exm_num2)
        self.input_text_by_name('condition.inqNo3', exm_num3)
        self.input_text_by_name('condition.inqNo4', exm_num4)
        self.input_text_by_name('condition.inqNo1', exm_num1)
        self.onclick_js('search()')

    def dl_metis_docs(self, odr_data):

        exm_list = odr_data.exm_num
        self.exm_list(exm_list)
        for exm_num in exm_list:
            self.search_exm_num(exm_num)


class CrawlDelisis(CrawlToTargetSite):
    def __init__(self):
        self.custom_path = r''
        self.dl_dir = r'C:\temp\test'
        self.url = 'http://www1.ina.melco.co.jp/delisis/Init.do'

    def login(self):
        pass

    def dl_delisis_docs(self, odr_data):
        pass


class CrawlOfficeStaff(CrawlToTargetSite):
    def __init__(self):
        self.custom_path = r''
        self.dl_dir = r'C:\temp\test'
        self.url = 'http://www.ina.melco.co.jp/OfficeSTAFF/GTRWeb/Login'


    def login(self):
        time.sleep(5)
        for i in range(5):
            if self.driver.current_url != self.url:
                # ★入力を促すため、ブラウザを全面に出す。
                # ★入力したID、PWをselfに格納
                pass

    def set_old_odrs(self, old_odr_list):
        self.old_odr_list = old_odr_list

    def search_old_odr(self, old_odr_num):
        self.click_element_by_xpath('//*[@id="ext-gen91"]/li[1]/ul/li[5]/div/img[1]')
        self.input_text_by_xpath('//*[@id="extCmbSearch"]', old_odr_num)
        self.click_element_by_xpath('//*[@id="extBtnSearch"]')
        self.down_iframe(0)
        # tdタグを検索して、品目表を含む文字列を取得
        titles = self.driver. find_elements_by_xpath()
        for title in titles:
            if '品目表' in title:
                self.click_element_by_id('ext-gen145')
                time.sleep(10)
                # ★DLファイルリネーム★
            if '摘要表' in title:
                self.click_element_by_id('ext-gen145')
                # ★DLファイルリネーム★
                time.sleep(10)

    def dl_old_itemtable(self, old_odr_list):
        self.get_root_html(self.url)
        self.login()
        self.set_old_odrs(old_odr_list)
        for old_odr in self.old_odr_list:
            self.search_old_odr(old_odr)


class Testsanshou:
    def __init__(self):
        pass

    def check(self, odr_data_list):
        for odr_data in odr_data_list:
            self.prog_odr = odr_data
            self.checkcheck()

    def checkcheck(self):
        self.prog_odr.odr_num = '23'


if __name__ == '__main__':
    pass
