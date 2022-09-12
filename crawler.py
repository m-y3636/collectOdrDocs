import os, time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from msedge.selenium_tools import EdgeOptions, Edge
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.alert import Alert
import pyautogui


# クローラークラス
class CrawlToTargetSite(object):
    def __init__(self):
        # webdriver起動
        self.url = ''
        self.custom_path = r'.'
        self.dl_dir = r'C:\temp\test'
        self.options = EdgeOptions()
        self.options.use_chromium = True

    # Edge設定
    def set_config(self):
        self.set_options(
            self.custom_path,
            '--ignore-ssl-errors',
            '--allow-running-insecure-content',
            '--disable-gpu',
        )
        self.options.add_experimental_option('prefs', {'download.default_directory': self.dl_dir})
        self.desired_capabilities = None

    # 初期値に含む。再設定用の関数
    def execute_Edge(self):
        # edgeの実行
        self.driver_path = EdgeChromiumDriverManager(path=self.custom_path).install()
        self.driver = Edge(executable_path=self.driver_path, options=self.options, desired_capabilities=self.desired_capabilities)
        self.wait = WebDriverWait(self.driver, 10)

    def set_options(self, *args, **kwargs):
        for opt in args:
            self.options.add_argument(opt)

    # # 初期値に含む。再設定用の関数
    # def execute_chrome(self, driver_path, options, desired_capabilities=None):
    #     self.driver = webdriver.Chrome(executable_path=driver_path, options=options, desired_capabilities=desired_capabilities)
    #     self.wait = WebDriverWait(self.driver, 10)
    #
    # # 初期値に含む。再設定用の関数(Chrome)
    # def chrome_basic_options(self, custom_path, *args, **kwargs):
    #     driver_path = ChromeDriverManager(path=custom_path).install()
    #     options = Options()
    #     for opt in args:
    #         options.add_argument(opt)
    #     return driver_path, options

    def set_root_url(self, url):
        self.url = url

    def get_root_html(self):
        # targetsiteを開く
        self.driver.get(self.url)
        return self.driver.current_url

    def close_browser(self):
        self.driver.close()

    def quit_browser(self):
        self.driver.quit()

    # 待機設定
    def wait_for_all(self):
        self.wait.until(EC.presence_of_all_elements_located)

    def wait_for_id(self, elem):
        self.wait.until(EC.presence_of_element_located((By.ID, elem)))

    def wait_for_name(self, elem):
        self.wait.until(EC.presence_of_element_located((By.NAME, elem)))

    def wait_for_xpath(self, elem):
        self.wait.until(EC.presence_of_element_located((By.XPATH, elem)))

    # 自動操作
    def get_html_attr(self, target_path, attr):
        link_elements = self.driver.find_elements_by_xpath(target_path)
        # 取得したURLを抜き出してリスト化
        link_attrs = []
        for elem in link_elements:
            link_attrs.append(elem.get_attribute(attr))
        return link_attrs

    def get_html_tag(self, tag):
        # targetsiteを開く
        tag_name_elements = self.driver.find_elements_by_tag_name(tag)
        tag_name_attrs = []
        for elem in tag_name_elements:
            tag_name_attrs.append(elem.get_attribute('textContent'))
        return tag_name_attrs

    def get_html_title(self):
        # targetsiteを開く
        tag_name_elements = self.driver.find_elements_by_tag_name('title')
        tag_name_attrs = []
        for elem in tag_name_elements:
            tag_name_attrs.append(elem.text)
        return tag_name_attrs

    def onclick_js(self, onclick_tag):
        script = f'javascript:{onclick_tag}'
        self.driver.execute_script(script)

    def click_element_by_id(self, target_id):
        self.wait_for_id(target_id)
        id_name_element = self.driver.find_element_by_id(target_id)
        id_name_element.click()
        return 0

    def input_text_by_id(self, target_id, input_text):
        self.wait_for_id(target_id)
        id_name_element = self.driver.find_element_by_id(target_id)
        id_name_element.send_keys(input_text)
        return 0

    def click_ele_by_text_of_tag(self, text, target):
        self.wait_for_all()
        xpath_element = self.driver.find_elements_by_tag_name(target)
        for ele in xpath_element:
            if ele.text == text:
                ele.click()
                break

    def click_ele_by_text_of_xpath(self, text, target):
        self.wait_for_all()
        xpath_element = self.driver.find_elements_by_xpath(target)
        for ele in xpath_element:
            if ele.text == text:
                ele.click()
                break

    def click_ele_by_text_of_id(self, text, target):
        self.wait_for_id(target)
        xpath_element = self.driver.find_element_by_xpath(target)
        for ele in xpath_element:
            if ele.text == text:
                ele.click()

    def click_element_by_xpath(self, target_xpath):
        self.wait_for_xpath(target_xpath)
        xpath_element = self.driver.find_element_by_xpath(target_xpath)
        xpath_element.click()
        time.sleep(1)
        return 0

    def input_text_by_xpath(self, target_xpath, input_text):
        self.wait_for_xpath(target_xpath)
        xpath_element = self.driver.find_element_by_xpath(target_xpath)
        xpath_element.send_keys(input_text)
        return 0

    def input_text_by_name(self, target_name, input_text):
        self.wait_for_name(target_name)
        xpath_element = self.driver.find_element_by_name(target_name)
        xpath_element.send_keys(input_text)
        return 0

    def text_clear_by_id(self, target_id):
        self.wait_for_id(target_id)
        id_name_element = self.driver.find_element_by_id(target_id)
        id_name_element.clear()
        return 0

    def text_clear_by_name(self, target_name):
        self.wait_for_name(target_name)
        name_element = self.driver.find_element_by_name(target_name)
        name_element.clear()
        return 0

    def text_clear_by_xpath(self, target_xpath):
        self.wait_for_xpath(target_xpath)
        xpath_element = self.driver.find_element_by_xpath(target_xpath)
        xpath_element.clear()
        return 0

    def text_clear_by_name(self, target_name):
        self.wait_for_name(target_name)
        id_element = self.driver.find_element_by_name(target_name)
        id_element.clear()
        return 0

    # プルダウン操作 ★
    def pulldown_chg_by_xpath(self, target_xpath, target_value):
        WebDriverWait(self.driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, target_xpath)))
        pulldown = self.driver.find_element_by_xpath(target_xpath)
        time.sleep(10)
        Select(pulldown).select_by_visible_text(target_value)

    def pulldown_chg_by_name(self, target_name, target_value):
        WebDriverWait(self.driver, 60).until(
            EC.element_to_be_clickable((By.NAME, target_name)))
        pulldown = self.driver.find_element_by_name(target_name)
        Select(pulldown).select_by_index(2)

    # チェックボックス操作
    def check_on_by_xpath(self, target_xpath):
        self.wait_for_xpath(target_xpath)
        status = self.driver.find_element_by_xpath(target_xpath)
        if not status.is_selected():
            status.click()
        return 0

    def check_on_by_name(self, target):
        self.wait_for_name(target)
        status = self.driver.find_element_by_name(target)
        if not status.is_selected():
            status.click()
        return 0

    def check_off_by_xpath(self, target_xpath):
        self.wait_for_xpath(target_xpath)
        status = self.driver.find_element_by_xpath(target_xpath)
        if status.is_selected():
            status.click()
        return 0

    def check_off_by_name(self, target):
        self.wait_for_name(target)
        status = self.driver.find_element_by_name(target)
        if status.is_selected():
            status.click()
        return 0

    # 警告クリック
    def accept_alert(self):
        Alert(self.driver).accept()
        time.sleep(1)

    # frameを下がる
    def down_frame(self, i):
        self.wait_for_xpath('//frame')
        frame = self.driver.find_elements_by_xpath('//frame')[i]
        self.driver.switch_to.frame(frame)

    def down_iframe(self, i):
        self.wait_for_xpath('//iframe')
        frame = self.driver.find_elements_by_xpath('//iframe')[i]
        self.driver.switch_to.frame(frame)

    def pulldown_reset(self, name):
        self.wait_for_name(name)
        pulldown = self.driver.find_elements_by_name(name)
        select = Select(pulldown[0])
        select.deselect_all()

    def pulldown_select_js(self, target_name, val):
        script1 = f'document.querySelector("body > form > div > table > tbody > tr > td:nth-child(1) > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > select:nth-child(1) > option:nth-child(3)").click()'
        self.driver.execute_script(script1)

    def move_to_new_tab(self):
        self.driver.switch_to.window(self.driver.window_handles[-1])


    # キーコントロール
    def send_a_key(self, target_x, key):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(key)

    def send_two_keys(self, target_x, key1, key2):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(key1, key2)

    def send_key_repetition(self, target_x, rep_num, key1):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        for i in range(rep_num):
            key_action.send_keys(key1)

    def send_key_up(self, target_x):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(Keys.UP)

    def send_key_down(self, target_x):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(Keys.DOWN)

    def send_keys_up(self, target_x):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(Keys.UP)

    def send_key_textalldel(self, target_x):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(Keys.CONTROL, 'a')
        key_action.send_keys(Keys.DELETE)

    def send_key_del(self, target_x):
        self.wait_for_xpath(target_x)
        key_action = self.driver.find_element_by_xpath(target_x)
        key_action.send_keys(Keys.DELETE)


class ctrl_key_selen:
    def __init__(self):
        self.up = Keys.UP
        self.down = Keys.DOWN
        self.ctl = Keys.CONTROL
        self.shift = Keys.SHIFT
        self.enter_key = Keys.ENTER
        self.f3 = Keys.F3
        self.f8 = Keys.F8
        self.f9 = Keys.F9



class testcrawl(CrawlToTargetSite):
    def __init__(self):
        custom_path = r'.'
        # Edge設定
        self.driver_path, self.options = self.edge_basic_options(custom_path, '--disable-gpu', '--kiosk-printing')
        self.desired_capabilities = None
        self.execute_Edge()


if __name__ == '__main__':
     test = testcrawl()
     # test.get_root_html('http://www.ina.melco.co.jp/OfficeSTAFF/GTRWeb/Login')
     # test.input_text_by_id('userid', 'LH21475')
     # test.input_text_by_id('passwd', '@Yuuchan362422')
     # test.click_element_by_id('btnlogin')
     # test.click_element_by_xpath('//*[@id="ext-gen91"]/li[1]/ul/li[5]/div/img[1]')
     # test.input_text_by_xpath('//*[@id="extCmbSearch"]', 'MV15450')
     # test.click_element_by_xpath('//*[@id="extBtnSearch"]')
     # test.down_iframe(0)
     # test.click_element_by_xpath('//*[@id="ext-gen50"]/div[1]/table/tbody/tr/td[7]/div/a')
     # # tdタグを検索して、品目表を含む文字列を取得
     # # リンクの数を数える。⇒繰り返し
     # # 同じ行のリンクをクリック（for）
     # handle_array = test.driver.window_handles
     # test.driver.switch_to.window(handle_array[-1])
     # test.click_element_by_xpath('//*[@id="ext-gen114"]')
     # # ポップアップウィンドウを閉じる
     # # ウィンドウ移動する。
     # test.quit_browser()
     # time.sleep(10)
     # # DLデータを名称毎に整理

     test.get_root_html('http://www1.ina.melco.co.jp/delisis/Init.do')
     test.input_text_by_name('username', 'teramoty')
     test.input_text_by_name('password', '&yuchan3')
     test.click_element_by_xpath('/html/body/center/form/input[1]')
     test.click_element_by_xpath('/html/body/table[3]/tbody/tr[1]/td[1]/li/a')
     test.input_text_by_xpath('/html/body/table[2]/tbody/tr/th/form/table/tbody/tr[2]/td[3]/input', 'M2RR6713')
     test.input_text_by_xpath('/html/body/table[2]/tbody/tr/th/form/table/tbody/tr[2]/td[5]/select', '電気')
     test.check_on_by_xpath('/html/body/table[2]/tbody/tr/th/form/table/tbody/tr[2]/td[15]/input[1]')
     test.click_element_by_xpath('/html/body/table[2]/tbody/tr/th/form/table/tbody/tr[2]/td[17]/input')
     test.move_to_new_tab()
     time.sleep(10)
     test.send_two_keys('/html/body/embed', Keys.CONTROL, 's')
     time.sleep(1)
     pyautogui.hotkey('ctrl', 's')
     time.sleep(1)
     for i in range(6):
         time.sleep(1)
         pyautogui.press('tab')
     pyautogui.press('enter')
     pyautogui.write(r'C:\Users\LH21475\Desktop')
     time.sleep(1)
     for i in range(3):
         pyautogui.hotkey('shift', 'tab')
         time.sleep(1)
     time.sleep(1)
     pyautogui.press('enter')
     time.sleep(1)
     test.quit_browser()
     time.sleep(1)
     time.sleep(10)
