import configparser
import csv
import datetime
import glob
import itertools
import os
import re
from selenium.webdriver.support.select import Select
from time import sleep

from g_browser.g_browser import GoogleBrowser
from connect_spreadsheet import open_sp

CONFIG_INI = configparser.ConfigParser()
CONFIG_INI.read('config.ini', encoding='utf-8')

class Do(GoogleBrowser):
    def __init__(self, headless=False):
        super().__init__(headless)
        self.sp = ''
        self.dl_path = r'C:\Users\ooaka\Downloads'

    def login(self):
        url = 'https://do3biz.do-furusato.com/deliveries'
        self.browser.get(url)

        id = CONFIG_INI['do_login']['username']
        self.browser.find_element_by_id('userName').send_keys(id)
        sleep(1)
        password = CONFIG_INI['do_login']['password']
        self.browser.find_element_by_name('password').send_keys(password)
        sleep(1)
        self.browser.find_element_by_id('loginBtn1').click()

    def search_status_request(self):
        url = 'https://do3biz.do-furusato.com/deliveries/request-shipping'
        self.browser.get(url)
        self.browser.find_element_by_css_selector("#delivery_status_chosen .search-choice-close").click()
        self.browser.find_element_by_css_selector("#delivery_status_chosen .chosen-search-input").click()
        self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "出荷依頼")]').click()
        self.browser.find_element_by_css_selector('#subscribe_year_chosen .search-field').click()
        try:
            self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "2021")]').click()
        except:
            pass
        try:
            self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "2022")]').click()
        except:
            pass
        # 定期便
        self.browser.find_element_by_xpath('//th[contains(text(), "定期便")]/following-sibling::td/span/span[@class="selection"]').click()
        self.browser.find_element_by_xpath('//ul[@class="select2-results__options"]/li[contains(text(), "定期便以外")]').click()
        # アースエイド
        self.browser.find_element_by_xpath('//th[contains(text(), "集荷先")]/following-sibling::td/div/ul/li/input').click()
        self.browser.find_element_by_xpath('//th[contains(text(), "集荷先")]/following-sibling::td/div/div/ul/li[contains(text(), "アースエイド")]').click()
        
        self.browser.find_element_by_id("searchBtn").click()

    def change_request_to_receipt(self):
        try:
            self.browser.find_element_by_xpath('//tr[@class="odd"]/td[contains(text(), "検索条件にマッチするデータがありません。")]')
            print('Not-order')
            return True
        except:
            pass
        self.browser.find_element_by_xpath('//input[contains(@type, "checkbox") and contains(@class, "checkAll") and contains(@class, "widthCheckbox")]').click()
        sleep(1)
        self.browser.find_element_by_id('configureBtn').click()
        sleep(1)
        self.browser.switch_to.frame(0)
        self.browser.find_element_by_css_selector("tr:nth-child(4) label").click()
        self.browser.find_element_by_css_selector("#configure_delivery_status_chosen span").click()
        self.browser.find_element_by_xpath('//li[contains(text(), "出荷依頼受領")]').click()
        self.browser.find_element_by_id('exportBtn').click()

    def filtering_export_csv(self):
        url = 'https://do3biz.do-furusato.com/deliveries/request-shipping'
        self.browser.get(url)
        self.browser.find_element_by_css_selector("#delivery_status_chosen .search-choice-close").click()
        self.browser.find_element_by_css_selector("#delivery_status_chosen .chosen-search-input").click()
        self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "出荷依頼受領")]').click()
        self.browser.find_element_by_css_selector('#subscribe_year_chosen .search-field').click()
        try:
            self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "2021")]').click()
        except:
            pass
        try:
            self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "2022")]').click()
        except:
            pass
        # 定期便
        self.browser.find_element_by_xpath('//th[contains(text(), "定期便")]/following-sibling::td/span/span[@class="selection"]').click()
        self.browser.find_element_by_xpath('//ul[@class="select2-results__options"]/li[contains(text(), "定期便以外")]').click()
        # アースエイド
        self.browser.find_element_by_xpath('//th[contains(text(), "集荷先")]/following-sibling::td/div/ul/li/input').click()
        self.browser.find_element_by_xpath('//th[contains(text(), "集荷先")]/following-sibling::td/div/div/ul/li[contains(text(), "アースエイド")]').click()
        today = datetime.date.today()
        delta = 30
        designated_day = today + datetime.timedelta(delta)
        month = designated_day.strftime('%m月')
        if month[0:1] == '0':
            month = month.replace('0', '')
        self.browser.find_element_by_name('delivery_info[delivery_date_to]').click()
        dropdown = self.browser.find_element_by_class_name('ui-datepicker-month')
        select = Select(dropdown)
        select.select_by_visible_text(month)
        date = designated_day.strftime('%d')
        if int(date) >= 26:
            self.browser.find_elements_by_link_text(date)[1].click()
        else:
            self.browser.find_element_by_link_text(date).click()
        
        self.browser.find_element_by_id("searchBtn").click()

    def export_csv(self):
        try:
            self.browser.find_element_by_xpath('//tr[@class="odd"]/td[contains(text(), "検索条件にマッチするデータがありません。")]')
            print('Not-order')
            return
        except:
            pass
        self.browser.find_element_by_xpath('//input[contains(@type, "checkbox") and contains(@class, "checkAll") and contains(@class, "widthCheckbox")]').click()
        sleep(1)
        self.browser.find_element_by_id('exportBtn').click()
        sleep(1)
        self.browser.switch_to.frame(0)
        self.browser.find_element_by_id('is_export_type_delivery_instruction').click()
        self.browser.find_element_by_css_selector("#export_type_delivery_instruction_chosen span").click()
        self.browser.find_element_by_xpath('//li[contains(text(), "ヤマトB2")]').click()
        self.browser.find_element_by_id('exportBtn').click()

    def import_csv(self):
        config_ini = configparser.ConfigParser()
        config_ini.read('config.ini')
        sheet_key = config_ini['sp_key']['do_sheet_key']
        self.sp = open_sp.open_sp(sp_key=sheet_key)
        import_sheet = self.sp.worksheet('import')
        last_row = len(import_sheet.col_values(1))
        path_choice = glob.glob(os.path.join(self.dl_path, r'yamato2022[0-9]*.csv'))

        for i, path_candidate in enumerate(path_choice):
            if i:
                return 'files.length >= 1'
            is_regex_match = bool(re.match(r'^yamato2022[0-9]{10}.csv', os.path.basename(path_candidate)))
        if is_regex_match:
            path = path_candidate
        
        with open(path, 'r') as csv_file:
            import_list = list(itertools.chain.from_iterable(list(csv.reader(csv_file))))
            row_count = round((len(import_list)) / 50)
        import_range = import_sheet.range(f'A{last_row + 1}:AY{last_row + row_count}')
        
        for i, cell in enumerate(import_range):
            cell.value = import_list[i]
        import_sheet.update_cells(import_range)
        os.remove(path)

    def export_yamato_csv(self):
        yamato_export_sheet = self.sp.worksheet('export')
        done_list = yamato_export_sheet.col_values(1)
        done_list.pop(0)
        export_values = yamato_export_sheet.get_all_values()
        now = datetime.datetime.now()
        date = now.strftime('%Y%m%d%H%M')
        export_path = os.path.join(self.dl_path, f'do_yamato_export_{date}.csv')
        with open(export_path, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(export_values)
        self._move_done_sheet(done_list)
        return export_path

    def _move_done_sheet(self, done_list):
        done_sheet = self.sp.worksheet('done')
        last_row = len(done_sheet.col_values(1)) + 1
        range_list = done_sheet.range(f'A{last_row}:A{last_row + len(done_list) -1}')
        for i, cell in enumerate(range_list):
            cell.value = done_list[i]
        done_sheet.update_cells(range_list)


class ImportYamato(GoogleBrowser):
    def __init__(self, headless=False):
        super().__init__(headless)

    def login(self):
        url = 'https://bmypage.kuronekoyamato.co.jp/bmypage/servlet/jp.co.kuronekoyamato.wur.hmp.servlet.user.HMPLGI0010JspServlet'
        self.browser.get(url)

        code1 = CONFIG_INI['yamato_login']['id']
        password = CONFIG_INI['yamato_login']['pass']

        self.browser.find_element_by_id('code1').send_keys(code1)
        self.browser.find_element_by_id('password').send_keys(password)
        self.browser.find_element_by_class_name('nav-login-btn').click()
        sleep(1)
        self.browser.find_element_by_link_text('送り状発行システムB2クラウド').click()
        sleep(1.5)

    def move_upload_page(self):
        issued_date_url = 'https://newb2web.kuronekoyamato.co.jp/ex_data_import.html'
        self.browser.get(issued_date_url)
        sleep(.5)

        from selenium.webdriver.support.select import Select
        doropdown = self.browser.find_element_by_id('torikomi_pattern')
        select = Select(doropdown)
        select.select_by_value('1')