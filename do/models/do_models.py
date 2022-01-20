import configparser
import datetime
from selenium.webdriver.support.select import Select
from time import sleep

from g_browser.g_browser import GoogleBrowser

class Do(GoogleBrowser):
    def __init__(self, headless=False):
        super().__init__(headless)

    def login(self):
        url = 'https://do3biz.do-furusato.com/deliveries'
        self.browser.get(url)
        config_ini = configparser.ConfigParser()
        config_ini.read('config.ini', encoding='utf-8')
        id = config_ini['do_login']['username']
        self.browser.find_element_by_id('userName').send_keys(id)
        sleep(1)
        password = config_ini['do_login']['password']
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
        self.browser.find_element_by_xpath('//li[contains(@class, "active-result") and contains(text(), "2022")]').click()
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
            # self.browser.find_element_by_xpath('//tr[@class="even"]')
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

    def export_csv(self):

        # 指定日
        today = datetime.date.today()
        youbi = datetime.datetime.weekday(today)
        delta = 9-youbi
        designated_day = today + datetime.timedelta(delta)
        month = designated_day.strftime('%m月')
        if month[0:1] == '0':
            month = month.replace('0', '')
        self.browser.find_element_by_name('delivery_info[specified_delivery_date_to]').click()
        dropdown = self.browser.find_element_by_class_name('ui-datepicker-month')
        select = Select(dropdown)
        select.select_by_visible_text(month)
        date = designated_day.strftime('%d')
        if int(date) >= 26:
            self.browser.find_elements_by_link_text(date)[1].click()
        else:
            self.browser.find_element_by_link_text(date).click()
        
