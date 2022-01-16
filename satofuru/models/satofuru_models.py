import glob
import configparser
import os
import re
import pyautogui
import sys
from time import sleep
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from g_browser.g_browser import GoogleBrowser
from connect_spreadsheet import open_sp

class Satofuru(GoogleBrowser):
    def __init__(self, headless=False):
        super().__init__(headless)
        self.config_ini = configparser.ConfigParser()
        self.config_ini.read('config.ini', encoding='utf-8')

    def login(self):
        url = 'https://partner.satofull.jp/login'
        self.browser.get(url)
        id = self.config_ini['first_login']['id']
        password = self.config_ini['first_login']['password']
        
        pyautogui.click(972, 195)
        pyautogui.write(id)
        pyautogui.click(972, 215)
        pyautogui.write(password)
        sleep(1)
        pyautogui.click(972, 270)

        user = self.config_ini['second_login']['username']
        user_password = self.config_ini['second_login']['user_password']
        self.browser.find_element_by_id('username').send_keys(user)
        self.browser.find_element_by_id('password').send_keys(user_password)
        self.browser.find_element_by_id('btn-login').click()
        sleep(2)

    def download(self):
        url = 'https://partner.satofull.jp/gift/jigyousya/pickup-calendar'
        self.browser.get(url)
        sleep(2)
        self.browser.find_element_by_xpath('//label[@for="target_date_by-2"]').click()
        sleep(2)
        self.browser.find_element_by_xpath('//button[contains(text(),"検索する")]').click()
        sleep(2)
        self.browser.find_element_by_xpath('//tr[@class="text-center"]/th[contains(@class, "bg-dblue") and contains(@class, "brw")]/a').click()
        sleep(3)

    def get_values(self):
        path = glob.glob(r'C:\Users\ooaka\Downloads\for_date.csv')
        with open(path[0], 'r', newline='') as file:
            data = file.read().split('\n')

        data = ','.join(data)
        r_data = set(re.findall(r'[1-9],[0-9]', data))
        r_dict = {v:v.replace(',', '') for v in r_data}       
        for k, v in r_dict.items():
            data = data.replace(k, v)
        self.data = data.replace('"','').split(',')
    
    def sp_import(self):
        sp_key = self.config_ini['sp_key']['satofuru_sheet_key']
        wb = open_sp.open_sp(sp_key)
        import_sheet = wb.worksheet('import')
        row_num = len(import_sheet.col_values(1))
        range_list = import_sheet.range(f'A{row_num+1}:K{row_num+1 + int(len(self.data)/12)}')

        for i, range in enumerate(range_list):
            range.value = self.data[i]
        import_sheet.update_cells(range_list)

        
