import os
from selenium.webdriver.chrome.options import Options
from selenium import webdriver

file_path = os.path.dirname(os.path.dirname(__file__))
file_path = os.path.join(file_path, r'g_browser\chromedriver.exe')

class GoogleBrowser(object):
    def __init__(self, headless=False):
        self.options = Options()
        if headless:
            self.options.add_argument("--headless")
        self.options.add_argument('--window-size=1920,1080') 
        self.chromedriver = file_path
        self.browser = webdriver.Chrome(self.chromedriver, options=self.options)
        self.browser.implicitly_wait(2)

    def quit(self):
        self.browser.quit()

