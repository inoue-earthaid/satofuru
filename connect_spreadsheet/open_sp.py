import configparser
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import sys

SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
path = os.path.dirname(__file__)

config_ini = configparser.ConfigParser()
config_ini.read('config.ini', encoding='utf-8')
sp_json = config_ini['sp_api']['project_key']
SERVICE_ACCOUNT_FILE = os.path.join(path, sp_json)
credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)


def open_sp(sp_key):
    gs = gspread.authorize(credentials)
    spreadsheet = gs.open_by_key(sp_key)
    return spreadsheet