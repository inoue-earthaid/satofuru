from connect_spreadsheet import open_sp
from agricolletto.files.secrets import auth
import re
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFResourceManager
import itertools
import configparser
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
CONFIG_INI =configparser.ConfigParser()
CONFIG_INI.read('config.ini', encoding='utf-8')

class Agricolletto(object):
    def __init__(self) -> None:
        CONFIG_INI = configparser.ConfigParser()
        CONFIG_INI.read('config.ini', encoding='utf-8')
        self.base_dir = r'C:\Users\ooaka\Downloads'
        self.service = auth.connect_gmail()
        self.download_path = ''
        self.pdf_txt_path = ''

    def download_file(self):
        folder_id = CONFIG_INI['g_drive']['download_order_google_drive_id']
        files = self.service.files().list(pageSize=10, fields="nextPageToken, files(id, name)", q=f'"{folder_id}" in parents').execute()
        files = files.get('files', [])
        return files

    def convert_pdf(self):
        self.pdf_txt_path = os.path.join(self.base_dir, 'pdf_text.txt')
        with open(self.pdf_txt_path, 'w', encoding='utf-8') as outfp:
            with open(self.download_path, 'rb') as fp:
                # 各種テキスト抽出に必要なPdfminer.sixのオブジェクトを取得する処理
                rmgr = PDFResourceManager()  # PDFResourceManagerオブジェクトの取得
                lprms = LAParams(
                    detect_vertical=False
                )
                with TextConverter(rmgr, outfp, laparams=lprms) as conv:
                    iprtr = PDFPageInterpreter(rmgr, conv)

                    for page in PDFPage.get_pages(fp):
                        iprtr.process_page(page)

    def shaping_pdf(self):
        delete_char = [
            '委託業者日次実績表\n',
            '1075 株式会社アースエイド\n',
            '仕入先：\n',
            '日付：\n',
            '時刻：\n',
            'ﾍﾟｰｼﾞ：\n',
            '仕入先計\n',
            '\n',
            '\x0c',
            '',
        ]
        head_line_list = [
            '売上日付\n',
            '商品名\n',
            '数量\n',
            '金額\n'
        ]

        with open(self.pdf_txt_path, 'r', encoding='utf-8') as f:
            pdf_text = f.readlines()
        pdf_text = [re.sub(r'[\n　]', '', chr)
                    for chr in pdf_text if chr not in delete_char + head_line_list]

        pattern = r'20[2-9]{2}/.*$'
        for i, chr in enumerate(pdf_text):
            if not re.match(pattern, chr):
                count = i
                pdf_text.pop(count*2+2)
                pdf_text.pop(count*2+1)
                pdf_text.pop(count*2)
                break
        for _ in range(count+2):
            pdf_text.pop(-1)
        self.wrapper = []
        for _ in range(count):
            new_list = []
            for i, v in enumerate(pdf_text):
                if i % count == _:
                    new_list.append(v)
            self.wrapper.append(new_list)
        os.remove(self.pdf_txt_path)

    def reflect_spreadsheet(self):
        sheet_key = CONFIG_INI['sp_key']['agricolletto_sheet_key']
        wb = open_sp.open_sp(sheet_key)
        sheet = wb.worksheet('売上')
        data_count = len(self.wrapper)
        last_row = len(sheet.col_values(2))
        reflection_range = sheet.range(
            f'B{last_row + 1}:D{last_row + data_count}')
        value_list = list(itertools.chain.from_iterable(self.wrapper))
        for i, range in enumerate(reflection_range):
            range.value = value_list[i]
        sheet.update_cells(reflection_range)
