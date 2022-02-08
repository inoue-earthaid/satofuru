import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

import configparser
import itertools
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams
import re
from pydrive.drive import GoogleDrive
from files import auth

from connect_spreadsheet import open_sp


class Agricolletto(object):
    def __init__(self) -> None:
        self.config_ini = configparser.ConfigParser()
        self.config_ini.read('config.ini', encoding='utf-8')
        self.dst_path = r'C:\Users\ooaka\Downloads'
        self.download_path = ''
        self.pdf_txt_path = ''

    def download_file(self):
        folder_id = self.config_ini['g_drive']['download_order_google_drive_id']
        gauth = auth.get_googleauth()
        gauth.LocalWebserverAuth()
        drive = GoogleDrive(gauth)
        files = drive.ListFile({'q': '"{}" in parents'.format(folder_id)}).GetList()
        for file in files:
            self.download_path = os.path.join(self.dst_path, file['title'])
            file.GetContentFile(self.download_path)

    def convert_pdf(self):
        self.pdf_txt_path = os.path.join(self.dst_path, 'pdf_text.txt') 
        with open(self.pdf_txt_path, 'w',encoding='utf-8') as outfp:
            with open(self.download_path, 'rb') as fp:
                # 各種テキスト抽出に必要なPdfminer.sixのオブジェクトを取得する処理
                rmgr = PDFResourceManager() # PDFResourceManagerオブジェクトの取得
                lprms = LAParams(
                    detect_vertical=False
                )
                with TextConverter(rmgr, outfp, laparams=lprms) as conv:
                    iprtr = PDFPageInterpreter(rmgr, conv)

                    for page in PDFPage.get_pages(fp):
                        iprtr.process_page(page)
        os.remove(self.download_path)

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
            ]
        head_line_list = [
            '売上日付\n',
            '商品名\n',
            '数量\n',
            '金額\n'
        ]

        with open(self.pdf_txt_path, 'r', encoding='utf-8') as f:
            pdf_text = f.readlines()
        pdf_text = [re.sub(r'[\n　]', '', chr) for chr in pdf_text if chr not in delete_char + head_line_list] 
        
        pattern = r'20[2-9]{2}/.*$'
        for i, chr in enumerate(pdf_text):
            if not re.match(pattern, chr):
                count = i
                pdf_text.pop(count*2+2)
                pdf_text.pop(count*2+1)
                pdf_text.pop(count*2)
                break
        
        total_amount =pdf_text.pop(-1)
        total_qty = pdf_text.pop(-(count+1))
        for _ in range(count):
            pdf_text.pop(-1)
        self.wrapper = []
        for _ in range(count):
            new_list = []
            for i, v in enumerate(pdf_text):
                if i % count == _:
                    new_list.append(v)
            self.wrapper.append(new_list)
        os.remove(self.pdf_txt_path)

        # wrapper_words = []
        # head_line_list = [head_line.replace('\n', '') for head_line in head_line_list]
        # for inner in self.wrapper:
        #     date = f'{head_line_list[0]} : {inner[0]}'
        #     item = f'{head_line_list[1]}　 : {inner[1]}'
        #     qty = f'{head_line_list[2]}　　 : {inner[2]}'
        #     words = [date, item, qty]
        #     r = '\n'.join(words)
        #     wrapper_words.append(r)
        
        # head_line = '■ アグリコレット売上'
        # delimiter = '\n' + '-' * 20 + '\n'
        # result = f'{head_line}{delimiter}' + f'{delimiter}'.join(wrapper_words)
        # return result

    def reflect_spreadsheet(self):
        sheet_key = self.config_ini['sp_key']['agricolletto_sheet_key']
        wb = open_sp.open_sp(sheet_key)
        sheet = wb.worksheet('売上')
        data_count = len(self.wrapper)
        last_row = len(sheet.col_values(2))
        reflection_range = sheet.range(f'B{last_row + 1}:D{last_row + data_count}')
        value_list = list(itertools.chain.from_iterable(self.wrapper))
        for i, range in enumerate(reflection_range):
            range.value = value_list[i]
        sheet.update_cells(reflection_range)
