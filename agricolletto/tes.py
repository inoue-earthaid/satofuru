# 必要なPdfminer.sixモジュールのクラスをインポート
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams


class Agricolletto(object):
    def read_pdf(self):
        with open(r'agricolletto\tes.txt', 'w',encoding='utf-8') as outfp:
            with open(r"agricolletto\委託業者日次実績表1075.PDF", 'rb') as fp:
                # 各種テキスト抽出に必要なPdfminer.sixのオブジェクトを取得する処理
                rmgr = PDFResourceManager() # PDFResourceManagerオブジェクトの取得
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
            ]
        path = r'agricolletto\tes.txt'
        with open(path, 'r', encoding='utf-8') as f:
            pdf_text = f.readlines()
        pdf_text = [chr.replace('\n', '') for chr in pdf_text if chr not in delete_char]
        
        pop_list = [9, 10, 11]
        for i in sorted(pop_list, reverse=True):
            pdf_text.pop(i)
        
        for i, chr in enumerate(pdf_text):
            if chr == '商品名':
                count = i
            if chr == '金額':
                insert_i = i
                insert_chr = pdf_text.pop(i)
        pdf_text.insert(insert_i+count, insert_chr)
        index_qty = pdf_text.index('数量')
        index_amount = pdf_text.index('金額')
        total_qty = pdf_text.pop(index_qty + count)
        total_amount = pdf_text.pop(index_amount + count - 1)

        wrapper = []
        for _ in range(count):
            new_list = []
            for i, v in enumerate(pdf_text):
                if i % count == _:
                    new_list.append(v)
            wrapper.append(new_list)
        
        head_lines = wrapper.pop(0)
        wrapper_words = []
        for inner in wrapper:
            date = f'{head_lines[0]} : {inner[0]}'
            item = f'{head_lines[1]}　 : {inner[1]}'
            qty = f'{head_lines[2]}　　 : {inner[2]}'
            amount = f'{head_lines[3]}　　 : {inner[3]}'
            words = [date, item, qty, amount]
            r = '\n'.join(words)
            wrapper_words.append(r)
        
        head_line = '■ アグリコレット売上'
        delimiter = '\n' + '-' * 20 + '\n'
        result = f'{head_line}{delimiter}' + f'{delimiter}'.join(wrapper_words)
        return result


if __name__ == '__main__':
    agri = Agricolletto()
    agri.read_pdf()
    result = agri.shaping_pdf()
    print(result)