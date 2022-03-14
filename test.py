from pathlib import Path
from unittest.case import DIFF_OMITTED

import numpy as np
import pandas as pd

from connect_spreadsheet import open_sp
import settings

def sample_def():
    sp_key = settings.CONFIG_INI['sp_key']['do_sheet_key']
    sp = open_sp.open_sp(sp_key)
    sheet = sp.worksheet('items')
    values = sheet.get_all_values()
    header = values.pop(0)
    df = pd.DataFrame(values, columns=header)
    df_filterd_columns = df.copy()
    df_filterd_columns.drop(columns=['SELSECT_SKU', '個数（マイナス可）', ''], inplace=True)
    target_sku_list = ['EA080', 'EA093', 'EA090']
    print(df_filterd_columns.query(f'{df_filterd_columns.columns[0]} in {target_sku_list} & 入庫 != ""')[['SKU', '商品名']])
    


if __name__ == '__main__':
    sample_def()