import xlsxwriter
import os
import numpy as np
from Txt_to_xlsx_programs import paste1
def copy(book,Sheet,path1):
    print(Sheet)
    ws = book.get_worksheet_by_name(Sheet)
    paste1.paste(ws,book,Sheet,path1)
