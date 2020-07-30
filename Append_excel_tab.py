import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os

## append excel tab without change the previous one
# script is in another folder
# test excel is aaa.xlsx, under the Archive File folder

path = os.path.abspath(os.path.join(os.getcwd(),'..'))
path=path + '/Archive File/aaa.xlsx'

book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

x3 = np.random.randn(100, 2)
df3 = pd.DataFrame(x3)

x4 = np.random.randn(100, 2)
df4 = pd.DataFrame(x4)

df3.to_excel(writer, sheet_name ='x3.xlsx')
df4.to_excel(writer, sheet_name ='x4.xlsx')
writer.save()
writer.close()
