import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_excel('ProdMAPP.xlsx')

print("Column Headings:")
print(df.columns)
