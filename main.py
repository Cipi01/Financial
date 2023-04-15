import pandas as pd
from TextFilter import pdf_extractor, table_modifier, dict_maker

table = pdf_extractor('SV08870913300_2023_03.pdf')
modified_table = table_modifier(table)
transactions = dict_maker(modified_table)

df = pd.DataFrame(transactions)
df.to_excel('transactions.xlsx', index=False)

print('transactions')
