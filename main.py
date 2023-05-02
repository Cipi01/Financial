import pandas as pd
import datetime
from TextFilter import pdf_extractor, table_modifier, dict_maker

table = pdf_extractor('SV08870913300_2023_03.pdf')
modified_table = table_modifier(table)
transactions = dict_maker(modified_table)

# Loop through the transactions and modify the 'Details' column as required
for i in range(1, len(transactions)):
    if transactions[i]["Details"].startswith("Card nr"):
        # Split the cell until "Card/Terminale" substring is encountered
        split_text = transactions[i]["Details"].split("Card/Terminale")[0]
        remaining_text = transactions[i]["Details"].replace(split_text, "").strip()
        # Append the split text to the cell above it
        transactions[i-1]["Details"] += split_text.strip()
        # Update the current cell with the remaining text
        transactions[i]["Details"] = remaining_text

# Convert the transactions to a DataFrame and write to an Excel file
df = pd.DataFrame(transactions)
df["Details"] = df["Details"].str.strip()
df.to_excel(f'C:/Users/cipri/OneDrive/Desktop/Financial/{datetime.datetime.now().strftime("%m_%y")}.xlsx', index=False)

print('transactions')
