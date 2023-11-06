"""Script to sort 3 columns (D, E, F) from a spreadsheet in ascending order
and save as 3 separate files.

    Input: 'Top_700_Coins_new.xlsx'

    Output:
        'Top700_1hr.xlsx
        Top700_24hr.xlsx
        Top700_7d.xlsx'

"""

import pandas as pd

# Read the Excel file
df = pd.read_excel('Top_700_coins_new.xlsx')

# Sort by column D and write to a new Excel file
df_sorted_by_d = df.sort_values(by=df.columns[3])
df_sorted_by_d.to_excel('Top700_1hr.xlsx', index=False)

# Sort by column E and write to a new Excel file
df_sorted_by_e = df.sort_values(by=df.columns[4])
df_sorted_by_e.to_excel('Top700_24hr.xlsx', index=False)

# Sort by column F and write to a new Excel file
df_sorted_by_f = df.sort_values(by=df.columns[5])
df_sorted_by_f.to_excel('Top700_7d.xlsx', index=False)