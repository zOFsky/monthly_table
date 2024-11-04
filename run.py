import pandas as pd

# Read the Excel file, specifying the sheet name
df = pd.read_excel('sep_new_corr.xlsx', sheet_name='таблиця')
df.columns.values[5] = '0'

def fill_empty_cells_with_previous_value(df, N=30):
    for col in range(1, N + 1):
        col2 = str(col)
        for idx, row in df.iterrows():
            # Check if the current column has NaN or empty value and the previous column has a valid string
            if pd.isna(row[col2]) or row[col2] == '':
                previous_value = row[str(int(col2) - 1)]
                if isinstance(previous_value, str) and not previous_value.endswith('2') and previous_value.strip():  # Check if the previous value is a non-empty string
                    df.at[idx, col2] = previous_value + '2'  # Fill the cell with the previous value + '2'
    return df

df2 = fill_empty_cells_with_previous_value(df)
df3 = df2.applymap(lambda x: 1 if pd.notna(x) and x != '' else x)
df3.iloc[:, :5] = df2.iloc[:, :5]

df2.to_excel('sep_new_prolonged.xlsx', index=False)
df3.to_excel('sep_new_prolonged_1.xlsx', index=False)