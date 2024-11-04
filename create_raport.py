import pandas as pd
from functions import squash_intervals
import ast

# Define the squash_intervals function to compress intervals

# Read the Excel file
df = pd.read_excel("october_edited_colored.xlsx", dtype=str)

# Set the number of days in the month
days_in_month = 31

# Step 1: Replace all "2" values with values from the previous column
for i in range(1, days_in_month + 1):
    prev_col = str(i - 1)
    curr_col = str(i)
    df[curr_col] = df.apply(lambda row: row[prev_col] if row[curr_col] == "2" else row[curr_col], axis=1)

# Step 2 & 3: Create dictionaries for each row and apply squash_intervals
def create_dict_for_row(row):
    values_dict = {}
    for i in range(1, days_in_month + 1):
        col_value = row[str(i)]
        if col_value != "0":  # Skip "0" values
            col_num = i
            if col_value not in values_dict:
                values_dict[col_value] = []
            values_dict[col_value].append(col_num)

    # Apply squash_intervals to each list of column numbers
    for key in values_dict:
        values_dict[key] = squash_intervals(values_dict[key])

    return values_dict

# Add a new column with the resulting dictionaries
df['values_dict'] = df.apply(create_dict_for_row, axis=1)

# Save the modified DataFrame to a new Excel file if needed
df.to_excel("october_edited_with_dicts.xlsx", index=False)


# Define mapping for key replacements
key_replacements = {
    "БЧ": "всебічне забезпечення розгорнутих пунктів управління та їх елементів",
    "МВГ": "в складі мобільно-вогневих груп",
    "ОВО": "охорона військових об'єктів",
    "МЗ": "медичне забезпечення",
    "ЛЗ": "логістичне забезпечення розгорнутих пунктів управління та їх елементів",
    "НО": "охорона, оборона та всебічне забезпечення розгорнутих пунктів та їх елементів",
    "Р": "розпорядження",
    "шпт": "шпиталь",
    "вдр": "відрядження",
    "вдп": "відпустка"
}


# Function to convert each sublist to the specified date format
def convert_to_date_ranges_string(data_dict, mm, yyyy):
    formatted_strings = []

    for key, sublists in data_dict.items():
        # Apply the date formatting to each sublist
        date_strings = []

        for sublist in sublists:
            if len(sublist) == 1 or sublist[0] == sublist[1]:
                # Case where both numbers are the same
                date_str = f"{sublist[0]:02}.{mm:02}.{yyyy}"
            else:
                # Case where there are two different numbers
                date_str = f"{sublist[0]:02}.{mm:02}.{yyyy}-{sublist[1]:02}.{yyyy}"

            date_strings.append(date_str)

        # Join all date strings for this key
        date_range_str = ", ".join(date_strings)

        # Get the replacement for the key, if it exists, otherwise keep the key
        replacement = key_replacements.get(key, key)

        # Format the string as "value - {string replacement}"
        formatted_strings.append(f"{date_range_str} - {replacement}")

    # Join all formatted strings into a single string
    return ", ".join(formatted_strings)


# Read the DataFrame from the Excel file
df = pd.read_excel("october_edited_with_dicts.xlsx")

# Set the month and year for the date formatting
mm = 10  # For October
yyyy = 2024

# Convert the 'values_dict' column from a string representation to an actual dictionary
df['values_dict'] = df['gitvalues_dict'].apply(lambda x: ast.literal_eval(x) if isinstance(x, str) else x)

# Apply the function to each row in the DataFrame's `values_dict` column
df['formatted_dates_string'] = df['values_dict'].apply(lambda d: convert_to_date_ranges_string(d, mm, yyyy))

# Save the modified DataFrame to a new Excel file if desired
df.to_excel("october_edited_with_formatted_dates_string.xlsx", index=False)
