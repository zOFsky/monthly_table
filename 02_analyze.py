import os
import json
import pandas as pd
from functions import analyze_dataframes

# Load the config file
with open('config.json', 'r') as file:
    config = json.load(file)

# Access values
month = config['month']
days_in_month = config['days_in_month']


def save_list_to_excel(lst, filename):
    """Save list to an Excel file with 1 value per row."""
    # Create a DataFrame from the list
    df = pd.DataFrame(lst, columns=['Values'])

    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)


# Ensure the 'analysis' folder exists
analysis_folder = 'analysis'
os.makedirs(analysis_folder, exist_ok=True)

# Initialize an empty set for the union of missing_in_roster
union_missing_in_roster = set()

# Iterate through the files aug_1 to aug_31
for N in range(days_in_month+1):
    filename = f'outputs/{month}_{N}_edited.csv'

    # Read the CSV file into a DataFrame
    df_aug = pd.read_csv(filename)
    df_aug.columns = ["name", "duty"]

    # Apply the analyze_dataframes function
    not1to1, one_to_one, missing_in_extracted, missing_in_roster = analyze_dataframes(df_aug, roster)

    # Save the lists to text files
    save_list_to_excel(missing_in_roster, os.path.join(analysis_folder, f'missing_in_roster_{N}.xlsx'))
    save_list_to_excel(one_to_one, os.path.join(analysis_folder, f'one_to_one_{N}.xlsx'))
    save_list_to_excel(not1to1, os.path.join(analysis_folder, f'duplicates_{N}.xlsx'))
    save_list_to_excel(missing_in_extracted, os.path.join(analysis_folder, f'not_in_extracted_{N}.xlsx'))


    # Update the union set with missing_in_roster
    union_missing_in_roster.update(missing_in_roster)

# Save the union of missing_in_roster to a text file
save_list_to_excel(union_missing_in_roster, os.path.join(analysis_folder, 'union_missing_in_roster.xlsx'))
