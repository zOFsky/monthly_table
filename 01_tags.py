import os
import pandas as pd
import functions as fun
import json

# Load the config file
with open('config.json', 'r', encoding='utf-8') as file:
    config = json.load(file)

# Access values
month = config['month']
days_in_month = config['days_in_month']
input_folder = config['input_folder']
output_folder = config['output_folder']
roster_file = config['roster_file']


fun.process_folder(input_folder, output_folder, month)

# Load the changenames.xlsx file
change_df = pd.read_excel('changenames.xlsx')

# Create a dictionary for the replacements
replacement_dict = dict(zip(change_df['rod'], change_df['dav']))
#########################
# Get all CSV files in the 'outputs' folder
csv_files = [f for f in os.listdir(output_folder) if f.endswith('.csv')]

# Loop through each CSV file in the 'outputs' folder
for csv_file in csv_files:
    # Load the CSV file into a DataFrame
    df = pd.read_csv(os.path.join(output_folder, csv_file), header=None)

    # Replace values in the DataFrame using the replacement dictionary
    df.replace(replacement_dict, inplace=True)

    # Save the modified DataFrame with "_edited" appended to the filename
    new_filename = os.path.splitext(csv_file)[0] + '_edited.csv'
    df.to_csv(os.path.join(output_folder, new_filename), index=False, header=False)
########################################


roster = pd.read_excel(roster_file)
roster["surname"] = roster["surname"].str.capitalize()
roster["name_initials"] = roster["name"].str[0] + "." + roster["surname"]
roster = roster[["title", "surname", "name", "middlename", "name_initials"]]

fun.update_roster_with_one_to_one(roster, output_folder, days_in_month, month)

# Print or use the updated 'roster' DataFrame as needed
roster.to_excel("roster_done_tags.xlsx", index=False)
#from this point adding dupes manually is necessary