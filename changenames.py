import os
import pandas as pd

# Load the changenames.xlsx file
change_df = pd.read_excel('changenames.xlsx')

# Create a dictionary for the replacements
replacement_dict = dict(zip(change_df['rod'], change_df['dav']))

# Get all CSV files in the 'outputs' folder
output_folder = 'outputs_tags'
csv_files = [f for f in os.listdir(output_folder) if f.endswith('.csv')]

# Loop through each CSV file in the 'outputs' folder
for csv_file in csv_files:
    # Load the CSV file into a DataFrame
    df = pd.read_csv(os.path.join(output_folder, csv_file), header=None)

    # Replace values in the DataFrame using the replacement dictionary
    df[0].replace(replacement_dict, inplace=True)

    # Save the modified DataFrame with "_edited" appended to the filename
    new_filename = os.path.splitext(csv_file)[0] + '_edited.csv'
    df.to_csv(os.path.join(output_folder, new_filename), index=False, header=False)
