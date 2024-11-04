import os
import re
import pandas as pd
from docx import Document
import csv

days_in_month = 31
input_folder = '10_october'  # Folder containing .docx files
month = 'oct'
output_folder = 'outputs'  # Folder where outputs will be saved

import re
import csv
from docx import Document


def extract_cyrillic_patterns(doc_path, output_txt_path, output_csv_path):
    # Open the Word document
    doc = Document(doc_path)

    # Define the regex pattern to include all Ukrainian Cyrillic letters
    # - Optional "Capital letter.-" before the main pattern
    # - Apostrophe (’) allowed after the second big letter or between small letters
    # - Whitespace allowed after the capital letter and dot
    pattern = re.compile(r'(?:[А-ЩЬЮЯҐЄІЇ]\.-)?([А-ЩЬЮЯҐЄІЇ])\.\s?([А-ЩЬЮЯҐЄІЇа-щьюяґєії]{1,}(?:’?[а-щьюяґєії]+)?)')

    # Regex to capture content inside tags (e.g., {МВГ} ... {/МВГ})
    tag_pattern = re.compile(r'\{([А-ЯҐЄІЇ]+)\}([\s\S]+?)\{\/\1\}', re.DOTALL)

    # Concatenate all paragraph text into a single string to allow for cross-paragraph tag searches
    full_text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])

    # List to store the extracted patterns along with the tag
    extracted_patterns = []

    # Search for text inside tags across the entire document text
    tag_matches = tag_pattern.findall(full_text)

    for tag, content in tag_matches:
        # Search for the Cyrillic pattern inside the content of each tag
        matches = pattern.findall(content)
        for match in matches:
            # Reconstruct the full match and pair it with the tag
            full_match = f"{match[0]}.{match[1]}"
            extracted_patterns.append((full_match, tag))

    # Write the extracted patterns to the output text file
    with open(output_txt_path, 'w', encoding='utf-8') as txt_file:
        for pattern, tag in extracted_patterns:
            txt_file.write(f"{pattern} ({tag})\n")

    # Write the extracted patterns to the output CSV file
    with open(output_csv_path, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        for pattern, tag in extracted_patterns:
            writer.writerow([pattern, tag])


def process_folder(input_folder, output_folder, month):
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Loop through each .docx file in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith('.docx'):
            doc_path = os.path.join(input_folder, filename)

            # Extract the day number from the filename
            day_number = int(filename.split('_')[1].split('.')[0])  # Extract day from format "БН_04.08.2024.docx"

            # Define output paths for .txt and .csv files with the format "aug_{N}"
            output_txt_path = os.path.join(output_folder, f"{month}_{day_number}.txt")
            output_csv_path = os.path.join(output_folder, f"{month}_{day_number}.csv")

            # Apply the extraction function
            extract_cyrillic_patterns(doc_path, output_txt_path, output_csv_path)
            print(f"Processed and saved: {filename} as aug_{day_number}")

def analyze_dataframes(extracted, roster):
    # 1) Lists of items present in both but not in 1:1 relation
    # Group by 'name' and 'name_initials' and count occurrences
    extracted_grouped = extracted['name'].value_counts()
    roster_grouped = roster['name_initials'].value_counts()

    # Find items that are present in both dataframes
    common_items = set(extracted_grouped.index).intersection(set(roster_grouped.index))

    # Filter for items where there's more than one occurrence in at least one of the dataframes
    not_one_to_one = [item for item in common_items if roster_grouped[item] > 1]

    # 2) List of items that are 1:1 (no duplicates in either dataframe)
    one_to_one = [item for item in common_items if extracted_grouped[item] >= 1 and roster_grouped[item] == 1]

    # 3) List of items present in 'roster' but absent in 'extracted'
    missing_in_extracted = list(set(roster['name_initials']) - set(extracted['name']))

    # 4) List of items present in 'extracted' but absent in 'roster'
    missing_in_roster = list(set(extracted['name']) - set(roster['name_initials']))

    return not_one_to_one, one_to_one, missing_in_extracted, missing_in_roster


process_folder(input_folder, output_folder, month)


def update_roster_with_one_to_one(roster, outputs_folder, days_in_month, month):
    for i in range(0, days_in_month + 1):  # Loop from 1 to the number of days in the month
        # Construct file path
        file_path = os.path.join(outputs_folder, f"{month}_{i}_edited.csv")

        # Check if the file exists
        if os.path.exists(file_path):
            # Read the CSV file into a DataFrame
            extracted = pd.read_csv(file_path, names=['name', 'type'])

            # Apply analyze_dataframes function
            _, one_to_one, _, _ = analyze_dataframes(extracted, roster)

            # Create a new column in the 'roster' DataFrame named i (as a string)
            column_name = str(i)
            roster[column_name] = 0

            # Update the column with the 'type' value where the item is present in one_to_one list
            for name in one_to_one:
                type_value = extracted.loc[extracted['name'] == name, 'type'].values[0]
                roster.loc[roster['name_initials'] == name, column_name] = type_value


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

# Example usage:
outputs_folder = 'outputs'  # Folder where the CSV files are stored

roster = pd.read_excel("state_oct.xlsx")
roster["surname"] = roster["surname"].str.capitalize()
roster["name_initials"] = roster["name"].str[0] + "." + roster["surname"]
roster = roster[["title", "surname", "name", "middlename", "name_initials"]]

update_roster_with_one_to_one(roster, outputs_folder, days_in_month, month)

# Print or use the updated 'roster' DataFrame as needed
roster.to_excel("roster_done_tags.xlsx", index=False)