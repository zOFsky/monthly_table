import os
import re
import pandas as pd
from docx import Document
import csv

def extract_cyrillic_patterns(doc_path, output_txt_path, output_csv_path):
    # Open the Word document
    doc = Document(doc_path)

    # Define the regex pattern to include all Ukrainian Cyrillic letters
    # - Optional "Capital letter.-" before the main pattern
    # - Apostrophe (’) allowed after the second big letter or between small letters
    # - Whitespace allowed after the capital letter and dot
    pattern = re.compile(r'(?:[А-ЩЬЮЯҐЄІЇ]\.-)?([А-ЩЬЮЯҐЄІЇ])\.\s?([А-ЩЬЮЯҐЄІЇа-щьюяґєії]{1,}(?:’?[а-щьюяґєії]+)?)')

    # List to store the extracted patterns
    extracted_patterns = []
    # Iterate through paragraphs in the document
    for paragraph in doc.paragraphs:
        # Search for the pattern in each paragraph
        matches = pattern.findall(paragraph.text)
        for match in matches:
            # Reconstruct the full match and add it to the list
            full_match = f"{match[0]}.{match[1]}" #if ' ' in paragraph.text else f"{match[0]}.{match[1]}"
            extracted_patterns.append(full_match)

    # Write the extracted patterns to the output text file
    with open(output_txt_path, 'w', encoding='utf-8') as txt_file:
        for pattern in extracted_patterns:
            txt_file.write(pattern + '\n')

    # Write the extracted patterns to the output CSV file
    with open(output_csv_path, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        for pattern in extracted_patterns:
            writer.writerow([pattern])


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
            print(f"Processed and saved: {filename} as {month}_{day_number}")


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


def update_roster_with_one_to_one(roster, outputs_folder, days_in_month, month):
    for i in range(days_in_month+1):  # Loop from 1 to 10
        # Construct file path
        file_path = os.path.join(outputs_folder, f"{month}_{i}_edited.csv")

        # Check if the file exists
        if os.path.exists(file_path):
            # Read the CSV file into a DataFrame
            extracted = pd.read_csv(file_path, names = ['name', 'type'])
            #extracted.columns = ['name']  # Ensure the column is named 'name'

            # Apply analyze_dataframes function
            _, one_to_one, _, _ = analyze_dataframes(extracted, roster)

            # Create a new column in the 'roster' DataFrame named i (as a string)
            column_name = str(i)
            roster[column_name] = 0

            # Update the column with 1 where the item is present in one_to_one list
            roster.loc[roster['name_initials'].isin(one_to_one), column_name] = 1

def squash_intervals(numbers):
    """
    Takes a list of numbers and returns a list of intervals.
    Each interval is represented by the first and last number in the consecutive sequence.
    If a number is not part of a sequence, it will still appear as a pair [number, number].

    Args:
        numbers (list): A list of sorted numbers.

    Returns:
        list: A list of lists, where each list is a pair representing an interval.
    """
    if not numbers:
        return []

    # Initialize the list to hold squashed intervals
    squashed = []

    # Initialize the start of the first interval
    start = numbers[0]

    for i in range(1, len(numbers)):
        # If the current number is not consecutive
        if numbers[i] != numbers[i - 1] + 1:
            # Append the interval to the result as a pair [start, end]
            squashed.append([start, numbers[i - 1]])
            # Update the start of the next interval
            start = numbers[i]

    # Add the final interval as a pair [start, end]
    squashed.append([start, numbers[-1]])

    return squashed
