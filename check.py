import pandas as pd
from docx import Document
import os
import re

# Load the DataFrame
df = pd.read_excel('roster_done_edited_almost.xlsx')  # Replace with your actual DataFrame file

# Directory where the Word documents are stored
docx_folder = '09_september'

# Process each file in the directory
for filename in os.listdir(docx_folder):
    if filename.endswith('.docx'):
        print(f"working with file {filename}")
        # Extract the day number from the filename
        day_match = re.search(r'_(\d{2})\.', filename)
        if day_match:
            day = int(day_match.group(1))
            print(f"Extracted day {day}")


            # Read the document
            doc_path = os.path.join(docx_folder, filename)
            print(f"entering file {doc_path}")
            doc = Document(doc_path)
            doc_text = ' '.join(paragraph.text for paragraph in doc.paragraphs)

            # Iterate over surnames and update the DataFrame
            for surname in df['surname']:
                column_name = str(day)  # Ensure day is two digits
                if column_name in df.columns:
                    if df.at[df[df['surname'] == surname].index[0], column_name] == 0:
                        if surname in doc_text:
                            df.at[df[df['surname'] == surname].index[0], column_name] = 33

# Save the updated DataFrame to an Excel file
df.to_excel('df_checked.xlsx', index=False)
