import os
import re
from docx import Document


def extract_patterns_from_docx(directory, output_file):
    # Define the regex pattern to match {some_text} and {/another_text}
    pattern = re.compile(r"\{(/?.+?)\}")

    with open(output_file, 'w', encoding='utf-8') as outfile:
        # Iterate through all files in the directory
        for filename in os.listdir(directory):
            if filename.endswith('.docx'):
                docx_path = os.path.join(directory, filename)

                try:
                    # Read the .docx file
                    document = Document(docx_path)

                    # Extract all text from the document
                    full_text = '\n'.join([paragraph.text for paragraph in document.paragraphs])

                    # Find all patterns in the text
                    matches = pattern.findall(full_text)

                    # Write the filename to the output file
                    outfile.write(f"File: {filename}\n")

                    # Write each matched pattern
                    for match in matches:
                        outfile.write(f"{{{match}}}")
                        # Add a breakline if the pattern starts with {/
                        if match.startswith('/'):
                            outfile.write("\n")

                    outfile.write("\n")  # Separate entries by a blank line

                except Exception as e:
                    print(f"Error processing file {filename}: {e}")


# Define the directory with .docx files and the output .txt file
docx_directory = '12_december'  # Change this to your directory path
output_txt_file = 'extracted_patterns.txt'

# Run the function
extract_patterns_from_docx(docx_directory, output_txt_file)

print(f"Patterns have been extracted and saved to {output_txt_file}.")
