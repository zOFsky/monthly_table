import pandas as pd
import json
import functions as fun

# Load the config file
with open('config.json', 'r', encoding='utf-8') as file:
    config = json.load(file)

# Access values
year = config['year']
month = config['month_int']
days_in_month = config['days_in_month']
duties = config['duties_2_days']


# Load the Excel file and the specific sheet into a DataFrame
df = pd.read_excel('with_calculated_days.xlsx')

# Create the 'days_list' column by iterating over each row
def create_days_list(row):
    days = []
    for col in range(1, days_in_month+1):  # Iterate over columns 1 to 31
        if row[str(col)] in (duties + ['АРБА',"2"]):
            days.append(col)
    return days

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


# Apply the function to each row to create the 'days_list' column
df['days_list'] = df.apply(create_days_list, axis=1)

df['intervals'] = df['days_list'].apply(squash_intervals)


# Apply the function to the DataFrame
fun.extract_dates(df, month, year)

# Save the updated DataFrame to Excel (if needed)
output_file = "result_int.xlsx"

# Write to Excel using xlsxwriter to enable text wrapping
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='пораховані дні', index=False)

    # Get the workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['пораховані дні']

    # Set text wrap for the columns with breaklines
    wrap_format = workbook.add_format({'text_wrap': True})

    # Apply text wrapping to the 'from' and 'to' columns
    worksheet.set_column('B:B', 20, wrap_format)  # Assuming 'from' is column B
    worksheet.set_column('C:C', 20, wrap_format)  # Assuming 'to' is column C

    # Adjust row heights dynamically based on text length (optional)
    for row_num, row_data in enumerate(df['from'], start=1):
        row_height = max(len(row_data.split('\n')), 1) * 15  # Set row height based on the number of lines
        worksheet.set_row(row_num, row_height)