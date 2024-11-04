import pandas as pd
from datetime import datetime

days_in_month = 31

# Load the Excel file and the specific sheet into a DataFrame
df = pd.read_excel('october_edited_colored.xlsx')

# Create the 'days_list' column by iterating over each row
def create_days_list(row):
    days = []
    for col in range(1, days_in_month+1):  # Iterate over columns 1 to 31
        if row[str(col)] in ['БЧ','МВГ', 'ОВО', 'Р', 'МЗ', 'ЛЗ','НО', "2"]:
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


# Sample dataframe (adjust as per your actual data)
# df = pd.DataFrame({'intervals': [[[1, 5], [10, 15]], [], [[20, 25]]]})

def extract_dates(intervals_column, month, year):
    # Helper function to convert day, month, and year into 'dd.mm.yyyy' format
    def convert_to_date(day, month, year):
        date_obj = datetime(year, month, day)
        return date_obj.strftime('%d.%m.%Y')

    # Function to extract dates from intervals
    def get_dates(intervals, month, year):
        from_dates = []
        to_dates = []

        for interval in intervals:
            if interval:  # Check if interval is not empty
                from_day = interval[0]  # First element of the interval
                to_day = interval[1]  # Second element of the interval

                # Convert days to date strings
                from_dates.append(convert_to_date(from_day, month, year))
                to_dates.append(convert_to_date(to_day, month, year))

        # Join dates with breaklines ('\n')
        return '\n'.join(from_dates), '\n'.join(to_dates)

    # Apply the get_dates function to each row
    df['from'] = df['intervals'].apply(lambda x: get_dates(x, month, year)[0])
    df['to'] = df['intervals'].apply(lambda x: get_dates(x, month, year)[1])


# Parameters for the function (adjust month and year as needed)
month = 10  # Example: September
year = 2024

# Apply the function to the DataFrame
extract_dates(df['intervals'], month, year)

# Save the updated DataFrame to Excel (if needed)
# df.to_excel('your_file_updated.xlsx', sheet_name='пораховані дні', index=False)
output_file = "result_int.xlsx"
# Write to Excel using xlsxwriter to enable text wrapping

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