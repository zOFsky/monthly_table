import json
import functions as fun

# Load the config file
with open('config.json', 'r', encoding='utf-8') as file:
    config = json.load(file)

# Access values
month = config['month']
days_in_month = config['days_in_month']
duties = config['duties_2_days']
arba_days = config['arba_days']


# requires manually colored table with added dupes and filtered days of arba
colors = fun.replace_values_by_color('roster_result_colored_dupes.xlsx',
                                     f'{month}_replaced.xlsx')


# Example usage
fun.process_excel_file(f'{month}_replaced.xlsx', duties, days_in_month)
# "with_calculated_days.xlsx" is created