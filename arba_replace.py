import pandas as pd

# Load the config file
with open('config.json', 'r', encoding='utf-8') as file:
    config = json.load(file)

# Access values
exclude_columns = config["arba_days"]

df = pd.read_excel("roster_done_tags.xlsx")
# Columns to exclude from replacement
#exclude_columns = ["8", "16", "25", "27", "28"]

# Replace "АРБА" with 0 in all columns except those in the exclude list
df = df.apply(lambda col: col.replace("АРБА", 0) if col.name not in exclude_columns else col)

df.to_excel("roster_result.xlsx", index=False)