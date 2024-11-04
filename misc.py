import pandas
import pandas as pd

state = pd.read_excel("state_october.xlsx", nrows=712)
state = state.dropna(subset=["Ім'я"])
# Clean all string columns from leading and trailing whitespaces
state = state[['Unnamed: 7', 'Прізвище', "Ім'я", 'По-батькові']]
state = state.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
state.to_excel("state_oct.xlsx", index = False)
