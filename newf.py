import pandas as pd

df = pd.read_excel("sem5_result.xlsx", skiprows=6)
print(df.columns)