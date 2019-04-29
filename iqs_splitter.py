import xlwings as xw
import pandas as pd


file_location = r'\\legros\Data\admin\leo121\processed\heineken\25-03-2019_13-21-19_heineken.xlsx'
sheet_name = 'IQS'

df = pd.read_excel(file_location, sheet_name, skiprows=1)

unique_agents = df.Pricing.unique()

for agent in unique_agents:
    print(agent)
    df2 = df.loc[df['Pricing'] == agent]
    df2.to_excel(f'{agent}_output.xlsx', sheet_name=sheet_name)

