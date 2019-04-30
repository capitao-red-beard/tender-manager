import pandas as pd

file_location = r'\\legros\Data\admin\leo121\processed\heineken\25-03-2019_13-21-19_heineken.xlsx'
sheet_name = 'IQS'

df = pd.read_excel(file_location, sheet_name, skiprows=1)

unique_lanes = df.Lane.unique()

matches = []

for end in unique_lanes:
    for start in unique_lanes:
        if end[5:] == start[:2]:
            matches.append(f'{end} -> {start}')

print(matches)
