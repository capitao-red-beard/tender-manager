import warnings

warnings.filterwarnings("ignore")

import numpy as np
import numpy.core.multiarray as multiarray
import pandas as pd

config_location = ''
file_location = ''

# import format is file_name,sheet_name,row_number,skip_list,payload_notation,transit_time_natation
kryon_input = ['file','sheet',1,[],'empty','empty']

file_name = file_location + str(kryon_input[0])
sheet_name = str(kryon_input[1])
row_number = int(kryon_input[2]) - 1
skip_list = kryon_input[3]
payload_notation = str(kryon_input[4]).lower()
transit_time_notation = str(kryon_input[5]).lower()



data_columns = ['customer lane id',
                  'origin country',
                  'origin city',
                  'origin postal code',
                  'destination country',
                  'destination city',
                  'destination postal code',
                  'requested equipment type',
                  'offered equipment type'
                  'shipments per year',
                  'transit time',
                  'payload',
                  'currency',
                  'round 1 offered rate',
                  'round 2 offered rate',
                  'round 3 offered rate']

data_columns = [i for i in data_columns if i not in skip_list]

def read_to_dict(file, sheet,column_keys = 'Original', columns_items = 'Samskip'):

    if (file[-4:] == 'xlsx') | (file[-4:] == '.xls'):
        df = pd.read_excel(file, sheet_name=sheet)
    elif file[-4:] == '.csv':
        df = pd.read_csv(file)
    else:
        return "Wrong file type"

    list_keys = [str(i).lower().strip().replace('\n', '') for i in df[column_keys]]
    list_items = [str(i).lower().strip().replace('\n', '') for i in df[columns_items]]
    dictionary_to_return = dict(zip(list_keys, list_items))

    return dictionary_to_return

def find_missing_columns(file, sheet, start_row=0, columns_needed=[],payload_type = '' , transit_type='' ):

    column_config_dictionary = read_to_dict(config_location + 'matched_columns.csv', 'matched_columns')

    try:
        original = pd.read_excel(file, sheet_name=sheet, skiprows=start_row, dtype=str)
    except FileNotFoundError:
        original = pd.read_csv(file, skiprows=start_row, dtype=str)

    original.columns = [i.strip().replace("'", '').replace('\n', '').replace('€', 'eur').replace('°c', 'celsius').replace('[','').replace(']', '') for i in original.columns.str.lower()]
    original = original.rename(columns=column_config_dictionary)

    missing_columns = [i for i in columns_needed if i not in list(original.columns)]
    if missing_columns or ((payload_notation == 'empty') & ('payload' in columns_needed)) or (( transit_time_notation == 'empty') & ('transit time' in columns_needed )):
        return True, missing_columns
    else:
        df = original[columns_needed]

        if payload_type.isin(['kilos','kilo']):
            try:
                original['payload'] = [int(i) / 1000 for i in original['payload (in ton)']]
            except ValueError:
                return 'Error in pyaload conversion'

        transit_df = pd.read_excel(config_location + 'Transit_mapping.xlsx',sheet_name=transit_type)
        transit_dict = dict(zip(transit_df['Original'], transit_df['Samskip']))
        df['transit time'].replace(transit_dict, inplace=True)

        # insert in table
        # insert original in blob

        return False

# find_missing_columns(file_name,sheet_name,row_number,data_columns,payload_notation,transit_time_notation)
