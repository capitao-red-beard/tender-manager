import warnings

warnings.filterwarnings("ignore")

from azure.storage.blob import BlockBlobService
import numpy as np
import numpy.core.multiarray as multiarray
import pandas as pd

matched_columns_location = '\\\\lebrun\\Data\\admin\\leo139\\config\\matched_columns.csv'
# matched_columns_location = br'$MATCHED_COLUMNS_LOCATION$'

# import format is file_name,sheet_name,row_number,skip_list,payload_notation,transit_time_natation
# kryon_input = [1,'Freight & Service',12,[],'empty','empty']
kryon_input = $parameters$

# file_name = '//lebrun/Data/admin/leo139/to_be_processed/HEINEKEN_European International Transport Tender 2019-2020.xlsx'
file_name = r'$excel_file_path$'
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

account_name = 'samsmdpblobdev02'
container_name = 'raw'
block_blob_service = BlockBlobService(account_name='samsmdpblobdev02',
                                      account_key='Dr3Qut1sQMqUdTFAZ8u4fKePFfoTgMETi4/RMURiT6wcyqCFC0m1l1bnYtDDXAaFBjs4IbcXY8Xt89dRYkNY6Q==')

def create_blob_from_path(blob_name, file_path):
    block_blob_service.create_blob_from_path(container_name, blob_name, file_path)

def read_to_dict(file, sheet, column_keys='Original', columns_items='Samskip'):
    if (file[-4:] == 'xlsx') | (file[-4:] == '.xls'):
        df = pd.read_excel(file, sheet_name=sheet)
    elif file[-4:] == '.csv':
        df = pd.read_csv(file)
    else:
        print("Wrong file type")
        return("Wrong file type")

    list_keys = [str(i).lower().strip().replace('\n', '') for i in df[column_keys]]
    list_items = [str(i).lower().strip().replace('\n', '') for i in df[columns_items]]
    dictionary_to_return = dict(zip(list_keys, list_items))

    return dictionary_to_return


def find_missing_columns(file, sheet, start_row=0, columns_needed=[], payload_type='', transit_type=''):
    column_config_dictionary = read_to_dict(matched_columns_location, 'matched_columns')

    try:
        original = pd.read_excel(file, sheet_name=sheet, skiprows=start_row, dtype=str, engine='xlrd')
    except FileNotFoundError:
        original = pd.read_csv(file, skiprows=start_row, dtype=str)

    original.columns = [
        i.strip().replace("'", '').replace('\n', '').replace('€', 'eur').replace('°c', 'celsius').replace('[',
                                                                                                          '').replace(
            ']', '') for i in original.columns.str.lower()]
    original.rename(columns=column_config_dictionary, inplace=True)

    missing_columns = [i for i in columns_needed if i not in list(original.columns)]
    if missing_columns or ((payload_notation == 'empty') & ('payload' in columns_needed)) or (
            (transit_time_notation == 'empty') & ('transit time' in columns_needed)):
        print(True, missing_columns)
    else:
        df = original[columns_needed]

        if payload_type.isin(['kilos', 'kilo']):
            try:
                original['payload'] = [int(i) / 1000 for i in original['payload (in ton)']]
            except ValueError:
                return 'Error in pyaload conversion'

        transit_df = pd.read_excel(config_location + 'Transit_mapping.xlsx', sheet_name=transit_type)
        transit_dict = dict(zip(transit_df['Original'], transit_df['Samskip']))
        df['transit time'].replace(transit_dict, inplace=True)

        # insert in table
        # insert original in blob
        # insert customer from input

        print(False)

find_missing_columns(file_name,sheet_name,row_number,data_columns,payload_notation,transit_time_notation)
