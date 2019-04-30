import numpy as np
import pandas as pd
import xlwings as xw

import warnings

warnings.filterwarnings('ignore')


def get_tender_data(log_file, iqs_file):
    with open(log_file) as file:
        data = file.readlines()

        for d in data:
            if iqs_file in d:
                return_list = [i.replace('[', '').replace(']', '').replace('\n', '') for i in d.split(',')]

                return return_list


def get_excel_column(number):
    # function to get column letter by giving an integer

    a_b_c_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                  'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    if 26 <= number <= 702:
        one = int(np.floor(number / 26) - 1)
        two = int(number % 26)

        return a_b_c_list[one] + a_b_c_list[two]

    elif number >= 702:
        one = int(np.floor(number / 702) - 1)
        two = int(np.floor((number - 702 - (one * 702)) / 26))
        three = int(number % 26)

        return a_b_c_list[one] + a_b_c_list[two] + a_b_c_list[three]

    else:
        return a_b_c_list[number]


def taf_to_blob_format(samskip_file=''):
    # function to prepare the taf file for the database input

    df_taf = pd.read_excel(samskip_file, sheet_name='TAF')
    df_iqs = pd.read_excel(samskip_file, sheet_name='IQS', skiprows=1)

    df_iqs.columns = [i.strip() for i in df_iqs.columns]

    df_iqs = df_iqs[['Customer Lane ID', 'Pricing',
                     'Department', 'Lane', 'No', 'Origin Country', 'Origin City',
                     'Origin Zip Code', 'Destination Country', 'Destination City',
                     'Destination Zip Code', 'Equipment Requested', 'Shipments / Year',
                     'Transit Time', 'Equipment Offered', 'Payload (in ton)',
                     'Currency', 'Total Costs', 'Margin Round 1', 'Margin % Round 1',
                     'Round 1 Rate', 'Total Revenue round 1', 'Total current margin Round 1',
                     'Round 1: Rank', 'Round 2: Target Rate', 'Margin Round 2',
                     'Margin % Round 2', 'Round 2 Rate', 'Total Revenue round 2',
                     'Total current margin Round 2', 'Round 2: Rank', 'Round 3: Target Rate',
                     'Margin Round 3', 'Margin % Round 3', 'Round 3 Rate',
                     'Total Revenue Round 3', 'Total current margin Round 3',
                     'Awards Volume', 'Awards Allocation %', 'Awards Rate', 'Feedback']]

    df_iqs['customer'] = df_taf.iloc[1, 2]
    df_iqs['tender'] = df_taf.iloc[2, 2]
    df_iqs['date'] = df_taf.iloc[3, 2]
    df_iqs['last_tender'] = df_taf.iloc[4, 2]
    df_iqs['account_manager'] = df_taf.iloc[5, 2]
    df_iqs['commodity'] = df_taf.iloc[27, 5]
    df_iqs['valid_from'] = df_taf.iloc[39, 5]
    df_iqs['valid_to'] = df_taf.iloc[40, 5]
    df_iqs['validity'] = df_taf.iloc[41, 5]
    # to check later on if it is the correct row number
    df_iqs['historic feedback'] = df_taf.iloc[77, 5]

    print('Ended Successfully')

    return df_iqs


def to_tender_format(tender_file='', tender_sheet_name='', row_number=0, samskip_file='', transit_time_file='',
                     config_file='', payload_unit='tonnes', transit_time_unit='days'):
    # function to transform IQS sheet to customer sheet

    app = xw.App()
    wb = app.books.open(tender_file)

    sht1 = wb.sheets[tender_sheet_name]
    sht1.select()

    df_sam = pd.read_excel(samskip_file, sheet_name='IQS', skiprows=1)
    df_sam.columns = [i.lower().strip().replace("'", '').replace('\n', '').replace('€', 'eur').replace('°c', 'celsius')
                          .replace('[', '').replace(']', '') for i in df_sam.columns]
    part1 = df_sam.loc[:, 'transit time':'total current margin round 1']
    part2 = df_sam.loc[:, 'round 1: rank':'feedback']
    df_sam = pd.concat([part1.iloc[:, 1:], part2], 1)
    sam_headers = list(df_sam.columns)

    df_cus = pd.read_excel(tender_file, sheet_name=tender_sheet_name, skiprows=int(row_number) - 1)
    df_cus.columns = [i.lower().strip().replace("'", '').replace('\n', '').replace('€', 'eur').replace('°c', 'celsius')
                          .replace('[', '').replace(']', '') for i in df_cus.columns]
    cus_headers = list(df_cus.columns)

    del df_cus

    df_config = pd.read_csv(config_file)
    config_dic = dict(zip(list(df_config['Original']), list(df_config['Samskip'])))

    match_dic = {}

    for col in cus_headers:
        if col in sam_headers:
            match_dic[col] = col

        else:
            try:
                match_dic[col] = config_dic[col]

            except KeyError:
                pass

    for col_cus in match_dic:
        print(col_cus)

        try:
            col_sam = match_dic[col_cus]

        except KeyError:
            col_sam = col_cus

        sam_data = df_sam[col_sam]

        if (col_sam == 'payload (in ton)') and (payload_unit == 'kg'):
            sam_data = [i / 1000 for i in sam_data]

            print('Payload changed from kg')

        if col_sam == 'transit time':
            transit_df = pd.read_excel(transit_time_file, sheet_name=transit_time_unit)
            transit_dic = dict(zip(list(transit_df['Samskip']), list(transit_df['Original'])))
            new_data = []

            for d in sam_data:
                try:
                    new_data.append(transit_dic[d])

                except KeyError:
                    new_data.append(np.nan)

            print('transit time notation changed')

        cell_letter = get_excel_column([i for i, x in enumerate(cus_headers) if x == col_cus][0])
        cell_list = [cell_letter + str(num) for num in range(int(row_number) + 1, int(row_number) + len(df_sam))]

        for x, cell in enumerate(cell_list):
            # print(cell)
            # print(sht1.range(cell).value)
            # print(list(df_sam[col_sam])[x])
            # print(pd.isnull(list(df_sam[col_sam])[x]))
            try:
                if (sht1.range(cell).value == list(df_sam[col_sam])[x]) or \
                        (pd.isnull(list(df_sam[col_sam])[x]) == True) or \
                        (list(df_sam[col_sam])[x] == None) or \
                        (list(df_sam[col_sam])[x] == 'nan'):
                    continue

                elif list(df_sam[col_sam])[x] != '(Select a value)':
                    # print('elif',list(df_sam[col_sam])[x])
                    sht1.range(cell).value = list(df_sam[col_sam])[x]

            except Exception as e:
                # print(list(df_sam[col_sam])[x])
                # print(e)
                pass

    wb.save()
    wb.close()
    app.kill()

    print('Ended Successfully')