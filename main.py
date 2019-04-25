import os
import shutil
from pathlib import Path
import sys

import pandas as pd
import json

import blob_manager
import tender_manager
import table_manager

# folder location of processed IQS
processed_folder = r'\\legros\Data\admin\leo121\reverse_processed'

# pre-defined name for the blob to store tender data
blob_name_iqs = 'iqs_data.parquet'

# variables to get from MS powershell
iqs_path = sys.argv[1]

# get the iqs file name
iqs_file = blob_manager.split_path(iqs_path)[4]

# files to get from leo folder
folder = Path('//legros/Data/admin/leo121/config')
log_file = Path(folder / 'file_names.txt')
transit_mapping_file = Path(folder / 'Transit_mapping.xlsx')
config_file = Path(folder / 'matched_columns.csv')

# get the correct parameters for the given tender file
parameters = tender_manager.get_tender_data(str(log_file), str(iqs_file))

# then we convert the iqs file to the original tender format
df = tender_manager.to_tender_format(parameters[0], parameters[2], parameters[3], parameters[1],
                                     str(transit_mapping_file), str(config_file), parameters[5], parameters[4])

# update the table with the data from IQS sheet
# blob_manager.update_blob_from_path(blob_name_iqs, tender_manager.taf_to_blob_format(parameters[1]))
table_manager.insert_batch_entity(tender_manager.taf_to_blob_format(parameters[1]))

# finally write the unmodified taf to the blob
blob_manager.create_blob_from_path(blob_manager.generate_blob_name(parameters[1], parameters[0]), parameters[0])

# move the IQS file to the processed folder
try:
    shutil.move(iqs_path, processed_folder)
    print('Ended with success')

except shutil.Error:
    try:
        os.remove(iqs_path)
        print('Ended with success and removed duplicate')

    except FileNotFoundError:
        print('Could not find the file to move')
        pass

# end
