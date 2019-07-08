# Built-in/Generic Imports
import os
import shutil
from pathlib import Path
import sys

# Own modules
from azure_connectors import blob_manager, table_manager
from excel_handlers import tender_manager

__author__ = '{Adriaan van der Valk}, {Jason van Pelt}'
__copyright__ = 'Copyright {2019}, {tender_manager}'
__credits__ = ['']
__license__ = '{MIT}'
__version__ = '{1}.{0}.{0}'
__maintainer__ = '{Adriaan van der Valk}, {Jason van Pelt}'
__email__ = '{adrioaan.van.der.valk@samskip.com}, {jason.van.pelt@samskip.com}'
__status__ = '{in development}'

# folder location of processed IQS
processed_folder = r'\\legros\Data\admin\leo121\reverse_processed'

# variables to get from MS powershell
iqs_path = sys.argv[1]

# get the iqs file name0
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
