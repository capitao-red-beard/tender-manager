import os

from azure.storage.blob import BlockBlobService

from utilities import key_manager

account_name = 'samsmdpblobdev02'
container_name = 'raw'
block_blob_service = BlockBlobService(account_name='samsmdpblobdev02',
                                      account_key=key_manager.get_password('azure', 'samsmdpblobdev02'))


def create_blob_from_path(blob_name, file_path):
    block_blob_service.create_blob_from_path(container_name, blob_name, file_path)


def delete_blob(blob_name):
    block_blob_service.delete_blob(container_name, blob_name)


def get_blob_list():
    blobs = []
    generator = block_blob_service.list_blobs(container_name)

    for blob in generator:
        blobs.append('Blob Name: {}'.format(blob.name))

    return blobs


def get_blob_url(blob_name):
    return block_blob_service.make_blob_url(container_name, blob_name)


def create_container():
    block_blob_service.create_container(container_name)


def delete_container():
    block_blob_service.delete_container(container_name)


def get_container_list():
    containers = []
    generator = block_blob_service.list_containers()

    for container in generator:
        containers.append('Container Name: {}'.format(container.name))

    return containers


def generate_blob_name(tender_file, original_file_name):
    path, file = os.path.split(original_file_name)
    all_parts = split_path(tender_file)

    return r'{}\{}\{}\{}\{}'.format(all_parts[4], all_parts[5][6:10], all_parts[5][3:5], all_parts[5][0:2], file)


def split_path(path):
    all_parts = []

    while 1:
        parts = os.path.split(path)

        if parts[0] == path:
            all_parts.insert(0, parts[0])
            break

        elif parts[1] == path:
            all_parts.insert(0, parts[1])
            break

        else:
            path = parts[0]
            all_parts.insert(0, parts[1])

    return all_parts


# no longer in use
# local_file = r'C:\Users\AV10\PycharmProjects\tender\iqs_data.parquet'
'''
def create_parquet_blob_from_path(blob_name):
    block_blob_service.create_blob_from_path(container_name, blob_name, local_file)


def update_blob_from_path(blob_name, data_frame):
    print(data_frame)

    try:
        block_blob_service.get_blob_to_path(container_name, blob_name, local_file)
        df_local = pd.read_parquet(local_file)
        df_to_submit = pd.concat([df_local, data_frame])

        try:
            write(local_file, df_to_submit)

        except ParquetException:
            print('Failed to convert pandas to parquet ' + str(ParquetException))

        block_blob_service.create_blob_from_path(container_name, blob_name, local_file)

    except Exception:
        data_frame.to_parquet(local_file, compression='gzip')
        create_parquet_blob_from_path(blob_name)
'''
