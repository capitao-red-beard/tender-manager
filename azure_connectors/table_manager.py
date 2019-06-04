from azure.cosmosdb.table.tableservice import TableService
from azure.cosmosdb.table.tablebatch import TableBatch

from utilities import key_manager

table_service = TableService(account_name='samsmdpblobdev02',
                             account_key=key_manager.get_password('azure', 'samsmdpblobdev02'))


def create_table(table_name):
    table_service.create_table(table_name)


def insert_entity(task, table_name='tender'):
    table_service.insert_entity(table_name, task)


def insert_batch_entity(df, table_name='tender'):
    df.columns = [i.lower().replace(' ', '_') for i in df.columns]
    df.columns = [i.replace('/', 'per') for i in df.columns]
    df.columns = [i.replace('(', '') for i in df.columns]
    df.columns = [i.replace(')', '') for i in df.columns]
    df.columns = [i.replace('%', 'percent') for i in df.columns]
    df.columns = [i.replace(':', '') for i in df.columns]
    df.columns = [i.replace('1', 'one') for i in df.columns]
    df.columns = [i.replace('2', 'two') for i in df.columns]
    df.columns = [i.replace('3', 'three') for i in df.columns]
    data = df.to_dict('records')

    base = get_max_entity_row()

    batch = TableBatch()

    for i, d in enumerate(data, 1):
        d['PartitionKey'] = 'iqs'
        d['RowKey'] = str(base + i)

        batch.insert_entity(d)

        if i % 100 is 0 or i is len(data) + 1:
            table_service.commit_batch(table_name, batch)
            batch = TableBatch()


def get_max_entity_row(partition_key='iqs', table_name='tender'):
    tasks = table_service.query_entities(table_name, filter=f"PartitionKey eq '{partition_key}'")

    current_max = 0

    try:
        for t in tasks:
            if int(t.RowKey) > int(current_max):
                current_max = int(t.RowKey)

    except AttributeError:
        return int(current_max)

    return int(current_max) + 1


def delete_table(table_name):
    table_service.delete_table(table_name)


def delete_entity(partition_key, row_key, table_name='tender'):
    table_service.delete_entity(table_name, partition_key, row_key)
