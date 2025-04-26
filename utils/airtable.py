import os
from pyairtable import api

AIRTABLE_API_KEY = os.getenv('AIRTABLE_API_KEY')
BASE_ID = os.getenv('AIRTABLE_BASE_ID')
AIRTABLE_API = api.Api(AIRTABLE_API_KEY)

TABLE_NAME = 'Outlook Emails 2nd Batch'


def get_entries() -> list[dict]:
    table = AIRTABLE_API.table(base_id=BASE_ID, table_name=TABLE_NAME)
    records = table.all()
    for record in records:
        record['fields']['id'] = record['id']
    entries = [record['fields'] for record in records if not record['fields'].get('Prepared')]
    return entries


def update_entry(entry_id: str, fields: dict) -> dict:
    table = AIRTABLE_API.table(base_id=BASE_ID, table_name=TABLE_NAME)
    record = table.update(entry_id, fields, typecast=True)
    return record['fields']