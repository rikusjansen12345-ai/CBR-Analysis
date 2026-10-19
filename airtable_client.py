import requests

from config import (
    AIRTABLE_TOKEN,
    AIRTABLE_BASE_ID,
    AIRTABLE_TABLE_NAME,
)


def get_records():
    """
    Retrieves all records from the Airtable table.
    """

    url = (
        f"https://api.airtable.com/v0/"
        f"{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
    )

    headers = {
        "Authorization": f"Bearer {AIRTABLE_TOKEN}"
    }

    records = []
    offset = None

    while True:

        params = {}

        if offset:
            params["offset"] = offset

        response = requests.get(
            url,
            headers=headers,
            params=params
        )

        response.raise_for_status()

        data = response.json()

        records.extend(data["records"])

        offset = data.get("offset")

        if not offset:
            break

    return records