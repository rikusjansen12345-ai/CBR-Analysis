from config import AIRTABLE_TOKEN

print(f"Token loaded: {AIRTABLE_TOKEN[:10]}...")

import os

from clients.airtable_client import get_records
from config import INCLUDED_STAGES

from services.report_service import build_report_data
from services.excel_service import create_workbook


def main():
    print("Connecting to Airtable...")

    records = get_records()

    print(f"Retrieved {len(records)} total records.")

    filtered_records = [
        record
        for record in records
        if record["fields"].get("Sales Stage") in INCLUDED_STAGES
    ]

    print(
        f"Found {len(filtered_records)} records in the selected sales stages."
    )

    print("Preparing report...")

    report = build_report_data(filtered_records)

    print("Creating workbook...")

    file_path = create_workbook(report)

    print("Workbook saved to:")
    print(file_path)

    print("Opening workbook...")

    os.startfile(file_path)


if __name__ == "__main__":
    main()