# config.py

import os

from dotenv import load_dotenv

load_dotenv()

AIRTABLE_TOKEN = os.getenv("AIRTABLE_TOKEN")
AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID")
AIRTABLE_TABLE_NAME = os.getenv("AIRTABLE_TABLE_NAME")

from pathlib import Path

REPORT_OUTPUT_FOLDER = Path(
    os.getenv("REPORT_OUTPUT_FOLDER", Path.home() / "Downloads" / "CBR Reports")
)

INCLUDED_STAGES = [
    "Approved",
    "Approved Not Here",
    "Strip - outside supplier",
    "Additional Damages/Inspect",
    "Carbon Repairs",
    "Resprays",
    "Restores / Block Sprays",
    "QC",
    "Rebuild",
    "Invoiced-Stock here",
]

REPORT_FILENAME = "CBR Analysis - {date}.xlsx"

REPORT_COLUMNS = [
    "Name",
    "Customer Name",
    "Sales Stage",
    "Date Created",
    "Expected Close Date",
    "Days in Shop",
    "Number of Days",
    "Forecast Value",
    "Deposit Paid",
    "Balance Due",
]

