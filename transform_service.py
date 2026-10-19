from datetime import date, datetime


def parse_date(date_string):
    """
    Converts an Airtable date string into a Python date object.
    """

    if not date_string:
        return None

    return datetime.fromisoformat(
        date_string.replace("Z", "+00:00")
    ).date()


def transform_record(record):
    """
    Converts an Airtable record into a clean report dictionary.
    """

    fields = record["fields"]

    date_created = parse_date(
        fields.get("Date Created")
    )

    expected_close_date = parse_date(
        fields.get("Expected Close Date")
    )

    today = date.today()

    forecast_value = fields.get("Forecast Value") or 0
    deposit_paid = fields.get("Deposit Paid") or 0

    balance_due = forecast_value - deposit_paid

    if (
        date_created
        and fields.get("Sales Stage") != "Approved Not Here"
    ):
        days_in_shop = (today - date_created).days
    else:
        days_in_shop = None

    if date_created and expected_close_date:
        number_of_days = (
            expected_close_date - date_created
        ).days
    else:
        number_of_days = None

    return {
        "name": fields.get("Name"),
        "customer_name": fields.get("Customer name"),
        "sales_stage": fields.get("Sales Stage"),
        "date_created": date_created,
        "expected_close_date": expected_close_date,
        "days_in_shop": days_in_shop,
        "number_of_days": number_of_days,
        "forecast_value": forecast_value,
        "deposit_paid": deposit_paid,
        "balance_due": balance_due,
    }