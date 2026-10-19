def build_report_data(records):
    """
    Converts Airtable records into a simplified format
    for the Excel report.
    """

    report = []

    for record in records:
        fields = record["fields"]

        report.append({
            "name": fields.get("Name", ""),
            "customer_name": fields.get("Customer name", ""),
            "sales_stage": fields.get("Sales Stage", ""),
            "rep": fields.get("Technician", ""),
            "date_created": fields.get("Date Created", ""),
            "expected_close_date": fields.get("Expected Close Date", ""),
            "forecast_value": fields.get("Forecast Value", 0),
            "deposit_paid": fields.get("Deposit Paid", 0),
        })

    return report