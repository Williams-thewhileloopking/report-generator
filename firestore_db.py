
from datetime import datetime, timezone

def save_report(db, report_name, reporting_month, file_name):
    db.collection("reports").add({
        "report_name": report_name,
        "reporting_month": reporting_month,
        "file_name": file_name,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "status": "Uploaded"
    })

def get_reports(db):
    reports = db.collection("reports").order_by(
        "created_at", direction="DESCENDING"
    ).stream()

    return [report.to_dict() for report in reports]
