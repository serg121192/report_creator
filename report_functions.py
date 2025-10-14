from datetime import date
import pandas as pd

from credentials import (
    connection,
    department_map,
    request,
    columns_rename
)

# Making an engine from the connection request
engine = connection()


# The function which do the replacement of the departments names from the dict
def department_replacements(_report):
    _report["department"] = _report["department"].replace(department_map)

    return _report


# The function which calculates users inside the exact department
def department_counters(_report):
    depart_counts = _report["department"].value_counts()
    _report["count"] = _report["department"].map(depart_counts)

    return _report


# The functions which saves the DataFrame table to the Excel-file
def save_report_to_excel(_report):
    _report.to_excel(
        f"Безпечне місто {date.today()}.xlsx",
        index=False,
        sheet_name="Користувачі"
    )


# The main function which edit the raw data taken from the DB
# and save it to the .xlsx-file
def report_merge():
    report = pd.read_sql(request, con=engine)

    report = report.rename(columns={"title": "department"})
    report = department_replacements(report)
    report = department_counters(report)
    report.insert(
        report.columns.get_loc("count") + 1,
        "security_level", "Звичайний"
    )
    save_report_to_excel(
        report[
            [
                "department",
                "count",
                "username",
                "security_level",
                "name"
            ]
        ].rename(columns=columns_rename)
    )
