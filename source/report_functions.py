import os
from datetime import date
import pandas as pd
from pandas import DataFrame

from source.credentials import (
    connect_to_database,
    department_map,
    request,
    columns_rename
)

# Getting an engine from the successfully established
# connection to the database
engine = connect_to_database()


# The function which do the replacement of the departments names from the dict
def department_replacements(report) -> DataFrame:
    report["department"] = report["department"].replace(department_map)

    return report


# The function which calculates users inside the current department
def department_counters(report) -> DataFrame:
    depart_counts = report["department"].value_counts()
    report["count"] = report["department"].map(depart_counts)

    return report


# The function which saves the DataFrame table to the Excel-file
def save_report_to_excel(report) -> None:
    os.makedirs("reports", exist_ok=True)
    report.to_excel(
        f"./reports/Безпечне місто {date.today()}.xlsx",
        index=False,
        sheet_name="Користувачі"
    )


# The main function which calls the functions above and edit
# the raw data taken from the DB and save it to the .xlsx-file
def report_merge() -> None:
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
