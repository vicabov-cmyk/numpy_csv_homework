from pathlib import Path

import numpy as np
import pandas as pd

csv_path = Path(__file__).resolve().parent / "data" / "employees.csv"

employee_dtype = np.dtype([
    ("id", np.int64),
    ("age", np.int64),
    ("experience", np.float64),
    ("salary", np.float64),
    ("department_id", np.int64),
    ("projects_count", np.int64),
    ("rating", np.float64),
    ("bonus", np.float64),
])

employees = np.loadtxt(
    csv_path,
    delimiter=",",
    skiprows=1,
    dtype=employee_dtype,
    ndmin=1,
)

print("Кількість рядків:", len(employees))
print(employees)
columns = [name for name in employees.dtype.names if name != "department_id"]
processed_employees = employees[columns].copy()

print("Кількість стовпців після обробки:", len(columns))
output_dir = Path(__file__).resolve().parent / "output"
output_dir.mkdir(exist_ok=True)
output_csv = output_dir / "employees_processed.csv"

np.savetxt(
    output_csv,
    processed_employees,
    delimiter=";",
    header=";".join(columns),
    comments="",
    fmt=["%d", "%d", "%.1f", "%.2f", "%d", "%.1f", "%.2f"],
)

saved_employees = np.loadtxt(
    output_csv,
    delimiter=";",
    skiprows=1,
    dtype=processed_employees.dtype,
    ndmin=1,
)

print("Рядків після повторного читання CSV:", len(saved_employees))
output_excel = output_dir / "employees_processed.xlsx"
table = pd.DataFrame.from_records(processed_employees)
excel_formats = ["0", "0", "0.0", "0.00", "0", "0.0", "0.00"]

with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    table.to_excel(writer, sheet_name="Employees", index=False)
    sheet = writer.sheets["Employees"]

    for cells, number_format in zip(sheet.iter_cols(min_row=2), excel_formats):
        for cell in cells:
            cell.number_format = number_format

print("Файл Excel збережено:", output_excel.name)