import csv
import json
import os
from typing import Optional


import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet import worksheet


xls_path_line_yello = "./line_yellow.xlsx"
xls_path_line_skyblue = "./line_skyblue.xlsx"

def create_timetable(excel_path: str, timetable_path: str, 
        sheet_weekdays_up: str, sheet_weekdays_down: str, sheet_weekends_up: str, sheet_weekends_down: str, 
        row_weekdays_up: int, row_weekdays_down: int, row_weekends_up: int, row_weekends_down: int,
        ) -> None:
    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    os.makedirs(f"{timetable_path}/weekdays", exist_ok=True)
    with open(f"{timetable_path}/weekdays/up.csv", "w") as f:
        writer = csv.writer(f)
        writer.writerows(parse_excel_timetable_vertical(workbook[sheet_weekdays_up], row_weekdays_up))
    with open(f"{timetable_path}/weekdays/down.csv", "w") as f:
        writer = csv.writer(f)
        writer.writerows(parse_excel_timetable_vertical(workbook[sheet_weekdays_down], row_weekdays_down))

    os.makedirs(f"{timetable_path}/weekends", exist_ok=True)
    with open(f"{timetable_path}/weekends/up.csv", "w") as f:
        writer = csv.writer(f)
        writer.writerows(parse_excel_timetable_vertical(workbook[sheet_weekends_up], row_weekends_up))
    with open(f"{timetable_path}/weekends/down.csv", "w") as f:
        writer = csv.writer(f)
        writer.writerows(parse_excel_timetable_vertical(workbook[sheet_weekends_down], row_weekends_down))
    workbook.close()


def parse_excel_timetable_horizontal(sheet: worksheet, column: str) -> list:
    heading_column = "B"
    result = []

    for row_index, row in enumerate(worksheet.rows):
        if row_index > 2:
            if row[column_index_from_string(column) - 1].value and str(row[column_index_from_string(column) - 1].value).strip():
                result.append([row[column_index_from_string(heading_column) - 1].value, str(row[column_index_from_string(column) - 1].value)])
    return result


def parse_excel_timetable_vertical(sheet: worksheet, station_row: int) -> list:
    heading_row = 2
    result = []

    for row_index, row in enumerate(sheet.rows):
        if row_index == heading_row - 1:
            heading_data = [cell.value for cell in row]
        elif row_index == station_row - 1:
            for cell_index, cell in enumerate(row):
                if cell.value is not None and ":" in str(cell.value):
                    result.append([heading_data[cell_index], cell.value])
    return result


create_timetable(xls_path_line_skyblue, "./skyblue", 
    "4호선(평일-상행)", "4호선(평일-하행)", "4호선(휴일-상행)", "4호선(휴일-하행)", 
    row_weekdays_up=19, row_weekdays_down=85, row_weekends_up=19, row_weekends_down=85
)

create_timetable(xls_path_line_yello, "./yellow", 
    "수인분당(평일-상행)", "수인분당(평일-하행)", "수인분당(휴일-상행)", "수인분당(휴일-하행)", 
    row_weekdays_up=45, row_weekdays_down=89, row_weekends_up=45, row_weekends_down=89
)