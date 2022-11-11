import datetime
import csv

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet import worksheet


xls_path_line_yello = "./line_yellow.xlsx"
xls_path_line_skyblue = "./line_skyblue.xlsx"

# 수인분당선
# 인천 -> 청량리 BU 행
workbook = openpyxl.load_workbook(xls_path_line_yello, read_only=True, data_only=True)
sheet = workbook["수인선 평일 상"]
date = datetime.date(1, 1, 1)
start_index = 272
with open(f"./yellow/station.csv", "w", encoding="utf-8", newline="") as f:
    writer = csv.writer(f)
    for row_index, row in enumerate(sheet.rows):
        if row_index == 3:
            station_name = row[column_index_from_string("A") - 1].value
            station_number = f"K{start_index}"
            writer.writerow([station_number, station_name, 0])
        elif row_index == 4:
            start_time = datetime.datetime.combine(date, row[column_index_from_string("BU") - 1].value)
        elif row_index > 4 and row_index % 2 == 1:
            passing_time = datetime.datetime.combine(date, row[column_index_from_string("BU") - 1].value)
            if start_index == 269:
                start_index -= 2
            else:
                start_index -= 1
            if passing_time is not None:
                station_name = row[column_index_from_string("A") - 1].value
                station_number = f"K{start_index}"
                writer.writerow([station_number, station_name, (passing_time - start_time).total_seconds() / 60])

# 4호선
# 인천 -> 청량리 BU 행
workbook = openpyxl.load_workbook(xls_path_line_skyblue, read_only=True, data_only=True)
sheet = workbook["안산과천(4호)선_평일_상행"]
date = datetime.date(1, 1, 1)
start_index = 456
with open(f"./skyblue/station.csv", "w", encoding="utf-8", newline="") as f:
    writer = csv.writer(f)
    for row_index, row in enumerate(sheet.rows):
        if row_index == 3:
            station_name = row[column_index_from_string("A") - 1].value
            station_number = f"K{start_index}"
            writer.writerow([station_number, station_name, 0])
        elif row_index == 4:
            start_time = datetime.datetime.combine(date, row[column_index_from_string("BU") - 1].value)
        elif row_index > 4 and row_index % 2 == 1:
            if type(row[column_index_from_string("BU") - 1].value) is datetime.time:
                passing_time = datetime.datetime.combine(date, row[column_index_from_string("BU") - 1].value)
                start_index -= 1
            
                station_name = row[column_index_from_string("A") - 1].value
                station_number = f"K{start_index}"
                writer.writerow([station_number, station_name, (passing_time - start_time).total_seconds() / 60])
            