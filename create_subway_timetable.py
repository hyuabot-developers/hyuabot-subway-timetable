import csv

import openpyxl
from openpyxl.worksheet import worksheet

xls_path_line_yello = "./1071.xlsx"
xls_path_line_skyblue = "./1004.xlsx"


def create_timetable(
        excel_path: str, route_id: int, sheet_weekdays_up: str, sheet_weekdays_down: str, sheet_weekends_up: str,
        sheet_weekends_down: str) -> None:
    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    rows = []
    for station_name, timetable in parse_data(workbook[sheet_weekdays_up]).items():
        for time in timetable:
            rows.append([station_name, "weekdays", "up",
                         time["start"], time["terminal"], time["time"].strftime("%H:%M:%S")])
    for station_name, timetable in parse_data(workbook[sheet_weekdays_down]).items():
        for time in timetable:
            rows.append([station_name, "weekdays", "down",
                         time["start"], time["terminal"], time["time"].strftime("%H:%M:%S")])
    for station_name, timetable in parse_data(workbook[sheet_weekends_up]).items():
        for time in timetable:
            rows.append([station_name, "weekends", "up",
                         time["start"], time["terminal"], time["time"].strftime("%H:%M:%S")])
    for station_name, timetable in parse_data(workbook[sheet_weekends_down]).items():
        for time in timetable:
            rows.append([station_name, "weekends", "down",
                         time["start"], time["terminal"], time["time"].strftime("%H:%M:%S")])
    with open(f"./{route_id}.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    workbook.close()


def parse_data(sheet: worksheet):
    start_list = []
    heading_list = []
    station_list = []
    timetable_list = []
    timetable_dict = {}
    for row_index, row in enumerate(sheet.rows):
        if row_index == 0:
            start_list = [cell.value for cell in row[1:]]
        elif row_index == 1:
            heading_list = [cell.value for cell in row[1:]]
        elif row_index > 2:
            if row_index % 2 == 1:
                station_list.append(row[0].value)
            else:
                timetable_list.append([cell.value for cell in row[1:]])
    # print(timetable_list)
    for station_index, station_name in enumerate(station_list):
        timetable_dict[station_name] = []
        for timetable_index, timetable in enumerate(timetable_list[station_index]):
            if str(timetable).strip() and timetable:
                timetable_dict[station_name].append({
                    "time": timetable,
                    "start": start_list[timetable_index],
                    "terminal": heading_list[timetable_index],
                })
    return timetable_dict


create_timetable(xls_path_line_skyblue, 1004,
                 "안산과천(4호)선_평일_상행", "안산과천(4호)선_평일_하행", "안산과천(4호)선_휴일_상행", "안산과천(4호)선_휴일_하행")
create_timetable(xls_path_line_yello, 1071,
                 "수인선 평일 상", "수인선 평일 하", "수인선 휴일 상", "수인선 휴일 하")
