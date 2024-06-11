import pandas as pd
import openpyxl

# ---------input số series sổ đã sử dụng từ file excel xuất từ phần mềm
data_series_uesed_file = pd.read_excel('used.xlsx', sheet_name="Sheet2", usecols="B")
series_used_file = list(data_series_uesed_file['so_stk'])  # usecols is used serials

class SeriesStk:
    def __init__(self, first2, start, end):
        self.first2 = first2
        self.start = start
        self.end = end


start_end_used_seri = [
    SeriesStk(first2="AB. 0", start=256041, end=257500),
    SeriesStk(first2="AC", start=4122001, end=4122101),
    SeriesStk(first2="AB ", start=5016501, end=5018500)
]


def make_series(start_end):
    series_stk= []
    for seri in start_end:
        amount_seri = seri.end - seri.start +1
        for i in range(amount_seri):
            series_stk.append(seri.first2 + str(seri.start))
            seri.start += 1
    return series_stk


def compare_series(list_series_1, list_series_2):  # xem series 1 trùng series 2 những số nào
    compared_series = []
    for seri in list_series_1:
        if seri not in list_series_2:
            compared_series.append(seri)
    return compared_series


def write_to_excel(sheet_name, series_name):
    workbook_excel = openpyxl.load_workbook("output.xlsx")
    worksheet = workbook_excel[sheet_name]
    last_row = worksheet.max_row
    if last_row > 1:
        rows_to_delete = range(1, last_row)  # Excludes the first row (index 0)
        # Delete the rows in reverse order to avoid shifting issues
        for row_index in sorted(rows_to_delete, reverse=True):
            worksheet.delete_rows(row_index + 1)  # Adding 1 to adjust for 0-indexing
    last_row = worksheet.max_row
    for i, j in enumerate(series_name, start=1):
        worksheet.cell(row=i + last_row, column=1, value=j)
    workbook_excel.save("output.xlsx")
    workbook_excel.close()


seri_used_start_end = make_series(start_end_used_seri)
so_hong = compare_series(seri_used_start_end, series_used_file)


write_to_excel(sheet_name='Sheet1', series_name=so_hong)
