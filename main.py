from openpyxl import load_workbook

wb = load_workbook("order_agas_may_june.xlsx")
print(wb.get_sheet_names())