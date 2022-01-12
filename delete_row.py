import xlwings as xw
from xlwings.constants import DeleteShiftDirection

wb = xw.Book('workbook.xlsx')
sheet = wb.sheets['miir_pptk_stat_col']
sheet_actual_columns = wb.sheets['actual_columns']
rng_columns = sheet_actual_columns.range((1, 'A'), (7026, 'M')).value

sheet_actual_columns.range('20:20').api.Delete()
