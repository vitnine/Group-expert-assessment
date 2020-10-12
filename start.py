# _*_ coding: utf-8 _*_

from models import ExcelShell
from preparations import work_book, work_sheets_names



# user_sheet = input("\nChoose your sheet: ")
#
# while user_sheet not in work_sheets_names:
#     print("Unknown sheet, try again")
#     user_sheet = input("Choose your sheet")
#
# print(f"{user_sheet} was chosen successfully")

sheet_name = "FirstLab"
wb = ExcelShell(work_book=work_book, current_sheet=sheet_name, sheet_names=work_sheets_names)
wb.get_matrix()
matrix = wb.matrix
print(matrix)
wb.get_max_in_lines()
wb.reduce_lines()
wb.get_group_assessment(border=0.001)
print(wb.start_matrix)


# wb.get_group_assessment(stop_border=0.001)

