# _*_ coding: utf-8 _*_

from models import ExcelShell, FirstLab
from preparations import work_book, work_sheets_names



# user_sheet = input("\nChoose your sheet: ")
#
# while user_sheet not in work_sheets_names:
#     print("Unknown sheet, try again")
#     user_sheet = input("Choose your sheet")
#
# print(f"{user_sheet} was chosen successfully")

sheet_name = "FirstLab"
wb = FirstLab(work_book=work_book, current_sheet=sheet_name)

matrix = getattr(wb, "matrix")
print(matrix)

wb.get_group_assessment(border=0.001)

print(matrix)

wb2 = FirstLab(work_book=work_book, current_sheet=sheet_name)

matrix2 = wb2.get_matrix()
print(matrix2)

wb2.get_group_assessment(border=0.001)

print(matrix2)
# print('\n')
# print(wb.start_matrix)
# print(wb.matrix)
# print(wb.Ki)
# print(wb.delta_Ki)
# print(wb.lamb)
# print(wb.Xj)
# print('\n')




# wb.get_group_assessment(stop_border=0.001)

