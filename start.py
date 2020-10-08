# _*_ coding: utf-8 _*_

from models import MatrixShell
from preparations import work_book, work_sheets_names

wb = MatrixShell(work_book=work_book, sheet_names=work_sheets_names)

# user_sheet = input("\nChoose your sheet: ")
#
# while user_sheet not in work_sheets_names:
#     print("Unknown sheet, try again")
#     user_sheet = input("Choose your sheet")
#
# print(f"{user_sheet} was chosen successfully")

sheet_name = "FirstLab"
matrix = wb.get_matrix(sheet_name=sheet_name)
print(wb.sheet_names)
print(matrix)


