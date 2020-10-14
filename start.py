# _*_ coding: utf-8 _*_

from models import ExcelShell, FirstLab
from preparations import work_book, work_sheets_names


print(work_sheets_names)
user_sheet = input("\nChoose your sheet: ")

while user_sheet != "FirstLab":
    print("Unknown sheet, try again")
    user_sheet = input("Choose your sheet")

print(f"{user_sheet} was chosen successfully")


tusk1 = FirstLab(work_book=work_book, current_sheet=user_sheet)

tusk1.get_group_assessment()
tusk1.get_cons_of_exp()
print(tusk1)

