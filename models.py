from typing import List
import addict


class ExcelShell:
    def __init__(
            self,
            work_book,
            current_sheet: str,
    ) -> None:
        self.work_book = work_book
        self.sheet_names = work_book.sheetnames
        self.current_sheet = current_sheet
        self.ws = addict.Dict(**{name: self.work_book[name] for name in self.sheet_names})
        self.start_matrix = []
        self.matrix = []
        self.max = []

    def get_matrix(self):
        if not self.matrix:
            self.matrix = []
            self.start_matrix = []
            for i, row in enumerate(self.ws[self.current_sheet].rows):
                if i == 0:
                    continue
                self.matrix.append([])
                self.start_matrix.append([])
                for j in range(len(row)):
                    if j == 0:
                        continue
                    self.matrix[i - 1].append(row[j].value)
                    self.start_matrix[i - 1].append(row[j].value)

        return self.matrix

    def get_max_in_lines(self):
        if not self.max:
            for i, row in enumerate(self.matrix):
                self.max.append(self.matrix[i][0])
                for j, col in enumerate(row):
                    if self.matrix[i][j] >= self.max[i]:
                        self.max[i] = self.matrix[i][j]

        return self.max

    def reduce_lines(self):
        for i, row in enumerate(self.matrix):
            for j, col in enumerate(row):
                self.matrix[i][j] = self.matrix[i][j] / self.max[i]

        return self.matrix


class FirstLab(ExcelShell):
    def __init__(self, work_book, current_sheet):
        ExcelShell.__init__(self, work_book=work_book, current_sheet=current_sheet)
        self.Ki = []
        self.Xj = []
        self.delta_Ki = [1]
        self.lamb = []

    def find_Ki(self, t):
        self.Ki.append([])
        lines_amount = len(self.matrix)
        if t == 0:
            for i in range(lines_amount):
                self.Ki[t].append(1 / lines_amount)
        else:
            for i in range(lines_amount):
                result = 0
                for j in range(len(self.matrix[i])):
                    result += (self.Xj[t - 1][j] * self.matrix[i][j]) / self.lamb[t - 1]
                self.Ki[t].append(result)

            self.delta_Ki.append(abs(self.Ki[t][0] - self.Ki[t - 1][0]))

        return self.Ki

    def find_lambda(self, t):
        lamb = 0
        for i in range(len(self.matrix)):
            for j in range(len(self.matrix[0])):
                lamb += self.matrix[i][j] * self.Xj[t][j]

        self.lamb.append(round(lamb, 10))
        return self.lamb

    def find_Xj(self, t):
        self.Xj.append([])
        for j in range(len(self.matrix[0])):
            result = 0
            for i in range(len(self.matrix)):
                result += (self.Ki[t][i] * self.matrix[i][j])

            self.Xj[t].append(round(result, 10))
        return self.Xj

    def get_group_assessment(self, border=0):
        self.get_matrix()
        self.get_max_in_lines()
        self.reduce_lines()
        t = 0
        while self.delta_Ki[-1] > border or self.delta_Ki[-1] == self.delta_Ki[-2]:
            Ki = self.find_Ki(t=t)
            Xj = self.find_Xj(t=t)
            lamb = self.find_lambda(t=t)
            t += 1

#     def __str__(self):
#         start_matrix = ""
#         current_matrix = ""
#
#         for row in self.start_matrix:
#             start_matrix += str(row) + "\n"
#
#         for row in self.matrix:
#             current_matrix += str(row) + "\n"
#
#         for iter in range(len(self.delta_Ki)-1):
#             pass
#
#
#         wrapper = f"""
# Название листа {self.current_sheet}
#
# Введённая матрица
# {start_matrix}
# """
#         body = f"""
# Сокращенная матрица
# {current_matrix}
# """
#         return wrapper + body