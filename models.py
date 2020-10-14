import addict
from scipy.stats import rankdata


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
                if i == 0 or i == 1:
                    continue
                self.matrix.append([])
                self.start_matrix.append([])
                for j in range(len(row)):
                    if j == 0:
                        continue
                    self.matrix[i - 2].append(row[j].value)
                    self.start_matrix[i - 2].append(row[j].value)

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

    def list_sum(self, list):
        sum = 0
        for i in list:
            sum += i
        return sum


class FirstLab(ExcelShell):
    def __init__(self, work_book, current_sheet):
        ExcelShell.__init__(self, work_book=work_book, current_sheet=current_sheet)
        self.Ki = []
        self.Xj = []
        self.delta_Ki = [1]
        self.lamb = []
        self.rang_matrix = []
        self.rang_sum = []
        self.deviation = []
        self.average = []
        self.squared_deviation = []
        self.consistency_of_experts = None
        self.sum_Ti = None
        self.sum_sqr_deviation = None

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

    def average_deviation(self):
        if not self.deviation:
            deviation = []
            sqr_deviation = []

            average = self.list_sum(self.rang_sum)/len(self.matrix[0])
            self.average = average
            for el in self.rang_sum:
                deviation.append(el-average)
                sqr_deviation.append((el-average)**2)

            self.deviation = deviation
            self.squared_deviation = sqr_deviation
            self.sum_sqr_deviation = self.list_sum(sqr_deviation)

    def find_cons(self):
        if self.consistency_of_experts is None:
            const = 1/12
            m = len(self.rang_matrix)
            n = len(self.rang_matrix[0])
            S = self.sum_sqr_deviation
            Ti = self.sum_Ti
            W = S / (const*(m**2)*(n**3 - n)-m * Ti)
            self.consistency_of_experts = W
        return self.consistency_of_experts

    def find_Ti(self):
        if self.sum_Ti is None:
            Ti = []
            Hi = []
            count = 0
            for k in range(len(self.rang_matrix)):
                Hi.append([])
                for i in range(len(self.rang_matrix[k])):
                    for el in self.rang_matrix[k][i::-1]:
                        if el in self.rang_matrix[k][i + 1::] and el not in Hi[k]:
                            Hi[k].append(el)
                sum = 0
                for i in range(len(Hi[k])):
                    h = 0
                    for el in self.rang_matrix[k]:
                        if el == Hi[k][i]:
                            h += 1
                    sum += (h ** 3 - h)
                Ti.append(sum)

            sum_Ti = self.list_sum(Ti)/12
            self.sum_Ti = sum_Ti
        return self.sum_Ti

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

    def get_rang_matrix(self):
        if not self.rang_matrix:
            array = []
            rang_sum = []
            for i in range(len(self.start_matrix)):
                array.append(self.start_matrix[i])
                array[i] = rankdata(array[i])
            self.rang_matrix = array

        return self.rang_matrix

    def get_rang_sum(self):
        if not self.rang_sum:
            for j in range(len(self.rang_matrix[0])):
                result = 0
                for i in range(len(self.rang_matrix)):
                    result += (self.rang_matrix[i][j])
                self.rang_sum.append(result)

        return self.rang_sum

    def get_cons_of_exp(self):
        if self.consistency_of_experts is None:
            self.get_matrix()
            self.get_rang_matrix()
            self.get_rang_sum()
            self.average_deviation()
            self.find_Ti()
            self.find_cons()

        return self.consistency_of_experts

    def __str__(self):
        start_matrix = ""
        current_matrix = ""
        group_assessment = "\n------------------------------------" \
                           "\nГРУППОВАЯ ОЦЕНКА КАЖДОЙ АЛЬТЕРНАТИВЫ" \
                           "\n------------------------------------\n"
        rang_matrix = ""
        matrix = []
        for row in self.start_matrix:
            start_matrix += str(row) + "\n"

        for row in self.matrix:
            current_matrix += str(row) + "\n"

        for i in range(len(self.Ki)):
            group_assessment += f"""
Номер итерации, t = {i}
Компетентность, Ki = {self.Ki[i]}
Лямбда, λ = {self.lamb[i]}
Разность компетенций, Δ = {self.delta_Ki[i]}
Результат оценки, Xj = {self.Xj[i]}
"""
        for row in range(len(self.rang_matrix)):
            matrix.append([])
            for el in range(len(self.rang_matrix[row])):
                matrix[row].append(self.rang_matrix[row][el])
            rang_matrix += str(matrix[row]) + "\n"



        wrapper = f"""
Название листа {self.current_sheet}

Введённая матрица
{start_matrix}
Сокращенная матрица
{current_matrix}"""

        consistency_of_experts = f"""
--------------------------------
СОГЛАСОВАННОСТЬ МНЕНИЙ ЭКСПЕРТОВ
--------------------------------

Матрица рангов
{rang_matrix}

Сумма рангов экспертов
{self.rang_sum}

Средний ранг экспертов
{self.average}

Квадраты отклонения рангов
{self.squared_deviation}

Сумма квадратов отклонения рангов
{self.sum_sqr_deviation}

Сумма показателей связанных рангов
{self.sum_Ti}

Коэффициента конкордации
{self.consistency_of_experts}"""

        return wrapper + group_assessment + consistency_of_experts
