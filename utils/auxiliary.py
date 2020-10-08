from os.path import join, abspath
from content import first_lab


class CorrectPath:
    matrix_path = abspath(join('.', first_lab.DataPath.matrix_path))
