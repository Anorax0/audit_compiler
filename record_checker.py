
class RecordChecker:
    def __init__(self, position):
        self.col_A = wb[f'A{position}'].value
        self.col_B = wb[f'B{position}'].value
        self.col_C = wb[f'C{position}'].value
        self.col_D = wb[f'D{position}'].value
        self.col_E = wb[f'E{position}'].value
        self.col_F = wb[f'F{position}'].value
        self.col_G = wb[f'G{position}'].value
        self.col_H = wb[f'H{position}'].value
        self.col_I = wb[f'I{position}'].value
        self.col_J = wb[f'J{position}'].value
        self.col_K = wb[f'K{position}'].value
        self.col_L = wb[f'L{position}'].value
        self.col_M = wb[f'M{position}'].value

    def check_category_FG(self):
        if self.col_F != '' or self.col_G != '':
            return True
        else:
            return False

    def check_category_ED(self):
        if self.col_E != '' or self.col_D != '':
            return True
        else:
            return False

    def check_category_J(self):
        if self.col_J != '':
            return True
        else:
            return False

    def check_category_M(self):
        if self.col_M != '':
            return True
        else:
            return False


if __name__ == '__main__':
    from file_manager import load_file
    file_path = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
    wb = load_file(file_path)
    for num in range(2, 30):
        test = RecordChecker(num)
        print(test.col_A, test.check_category_FG())
