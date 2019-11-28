class RecordChecker:
    def __init__(self, workbook, position):
        self.wb = workbook
        self.col_A = self.wb[f'A{position}'].value
        self.col_B = self.wb[f'B{position}'].value
        self.col_C = self.wb[f'C{position}'].value
        self.col_D = self.wb[f'D{position}'].value
        self.col_E = self.wb[f'E{position}'].value
        self.col_F = self.wb[f'F{position}'].value
        self.col_G = self.wb[f'G{position}'].value
        self.col_H = self.wb[f'H{position}'].value
        self.col_I = self.wb[f'I{position}'].value
        self.col_J = self.wb[f'J{position}'].value
        self.col_K = self.wb[f'K{position}'].value
        self.col_L = self.wb[f'L{position}'].value
        self.col_M = self.wb[f'M{position}'].value

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
        if self.col_J != '' or self.col_J == 'PENDING' or self.col_J == 'SUBMITTED':
            return True
        else:
            return False

    def check_category_M(self):
        if self.col_M != '' and self.col_M != 'None':
            return True
        else:
            return False


if __name__ == '__main__':
    from file_manager import FileManager
    file_path = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
    wb = FileManager(file_path).load_file()
    for num in range(2, 20):
        test = RecordChecker(wb, num)
        print(test.col_A, test.check_category_FG(), test.check_category_ED(),
              test.check_category_J(), test.check_category_M())
