class RecordChecker:
    def __init__(self, workbook, position):
        self.wb = workbook
        self.col_a = self.wb[f'A{position}'].value
        self.col_b = self.wb[f'B{position}'].value
        self.col_c = self.wb[f'C{position}'].value
        self.col_d = self.wb[f'D{position}'].value
        self.col_e = self.wb[f'E{position}'].value
        self.col_f = self.wb[f'F{position}'].value
        self.col_g = self.wb[f'G{position}'].value
        self.col_h = self.wb[f'H{position}'].value
        self.col_i = self.wb[f'I{position}'].value
        self.col_j = self.wb[f'J{position}'].value
        self.col_k = self.wb[f'K{position}'].value
        self.col_l = self.wb[f'L{position}'].value
        self.col_m = self.wb[f'M{position}'].value

    def check_category_fg(self):
        if self.col_f != '' or self.col_g != '':
            return True
        else:
            return False

    def check_category_ed(self):
        if self.col_e != '' or self.col_d != '':
            return True
        else:
            return False

    def check_category_j(self):
        if self.col_j != '' or self.col_j == 'PENDING' or self.col_j == 'SUBMITTED':
            return True
        else:
            return False

    def check_category_m(self):
        if self.col_m != '' and self.col_m != 'None':
            return True
        else:
            return False


if __name__ == '__main__':
    from file_manager import FileManager
    file_path = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
    wb = FileManager(file_path).load_file()
    for num in range(2, 20):
        test = RecordChecker(wb, num)
        print(test.col_a, test.check_category_fg(), test.check_category_ed(),
              test.check_category_j(), test.check_category_m())
