import sys
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import colorama
from record_checker import RecordChecker

# initialize colorama
colorama.init()


class FileManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = Workbook()

    def load_file(self):
        try:
            import warnings

            warnings.simplefilter("ignore")
            workbook = load_workbook(filename=self.file_path)
            workbook = workbook['Sheet1']
            warnings.simplefilter("default")

            return workbook

        except FileNotFoundError:
            sys.exit(colorama.Fore.RED + 'File not found.')

    def check_col_JK(self):
        workbook = load_workbook(self.file_path)
        ws = workbook['Sheet1']

        # workbook = self.load_file()

        for row in range(2, self.get_max_row() + 1):
            if str(ws[f'J{row}'].value) == 'REVIEWED' or str(ws[f'J{row}'].value) == 'PENDING':
                pass
            else:
                pattern = PatternFill("solid", fgColor="FF0000")
                for cell in ws[row:row]:
                    cell.fill = pattern
        workbook.save(self.file_path)

    def get_max_row(self):
        workbook = self.load_file()
        return workbook.max_row

    def save_workbook(self, records_list):
        import warnings
        warnings.simplefilter("ignore")

        try:
            ws = load_workbook(self.file_path)
            if 'Audit' in ws.sheetnames:
                to_remove = ws.get_sheet_by_name('Audit')
                ws.remove_sheet(to_remove)
                ws.create_sheet('Audit')
                audit_sheet = ws['Audit']
            else:
                ws.create_sheet('Audit')
                audit_sheet = ws['Audit']

            audit_sheet['A1'] = 'ASIN'
            audit_sheet['B1'] = 'Subject to Clp'
            audit_sheet['C1'] = 'Is kit'
            audit_sheet['D1'] = 'eu2008_labeling_risk'
            audit_sheet['E1'] = 'eu2008_labeling_hazard'
            audit_sheet['F1'] = 'SDS'
            audit_sheet['G1'] = 'TT'
            audit_sheet['H1'] = 'Reviewer Login'
            audit_sheet['I1'] = 'Review Date'
            audit_sheet['J1'] = 'State'
            audit_sheet['K1'] = 'Utc Reason'
            audit_sheet['L1'] = 'Is Dg Utc'
            audit_sheet['M1'] = 'CLP Exemption'

            for i, row in enumerate(records_list, start=2):
                checker = RecordChecker(ws['Sheet1'], row)
                audit_sheet[f'A{i}'] = checker.col_A
                audit_sheet[f'B{i}'] = checker.col_B
                audit_sheet[f'C{i}'] = checker.col_C
                audit_sheet[f'D{i}'] = checker.col_D
                audit_sheet[f'E{i}'] = checker.col_E
                audit_sheet[f'F{i}'] = checker.col_F
                audit_sheet[f'G{i}'] = checker.col_G
                audit_sheet[f'H{i}'] = checker.col_H
                audit_sheet[f'I{i}'] = checker.col_I
                audit_sheet[f'J{i}'] = checker.col_J
                audit_sheet[f'K{i}'] = checker.col_K
                audit_sheet[f'L{i}'] = checker.col_L
                audit_sheet[f'M{i}'] = checker.col_M
            ws.save(self.file_path)
        except BaseException as e:
            print(e)
            sys.exit(colorama.Fore.RED + 'Something gone wrong during saving records to new worksheet.')
        finally:
            warnings.simplefilter("default")


if __name__ == '__main__':
    file = 'test.xlsx'
    wb = FileManager(file)
    wkb = wb.load_file()
    records = [168, 147, 144, 107, 135, 142, 248]
    wb.save_workbook(records)
