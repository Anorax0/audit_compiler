import sys
from openpyxl import load_workbook
from colorama import Fore
import colorama

# initialize colorama
colorama.init()


class FileManager:
    def __init__(self, file_path):
        self.file_path = file_path

    def load_file(self):
        try:
            import warnings

            warnings.simplefilter("ignore")
            workbook = load_workbook(filename=self.file_path)
            workbook = workbook['Sheet1']
            warnings.simplefilter("default")

            return workbook

        except FileNotFoundError:
            sys.exit(Fore.RED + 'File not found.')

    def get_max_row(self):
        workbook = self.load_file()
        return workbook.max_row


if __name__ == '__main__':
    file = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
    wb = FileManager(file)
    print(wb.get_max_row())
