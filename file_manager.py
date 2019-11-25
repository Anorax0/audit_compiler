import sys
from openpyxl import load_workbook
from colorama import Fore
from colorama import init

# initialize colorama
init()


def load_file(file_path):
    try:
        import warnings

        warnings.simplefilter("ignore")
        workbook = load_workbook(filename=file_path, read_only=True)
        workbook = workbook['Sheet1']
        warnings.simplefilter("default")

        return workbook

    except FileNotFoundError:
        sys.exit(Fore.RED + 'File not found.')


if __name__ == '__main__':
    file = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
    load_file(file)
