import sys
from time import time
import openpyxl

from file_manager import load_file

start_time = time()

if len(sys.argv) != 3:
    sys.exit('File path or person name is empty. Please provide complete data.')
else:
    file_path = sys.argv[1]
    person_to_audit = sys.argv[2]

# file_path = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
# person_to_audit = 'garbud'

wb = load_file(file_path)

records_list = []

category_one = []
category_two = []
category_three = []
category_four = []


print(len(records_list))

print(f'\nExecuting time: {time() - start_time:.3f}s')

input('Press Enter to close.')
