import sys
from time import time

from record_checker import RecordChecker
from file_manager import FileManager

start_time = time()

if len(sys.argv) != 3:
    sys.exit('File path or person name is empty. Please provide complete data.')
else:
    file_path = sys.argv[1]
    person_to_audit = sys.argv[2]

# file_path = 'Classifications 2019-11-18 23-00-00_2019-11-19 22-59-59.xlsx'
# person_to_audit = 'garbud'

wb = FileManager(file_path).load_file()

print(f'\nExecuting time: {time() - start_time:.3f}s')

input('Press Enter to close.')
