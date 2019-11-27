import sys
from time import time
from itertools import chain
from random import choice
from file_manager import FileManager
from record_checker import RecordChecker

start_time = time()

# if len(sys.argv) != 3:
#     sys.exit('File path or person name is empty. Please provide complete data.')
# else:
#     file_path = sys.argv[1]
#     person_to_audit = sys.argv[2]

file_path = 'Classifications 2019-11-19 23-00-00_2019-11-20 22-59-59.xlsx'
person_to_audit = 'kozimalg'

input_percent = 0.2

records_length = 0

workbook = FileManager(file_path)
wb = workbook.load_file()

for row in range(2, workbook.get_max_row()+1):
    record = RecordChecker(wb, row)
    if record.col_H == person_to_audit:
        records_length += 1

# do math.ceil(number) to round up the float to whole integer
records_length_percent = round(records_length * input_percent)
category_ED_percent = round(records_length_percent * 0.4)
category_FG_percent = round(records_length_percent * 0.3)
category_J_percent = round(records_length_percent * 0.2)
category_M_percent = round(records_length_percent * 0.1)

print('to check ED', category_ED_percent)
print('to check FG', category_FG_percent)
print('to check M', category_M_percent)

category_ED = []
category_FG = []
category_J = []
category_M = []

to_check_category_ED = []
to_check_category_FG = []
to_check_category_J = []
to_check_category_M = []
all_categories = category_ED, category_FG, category_J, category_M

for row in range(2, workbook.get_max_row()+1):
    record = RecordChecker(wb, row)
    if record.col_H == person_to_audit:
        if record.check_category_ED() and row not in chain(*all_categories):
            category_ED.append(row)

        if record.check_category_FG() and row not in chain(*all_categories):
            category_FG.append(row)

        # if record.check_category_J() and row not in chain(*all_categories):
        #     category_J.append(row)
        if record.check_category_M() and row not in chain(*all_categories):
            category_M.append(row)

if category_ED_percent > len(category_ED):
    category_FG_percent += category_ED_percent-len(category_ED)
if category_FG_percent > len(category_FG):
    category_M_percent += category_FG_percent-len(category_FG)

print('to check ED', category_ED_percent)
print('to check FG', category_FG_percent)
print('to check M', category_M_percent)

while category_ED_percent != 0:
    if len(category_ED) == 0:
        break
    if len(to_check_category_ED) == category_ED_percent:
        break
    record_to_add = choice(category_ED)
    if record_to_add not in to_check_category_ED:
        to_check_category_ED.append(record_to_add)
        category_ED_percent -= 1

while category_FG_percent != 0:
    if len(category_FG) == 0:
        break
    if len(to_check_category_FG) == category_FG_percent:
        break
    record_to_add = choice(category_FG)
    if record_to_add not in to_check_category_FG:
        to_check_category_FG.append(record_to_add)
        category_FG_percent -= 1

while category_M_percent != 0:
    if len(category_M) == 0:
        break
    if len(to_check_category_M) == category_M_percent:
        break
    record_to_add = choice(category_M)
    if record_to_add not in to_check_category_M:
        to_check_category_M.append(record_to_add)
        category_M_percent -= 1

# print(records_length_percent)
print('entire ED', len(category_ED))
print('entire FG', len(category_FG))
# # print(category_J)
print('entire M', len(category_M))
#
# print()
#
print('ED-', len(to_check_category_ED), to_check_category_ED)
print('FG-', len(to_check_category_FG), to_check_category_FG)
# # print(category_J)
print('M-', len(to_check_category_M), to_check_category_M)

print(f'\n Executing time: { (time() - start_time):.3f}s')

# input('Press Enter to close.')
