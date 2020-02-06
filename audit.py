import sys
from time import time
from itertools import chain
from random import choice
from progress.bar import Bar
from file_manager import FileManager
from record_checker import RecordChecker
import colorama

# initialize colorama
colorama.init()

start_time = time()

# if len(sys.argv) < 3:
#     sys.exit('File path or person name is empty. Please provide complete data.')
# else:
#     file_path = sys.argv[1]
#     person_to_audit = sys.argv[2]
#     input_percent = 0.2
#     try:
#         if sys.argv[3]:
#             input_percent = float(sys.argv[3])
#     except IndexError:
#         pass

file_path = 'test.xlsx'
input_percent = 0.2
person_to_audit = 'srodbeat'

if input_percent > 1:
    input_percent = 1
if input_percent < 0.1:
    input_percent = 0.1

records_length = 0

workbook = FileManager(file_path)
workbook.check_col_JK()

wb = workbook.load_file()

bar_row_length = Bar('Setting length of workbook', max=workbook.get_max_row())
for row in range(2, workbook.get_max_row()+1):
    record = RecordChecker(wb, row)
    bar_row_length.next()
    if record.col_H == person_to_audit:
        records_length += 1
bar_row_length.next()
bar_row_length.finish()

print('all records', records_length)

# do math.ceil(number) to round up the float to whole integer
records_length_percent = round(records_length * input_percent)
category_ED_percent = round(records_length_percent * 0.4)
category_FG_percent = round(records_length_percent * 0.3)
category_J_percent = round(records_length_percent * 0.2)
category_M_percent = round(records_length_percent * 0.1)

print('all', records_length_percent)
print('ed: ', category_ED_percent)
print('fg: ', category_FG_percent)
print('j: ', category_J_percent)
print('m: ', category_M_percent)

category_ED = []
category_FG = []
# category_J = []
category_M = []

to_check_category_ED = []
to_check_category_FG = []
# to_check_category_J = []
to_check_category_M = []
# all_categories = category_ED, category_FG, category_M

bar_record_gathering = Bar('Gathering list of user\'s records', max=records_length)
for row in range(2, workbook.get_max_row()+1):
    record = RecordChecker(wb, row)
    if record.col_H == person_to_audit:
        # if record.check_category_ED() and row not in chain(*all_categories):
        if record.check_category_ed():
            category_ED.append(row)
        if record.check_category_fg():
            category_FG.append(row)
        # if record.check_category_J() and row not in chain(*all_categories):
        #     category_J.append(row)
        if record.check_category_m():
            category_M.append(row)
        bar_record_gathering.next()
bar_record_gathering.finish()

print('ed ', category_ED)
print('fg ', category_FG)
# category_J = []
print('m ', category_M)

if category_ED_percent > len(category_ED):
    category_FG_percent += category_ED_percent-len(category_ED)
    category_ED_percent = len(category_ED)
if category_FG_percent > len(category_FG):
    category_M_percent += category_FG_percent-len(category_FG)
    category_FG_percent = len(category_FG)

all_categories_to_check = []

bar_record_randomize = Bar('Randomizing user\'s records', max=records_length_percent)
while category_ED_percent != 0:
    if len(category_ED) == 0:
        break
    # if len(to_check_category_ED) == category_ED_percent:
    #     break
    record_to_add = choice(category_ED)
    if record_to_add not in to_check_category_ED and record_to_add not in all_categories_to_check:
        to_check_category_ED.append(record_to_add)
        all_categories_to_check.append(record_to_add)
        category_ED_percent -= 1
        bar_record_randomize.next()

print('----', category_FG_percent)

while category_FG_percent != 0:
    if len(category_FG) == 0:
        break
    # if len(to_check_category_FG) == category_FG_percent:
    #     break
    record_to_add = choice(category_FG)
    if record_to_add not in to_check_category_FG and record_to_add not in all_categories_to_check:
        to_check_category_FG.append(record_to_add)
        all_categories_to_check.append(record_to_add)
        category_FG_percent -= 1
        bar_record_randomize.next()

print('-----', len(to_check_category_FG), category_FG_percent)
while category_M_percent != 0:
    if len(category_M) == 0:
        break
    # if len(to_check_category_M) == category_M_percent:
    #     break
    record_to_add = choice(category_M)
    if record_to_add not in to_check_category_M and record_to_add not in all_categories_to_check:
        to_check_category_M.append(record_to_add)
        all_categories_to_check.append(record_to_add)
        category_M_percent -= 1
        bar_record_randomize.next()

bar_record_randomize.finish()
if bar_record_randomize.max > bar_record_randomize.index:
    print(colorama.Fore.RED + 'Cannot find enough records to fill criteria.')
    print(colorama.Style.RESET_ALL, end='')

# print(to_check_category_ED, to_check_category_FG, to_check_category_M)
print(len(all_categories_to_check))
print('Rows with choosen records to check: ')
print('ED rows: ', to_check_category_ED)
print('FG rows: ', to_check_category_FG)
print('M rows :', to_check_category_M)


records_to_check_list = [i for i in chain(to_check_category_ED, to_check_category_FG, to_check_category_M)]
workbook.save_workbook(records_to_check_list)

print(f'\n Executing time: { (time() - start_time):.3f}s')

# input('Press Enter to close.')
