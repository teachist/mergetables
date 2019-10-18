import xlrd
from xlrd import open_workbook
from xlwt import Workbook, easyxf
import sys
import os

# start_rwo, verify_key, sheet_key, fname


def get_row(fname, index_of_book, start_row=0, verify_key=-1, sheet_key=0):
    global results
    current_file = os.path.split(fname)[1]
    # print(filename)
    if verify_key == -1:
        verify_key = 0
    sh = open_workbook(fname).sheet_by_index(sheet_key)
    nrows, ncols = sh.nrows, sh.ncols
    for rowx in range(start_row - 1, nrows):
        if sh.row_types(rowx)[verify_key] == xlrd.XL_CELL_EMPTY:
            nrows = rowx
            break
        results.append(sh.row_values(rowx) + [current_file[:4]])
        # print(sh.row_values(rowx) + [current_file[:4]])

    if current_file[2:4] == '01':
        print()
    print(
        f'\t{index_of_book+1:02}\t{current_file[:2]}级{current_file[2:4]}班\t共统计到{nrows- start_row + 1}条记录')

    return results


def create_sheet(rows, total_files, to_save_fname):
    book = Workbook()
    sheet1 = book.add_sheet('summary')

    cols = len(rows[0])

    for row in range(len(rows)):
        for col in range(cols):
            sheet1.write(row, col, rows[row][col])
    print("\n"+'='*60)
    print(f"\t本校共42个班，目前统计班级{total_files}")
    print(f"\t一共写入数据：{len(rows)}.\n")
    book.save(to_save_fname+'.xls')


def get_filelist(root_dir):
    file_list = []
    for home, dirs, files in os.walk(root_dir):
        for filename in files:
            if filename == '.DS_Store':
                os.remove(os.path.join(home, filename))
            file_list.append(os.path.join(home, filename))
    return sorted(file_list)


args = sys.argv[1:]
forder_name = args[0]
to_save_fname = args[1]
sheet_key = int(args[2])
start_row = int(args[3])
verify_key = int(args[4])

results = []
file_list = get_filelist(forder_name)
for filename in file_list:
    if os.path.isfile(filename):
        index_of_book = file_list.index(filename)
        results = get_row(filename, index_of_book, start_row=start_row,
                          verify_key=verify_key, sheet_key=sheet_key)

create_sheet(results, len(file_list), to_save_fname)
