import os
from xlwt import Workbook
from xlrd import open_workbook
import xlrd
from .utils import verifyIDcardValid, get_filelist, print_data
from xlutils.view import View


class Mergetable:
    def __init__(self, folder_name, output_file_name):
        self._result_rows = []  # data that store
        self._folder_name = folder_name  # process folder
        self._output_file_name = output_file_name  # output file name

        # store all the files from folder
        self._file_list = get_filelist(self._folder_name)

        self.template_detect(self._file_list[0])

        for index_of_book, filename in enumerate(self._file_list):
            if os.path.isfile(filename):
                self.get_row(filename, index_of_book)

        self.create_sheet()

    def template_detect(self, template_file):
        """template_detect(self, template_file)

        It is boring to feed too much arguments to a terminal. If it
        could automatically detect template and figure out the needed
        parameter by user, that would be an awesome work!
        """
        print('++++++ 现在是模板检测，你将回答几个简单问题 ++++++')
        try:
            sheet_key = int(input('请输入你想检测的表的序号(1-n)：')) - 1
        except ValueError:
            sheet_key = 0

        try:
            view = View(template_file)
            print_data(view[sheet_key])
        except IndexError:
            print('你输入的表序号超过最大现有表的个数，为你处理默认序号为1的表')
            sheet_key = 0
            print_data(view[sheet_key])

        try:
            header_row = int(input('请你输入标题行在模板中的行号：').strip())
            verify_key = int(input('请你输入作为模板唯一识别的列号：').strip())
            is_idno = int(input('是否进行身份证校验？（0-否 1-是）').strip())
            end_col = int(input('请你输入你想要统计的有效数据的最后列号：').strip())
            self._sheet_key = sheet_key
            self._start_row = header_row
            self._verify_key = verify_key - 1
            self._end_col = end_col
            self._is_idno = is_idno
        except ValueError:
            print('输入有误，请重新运行！')

    def get_row(self, file_name, index_of_book):
        """get_row(self, file_name, index_of_book)

        In this function, it retrive all the rows in the given file.
        It will make some clearnings and verififcations, after that
        append them to the global list.
        """
        current_file = os.path.split(file_name)[1]
        sh = open_workbook(file_name).sheet_by_index(self._sheet_key)

        nrows, ncols = sh.nrows, sh.ncols

        if index_of_book == 0:
            self.get_header_fileds(sh)
            self._start_row = self._start_row + 1

        for rowx in range(self._start_row - 1, nrows):
            if sh.row_types(rowx)[self._verify_key] == xlrd.XL_CELL_EMPTY:
                # nrows = rowx
                break
            row_data = sh.row_values(rowx, end_colx=self._end_col)
            # print(data)
            # Fill the empty cell with nothing
            if len(row_data) < self._end_col:
                empty_cells = self._end_col - len(row_data)
                row_data = row_data + ['' for _ in range(empty_cells)]

            # verify id cardNo
            if self._is_idno == 1:
                row_data[self._verify_key] = verifyIDcardValid(
                    row_data[self._verify_key])

            # print("---", row_data, len(row_data))
            # append the class identifier
            row_data.append(current_file[:4])
            # print("---", row_data)
            self._result_rows.append(row_data)
            # print("---", self._result_rows)

        if int(current_file[:4]):
            if current_file[2:4] == '01':
                print()
            print(
                f'\t{index_of_book+1:02}\t{current_file[:2]}级{current_file[2:4]}班\t共统计到{nrows- self._start_row + 1:02}条记录')
        else:
            print(
                f'\t{index_of_book+1:02}\t{current_file[:4]}\t共统计到{nrows- self._start_row + 1}条记录，请注意命名规则！！')

    def get_header_fileds(self, spreadsheet):
        row_data = spreadsheet.row_values(
            self._start_row-1, end_colx=self._end_col)
        # print(data)
        # Fill the empty cell with nothing
        if len(row_data) < self._end_col:
            empty_cells = self._end_col - len(row_data)
            row_data = row_data + ['' for _ in range(empty_cells)]

        row_data.append("标志")
        self._result_rows.append(row_data)

    def create_sheet(self):
        book = Workbook()
        sheet1 = book.add_sheet('summary')

        cols = len(self._result_rows[0])

        for row in range(len(self._result_rows)):
            for col in range(cols):
                sheet1.write(row, col, self._result_rows[row][col])

        print("\n"+'='*60)
        print(f"\t本校共42个班，目前统计班级{len(self._file_list)}个。")
        print(f"\t一共写入数据：{len(self._result_rows)-1}条数据.\n")

        book.save(self._output_file_name+'.xls')
