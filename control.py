import random
import time
from tkinter import filedialog
import string
from openpyxl import load_workbook
import xlrd

class Controller:
    ui: object
    def __init__(self):
        self.load_if = None
        pass
    def init(self, ui):
        self.ui = ui

    def save_txt(self):
    # 弹出文件保存对话框
        file_path = filedialog.asksaveasfilename(defaultextension='.txt',
                                                filetypes=[('Text Files', '*.txt')],
                                                title='保存为TXT文件')
        # 如果用户取消了保存，则直接返回
        if not file_path:
            return
        try:
            # 打开选定的文件路径，以写入模式保存数据
            with open(file_path, 'w', encoding='utf-8') as f:
                # 获取表格数据，这里假设表格数据是字符串格式
                for row in self.ui.tk_table_num_collect.get_children():
                    row_data = self.ui.tk_table_num_collect.item(row)['values']
                # 将每行数据转换为字符串并写入文件
                    row_str = '\t'.join(str(item) for item in row_data) + '\n'
                    f.write(row_str)
        except Exception as e:
            return

    def limit_number(self,value, max_value):
        return min(value, max_value)

    def get_characters_range(self,method, start_num, end_num):
        if method == "小写字母":
            characters = string.ascii_lowercase
        elif method == "大写字母":
            characters = string.ascii_uppercase
        else:
            return None
        if end_num > 26:
            end_num = 26
        return characters[start_num-1:end_num]

    def handle_excel_data(self, method, end_num, load_if):
        if load_if.endswith('.xlsx'):
        # 处理 .xlsx 文件
            wb = load_workbook(load_if, read_only=True)
            sheet = wb.active
            columns = int(method[1:-1])
            names_from_excel = [
                sheet.cell(row=i, column=columns).value
                for i in range(1, end_num + 1)
                if sheet.cell(row=i, column=columns).value is not None
            ]
            wb.close()
        elif load_if.endswith('.xls'):
            # 处理 .xls 文件
            wb = xlrd.open_workbook(load_if)
            sheet = wb.sheet_by_index(0)  # 获取第一个工作表
            columns = int(method[1:-1]) - 1  # xlrd 的列索引从 0 开始
            names_from_excel = [
                sheet.cell_value(rowx=i, colx=columns)
                for i in range(end_num)
                if sheet.cell_value(rowx=i, colx=columns) is not None
            ]
        else:
            raise ValueError("Unsupported file format")

        return names_from_excel

    def generate_numbers_from_characters(self,start_num, end_num, generation_num, repeat_if, characters_range):
        available_range_size = len(characters_range)
        if repeat_if == "不重复" and generation_num > available_range_size:
            generation_num = available_range_size
        if repeat_if == "重复":
            return [random.choice(characters_range) for _ in range(generation_num)]
        elif repeat_if == "不重复":
            return random.sample(characters_range, generation_num)

    def generate_numbers_from_range(self,start_num, end_num, generation_num, repeat_if, load_num=None):
        if load_num is not None:
            if load_num < start_num:
                return []
            if repeat_if == "重复":
                if load_num < end_num:
                    end_num = load_num
                return [random.randint(start_num, end_num) for _ in range(generation_num)]
            elif repeat_if == "不重复":
                available_range_size = end_num - start_num + 1
                if generation_num > available_range_size:
                    generation_num = available_range_size
                if load_num < generation_num:
                    generation_num = load_num
                if load_num < end_num:
                    end_num = load_num
                return random.sample(range(start_num, end_num + 1), generation_num)
        else:
            if repeat_if == "重复":
                return [random.randint(start_num, end_num) for _ in range(generation_num)]
            elif repeat_if == "不重复":
                available_range_size = end_num - start_num + 1
                if generation_num > available_range_size:
                    generation_num = available_range_size
                return random.sample(range(start_num, end_num + 1), generation_num)

    def create_number(self, start_num, end_num, generation_num, repeat_if, method, clear_if):
        if start_num > end_num:
            start_num, end_num = end_num, start_num

        if method in ["小写字母", "大写字母"]:
            characters_range = self.get_characters_range(method, start_num, end_num)
            generated_numbers = self.generate_numbers_from_characters(start_num, end_num, generation_num, repeat_if, characters_range)
        elif method not in ["数字", "大写字母", "小写字母"] and self.load_if:
            start_num = max(1, start_num)  # 如果 start_num 是负数 设置为1
            end_num = max(1, end_num)
            names_from_excel = self.handle_excel_data(method, end_num, self.load_if)
            generated_numbers = self.generate_numbers_from_range(start_num, end_num, generation_num, repeat_if, load_num=len(names_from_excel))
        else:
            generated_numbers = self.generate_numbers_from_range(start_num, end_num, generation_num, repeat_if)

        if clear_if == "自动清除":
            self.ui.tk_table_num_collect.delete(*self.ui.tk_table_num_collect.get_children())

        for i, num in enumerate(generated_numbers, start=1):
            if method in ["小写字母", "大写字母"]:
                display_value = num
            elif method != "数字" and self.load_if:
                display_value = names_from_excel[num - 1] if num - 1 < len(names_from_excel) else ''
            else:
                display_value = num

            self.ui.tk_label_now_num.config(text=display_value)
            self.ui.update()
            time.sleep(0.02)
            self.ui.tk_table_num_collect.insert('', 'end', values=(i, display_value))


    def clear_table(self,evt):
        self.ui.tk_table_num_collect.delete(*self.ui.tk_table_num_collect.get_children())

    def load_execl(self, evt):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.load_if = file_path
            self.ui.tk_label_load_label.config(text=file_path)

            # 根据文件扩展名决定使用哪个库
            if file_path.endswith('.xlsx'):
                wb = load_workbook(file_path)
                sheet = wb.active  # 获取活动工作表
                column_count = sheet.max_column
            elif file_path.endswith('.xls'):
                wb = xlrd.open_workbook(file_path)
                sheet = wb.sheet_by_index(0)  # 获取第一个工作表
                column_count = sheet.ncols  # 获取列数
            else:
                self.ui.tk_label_load_label.config(text="不支持的文件格式")
                return

            cb = self.ui.tk_select_box_select_method  # 调用创建Combobox的方法
            column_options = ["数字", "小写字母", "大写字母"] + [f"第{i}列" for i in range(1, column_count + 1)]
            cb['values'] = column_options  # 设置Combobox的选项为列数的范围
            cb.set("第1列")  # 默认选择第一列
        else:
            self.load_if = None
            self.ui.tk_label_load_label.config(text="目前读取内容：空")
            cb = self.ui.tk_select_box_select_method
            cb['values'] = ["数字", "小写字母", "大写字母"]
            cb.set("数字")
