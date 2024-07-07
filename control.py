import random
import time
from tkinter import filedialog
from openpyxl import load_workbook

class Controller:
    ui: object
    def __init__(self):
        self.load_if = None
        pass
    def init(self, ui):
        self.ui = ui
        # TODO 组件初始化 赋值操作
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


    def create_number(self, start_num, end_num, generation_num, repeat_if, method, clear_if):
        if start_num > end_num:
            j = start_num
            start_num = end_num
            end_num = j
        # start_num 是随机数最小值, end_num 是最大值, generation_num 是生成随机数个数, repeat_if 是是否生成重复随机数
        # method是扫描方法
        if method != "数字" and self.load_if:
            wb = load_workbook(self.load_if)
            sheet = wb.active
            columns = int(method[:-1])
            names_from_excel = [sheet.cell(row=i, column=columns).value for i in range(1, end_num + 1)if sheet.cell(row=i, column=columns).value is not None]
            wb.close()
            generated_numbers = self.generate_numbers(start_num, end_num, generation_num, repeat_if, load_num = len(names_from_excel))
            if clear_if == "自动清除":
                self.ui.tk_table_num_collect.delete(*self.ui.tk_table_num_collect.get_children())
            for i, num in enumerate(generated_numbers, start=1):
                name = names_from_excel[num - 1]
                self.ui.tk_label_now_num.config(text=num)
                self.ui.update()
                time.sleep(0.02)
                self.ui.tk_table_num_collect.insert('', 'end', values=(i, name))
        else:
            generated_numbers = self.generate_numbers(start_num, end_num, generation_num, repeat_if)
            if clear_if == "自动清除":
                self.ui.tk_table_num_collect.delete(*self.ui.tk_table_num_collect.get_children())
            for i, num in enumerate(generated_numbers, start=1):
                self.ui.tk_label_now_num.config(text=num)
                self.ui.update()
                time.sleep(0.02)
                self.ui.tk_table_num_collect.insert('', 'end', values=(i, num))

    def generate_numbers(self, start_num, end_num, generation_num, repeat_if, load_num = None):
        if load_num is not None:
            if load_num < start_num:
                self.ui.tk_label_now_num.config(text="err")
                return []
            if repeat_if == "重复":
                if load_num < end_num:
                    end_num = load_num
                generated_numbers = [random.randint(start_num, end_num) for _ in range(generation_num)]
            elif repeat_if == "不重复":
                if load_num < generation_num:
                    generation_num = load_num
                if (end_num - start_num + 1) < generation_num:
                    generation_num = (end_num - start_num + 1)
                if load_num < end_num:
                    end_num = load_num
                generated_numbers = random.sample(range(start_num, end_num + 1), generation_num)
            return generated_numbers
        else:
            if repeat_if == "重复":
                generated_numbers = [random.randint(start_num, end_num) for _ in range(generation_num)]
            elif repeat_if == "不重复":
                if (end_num - start_num + 1) < generation_num:
                    generation_num = (end_num - start_num + 1)
                generated_numbers = random.sample(range(start_num, end_num + 1), generation_num)
            return generated_numbers
    

    def clear_table(self,evt):
        self.ui.tk_table_num_collect.delete(*self.ui.tk_table_num_collect.get_children())
        # self.load_if = None
        # self.ui.tk_label_load_label.config(text="目前读取内容：空")
        # cb = self.ui.tk_select_box_select_method
        # cb['values'] =  ["数字"]  # 设置Combobox的选项为列数的范围
        # cb.set("数字")  # 默认选择第一列

    def load_execl(self,evt):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.load_if = file_path
            self.ui.tk_label_load_label.config(text=file_path)
            wb = load_workbook(file_path)
            sheet = wb.active  # 获取活动工作表，你也可以根据需要选择特定的工作表
            # 获取列数
            column_count = sheet.max_column
            cb = self.ui.tk_select_box_select_method # 调用创建Combobox的方法
            column_options = ["数字"] + [f"{i}列" for i in range(1, column_count + 1)]
            cb['values'] = column_options  # 设置Combobox的选项为列数的范围
            cb.set("1列")  # 默认选择第一列
        else:
            return None

