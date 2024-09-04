import random
from tkinter import *
from tkinter.ttk import *
from ttkbootstrap import *
from pytkUI.widgets import *
from tkinter import messagebox
import json
import re

class WinGUI(Window):
    def __init__(self):
        super().__init__(themename="cosmo", hdpi=False)
        self.__win()
        self.tk_input_start_num = self.__tk_input_start_num(self)
        self.tk_label_from_label = self.__tk_label_from_label(self)
        self.tk_input_end_num = self.__tk_input_end_num(self)
        self.tk_label_now_num = self.__tk_label_now_num(self)
        self.tk_button_create_num = self.__tk_button_create_num(self)
        self.tk_table_num_collect = self.__tk_table_num_collect(self)
        self.tk_label_start_label = self.__tk_label_start_label(self)
        self.tk_label_end_label = self.__tk_label_end_label(self)
        self.tk_input_generation_num = self.__tk_input_generation_num(self)
        self.tk_label_generation_label = self.__tk_label_generation_label(self)
        self.tk_button_clear_tab = self.__tk_button_clear_tab(self)
        self.tk_button_load_tab = self.__tk_button_load_tab(self)
        self.tk_label_load_label = self.__tk_label_load_label(self)
        self.tk_select_box_repeat_if = self.__tk_select_box_repeat_if(self)
        self.tk_select_box_save_clear_box = self.__tk_select_box_save_clear_box(self)
        self.tk_select_box_select_method = self.__tk_select_box_select_method(self)
    def __win(self):
        self.title("随机数生成器")
        # 设置窗口大小、居中
        width = 540
        height = 400
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)

        self.minsize(width=width, height=height)

    def scrollbar_autohide(self,vbar, hbar, widget):
        """自动隐藏滚动条"""
        def show():
            if vbar: vbar.lift(widget)
            if hbar: hbar.lift(widget)
        def hide():
            if vbar: vbar.lower(widget)
            if hbar: hbar.lower(widget)
        hide()
        widget.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Leave>", lambda e: hide())
        if hbar: hbar.bind("<Enter>", lambda e: show())
        if hbar: hbar.bind("<Leave>", lambda e: hide())
        widget.bind("<Leave>", lambda e: hide())

    def v_scrollbar(self,vbar, widget, x, y, w, h, pw, ph):
        widget.configure(yscrollcommand=vbar.set)
        vbar.config(command=widget.yview)
        vbar.place(relx=(w + x) / pw, rely=y / ph, relheight=h / ph, anchor='ne')
    def h_scrollbar(self,hbar, widget, x, y, w, h, pw, ph):
        widget.configure(xscrollcommand=hbar.set)
        hbar.config(command=widget.xview)
        hbar.place(relx=x / pw, rely=(y + h) / ph, relwidth=w / pw, anchor='sw')
    def create_bar(self,master, widget,is_vbar,is_hbar, x, y, w, h, pw, ph):
        vbar, hbar = None, None
        if is_vbar:
            vbar = Scrollbar(master)
            self.v_scrollbar(vbar, widget, x, y, w, h, pw, ph)
        if is_hbar:
            hbar = Scrollbar(master, orient="horizontal")
            self.h_scrollbar(hbar, widget, x, y, w, h, pw, ph)
        self.scrollbar_autohide(vbar, hbar, widget)
    def new_style(self,widget):
        ctl = widget.cget('style')
        ctl = "".join(random.sample('0123456789',5)) + "." + ctl
        widget.configure(style=ctl)
        return ctl
    def __tk_input_start_num(self,parent):
        ipt = Entry(parent, bootstyle="primary")
        ipt.insert(0, "1")
        ipt.place(relx=0.1204, rely=0.0375, relwidth=0.1852, relheight=0.0750)
        return ipt
    def __tk_label_from_label(self,parent):
        label = Label(parent,text="~",anchor="center", bootstyle="primary")
        label.place(relx=0.3333, rely=0.0375, relwidth=0.0926, relheight=0.0750)
        return label
    def __tk_input_end_num(self,parent):
        ipt = Entry(parent, bootstyle="primary")
        ipt.insert(0, "10")
        ipt.place(relx=0.5463, rely=0.0375, relwidth=0.1852, relheight=0.0750)
        return ipt
    def __tk_label_now_num(self,parent):
        label = Label(parent,text="0",anchor="center", bootstyle="info")
        label.place(relx=0.0259, rely=0.2700, relwidth=0.1852, relheight=0.3625)
        return label
    def __tk_button_create_num(self,parent):
        btn = Button(parent, text="生成", takefocus=False,bootstyle="default")
        btn.place(relx=0.8148, rely=0.7750, relwidth=0.1481, relheight=0.1875)
        return btn
    def __tk_table_num_collect(self,parent):
        # 表头字段 表头宽度
        columns = {"Index":77,"Content":311}
        tk_table = Treeview(parent, show="headings", columns=list(columns),bootstyle="primary")
        for text, width in columns.items():  # 批量设置列属性
            tk_table.heading(text, text=text, anchor='center')
            tk_table.column(text, anchor='center', width=width, stretch=True)  # stretch 不自动拉伸
        tk_table.place(relx=0.2407, rely=0.1750, relwidth=0.7222, relheight=0.5500)
        self.create_bar(parent, tk_table,True, False,130, 70, 390,220,540,400)
        return tk_table
    def __tk_label_start_label(self,parent):
        label = Label(parent,text="最小:",anchor="center", bootstyle="primary")
        label.place(relx=0.0278, rely=0.0375, relwidth=0.0926, relheight=0.0750)
        return label
    def __tk_label_end_label(self,parent):
        label = Label(parent,text="最大:",anchor="center", bootstyle="primary")
        label.place(relx=0.4537, rely=0.0375, relwidth=0.0926, relheight=0.0750)
        return label
    def __tk_input_generation_num(self,parent):
        ipt = Entry(parent, bootstyle="default")
        ipt.place(relx=0.1204, rely=0.7750, relwidth=0.1852, relheight=0.0750)
        ipt.insert(0, "1")
        return ipt
    def __tk_label_generation_label(self,parent):
        label = Label(parent,text="个数:",anchor="center", bootstyle="primary")
        label.place(relx=0.0278, rely=0.7750, relwidth=0.0926, relheight=0.0750)
        return label
    def __tk_button_clear_tab(self,parent):
        btn = Button(parent, text="清除", takefocus=False,bootstyle="info")
        btn.place(relx=0.6167, rely=0.7750, relwidth=0.1778, relheight=0.0750)
        return btn
    def __tk_button_load_tab(self,parent):
        btn = Button(parent, text="读取文件/初始化", takefocus=False,bootstyle="default")
        btn.place(relx=0.3241, rely=0.7750, relwidth=0.2778, relheight=0.0750)
        return btn
    def __tk_label_load_label(self,parent):
        label = Label(parent,text="目前读取内容：空",anchor="center", bootstyle="secondary")
        label.place(relx=0.0296, rely=0.8875, relwidth=0.7593, relheight=0.0750)
        return label
    def __tk_select_box_repeat_if(self,parent):
        cb = Combobox(parent, state="readonly", bootstyle="info")
        cb['values'] = ("不重复","重复")
        cb.set("不重复")
        cb.place(relx=0.7593, rely=0.0375, relwidth=0.2037, relheight=0.0750)
        return cb
    def __tk_select_box_save_clear_box(self,parent):
        cb = Combobox(parent, state="readonly", bootstyle="primary")
        cb['values'] = ("自动清除","保留")
        cb.set("自动清除")
        cb.place(relx=0.0259, rely=0.6550, relwidth=0.1852, relheight=0.0750)
        return cb
    def __tk_select_box_select_method(self,parent):
        cb = Combobox(parent, state="readonly", bootstyle="primary")
        cb['values'] = ("数字","字母","密码", "自定义")
        cb.set("数字")
        cb.place(relx=0.0259, rely=0.1750, relwidth=0.1852, relheight=0.0750)
        return cb
class Win(WinGUI):
    def __init__(self, controller):
        self.ctl = controller
        self.current_settings = {
            "number": {
                "float_precision": 0
            },
            "letter": {
                "letter_case": "lower"
            },
            "password": {
                "include_numbers": True,
                "include_lower": True,
                "include_character": True,
                "include_special_chars": True,
                "exclude_similar_chars": False,
                "remove_least": False
            },
            'random_settings':{
            'distribution_type': 'uniform',
            'generate_content': '',
            'append_unit': '',
            "append_prefix":"",
            "append_interval":"",
            'enable_cross_generation': False,
            'enable_content_fix': True,
            "enable_random_cross":False
            }
        }
        self.settings_window = None
        self.content_window = None
        super().__init__()
        self.__event_bind()
        self.__style_config()
        self.config(menu=self.create_menu())
        self.ctl.init(self)

    def create_menu(self):
        menu_bar = Menu(self)
        # 创建“设置”菜单
        settings_menu = Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="导出到TXT", command=self.ctl.save_txt)
        settings_menu.add_command(label="当前模式设置", command=self.show_settings)
        menu_bar.add_cascade(label="菜单", menu=settings_menu)

        return menu_bar

    def show_settings(self):
        selected_method = self.tk_select_box_select_method.get()
        if selected_method == '数字':
            self.show_number_settings()
        elif selected_method == '字母':
            self.show_letter_settings()
        elif selected_method == '密码':
            self.show_password_settings()
        elif selected_method == '自定义':
            self.show_self_settings()
        else:
            messagebox.showinfo("提示", "当前项目无配置选项")

    def on_close_settings_window(self):
        self.settings_window.destroy()
        self.settings_window = None

    def show_number_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            # 如果窗口已经存在，将其置顶
            self.settings_window.lift()
            return
        # 数字模式的设置窗口
        self.settings_window = Toplevel(self)
        self.settings_window.title("数字设置")
        self.settings_window.geometry("250x120")

        Label(self.settings_window, text="浮点数位数:").place(x=20, y=20, width=80, height=25)
        float_precision = Entry(self.settings_window)
        float_precision.insert(0, self.current_settings.get("number", {}).get("float_precision", 0))
        float_precision.place(x=110, y=20, width=100, height=25)

        Button(self.settings_window, text="保存", command=lambda: self.save_settings(
            {"mode": "number", "float_precision": max(int(float_precision.get()),0)},
        self.settings_window
        )).place(x=75, y=70, width=100, height=30)
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_close_settings_window)

    def show_letter_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            # 如果窗口已经存在，将其置顶
            self.settings_window.lift()
            return
        # 字母模式的设置窗口
        self.settings_window = Toplevel(self)
        self.settings_window.title("字母设置")
        self.settings_window.geometry("300x120")

        letter_case = StringVar(value=self.current_settings.get("letter", {}).get("letter_case", "lower"))
        Radiobutton(self.settings_window, text="小写字母", variable=letter_case, value="lower").place(x=20, y=20, width=80, height=25)
        Radiobutton(self.settings_window, text="大写字母", variable=letter_case, value="upper").place(x=110, y=20, width=80, height=25)
        Radiobutton(self.settings_window, text="大小写混合", variable=letter_case, value="mixed").place(x=200, y=20, width=80, height=25)

        Button(self.settings_window, text="保存", command=lambda: self.save_settings(
            {"mode": "letter", "letter_case": letter_case.get()},
        self.settings_window
        )).place(x=100, y=70, width=100, height=30)
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_close_settings_window)

    def show_password_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            # 如果窗口已经存在，将其置顶
            self.settings_window.lift()
            return
        # 密码模式的设置窗口
        self.settings_window = Toplevel(self)
        self.settings_window.title("密码设置")
        self.settings_window.geometry("280x300")

        include_numbers = BooleanVar(value=self.current_settings.get("password", {}).get("include_numbers", "True"))
        include_lower = BooleanVar(value=self.current_settings.get("password", {}).get("include_lower", "True"))
        include_character = BooleanVar(value=self.current_settings.get("password", {}).get("include_character", "True"))
        include_special_chars = BooleanVar(value=self.current_settings.get("password", {}).get("include_special_chars", "True"))
        exclude_similar_chars = BooleanVar(value=self.current_settings.get("password", {}).get("exclude_similar_chars", "False"))
        remove_least = BooleanVar(value=self.current_settings.get("password", {}).get("remove_least", "False"))

        # 使用 place 方法定位组件
        Checkbutton(self.settings_window, text="包含数字", variable=include_numbers).place(x=20, y=20, width=120, height=25)
        Checkbutton(self.settings_window, text="包含小写字母", variable=include_lower).place(x=20, y=60, width=120, height=25)
        Checkbutton(self.settings_window, text="包含大写字母", variable=include_character).place(x=20, y=100, width=120, height=25)
        Checkbutton(self.settings_window, text="包含特殊符号", variable=include_special_chars).place(x=20, y=140, width=120, height=25)
        Checkbutton(self.settings_window, text="提高密码强度", variable=exclude_similar_chars).place(x=20, y=180, width=120, height=25)
        Checkbutton(self.settings_window, text="去除最小密码长度", variable=remove_least).place(x=20, y=220, width=120, height=25)

        Button(self.settings_window, text="保存", command=lambda: self.save_settings(
            {"mode": "password",
            "include_numbers": include_numbers.get(),
            "include_lower": include_lower.get(),
            "include_character": include_character.get(),
            "include_special_chars": include_special_chars.get(),
            "exclude_similar_chars": exclude_similar_chars.get(),
            "remove_least": remove_least.get(),},
        self.settings_window
        )).place(x=100, y=260, width=100, height=30)
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_close_settings_window)

    def show_self_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            # 如果窗口已经存在，将其置顶
            self.settings_window.lift()
            return
        def get_current_random_settings():
            return {
                "mode": "random_settings",
                "distribution_type": distribution_type.get(),
                "generate_content": generate_content.get(),
                "append_unit": append_unit.get(),
                "append_prefix": append_prefix.get(),
                "append_interval": append_interval.get(),
                "enable_cross_generation": enable_cross_generation.get(),
                "enable_content_fix": enable_content_fix.get(),
                "enable_random_cross": enable_random_cross.get()
            }

        # 启用组合生成选项
        def toggle_cross_generation():
            if enable_cross_generation.get():
                interval_entry.config(state=NORMAL)
                random_cross_radiobutton.config(state=NORMAL)
            else:
                interval_entry.config(state=DISABLED)
                random_cross_radiobutton.config(state=DISABLED)
                enable_random_cross.set(False)
                interval_entry.delete(0, END)
            self.save_settings(get_current_random_settings(), self.settings_window, False)

        # 随机数模式的设置窗口
        self.settings_window = Toplevel(self)
        self.settings_window.title("随机数设置")
        self.settings_window.geometry("350x390")  # 调整窗口大小以适应新增内容

        # 定义变量
        distribution_type = StringVar(value=self.current_settings.get("random_settings", {}).get("distribution_type", "uniform"))
        generate_content = StringVar(value=self.current_settings.get("random_settings", {}).get("generate_content", ""))
        append_unit = StringVar(value=self.current_settings.get("random_settings", {}).get("append_unit", ""))
        append_prefix = StringVar(value=self.current_settings.get("random_settings", {}).get("append_prefix", ""))
        append_interval = StringVar(value=self.current_settings.get("random_settings", {}).get("append_interval", ""))
        enable_cross_generation = BooleanVar(value=self.current_settings.get("random_settings", {}).get("enable_cross_generation", False))
        enable_content_fix = BooleanVar(value=self.current_settings.get("random_settings", {}).get("enable_content_fix", True))
        enable_random_cross = BooleanVar(value=self.current_settings.get("random_settings", {}).get("enable_random_cross", False))

        # 定位组件
        Label(self.settings_window, text="选择分布类型:").place(x=20, y=20, width=120, height=25)
        Radiobutton(self.settings_window, text="均匀分布", variable=distribution_type, value="uniform").place(x=140, y=20, width=120, height=25)
        Radiobutton(self.settings_window, text="正态分布", variable=distribution_type, value="normal").place(x=140, y=60, width=120, height=25)

        Label(self.settings_window, text="附加开头:").place(x=20, y=100, width=120, height=25)
        Entry(self.settings_window, textvariable=append_prefix).place(x=140, y=100, width=120, height=25)

        Label(self.settings_window, text="附加末尾:").place(x=20, y=140, width=120, height=25)
        Entry(self.settings_window, textvariable=append_unit).place(x=140, y=140, width=120, height=25)

        Label(self.settings_window, text="附加间隔:").place(x=20, y=180, width=120, height=25)
        interval_entry = Entry(self.settings_window, textvariable=append_interval)
        interval_entry.place(x=140, y=180, width=120, height=25)

        Checkbutton(self.settings_window, text="启用组合生成", variable=enable_cross_generation, command=toggle_cross_generation).place(x=20, y=220, width=150, height=25)
        # 随机交叉选项
        random_cross_radiobutton = Checkbutton(self.settings_window, text="随机组合", variable=enable_random_cross,command=lambda:self.save_settings(get_current_random_settings(), self.settings_window, False))
        random_cross_radiobutton.place(x=190, y=220, width=150, height=25)

        if not enable_cross_generation.get():
            interval_entry.config(state=DISABLED)
            random_cross_radiobutton.config(state=DISABLED)  # 默认禁用
        else:
            interval_entry.config(state=NORMAL)
            random_cross_radiobutton.config(state=NORMAL)  # 默认禁用

        Checkbutton(self.settings_window, text="生成内容纠正", variable=enable_content_fix,command=
        lambda: self.save_settings(get_current_random_settings(), self.settings_window, False)).place(x=20, y=260, width=150, height=25)

        # 生成内容输入按钮，设置为居中
        Button(self.settings_window, text="生成内容", command=lambda: self.show_content_input(generate_content)).place(x=0, y=300, width=350, height=40)

        # 保存按钮，设置为居中并加宽
        Button(self.settings_window, text="保存", command=lambda: self.save_settings(
            get_current_random_settings(), self.settings_window
        )).place(x=0, y=350, width=170, height=40)

        Button(self.settings_window, text="保存并加载", command=lambda: (
            self.save_settings(
                get_current_random_settings(), self.settings_window
            ),
            self.set_content_list(generate_content, self.settings_window)
        )).place(x=180, y=350, width=170, height=40)
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_close_settings_window)

    def show_content_input(self, content_var):
        if self.content_window is not None and self.content_window.winfo_exists():
            # 如果窗口已经存在，将其置顶
            self.content_window.lift()
            return
        # 创建一个输入窗口让用户填写生成内容
        self.content_window = Toplevel(self)
        self.content_window.title("生成内容")
        self.content_window.geometry("300x420")

        Label(self.content_window, text="输入生成内容(内容格式如下):").place(x=20, y=20, width=280, height=25)

        content_text = Text(self.content_window, wrap="word")
        content_text.place(x=20, y=60, width=260, height=300)

        content_text.insert(END, self.current_settings["random_settings"]["generate_content"])

        default_text = '''{
        "日期": ["2024-09-03", "2024-09-23", "2024-10-08"],
        "节日": ["中秋节", "国庆节", "春节"],
        "节气": ["白露", "秋分", "寒露"],
        "注意":["使用json的格式 键:[值,值,值,···], 不同的行之间也要写逗号分开"]
        }'''

        # 插入默认文字
        if not content_text.get("1.0", END).strip():
            content_text.insert(END, default_text)
            content_text.config(fg="gray")

        def on_focus_in(event):
            # 当用户点击文本框时，如果当前内容是默认提示文字，则清空文本框并设置文字颜色为黑色
            if content_text.get("1.0", END).strip() == default_text:
                content_text.delete("1.0", END)
                content_text.config(fg="black")

        def on_focus_out(event):
            # 当文本框失去焦点时，如果文本框为空，则重新插入默认提示文字并设置文字颜色为灰色
            if not content_text.get("1.0", END).strip():
                content_text.insert(END, default_text)
                content_text.config(fg="gray")

        # 绑定事件
        content_text.bind("<FocusIn>", on_focus_in)
        content_text.bind("<FocusOut>", on_focus_out)

        # 确定按钮
        Button(self.content_window, text="确定", command=lambda: self.save_content(content_var, content_text, self.content_window,default_text)).place(x=100, y=380, width=100, height=30)

        self.content_window.protocol("WM_DELETE_WINDOW", self.on_close_content_window)

    def on_close_content_window(self):
        self.content_window.destroy()
        self.content_window = None  # 重置为None 表示窗口已关闭

    def set_content_list(self, generate_content,setting_window):
        generate_content = generate_content.get()
        try:
            generate_content = json.loads(generate_content)
        except json.JSONDecodeError:
            messagebox.showerror("错误", "生成内容格式错误,无法解析为JSON")
            return
        setting_window.destroy()
        # 获取键列表
        keys = list(generate_content.keys())

        if not keys:
            messagebox.showerror("错误", "生成内容为空, 无法加载")
            return

        # 设置tk_select_box_select_method的选项
        cb = self.tk_select_box_select_method
        cb['values'] = ["数字", "字母", "密码", "自定义"]+keys
        cb.set(keys[0])  # 默认选择第一个选项

        messagebox.showinfo("提示", "内容已成功加载到列表")

    def save_content(self, content_var, content_text, content_window, default_text):
        fix_enable = self.current_settings.get("random_settings",{}).get("enable_content_fix",True)
        # 获取文本框中的内容并保存
        if content_text.get("1.0", "end-1c").strip() == default_text:
            content_text.delete("1.0", "end")
            content_var.set(content_text.get("1.0", "end-1c").strip())
            self.current_settings["random_settings"]["generate_content"] = content_var.get()
            content_window.destroy()
            return
        str_content = content_text.get("1.0", "end-1c").strip()
        if fix_enable:
            str_content = content_text.get("1.0", "end-1c").strip().rstrip(',')

            if not str_content.startswith('{'):
                str_content = '{' + str_content
            if not str_content.endswith('}'):
                str_content = str_content + '}'

            symbol_map = {
                '：': ':', '（': '(', '）': ')',
                '【': '[', '】': ']', '｛': '{',
                '｝': '}', '，': ','
            }
            for zh_symbol, en_symbol in symbol_map.items():
                str_content = re.sub(zh_symbol, en_symbol, str_content)
            # 将中文冒号替换为英文冒号
            str_content = re.sub(r'：', ':', str_content)

            # 自动为键添加引号（支持中文或其他非字母字符）
            str_content = re.sub(r'([^"\{\}\[\]\s:,]+)\s*:', r'"\1":', str_content)

            # 自动为方括号中的字符串添加引号，并处理冒号后的空格
            def add_quotes(match):
                # 移除冒号后方括号内的空格
                content = match.group(0)
                content = re.sub(r':\s*\[', ': [', content)
                # 添加引号
                items = content[content.index('[') + 1:content.index(']')].split(',')
                items = [f'"{item.strip()}"' if not (item.strip().startswith('"') and item.strip().endswith('"')) else item.strip() for item in items]
                return '[' + ', '.join(items) + ']'

            str_content = re.sub(r'\[[^\[\]]*\]', add_quotes, str_content)

        content_var.set(str_content)
        content_window.destroy()
        self.current_settings["random_settings"]["generate_content"] = content_var.get()

    def save_settings(self, settings, window,close=True):
        # 将设置保存为 JSON 字符串
        mode = settings.get("mode")
        if mode:
            self.current_settings[mode] = settings
            del self.current_settings[mode]["mode"]
        if close:
            window.destroy()

    def convert_letter_to_number(self,value):
    # 如果输入的是单个字母（忽略大小写），将其转换为对应的数字
        if isinstance(value, str) and len(value) == 1 and value.isalpha():
            return ord(value.lower()) - ord('a') + 1
        return int(value)
    def __event_bind(self):
        self.tk_button_create_num.bind('<Button-1>', lambda event: self.ctl.create_number(
        self.convert_letter_to_number(self.tk_input_start_num.get()),  # 最小值
        self.convert_letter_to_number(self.tk_input_end_num.get()),    # 最大值
        int(self.tk_input_generation_num.get()),     # 生成数量
        self.tk_select_box_repeat_if.get(),          # 是否重复
        self.tk_select_box_select_method.get(),      # 生成方法
        self.tk_select_box_save_clear_box.get(),
        self.current_settings))   # 是否自动清除

        self.tk_button_clear_tab.bind('<Button-1>',self.ctl.clear_table)
        self.tk_button_load_tab.bind('<Button-1>',self.ctl.load_execl)
        self.tk_select_box_select_method.bind("<<ComboboxSelected>>", self.on_select_method_change)
        self.tk_table_num_collect.bind("<Button-3>", self.show_context_menu)  # 右键点击
        self.tk_table_num_collect.bind("<Button-1>", self.copy_selection)  # 左键点击
        self.tk_table_num_collect.bind("<Control-c>", self.copy_selection)  # ctrlc
        self.tk_table_num_collect.bind("<Motion>", self.on_mouse_motion)    #鼠标移动
        pass
    def on_mouse_motion(self, event):
        # 获取鼠标下的表格行号
        item = self.tk_table_num_collect.identify_row(event.y)
        if item:
            # 取消所有当前选中的项
            self.tk_table_num_collect.selection_remove(self.tk_table_num_collect.selection())
            # 选中鼠标下的行
            self.tk_table_num_collect.selection_add(item)
    def show_context_menu(self, event):
        context_menu = Menu(self.tk_table_num_collect, tearoff=0)
        context_menu.add_command(label="复制", command=self.copy_selection)
        context_menu.add_command(label="删除", command=self.delete_selection)
        context_menu.post(event.x_root, event.y_root)
    def copy_selection(self, event=None):
        selected_items = self.tk_table_num_collect.selection()
        if selected_items:
            # 将选中的内容拼接为字符串，列之间用制表符，行之间用换行符
            selected_text = "\n".join(
                [str(self.tk_table_num_collect.item(item, 'values')[1]) for item in selected_items]
            )
            # 将拼接好的字符串复制到剪贴板
            self.tk_table_num_collect.clipboard_clear()
            self.tk_table_num_collect.clipboard_append(selected_text)
            # 更新系统剪贴板
            self.tk_table_num_collect.update()  # 确保剪贴板更新
            for item in selected_items:
                original_value = self.tk_table_num_collect.item(item, 'values')
                new_value = (original_value[0], "已复制到剪贴板")
                self.tk_table_num_collect.item(item, values=new_value)
                # 设置延迟，之后恢复原始内容
                self.tk_table_num_collect.after(500, lambda i=item, v=original_value: self.tk_table_num_collect.item(i, values=v))
    def delete_selection(self):
        selected_items = self.tk_table_num_collect.selection()
        if selected_items:
            for item in selected_items:
                self.tk_table_num_collect.delete(item)
    def on_select_method_change(self,event):
        method = self.tk_select_box_select_method.get()
        if method == "自定义" and self.current_settings["random_settings"]["generate_content"] =="":
            self.show_self_settings()

    def __style_config(self):
        sty = Style()
        sty.configure(self.new_style(self.tk_label_from_label),font=("微软雅黑",-22,"bold"))
        sty.configure(self.new_style(self.tk_label_now_num),font=("微软雅黑",-25,"bold underline"))
        sty.configure(self.new_style(self.tk_button_create_num),font=("微软雅黑",-20,"bold"))
        sty.configure(self.new_style(self.tk_label_start_label),font=("微软雅黑",-18,"bold"))
        sty.configure(self.new_style(self.tk_label_end_label),font=("微软雅黑",-18,"bold"))
        sty.configure(self.new_style(self.tk_label_generation_label),font=("微软雅黑",-18,"bold"))
        sty.configure(self.new_style(self.tk_button_clear_tab),font=("微软雅黑",-15,"bold"))
        sty.configure(self.new_style(self.tk_button_load_tab),font=("微软雅黑",-14,"bold"))
        sty.configure(self.new_style(self.tk_label_load_label),font=("微软雅黑",-13,"underline"))
        sty.configure("Treeview", font=("微软雅黑", 11))  # 单元格字体大小
        pass
if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()