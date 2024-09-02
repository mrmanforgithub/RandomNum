import random
from tkinter import *
from tkinter.ttk import *
from ttkbootstrap import *
from pytkUI.widgets import *
from tkinter import messagebox
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
        btn = Button(parent, text="读取Excel/初始化", takefocus=False,bootstyle="default")
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
        cb['values'] = ("数字","小写字母","大写字母","密码")
        cb.set("数字")
        cb.place(relx=0.0259, rely=0.1750, relwidth=0.1852, relheight=0.0750)
        return cb
class Win(WinGUI):
    def __init__(self, controller):
        self.ctl = controller
        super().__init__()
        self.__event_bind()
        self.__style_config()
        self.config(menu=self.create_menu())
        self.ctl.init(self)
    def create_menu(self):
        menu = Menu(self,tearoff=False)
        menu.add_command(label="导出结果",command=self.ctl.save_txt)
        return menu
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
        self.tk_select_box_save_clear_box.get() ))   # 是否自动清除
        self.tk_button_clear_tab.bind('<Button-1>',self.ctl.clear_table)
        self.tk_button_load_tab.bind('<Button-1>',self.ctl.load_execl)
        self.tk_table_num_collect.bind("<Button-3>", self.show_context_menu)  # 右键点击
        self.tk_table_num_collect.bind("<Button-1>", self.copy_selection)  # 右键点击
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