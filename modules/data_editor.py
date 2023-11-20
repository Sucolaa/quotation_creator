import sys
sys.path.append("./Lib/")
import tkinter as tk
from tkinter import Toplevel,ttk
import pandas as pd

class DataEditor:

    def __init__(self, root, data, columns, filename):
        self.root = root
        self.data = data
        self.columns = columns
        self.filename = filename
        self.item_id_to_index = {}

    def setup_ui(self, title):
        self.window = tk.Toplevel(self.root)
        self.window.title(title)

        self.tree = ttk.Treeview(self.window, columns=self.columns, show="headings", height=30)

        for col in self.columns:
            self.tree.heading(col, text=col)
        self.tree.grid(row=0, column=0, sticky="nsew")

        for i, row in self.data.iterrows():
            self.tree.insert("", "end", iid=i, values=row.tolist())

        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=0)
        self.window.grid_columnconfigure(0, weight=1)

        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        self.window.geometry(f"{screen_width}x{screen_height-150}")

        vscrollbar = ttk.Scrollbar(self.window, orient="vertical", command=self.tree.yview)
        vscrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vscrollbar.set)
        hscrollbar = ttk.Scrollbar(self.window, orient="horizontal", command=self.tree.xview)
        hscrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=hscrollbar.set)

        # self.window.grid_rowconfigure(0, weight=1)
        # self.window.grid_columnconfigure(0, weight=1)


        self.change_button = tk.Button(self.window, text="修改", command=self.change_item)
        self.change_button.grid()

        self.add_button = tk.Button(self.window, text="新增", command=self.add_item)
        self.add_button.grid()
        
        self.delete_button = tk.Button(self.window, text="删除", command=self.delete_item)
        self.delete_button.grid()

        self.save_button = tk.Button(self.window, text="保存", command=self.save_changes)
        self.save_button.grid()

    def change_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_id = selected_item[0]
            row_index = self.tree.index(item_id)
            selected_row = self.data.iloc[row_index]

            edit_window = tk.Toplevel(self.root)
            edit_window.title("修改信息")

            entry_vars = []
            for col in self.columns:
                label = tk.Label(edit_window, text=col)
                label.grid(row=self.columns.index(col), column=0)
                entry_var = tk.StringVar()
                entry_var.set(selected_row[col])
                entry = tk.Entry(edit_window, textvariable=entry_var)
                entry.grid(row=self.columns.index(col), column=1)
                entry_vars.append(entry_var)

            def save_changes():
                for col, entry_var in zip(self.columns, entry_vars):
                    self.data.loc[row_index, col] = entry_var.get()
                self.load_data()
                edit_window.destroy()

            save_button = tk.Button(edit_window, text="保存", command=save_changes)
            save_button.grid(row=len(self.columns), columnspan=2)

    def add_item(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("新增")

        entry_vars = {}
        for col in self.columns:
            label = tk.Label(add_window, text=col)
            label.grid(row=self.columns.index(col), column=0)
            entry_var = tk.StringVar()
            entry = tk.Entry(add_window, textvariable=entry_var)
            entry.grid(row=self.columns.index(col), column=1)
            entry_vars[col] = entry_var

        def save_new_row():
            new_row = {col: entry_var.get() for col, entry_var in entry_vars.items()}
            self.data = self.data._append(new_row, ignore_index=True)
            self.load_data()
            add_window.destroy()

        save_button = tk.Button(add_window, text="保存", command=save_new_row)
        save_button.grid(row=len(self.columns), columnspan=2)

    def load_data(self):
        self.tree.delete(*self.tree.get_children())
        for i, row in self.data.iterrows():
            item_id = i
            self.item_id_to_index[item_id] = i
            self.tree.insert("", "end", iid=item_id, values=row.tolist())

    def delete_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_id = selected_item[0]
            index = self.tree.index(item_id)
            self.data = self.data.drop(index)
            self.load_data()

    def save_changes(self):
        for item_id in self.tree.get_children():
            item_data = self.tree.item(item_id)["values"]
            row_index = self.item_id_to_index[int(item_id)]
            self.data.loc[row_index] = item_data
        try:
            self.data.to_csv(self.filename, index=False, encoding = "gbk")
            self.success_save_frame()
        except:
            self.fail_save_frame()

    #三张表的更改成功/失败提示弹窗
    def success_save_frame(self):
        success_frame = Toplevel(self.root)
        success_frame.title("Info")
        success_label = tk.Label(success_frame, text="成功修改表内容")
        success_label.pack()
        close_button = tk.Button(success_frame, text="关闭", command=success_frame.destroy)
        close_button.pack()

    def fail_save_frame(self):
        fail_frame = Toplevel(self.root)
        fail_frame.title("Info")
        fail_label = tk.Label(fail_frame, text = "未能修改表内容")
        fail_label.pack()
        close_button = tk.Button(fail_frame, text="关闭", command=fail_frame.destroy)
        close_button.pack()

    def manu_brand_list(self):
        self.setup_ui("品牌/公司/地址")

    def manu_package_list(self):
        self.setup_ui("套餐")

    def manu_record_list(self):
        self.setup_ui("历史记录")