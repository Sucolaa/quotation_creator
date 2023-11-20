import sys
sys.path.append("./Lib/")
import tkinter as tk
from tkinter import Toplevel,ttk
import pandas as pd
import time
from data_editor import DataEditor
from quotation import quotation_en, quotation_cn


class Quotation_Generator:
#总表样式
    def __init__(self, root):
        self.root = root
        self.root.title("Quotation Generator")

        self.yex_package = pd.read_csv("../data/package.csv", encoding='gbk')
        self.yex_brand = pd.read_csv("../data/brand_company_address.csv", encoding='gbk')
        self.record = pd.read_csv("../data/record.csv", encoding='gbk')

        self.new_number = (self.record['Quotation Date'] == time.strftime("%Y%m%d")).sum() + 1
        # self.record = pd.read_excel("../data/record.xlsx", sheet_name='Sheet1')

        self.lang_label = tk.Label(root, text="语言：")
        self.lang_label.grid(row=0, column=0, sticky="w")
        self.lang_var = tk.StringVar()
        self.lang_choices = ['中文CN',"英文EN","双语"]
        self.lang_menu = tk.OptionMenu(root, self.lang_var, *self.lang_choices)
        self.lang_menu.grid(row=0, column=1, sticky="w")
        self.lang_var.set(self.lang_choices[0])

        self.brand_label = tk.Label(root, text="品牌：")
        self.brand_label.grid(row=1, column=0, sticky="w")
        self.brand_var = tk.StringVar()
        self.brand_choices = list(self.yex_brand.drop_duplicates(subset=['Brand_Name'], keep="first")['Brand_Name'])
        self.brand_choices.append("(新建)")
        self.brand_menu = tk.OptionMenu(root, self.brand_var, *self.brand_choices)
        self.brand_menu.grid(row=1, column=1, sticky="w")

        self.company_label = tk.Label(root, text="公司名：")
        self.company_label.grid(row=2, column=0, sticky="w")
        self.company_var = tk.StringVar()
        self.company_choices = list(self.yex_brand['Company_Name'])
        self.company_choices.append("(新建)")
        self.company_menu = tk.OptionMenu(root, self.company_var, *self.company_choices)
        self.company_menu.grid(row=2, column=1, sticky="w")

        self.address_label = tk.Label(root, text="公司地址：")
        self.address_label.grid(row=3, column=0, sticky="w")
        self.address_var = tk.StringVar()
        self.address_choices = list(self.yex_brand['Company_Address'])
        self.address_choices.append("(新建)")
        self.address_menu = tk.OptionMenu(root, self.address_var, *self.address_choices)
        self.address_menu.grid(row=3, column=1, sticky="w")

        self.package_label = tk.Label(root, text="套餐等级：")
        self.package_label.grid(row=4, column=0, sticky="w")
        self.package_var = tk.StringVar()
        self.package_choices = self.yex_package['套餐']
        self.package_menu = tk.OptionMenu(root, self.package_var, *self.package_choices)
        self.package_menu.grid(row=4, column=1, sticky="w")
        self.package_var.set(self.package_choices[0])

        self.currency_label = tk.Label(root, text="结算货币：")
        self.currency_label.grid(row=5, column=0, sticky="w")
        self.currency_var = tk.StringVar()
        self.currency_choices = ['CNY','USD']
        self.currency_menu = tk.OptionMenu(root, self.currency_var, *self.currency_choices)
        self.currency_menu.grid(row=5, column=1, sticky="w")
        self.currency_var.set(self.currency_choices[0])

        self.quantity_label = tk.Label(root, text="数量：")
        self.quantity_label.grid(row=6, column=0, sticky="w")
        self.quantity_var = tk.StringVar()
        self.quantity_entry = tk.Entry(root, textvariable=self.quantity_var)
        self.quantity_entry.grid(row=6, column=1, sticky="w")
        self.quantity_var.set("0")

        self.tax_label = tk.Label(root, text="税率：")
        self.tax_label.grid(row=7, column=0, sticky="w")
        self.tax_var = tk.StringVar()
        self.tax_entry = tk.Entry(root, textvariable=self.tax_var)
        self.tax_entry.grid(row=7, column=1, sticky="w")
        self.tax_var.set("1")

        self.period_label = tk.Label(root, text="付款周期：")
        self.period_label.grid(row=8, column=0, sticky="w")
        self.period_var = tk.StringVar()
        self.period_entry = tk.Entry(root, textvariable=self.period_var)
        self.period_entry.grid(row=8, column=1, sticky="w")
        self.period_var.set("0")

        self.generate_button = tk.Button(root, text="生成PDF", command=self.generate_pdf)
        self.generate_button.grid(row=9, column=0, sticky="w")

        brand_columns = ['Brand_Name', 'Company_Name', 'Company_Address']
        self.brand_editor = DataEditor(self.root, self.yex_brand, brand_columns, '../data/brand_company_address.csv')
        self.brand_button = tk.Button(self.root, text="修改品牌表", command=self.brand_editor.manu_brand_list)
        self.brand_button.grid(row=9, column=1, sticky="w")

        package_columns = ["Package", "套餐", "USD", "CNY"]
        self.package_editor = DataEditor(self.root, self.yex_package, package_columns, '../data/package.csv')
        self.package_button = tk.Button(self.root, text="修改套餐表", command=self.package_editor.manu_package_list)
        self.package_button.grid(row=9, column=2, sticky="w")

        record_columns = ["Brand Name", "Currency", "Available", "Quotation Number", "Quotation Date", "Number", 
                          "Company Name", "Company Address", "Offering Date", "Payment Period", "Package Item", 
                          "Service Period", "Quantity", "Unit Price", "Total Amount", "Tax", "Total Amount with Tax", 
                          "PO Number", "回单编号", "备注1", "备注2"]
        self.record_editor = DataEditor(self.root, self.record, record_columns, "../data/record.csv")
        self.record_button = tk.Button(self.root, text="修改记录表", command=self.record_editor.manu_record_list)
        self.record_button.grid(row=9, column=3, sticky="w")
        


        self.lang_var.trace_add("write", self.update_package_choices)
        self.brand_var.trace_add("write", self.update_company_choices)
        self.company_var.trace_add("write", self.update_address_choices)

        self.company_var.trace_add("write", self.check_create_new_company)
        self.brand_var.trace_add("write", self.check_create_new_brand)
        self.address_var.trace_add("write", self.check_create_new_address)

        self.item_id_to_index = {}


#根据上一选项更新下一选项内容
    def update_package_choices(self, *args):
        lang_choice = self.lang_var.get()

        if lang_choice == '中文CN':
            self.package_choices = self.yex_package['套餐']
        elif lang_choice == '英文EN':
            self.package_choices = self.yex_package['Package']
        elif lang_choice == '双语':
            self.package_choices = self.yex_package['套餐']
        self.package_menu['menu'].delete(0, 'end')
        for package in self.package_choices:
            self.package_menu['menu'].add_command(label=package, command=tk._setit(self.package_var, package))

    def update_company_choices(self, *args):
        brand = self.brand_var.get()
        company_choice = self.yex_brand.loc[self.yex_brand['Brand_Name'] == brand]
        self.company_choices = list(company_choice['Company_Name'])
        self.company_choices.append("(新建)")
        self.company_menu['menu'].delete(0,'end')
        for company_name in self.company_choices:
            self.company_menu['menu'].add_command(label=company_name, command=tk._setit(self.company_var, company_name))

    def update_address_choices(self, *args):
        company = self.company_var.get()
        address_choice = self.yex_brand.loc[self.yex_brand['Company_Name'] == company]
        self.address_choices = list(address_choice['Company_Address'])
        self.address_choices.append("(新建)")
        self.address_menu['menu'].delete(0,'end')
        for address in self.address_choices:
            self.address_menu['menu'].add_command(label=address, command=tk._setit(self.address_var, address))

#手动新增单项选择（单次，不会保存）
    def check_create_new_brand(self, *args):
        if self.brand_var.get() == "(新建)":
            self.create_new_item("品牌", self.brand_choices)

    def check_create_new_company(self, *args):
        if self.company_var.get() == "(新建)":
            self.create_new_item("公司", self.company_choices)

    def check_create_new_address(self, *args):
        if self.address_var.get() == "(新建)":
            self.create_new_item("地址", self.address_choices)

    def create_new_item(self, item_type, choices_list):
        new_item_window = tk.Toplevel(self.root)
        new_item_window.title(f"新{item_type}")
        
        new_item_label = tk.Label(new_item_window, text=f"新{item_type}:")
        new_item_label.pack()

        new_item_var = tk.StringVar()
        new_item_entry = tk.Entry(new_item_window, textvariable=new_item_var)
        new_item_entry.pack()

        save_button = tk.Button(new_item_window, text="添加", command=lambda: self.save_new_item(item_type,choices_list, new_item_var, new_item_window))
        save_button.pack()

    def save_new_item(self, item_type, choices_list, new_item_var,new_item_window):
        new_item = new_item_var.get()
        choices_list.append(new_item)
        self.update_dropdown_menu(item_type)
        new_item_window.destroy()
        self.root.update()

    def update_dropdown_menu(self, item_type):
        if item_type == "品牌":
            self.brand_menu['menu'].delete(0, 'end')
            for brand in self.brand_choices:
                self.brand_menu['menu'].add_command(label=brand, command=tk._setit(self.brand_var, brand))
        elif item_type == "公司":
            self.company_menu['menu'].delete(0, 'end')
            for company in self.company_choices:
                self.company_menu['menu'].add_command(label=company, command=tk._setit(self.company_var, company))
        elif item_type == "地址":
            self.address_menu['menu'].delete(0, 'end')
            for address in self.address_choices:
                self.address_menu['menu'].add_command(label=address, command=tk._setit(self.address_var, address))


#成功/失败输出pdf文档的提示窗
    def show_success_frame(self):
        success_frame = Toplevel(self.root)
        success_frame.title("Info")
        success_label = tk.Label(success_frame, text="成功创建报价单,记录已新增")
        success_label.pack()
        close_button = tk.Button(success_frame, text="关闭", command=success_frame.destroy)
        close_button.pack()

    def show_fail_frame(self):
        fail_frame = Toplevel(self.root)
        fail_frame.title("Info")
        fail_label = tk.Label(fail_frame, text = "未能创建报价单，请确认信息是否齐全，所有项都为必填项")
        fail_label.pack()
        close_button = tk.Button(fail_frame, text="关闭", command=fail_frame.destroy)
        close_button.pack()

#将本次生成报价单新增到记录表
    def add_record(brand,currency,tax,company,address,period,package, price,quantity):
        record = pd.read_csv("../data/record.csv", encoding = 'gbk')
        new_number = (record['Quotation Date'] == time.strftime("%Y%m%d")).sum()
        new_number = new_number + 1
        current_date_struct = time.strptime(time.strftime('%Y%m%d'), '%Y%m%d')
        one_year_later_struct = time.struct_time((
            current_date_struct.tm_year + 1,
            current_date_struct.tm_mon, 
            current_date_struct.tm_mday,
            current_date_struct.tm_hour,
            current_date_struct.tm_min,
            current_date_struct.tm_sec,
            current_date_struct.tm_wday,
            current_date_struct.tm_yday,
            current_date_struct.tm_isdst))
        one_year_later_str = time.strftime('%Y%m%d', one_year_later_struct)

        price = float(price)
        quantity = float(quantity)
        total_amount = float(quantity*price)
        tax_cal = (int(tax)/100)+1
        total_with_tax = float(total_amount*tax_cal)
        address = address.replace("nan","")
        new_quotation = {
            'Brand Name' : brand,
            'Currency' : currency,
            'Available' : "",
            'Quotation Number' : f"YEXSYS{time.strftime('%Y%m%d')}{str(new_number).zfill(2)}",
            'Quotation Date' : f"{time.strftime('%Y%m%d')}",
            'Number' : f"{new_number}",
            'Company Name' : f"{company}",
            'Company Address' : f"{address}",
            'Offering Date' : f"{time.strftime('%B %d, %Y')}",
            'Payment Period' : f"NET {period}",
            'Package Item' : f"{package}",
            'Service Period' : f"{time.strftime('%Y%m%d')} ~ {one_year_later_str}",
            'Quantity' : f"{int(quantity)}",
            'Unit Price' : f"{price}",
            'Total Amount' : f"{total_amount}",
            'Tax' : f"{tax}%",
            'Total Amount with Tax' : f"{total_with_tax}"
        }
        record = record._append(new_quotation, ignore_index = True)
        record.to_csv("../data/record.csv", index = False, encoding='gbk')

#生成pdf，相关代码在另一个py里
    def generate_pdf(self, *args):
        lang = self.lang_var.get()
        brand = self.brand_var.get()
        company = self.company_var.get()
        address = self.address_var.get()
        package = self.package_var.get()
        currency = self.currency_var.get()
        quantity = self.quantity_var.get()
        tax = self.tax_var.get()
        period = self.period_var.get()

        # out_record = pd.read_csv("../data/record.csv", encoding='gbk')
        # new_number = (self.record['Quotation Date'] == time.strftime("%Y%m%d")).sum() + 1

        if company and package and address and quantity and (lang == "中文CN"):
            selected_package = self.yex_package['套餐'][self.yex_package['套餐'] == package].values[0]
            price = self.yex_package[f'{currency}'][self.yex_package['套餐'] == package].values[0]
            quotation_cn(company=company,address=address,quantity = quantity, price = price, tax = tax, period = period,currency = currency,number = self.new_number,package=selected_package, file_path=f"../YEXSYS{time.strftime('%Y%m%d')}{str(self.new_number).zfill(2)}.pdf")
            Quotation_Generator.add_record(brand,currency,tax,company,address,period,package,price,quantity)
            self.new_number += 1
            self.show_success_frame()

        elif company and package and quantity and (lang == "英文EN"):
            selected_package = self.yex_package['Package'][self.yex_package['Package'] == package].values[0]
            price = self.yex_package[f'{currency}'][self.yex_package['Package'] == package].values[0]
            quotation_en(company=company,address=address,quantity = quantity, price = price, tax = tax, period = period,currency = currency,number=self.new_number,package=selected_package, file_path=f"../YEXSYS{time.strftime('%Y%m%d')}{str(self.new_number).zfill(2)}.pdf")
            Quotation_Generator.add_record(brand,currency,tax,company,address,period,package,price,quantity)
            self.new_number += 1
            self.show_success_frame()

        elif company and package and address and quantity and (lang == "双语"):
            package_cn = self.yex_package['套餐'][self.yex_package['套餐'] == package].values[0]
            package_en = self.yex_package['Package'][self.yex_package['套餐'] == package].values[0]
            price = self.yex_package[f'{currency}'][self.yex_package['套餐'] == package].values[0]
            quotation_cn(company=company,address=address,quantity = quantity, price = price, tax = tax, period = period,currency = currency,number=self.new_number,package=package_cn, file_path=f"../YEXSYS{time.strftime('%Y%m%d')}{str(self.new_number).zfill(2)}.pdf")
            Quotation_Generator.add_record(brand,currency,tax,company,address,period,package,price,quantity)
            self.new_number += 1
            quotation_en(company=company,address=address,quantity = quantity, price = price, tax = tax, period = period,currency = currency,number=self.new_number,package=package_en, file_path=f"../YEXSYS{time.strftime('%Y%m%d')}{str(self.new_number).zfill(2)}.pdf")
            Quotation_Generator.add_record(brand,currency,tax,company,address,period,package,price,quantity)
            self.new_number += 1
            self.show_success_frame()

        else:
            self.show_fail_frame()

if __name__ == "__main__":
    root = tk.Tk()
    app = Quotation_Generator(root)
    root.mainloop()