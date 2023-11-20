import sys
sys.path.append("./Lib/")
import reportlab
import re
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import Image
import time

def quotation_cn(package,company,address,quantity,price, tax,period, currency,number, file_path):
    width, height = A4
    c = canvas.Canvas(file_path, pagesize=A4)

    if currency == "USD":
        currency_sign = "$"
    elif currency == "CNY":
        currency_sign = "￥"

    address = address.replace("nan","")

    price = float(price)
    quantity = float(quantity)
    total_amount = float(quantity*price)
    tax_cal = (int(tax)/100)+1
    total_with_tax = float(total_amount*tax_cal)

    price = '{:,.2f}'.format(price)
    quantity = int(quantity)
    total_amount = '{:,.2f}'.format(total_amount)
    total_with_tax = '{:,.2f}'.format(total_with_tax)

    number = str(number).zfill(2)
    
    yexsys_logo_path = "../resources/yexsys_logo.png"
    c.drawImage(yexsys_logo_path, 50, height - 150, 121.68, 121.68)

    pdfmetrics.registerFont(TTFont("dengxian", "../resources/dengxian.ttf"))
    pdfmetrics.registerFont(TTFont("dengxian_bold", "../resources/dengxian_bold.ttf"))

    c.setFont("dengxian_bold", 11)
    c.drawString(50, height - 180, "上海业询科技有限公司")

    c.setFont("dengxian",11)
    c.drawString(50, height - 195, "中国(上海)自贸试验区")
    c.drawString(50, height - 210, "环龙路65弄1号3层、4层 ")

    c.setFont("dengxian_bold", 11)
    c.drawString(50, height - 270, "报价至:")

    c.setFont("dengxian",11)
    c.drawString(50, height - 285, f"{company}")
    c.drawString(50, height - 300, f"{address}")

    c.setFont("dengxian_bold", 28)
    c.drawString(420, height - 80, "报价单")
    c.setFont("dengxian",14)
    #yexsys 年月日 第几次报价 做替换
    c.drawString(400, height - 105, f"#YEXSYS{time.strftime('%Y%m%d')}{number}")

    c.setFont("dengxian_bold",11)
    c.drawString(400, height - 180, "报价日期:")
    c.drawString(400, height - 195, "付款周期:")

    c.setFont("dengxian",11)
    #做替换
    c.drawString(470, height - 180, f"{time.strftime('%B %d, %Y')}")
    c.drawString(512, height - 195, f"NET {period}")

    c.setFillColor("#f7f7f7")
    c.rect(370, height - 230, 180, 30, fill=True, stroke=False)

    c.setFillColor("#000000")
    c.setFont("dengxian_bold",11)
    c.drawString(422, height - 218, "总计:")
    #做替换
    c.setFont("dengxian",11)
    c.drawString(494, height - 218, f"{currency_sign}{total_with_tax}")

    c.rect(50, height - 405, width-100, 30, fill = True)

    c.setFillColor("#ffffff")
    c.setFont("dengxian_bold",11)
    c.drawString(100, height - 395, "项目")
    c.drawString(240, height - 395, "服务期")
    c.drawString(340, height - 395, "数量")
    c.drawString(420, height - 395, "单价")
    c.drawString(500, height - 395, "金额")

    c.setStrokeColor("#f7f7f7")
    c.line(50, height - 435, width-50, height - 435)
    c.line(50, height - 470, width-50, height - 470)
    c.line(50, height - 505, width-50, height - 505)

    c.setFont("dengxian",11)
    c.setFillColor("#000000")
    #替换
    c.drawString(55, height-430, f"{package}")
    c.drawString(342.5, height-430, f"{quantity}")
    c.drawString(415, height-430, f"{currency_sign}{price}")
    c.drawString(485, height-430, f"{currency_sign}{total_amount}")

    c.drawString(425, height-530, f"金额：         {currency_sign}{total_amount}")
    c.drawString(425, height-550, f"税点：                      {tax}%")
    c.drawString(485, height-570, f"{currency_sign}{total_with_tax}")
    c.setFont("dengxian_bold",11)
    c.drawString(425, height-570, "总计：")

    c.drawString(50, height-650,"Notes:")
    c.drawString(50, height-695,"Terms:")

    c.setFont("dengxian",11)
    c.drawString(50, height-665,"对于所有与发票相关的问题，请联系我们的应收账款团队：ar@yexsys.com.")
    c.drawString(50, height-710,"请将付款发送至: ")
    c.drawString(50, height-725,"账户名: 上海业询科技有限公司")
    c.drawString(50, height-740,"银行名: 交通银行交银大厦支行")
    c.drawString(50, height-755,"银行地址: 中国上海市浦东新区银城中路188号")
    c.drawString(50, height-770,"SWIFT代码: COMMCNSHSHI")
    c.drawString(50, height-785,"银行账号: 310066577013003715694")

    c.save()

def quotation_en(package,company,address,quantity,price,tax,period,currency,number, file_path = "../quotation_en.pdf"):

    width, height = A4
    c = canvas.Canvas(file_path, pagesize=A4)

    if currency == "USD":
        currency_sign = "$"
    elif currency == "CNY":
        currency_sign = "￥"
    address = address.replace("nan","")

    price = float(price)
    quantity = float(quantity)
    total_amount = float(quantity*price)
    tax_cal = (int(tax)/100)+1
    total_with_tax = float(total_amount*tax_cal)

    price = '{:,.2f}'.format(price)
    quantity = int(quantity)
    total_amount = '{:,.2f}'.format(total_amount)
    total_with_tax = '{:,.2f}'.format(total_with_tax)
    number = str(number).zfill(2)
    
    yexsys_logo_path = "../resources/yexsys_logo.png"
    c.drawImage(yexsys_logo_path, 50, height - 150, 121.68, 121.68)

    pdfmetrics.registerFont(TTFont("dengxian", "../resources/dengxian.ttf"))
    pdfmetrics.registerFont(TTFont("dengxian_bold", "../resources/dengxian_bold.ttf"))

    c.setFont("dengxian_bold", 11)
    c.drawString(50, height - 180, "Shanghai YexSys Technology Co., Ltd.")

    c.setFont("dengxian",11)
    c.drawString(50, height - 195, "No. 608, KYMS Building")
    c.drawString(50, height - 210, "555 Wuding Road, Jing'an District ")
    c.drawString(50, height - 225, "Shanghai, 200040 P. R. China. ")

    c.setFont("dengxian_bold", 11)
    c.drawString(50, height - 270, "Bill To :")

    c.setFont("dengxian",11)
    #公司名和地址用选项替换
    c.drawString(50, height - 285, f"{company}")
    c.drawString(50, height - 300, f"{address}")

    c.setFont("dengxian_bold", 28)
    c.drawString(380, height - 80, "QUOTATION")
    c.setFont("dengxian",14)
    #yexsys 年月日 第几次报价 做替换
    c.drawString(400, height - 105, f"#YEXSYS{time.strftime('%Y%m%d')}{number}")

    c.setFont("dengxian_bold",11)
    c.drawString(420, height - 180, "DATE :")
    c.drawString(360, height - 195, "PAYMENT TERMS :")

    c.setFont("dengxian",11)
    #做替换
    c.drawString(470, height - 180, f"{time.strftime('%B %d, %Y')}")
    c.drawString(512, height - 195, f"NET {period}")

    c.setFillColor("#f7f7f7")
    c.rect(370, height - 230, 180, 30, fill=True, stroke=False)

    c.setFillColor("#000000")
    c.setFont("dengxian_bold",11)
    c.drawString(415, height - 218, "TOTAL :")
    #做替换
    c.setFont("dengxian",11)
    c.drawString(494, height - 218, f"{currency_sign}{total_with_tax}")

    c.rect(50, height - 405, width-100, 30, fill = True)

    c.setFillColor("#ffffff")
    c.setFont("dengxian_bold",11)
    c.drawString(100, height - 395, "Item")
    c.drawString(240, height - 395, "Service Period")
    c.drawString(340, height - 395, "Quantity")
    c.drawString(420, height - 395, "Unit Price")
    c.drawString(500, height - 395, "Amount")

    c.setStrokeColor("#f7f7f7")
    c.line(50, height - 435, width-50, height - 435)
    c.line(50, height - 470, width-50, height - 470)
    c.line(50, height - 505, width-50, height - 505)

    c.setFont("dengxian",11)
    c.setFillColor("#000000")
    #替换
    c.drawString(57.5, height-430, f"{package}")
    c.drawString(347.5, height-430, f"{quantity}")
    c.drawString(420, height-430, f"{price}")
    c.drawString(490, height-430, f"{currency_sign}{total_amount}")

    c.drawString(425, height-530, "Subtotal :")
    c.drawString(490, height-530, f"{currency_sign}{total_amount}")
    c.drawString(425, height-550, "Tax :")
    c.drawString(530, height-550, f"{tax}%")
    c.drawString(490, height-570, f"{currency_sign}{total_with_tax}")
    c.setFont("dengxian_bold",11)
    c.drawString(425, height-570, "Total :")

    c.drawString(50, height-650,"Notes :")
    c.drawString(50, height-695,"Terms :")

    c.setFont("dengxian",11)
    c.drawString(50, height-665,"For all invoice related questions, please contact our Accounts Receivable team at ar@yexsys.com. ")
    c.drawString(50, height-710,"Please send WIRE payments to: ")
    c.drawString(50, height-725,"BENEFICIARY NANE: Shanghai YexSys Technology Co., Ltd. ")
    c.drawString(50, height-740,"BENEFICIARY BANK: BANK OF COMMUNICATIONS SHANGHAI ")
    c.drawString(50, height-755,"BRANCH BANK ADDRESS: No188, Yinchengzhong Rd, Shanghai, China ")
    c.drawString(50, height-770,"SWIFT BIC: COMMCNSHSHI")
    c.drawString(50, height-785,"A/C NO#: 310066577013003715694")

    c.save()

if __name__ == "__main__":
    quotation_cn()
    quotation_en()


