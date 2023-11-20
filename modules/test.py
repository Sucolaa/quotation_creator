import pandas as pd
import time

record = pd.read_csv("../data/record.csv", encoding = 'gbk')
new_number = (record['Quotation Date'] == time.strftime("%Y%m%d")).sum()
print(new_number)





