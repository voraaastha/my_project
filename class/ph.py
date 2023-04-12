import requests
import json
from xlsxwriter import Workbook
import pandas as pd

new_url = 'https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY'

headers = {'User-Agent': 'Mozilla/5.0'}
page = requests.get(new_url,headers=headers)
d = json.loads(page.text)
option_data = d

header=['strikePrice','expiryDate','underlying','openInterest','changeinOpenInterest','lastPrice']

wb = Workbook('op.xlsx')
put_sheet = wb.add_worksheet("PE")
call_sheet = wb.add_worksheet("CE")
first_row = 0
for h in header:
    col = header.index(h)
    put_sheet.write(first_row,col,h)
    call_sheet.write(first_row,col,h)

row = 1
for i in option_data:
    for key,values in i.items():
        col = header.index(key)
        put_sheet(row,col,values)
    row+=1
wb.close()


