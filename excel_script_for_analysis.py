import pandas as pd
import os
import xlsxwriter
from openpyxl import load_workbook
from pandas import ExcelWriter


print('Ensure file editing is activated')
file = input('Select file: ')
data = pd.read_excel(file)
data.columns = data.iloc[0]
data = data[1:]
data[" Purchase Time "] = pd.to_datetime(data[" Purchase Time "])
data[' Payment Time '] = pd.to_datetime(data[' Payment Time '])
data['Purchase Date'] = data[' Purchase Time '].dt.date
data['Payment Date'] = data[' Payment Time '].dt.date
del (data[' Comments '], data[' Shipping Address '], data['Agent Address'], data[' Payment Time '], data[' Purchase Time '])
data = data[['Purchase Date', 'Payment Date', 'Member ID', 'Member Name', ' Order Flow ', ' Order Type ',
       ' Order Status ', ' total amount ', ' discount amount ', ' Total PV ',
       ' Amount due ', ' Actual Payment Amount ', ' Invoice Number ',
       ' Logistics Status ', 'product name', 'product quantity', 'Unit Price',
       'PV', 'Actual payment unit price']]


filename = os.path.splitext(file)[0]
extension = os.path.splitext(file)[1]
pth = os.path.dirname(file)
new_file = os.path.join(pth, filename+'_edited'+extension)
writer = pd.ExcelWriter(new_file, engine = 'openpyxl')
data.to_excel(writer, sheet_name = 'Order by invoice', index = False)
data.to_excel(writer, sheet_name = 'Order by Product', index = False)

writer.save()

obi = pd.read_excel(new_file, sheet_name="Order by invoice")
obp = pd.read_excel(new_file, sheet_name="Order by Product")
del (obi['product name'], obi['product quantity'], obi['Unit Price'], obi['Actual payment unit price'], obi['PV'])
del (obp[' Amount due '])

obi[' total amount '] = obi[' total amount '].str[1:].astype('float64')
obi[' Amount due '] = obi[' Amount due '].str[1:].astype('float64')
obi[' Actual Payment Amount '] = obi[' Actual Payment Amount '].str[1:].astype('float64')

obp['Unit Price'] = obp['Unit Price'].str[1:].astype('float64')
obp[' Actual Payment Amount '] = obp[' Actual Payment Amount '].str[1:].astype('float64')
obp['Actual payment unit price'] = obp['Actual payment unit price'].str[1:].astype('float64')

obi[' discount amount '] = obi[' total amount '] - obi[' Actual Payment Amount ']

obp[' Total PV '] = obp['product quantity'] * obp['PV']
obp[' total amount '] = obp['product quantity'] * obp['Unit Price']
obp[' Actual Payment Amount '] = obp['product quantity'] * obp['Actual payment unit price']
obp[' discount amount '] = obp[' total amount '] - obp[' Actual Payment Amount ']

obi = obi.drop_duplicates([' Order Flow '])

obi = obi[['Purchase Date', 'Payment Date', 'Member ID', 'Member Name',
       ' Order Flow ', ' Order Type ', ' Order Status ', ' Invoice Number ', ' Logistics Status ', ' total amount ',
       ' Total PV ', ' Amount due ', ' Actual Payment Amount ', ' discount amount ']]

obp = obp[['Purchase Date', 'Payment Date', 'Member ID', 'Member Name',
       ' Order Flow ', ' Order Type ', ' Order Status ', ' Invoice Number ', ' Logistics Status ',
       'product name', 'product quantity', 'Unit Price', 'PV',
       'Actual payment unit price', ' Total PV ', ' total amount ',
       ' discount amount ', 
       ' Actual Payment Amount ', ]]

writer = ExcelWriter(new_file)
obi.to_excel(writer,'Order by invoice')
obp.to_excel(writer,'Order by Product')
writer.save()

print('Completed')
print('Thank you for using this program')
