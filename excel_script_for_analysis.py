#import necessary libraries

import pandas as pd
import numpy as np
import os
import xlsxwriter
from openpyxl import load_workbook
from pandas import ExcelWriter
import matplotlib.pyplot as plt
import matplotlib.style as style
style.use('fivethirtyeight')

print('Ensure file editing is activated')
print('For clarity, ensure file is in a near-empty folder')

file = input('Select file: ')

#Load file into pandas df
data = pd.read_excel(file)

#Set the column names to the first row
data.columns = data.iloc[0]

#Remove the first row of the data
data = data[1:]

#Return only paid and completed orders
lists = ['Completed', 'Paid']
data = data[data[' Order Status '].isin(lists)]

#Strip column names off white space for easy analysis
columns = data.columns
columns = list(columns)
data.columns = [x.strip() for x in columns]
data.columns

#Set two columns to datetime objects
data["Purchase Time"] = pd.to_datetime(data["Purchase Time"])
data['Payment Time'] = pd.to_datetime(data['Payment Time'])

#Return the long date format that is easily processed with Excel
data['Purchase Date'] = data['Purchase Time'].dt.date
data['Payment Date'] = data['Payment Time'].dt.date

#Delete specific columns that are unnecessary
del (data['Comments'], data['Shipping Address'], data['Agent Address'], data['Payment Time'], data['Purchase Time'])

#Order the columns in a desired way
data = data[['Purchase Date', 'Payment Date', 'Member ID', 'Member Name', 'Order Flow', 'Order Type',
       'Order Status', 'total amount', 'discount amount', 'Total PV',
       'Amount due', 'Actual Payment Amount', 'Invoice Number',
       'Logistics Status', 'product name', 'product quantity', 'Unit Price',
       'PV', 'Actual payment unit price']]

#Extract the filename and extension
filename = os.path.splitext(file)[0]
extension = os.path.splitext(file)[1]
pth = os.path.dirname(file)
new_file = os.path.join(pth, filename+'_edited'+extension)
writer = pd.ExcelWriter(new_file, engine = 'openpyxl')

#Create new sheets in the new file created above
data.to_excel(writer, sheet_name = 'Order by Invoice', index = False)
data.to_excel(writer, sheet_name = 'Order by Product', index = False)
data.to_excel(writer, sheet_name = 'member pivot', index = False)
data.to_excel(writer, sheet_name = 'Top products', index = False)
data.to_excel(writer, sheet_name = 'Bottom products', index = False)
data.to_excel(writer, sheet_name = 'Category Pivot', index = False)
data.to_excel(writer, sheet_name = 'Fan-AC Pivot', index = False)
data.to_excel(writer, sheet_name = 'Fan-AC Pivot', index = False)
data.to_excel(writer, sheet_name = 'Lighting Pivot', index = False)

writer.save()

#Save two of the sheets into two different dataframes
obi = pd.read_excel(new_file, sheet_name="Order by Invoice")
obp = pd.read_excel(new_file, sheet_name="Order by Product")

#Delete the columns that are unnecessary for each sheet
del (obi['product name'], obi['product quantity'], obi['Unit Price'], obi['Actual payment unit price'], obi['PV'])
del (obp['Amount due'])

#Set the columns to float type object after removing the currency sign from each of the specified columns
obi['total amount'] = obi['total amount'].str[1:].astype('float64')
obi['Amount due'] = obi['Amount due'].str[1:].astype('float64')
obi['Actual Payment Amount'] = obi['Actual Payment Amount'].str[1:].astype('float64')

obp['Unit Price'] = obp['Unit Price'].str[1:].astype('float64')
obp['Actual Payment Amount'] = obp['Actual Payment Amount'].str[1:].astype('float64')
obp['Actual payment unit price'] = obp['Actual payment unit price'].str[1:].astype('float64')

#Perform some calculations on some specific columns
obi['discount amount'] = obi['total amount'] - obi['Actual Payment Amount']

obp['Total PV'] = obp['product quantity'] * obp['PV']
obp['total amount'] = obp['product quantity'] * obp['Unit Price']
obp['Actual Payment Amount'] = obp['product quantity'] * obp['Actual payment unit price']
obp['discount amount'] = obp['total amount'] - obp['Actual Payment Amount']

#Drop the rows with duplicates from the "Order flow" column, do this for only one dataframe
obi = obi.drop_duplicates(['Order Flow'])

#Order the columns in a desired way
obi = obi[['Purchase Date', 'Payment Date', 'Member ID', 'Member Name',
       'Order Flow', 'Order Type', 'Order Status', 'Invoice Number', 'Logistics Status', 'total amount',
       'Total PV', 'Amount due', 'Actual Payment Amount', 'discount amount']]

obp = obp[['Purchase Date', 'Payment Date', 'Member ID', 'Member Name',
       'Order Flow', 'Order Type', 'Order Status', 'Invoice Number', 'Logistics Status',
       'product name', 'product quantity', 'Unit Price', 'PV',
       'Actual payment unit price', 'Total PV', 'total amount',
       'discount amount', 
       'Actual Payment Amount']]

#Account for duplicate members who opened two accounts(e.g Haruna 1 and Haruna 2)
modified = obp.copy()
modified['Member Name'] = modified['Member Name'].str.replace('[0-9]+','').str.strip()

#Create Pivots

#Member Pivot
member_pv = modified.pivot_table(values=['Actual Payment Amount','product quantity'], 
                             index = 'Member Name', aggfunc=np.sum).sort_values(
                                                                    by = 'Actual Payment Amount', ascending = False)

#Product Pivot
prod_pv = modified.pivot_table(values='Actual Payment Amount', 
                                index='product name', aggfunc=np.sum).sort_values(by = 'Actual Payment Amount',
                                                                                    ascending = False
                                                                                  )
top_ten_prod = prod_pv.head(10)
bot_ten_prod = prod_pv.tail(10)

#Category pivot
modified['cat'] = modified['product name'].str[4:6]
modified['cat'] = modified['cat'].str.upper().str.strip()
cat_pv = modified.pivot_table(values=['Actual Payment Amount','discount amount'], 
                          index = 'cat', aggfunc=np.sum).sort_values(by = 'Actual Payment Amount',
                                                                     ascending = False                                 
                                                                    )

fan_ac = cat_pv.loc[['CF','GR'],:]
lighting_pv = cat_pv.loc[['OL','EL','SL','RL','LB'],:]

#Assign the result of these datatframes and pivots to their appropraite sheet and save

writer = ExcelWriter(new_file)
obi.to_excel(writer,'Order by Invoice')
obp.to_excel(writer,'Order by Product')
member_pv.to_excel(writer,'member pivot')
top_ten_prod.to_excel(writer,'Top products')
bot_ten_prod.to_excel(writer,'Bottom products')
cat_pv.to_excel(writer,'Category Pivot')
fan_ac.to_excel(writer,'Fan-AC Pivot')  
lighting_pv.to_excel(writer,'Lighting Pivot')

writer.save()

#Create directory for images
the_folder = os.path.join(pth, filename+'_images')
os.mkdir(the_folder)
os.chdir(the_folder)

#Top 5 Members
#Save the result of the pivot as an image
member_pv['Actual Payment Amount'].head(5).plot(kind = 'bar',
               title = 'Top Five Members by Sales',
               figsize = (16,12),
               legend = False,
               rot = 25,
                )

plt.grid()
plt.savefig('Member.png')

#Daily sales income image
by_date = modified.pivot_table(values=['Actual Payment Amount','product quantity'], 
                             index = 'Payment Date', aggfunc=np.sum).sort_index()

by_date['Actual Payment Amount'].plot(figsize=(16,12))
plt.grid()
plt.title('Total Sales Done Daily')
plt.axhline(max(by_date['Actual Payment Amount']), label='Maximum value', color='Red')
plt.legend()
plt.savefig('Sales Income by Date.png')

#Top 5 Products
top_ten_prod.head(5).plot(kind = 'bar',
               title = 'Top Five Selling Products',
               figsize = (16,12),
               colormap = plt.cm.BuPu_r,
               legend = False,
               rot = 25
                )
plt.xlabel('Product SKU')
plt.grid()
plt.legend()
plt.savefig('Best selling Products.png')

#bottom ten product image
bot_ten_prod.plot(kind = 'bar',
               title = 'Bottom Ten Selling Products',
               figsize = (16,12),
               legend = False,
               rot = 25,
               colormap = plt.cm.Reds_r
                )
plt.xlabel('Product SKU')
plt.grid()
plt.legend()
plt.savefig('Bottom selling Products.png')

#fan category image
fan_ac['Actual Payment Amount'].plot(kind = 'bar',
               title = 'Sales From Fan and AC Category',
               figsize = (16,12),
               legend = False,
               colormap = plt.cm.Blues_r,
               rot = 0
                )
plt.xlabel('Category')
plt.ylabel('Sales Income (Tens of milliions)')
plt.grid()
plt.savefig('Fan and AC.png')

#lighting image
lighting_pv.plot(kind = 'bar',
               title = 'Sales From Lighting Category',
               figsize = (16,12),
               legend = False,
               rot = 0,
               colormap = plt.cm.RdYlBu_r
                )
plt.xlabel('Category')
plt.ylabel('Income(Tens of millions)')
plt.grid()
plt.legend()
plt.savefig('lighting_cat.png')

print('Completed')
print('Excel workbook and folder with images in them have been created')
print('Thank you for using this program')
