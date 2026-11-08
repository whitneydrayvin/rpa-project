# Insert your python code here
#Imports
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

#1
while True:    
    directory = single_input("Please input the directory where the file is located.") + '\\'
    directory1 = directory + 'ProductData.xlsx'
    directory2 = directory + 'MasterInvoice.pdf'
    directory3 = directory + 'Order1.xlsx'
    directory4 = directory + 'Order2.xlsx'
    directory5 = directory + 'Order3.xlsx'
    directory6 = directory + 'Order4.xlsx'
    if os.path.exists(directory) == True and os.path.exists(directory1) == True and os.path.exists(directory2) == True and os.path.exists(directory3) == True and os.path.exists(directory4) == True and os.path.exists(directory5) == True and os.path.exists(directory6) == True:
        break
    else:
        show_message("Error", "The directory you entered does not exist or does not conatin the files. Please try again.")

#2
df1 = pd.read_excel(directory + 'ProductData.xlsx', 'Descriptions')
df2 = pd.read_excel(directory + 'ProductData.xlsx', 'Prices')
product_descriptions = {}
product_prices = {}
for i, rows in df1.iterrows():
    product_descriptions[rows['Unit number']] = rows.Description
for i, rows in df2.iterrows():
    product_prices[rows['Unit number']] = rows['Price per unit']

#3
def get_order_data(invoice_num, filename, directory):
    filepath = directory + filename
    file = openpyxl.load_workbook(filepath, data_only=True)
    sheet = file['OrderForm']

    update_dict = {}
    update_dict['InvoiceNum'] = invoice_num
    update_dict['Date'] = sheet['C9'].value.strftime('%m/%d/%Y')
    if pd.isnull(sheet['B24'].value) == False:
        update_dict['Total'] = str(f'${(product_prices[sheet['B24'].value] * sheet['C24'].value):,.2f}')
    if pd.isnull(sheet['B25'].value) == False:
        update_dict['Total'] = str(f'${(product_prices[sheet['B24'].value] * sheet['C24'].value) + (product_prices[sheet['B25'].value] * sheet['C25'].value):,.2f}')
    if pd.isnull(sheet['B26'].value) == False:
        update_dict['Total'] = str(f'${(product_prices[sheet['B24'].value] * sheet['C24'].value) + (product_prices[sheet['B25'].value] * sheet['C25'].value) + (product_prices[sheet['B26'].value] * sheet['C26'].value):,.2f}')
    if pd.isnull(sheet['B27'].value) == False:
        update_dict['Total'] = str(f'${(product_prices[sheet['B24'].value] * sheet['C24'].value) + (product_prices[sheet['B25'].value] * sheet['C25'].value) + (product_prices[sheet['B26'].value] * sheet['C26'].value) + (product_prices[sheet['B27'].value] * sheet['C27'].value):,.2f}')
    if pd.isnull(sheet['B28'].value) == False:
        update_dict['Total'] = str(f'${(product_prices[sheet['B24'].value] * sheet['C24'].value) + (product_prices[sheet['B25'].value] * sheet['C25'].value) + (product_prices[sheet['B26'].value] * sheet['C26'].value) + (product_prices[sheet['B27'].value] * sheet['C27'].value) + (product_prices[sheet['B28'].value] * sheet['C28'].value):,.2f}')
    update_dict['CustomerAddress1'] = str(sheet['B4'].value)
    update_dict['CustomerName'] = str(sheet['C14'].value)
    update_dict['CustomerCompany'] = str(sheet['C15'].value)
    update_dict['CustomerAddress2'] = str(sheet['B5'].value)
    update_dict['Price0'] = str(f'${product_prices[sheet['B24'].value]:,.2f}')
    update_dict['Quantity0'] = str(sheet['C24'].value)
    update_dict['Amount0'] = str(f'${product_prices[sheet['B24'].value] * sheet['C24'].value:,.2f}')
    if pd.isnull(sheet['B25'].value) == False:
        update_dict['Price1'] = str(f'${product_prices[sheet['B25'].value]:,.2f}')
        update_dict['Quantity1'] = str(sheet['C25'].value)
        update_dict['Amount1'] = str(f'${product_prices[sheet['B25'].value] * sheet['C25'].value:,.2f}')
    if pd.isnull(sheet['B26'].value) == False:
        update_dict['Price2'] = str(f'${product_prices[sheet['B26'].value]:,.2f}')
        update_dict['Quantity2'] = str(sheet['C26'].value)
        update_dict['Amount2'] = str(f'${product_prices[sheet['B26'].value] * sheet['C26'].value:,.2f}')
    if pd.isnull(sheet['B27'].value) == False:
        update_dict['Price3'] = str(f'${product_prices[sheet['B27'].value]:,.2f}')
        update_dict['Quantity3'] = str(sheet['C27'].value)
        update_dict['Amount3'] = str(f'${product_prices[sheet['B27'].value] * sheet['C27'].value:,.2f}')
    if pd.isnull(sheet['B28'].value) == False:
        update_dict['Price4'] = str(f'${product_prices[sheet['B28'].value]:,.2f}')
        update_dict['Quantity4'] = str(sheet['C28'].value)
        update_dict['Amount4'] = str(f'${product_prices[sheet['B28'].value] * sheet['C28'].value:,.2f}')
    update_dict['Description0'] = str(product_descriptions[sheet['B24'].value])
    if pd.isnull(sheet['B25'].value) == False:
        update_dict['Description1'] = str(product_descriptions[sheet['B25'].value])
    if pd.isnull(sheet['B26'].value) == False:
        update_dict['Description2'] = str(product_descriptions[sheet['B26'].value])
    if pd.isnull(sheet['B27'].value) == False:
        update_dict['Description3'] = str(product_descriptions[sheet['B27'].value])
    if pd.isnull(sheet['B28'].value) == False:
        update_dict['Description4'] = str(product_descriptions[sheet['B28'].value])

    custname = str(sheet['C14'].value)
    custmail = str(sheet['C13'].value)
    orderdate = sheet['C9'].value.strftime('%m/%d/%Y')
    return update_dict, custname, custmail, orderdate

#4
def create_invoice(invoice_num,order_data,directory):
    input_pdf = directory + 'MasterInvoice.pdf'
    output_pdf = directory + f'Invoice {invoice_num}.pdf'
    create_pdf(input_pdf, output_pdf, order_data)

#5
    gmail_user = "drayvin.m.whitney@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f'''Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech CO.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com'''

    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)        

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()

#6
files = os.listdir(directory)
invoice_num = 8515

for file in files:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1