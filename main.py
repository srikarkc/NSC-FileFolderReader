# Importing the os and regex module
import os
import re

# Importing openpyxl
import openpyxl

# Creating a workbook
wb = openpyxl.Workbook()

# Setting worksheet
sheet = wb.active
sheet.title = 'GondolaCarsSummary'

# Get folder names
path = r'N:\common\Industrial Engineering\Car Orders\GONDOLA CARS'
order_list = os.listdir(path)

# Use the below line to get the # of orders
# print('Number of orders in the current directory: ' + str(len(order_list)))

# Regular expression to match order #
ord_list_regex = r'P\d{4}-\d{2}'

# Get order number from the list
order_numbers_list = []
for order in order_list:
    ord_number = re.match(ord_list_regex, order).group(0)
    order_numbers_list.append(ord_number)    
#print(order_numbers_list)

# Get car type from the list
car_type_list = []
for order in order_list:
    order_split_list = order.split('- ')
    car_type_list.append(order_split_list[1])
#print(car_type_list)

# Writing order list to the first row in Excel
sheet['A1'] = 'Orders'
sheet['C1'] = 'timestudiesPerformed'
for i in range(len(order_numbers_list)):
    sheet['A' + str(i+2)] = order_numbers_list[i]
    line_studies_path = path + r'\\' + order_list[i] + r'\Line Studies'
    line_studies = []
    if os.path.isdir(line_studies_path):
        line_studies_list = os.listdir(line_studies_path)
        sheet['C' + str(i+2)] = str(line_studies_list)

# Writing Car Type to the second row in Excel
sheet['B1'] = 'Car Type'
for i in range(len(car_type_list)):
    ts = car_type_list[i]
    sheet['B' + str(i+2)] = str(ts)

# Saving workbook as 'carTimeStudies.xlsx'
wb.save('carTimeStudies.xlsx')
