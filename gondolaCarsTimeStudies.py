# Importing the openpyxl module
import openpyxl

# Setting the workbook to work in
wb = openpyxl.load_workbook('gondolaCarTimeStudies.xlsx')

# Setting worksheet
sheet = wb.active
sheet.title = 'GondolaCarsSummary'

# The following is the data list for different orders
orderList = {redacted} # confidential information (see local drive)

# Storing order values in a new list
new_order_list = []
for key in orderList:
    new_order_list.append(key)

# Writing order list to the first row in Excel
sheet['A1'] = 'Orders'
for i in range(len(new_order_list)):
    sheet['A' + str(i+2)] = new_order_list[i]

# Writing Timestudies conducted to the second row in Excel
sheet['B1'] = 'Timestudies Performed'
for i in range(len(new_order_list)):
    ts = orderList[new_order_list[i]]['timestudies']
    sheet['B' + str(i+2)] = str(ts)

sheet['C1'] = 'Type of Car'
for i in range(len(new_order_list)):
    typeOfCar = orderList[new_order_list[i]]['capacity']
    sheet['C' + str(i+2)] = str(typeOfCar)

# Saving workbook
wb.save('gondolaCarTimeStudies.xlsx')
