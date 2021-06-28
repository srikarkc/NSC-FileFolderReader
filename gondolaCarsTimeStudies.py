# orderList = [P420021, P417020, P413520, P414520, P410520, P411020, P796018, P407519, P409519, P405519, P406019, P795518, P777017, P775517, P403019, P780517, P751515, P751015]

# Importing the openpyxl module
import openpyxl

# Setting the workbook to work in
wb = openpyxl.load_workbook('gondolaCarTimeStudies.xlsx')

# Setting worksheet
sheet = wb.active
sheet.title = 'GondolaCarsSummary'

# The following is the data list for different orders
orderList = {'P7510-15': {'customer': 'AIM', 'capacity': '6400 cu. ft.', 'timestudies': ['End Wall Assy', 'End Wall Sub Assy', 'Top Chord Assy, Side', 'Line Summary from Previous Studies']}, 'P7515-15': {'customer': 'Tunnel Hill', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P7805-17' : {'customer': 'Murphy Road Recycling', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P4030-19' : {'customer': 'Trojan Recycle', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P7755-17' : {'customer': 'Cardella Waste', 'capacity': '6400 cu. ft.', 'timestudies': ['Bolster Assy', 'Bolster Assy w/ E/L Ext.']}, 'P7770-17' : {'customer': 'Tunnel Hill', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P7955-18' : {'customer': 'AIM', 'capacity': '6400 cu. ft.', 'timestudies': ['Apply Ends', 'Apply Sides', 'Center Sill Assy', 'End Sheet Assy', 'End Wall Assy', 'Side Lines', 'Skid', 'Underframe', 'Weld Exterior']}, 'P4060-19' : {'customer': 'Frost Bridge', 'capacity': '6400 cu. ft.', 'timestudies': ['Center Sill Assy', 'End Wall Assy']}, 'P4055-19' : {'customer': 'Cardella Waste', 'capacity': '6400 cu. ft.', 'timestudies': ['Underframe']}, 'P4095-19' : {'customer': 'Midwest Railcar', 'capacity': '6400 cu. ft.', 'timestudies': ['Safety Appliance', 'Skid']}, 'P4075-19' : {'customer': 'Residco', 'capacity': '6400 cu. ft.', 'timestudies': ['Apply Ends', 'Apply Sides', 'OK Position', 'Side Lines', 'Side Post Extension Assy', 'Underframe', 'Weld Exterior']}, 'P7960-18': {'customer': 'Aim', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P411020' : {'customer': 'Midwest Rail', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P410520' : {'customer': 'PNC', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P414520' : {'customer': 'Midwest Railcar', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P413520' : {'customer': 'Frost Bridge', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P417020' : {'customer': 'Enkay for Gerdau Long Steel North America', 'capacity': '6400 cu. ft.', 'timestudies': None}, 'P420021' : {'customer': 'Enkay for Gerdau', 'capacity': '6400 cu. ft.', 'timestudies': None}}

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
