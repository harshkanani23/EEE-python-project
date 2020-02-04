# importing required packages for this program

import xlwt                         # xlwt package is used for connection of excel and data entered
from xlwt import Workbook
import pandas as pd                 # panda is used for data science and making any dataframe
import xlrd                         # for storing data in excel sheet
import os
import math                         # for doing mathematical operation
import re                           # re = regular expression


# Workbook is created
wb = Workbook()


# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')


# Dictinary of Data of devices and brand

# data is tentative for project purpose
d = {
    'devices' : ['fan', 'tv', 'pc', 'laptop', 'vacuumcleaner', 'tubelight', 'refrigerator', 'washingmachine', 'ac', 'oven', 'heater', 'iron'],
    'fan' : ['havells', 'crompton', 'usha', 'bajaj', 'orient'],
    'tv' : ['lg', 'samsung', 'mi', 'sony', 'oneplus'],
    'pc' : ['samsung', 'dell', 'hp', 'apple', 'lenovo'],
    'laptop': ['dell', 'macbook', 'hp', 'lenovo', 'asus'],
    'vacuumcleaner': ['eurekaforbes', 'irobot', 'electroflux', 'honeywell', 'sanitare'],
    'refrigerator': ['samsung', 'lg', 'haier', 'bosch', 'whirpool'],
    'tubelight': ['havells', 'philips', 'hpl', 'bajaj' 'syska' ],
    'washingmachine': ['ifb', 'bosch', 'samsung', 'lg', 'whirpool'],
    'ac' : ['haier', 'mitsubisi', 'ogeneral', 'lg', 'samsung'],
    'oven': ['lg', 'samsung', 'ifb', 'kenstar', 'onida'],
    'heater': ['jaguar', 'bajaj', 'crompton', 'venus', 'havells'],
    'iron' : ['philips', 'bajaj'],
    'tv_vol': ['55', '22', '14', '20', '17'],
    'tv_crr' : ['12', '10', '13', '10', '14'],
    'fan_vol': ['5', '12', '24', '32', '41'],
    'fan_crr' : ['1', '2', '3', '4', '2'],
    'pc_vol' : ['15', '22', '14', '20', '17'],
    'pc_crr' : ['12', '10', '13', '10', '14'],
    'laptop_vol' : ['22', '24', '27', '24', '41'],
    'laptop_crr' : ['3', '2', '3', '2', '7'],
    'vacuumcleaner_vol' : ['15', '22', '14', '20', '17'],
    'vacuumcleaner_crr' : ['12', '10', '13', '10', '14'],
    'refrigerator_vol' : ['5', '12', '24', '32', '41'],
    'refrigeratot_crr' : ['1', '2', '3', '4', '2'],
    'tubelight_vol' : ['15', '22', '14', '20', '17'],
    'tubelight_crr' : ['12', '10', '13', '10', '14'],
    'washingmachine_vol' : ['15', '12', '24', '32', '41'],
    'washingmachine_crr' : ['7', '5', '13', '13', '22'],
    'ac_vol' : ['15', '22', '14', '20', '17'],
    'ac_crr' : ['12', '10', '13', '10', '14'],
    'oven_vol' : ['55', '52', '44', '32', '41'],
    'oven_crr' : ['11', '23', '32', '24', '28'],
    'heater_vol' : ['15', '22', '14', '20', '17'],
    'heater_crr' : ['12', '10', '13', '10', '14'],
    'iron_vol' : ['19', '22', '25', '32', '41'],
    'iron_crr' : ['11', '9', '10', '17', '18'],
    }


# dictionary for units in different state

state_units = {'gujarat':'1.5', 'tamilnadu': '1.75', 'punjab': '2', 'jammukashmir': '2.5',
               'Delhi':'1.35', 'maharashtra': '1.21', 'karnataka': '1.5', 'westbengal':'1.4',
               'haryana': '1.1', 'madhyapradesh': '1.4', 'uttarpradesh': '0.9', 'orrisa': '1.0',
               'kerala': '1.23', 'bihar': '0.9'}



# For which project you want to estimate a electricity bill

print("Enter the name of project from [HOME], [COMMERCIAL], [SCHOOL], [COLLEGE], [SHOPPING MALL]")

project = input("Enter the name of project: ")
state = input('Enter the name of state: ')

if project == 'home':
    bed = input("Enter the total number of BHK: ")
if project == 'commercial':
    offices = input("Enter the total number of shops: ")
if project == 'school':
    class_room = input("Enter the total number of classrooms: ")
if project == 'college':
    class_room1 = input("Enter the total number of classrooms: ")
if project == 'Shopping Mall':
    shops = input('Enter the total number of shops: ')


# How many devices you want to add
total_devices = input("Enter Total Number of Devices you want to add: ")



# storing heading in excel
sheet1.write(0, 1, 'Brand')
sheet1.write(0, 0, 'Device')
sheet1.write(0, 2, 'Quantity')
sheet1.write(0, 3, 'Voltage')
sheet1.write(0, 4, 'Current')
sheet1.write(0, 5, 'Power')
sheet1.write(0, 6, 'Hour/day')
sheet1.write(0, 7, 'Unit')
sheet1.write(0, 8, 'Total bill')
sheet1.write(0, 9, 'Remarks')               # remarks is for optimization of bill by using less
                                            #  consuming device suggested by program



# algorithm
for i in range(int(total_devices)):
  d_name = input("Enter the name of device: ")
  device_name = sheet1.write(i+1, 0, d_name )
  b_name = input("Enter the name of brand: ")
  brand = sheet1.write(i+1, 1, b_name )
  l = d['devices']
  if d_name in l and b_name in d[d_name]:
    v = d_name + '_vol'
    current = d_name + '_crr'
    index = d[d_name].index(b_name)
    V = d[v][index]
    I = d[current][index]
  else:
    V = input('Enter the voltage of device:')
    I = input('Enter the current of device:')

  qua = input("Enter the Quantity: ")
  quantity = sheet1.write(i+1, 2, qua)
  vlt = sheet1.write(i + 1, 3, V)
  crt = sheet1.write(i + 1, 4, I)
  usage = input('Enter your usage as hour/day: ')
  use = sheet1.write(i+1, 6, usage)
  power = int(V) * int(I)
  pwr = sheet1.write(i+1, 5, power)
  units = (power * int(usage) * 30 * int(qua)) / 1000
  total_units = sheet1.write(i+1, 7, units)
  bill = units * float(str(state_units[state]))
  total_bill = sheet1.write(i+1, 8, bill)
  li = d[v]
  v1 = li.index(min(li))
  sheet1.write(i+1, 9, d[d_name][v1])


# saving data in excel
wb.save('xlwt example.xls')

# creating a data frame in python terminal
df = pd.read_excel('xlwt example.xls')
print(df)
