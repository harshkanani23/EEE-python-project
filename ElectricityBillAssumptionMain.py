# Importing required modules.
from os import system
from traceback import print_exc

try:
  from UtilityCore import getData, printDataFrame, writeToFile, createWorkbook
except ModuleNotFoundError:
  print("UtilityCore.py not found!") 

# File where data will be stored.
filename = 'ElectricityBill.xls'

# Import relevant data from .json files.
brandsdevices = getData('BrandDevices.json')
state_units = getData('StateUnits.json')

def start():
  print('**************** WELCOME ****************')
  # Create a workbook to work with.
  try:
    wb = createWorkbook()
    # Add sheet to Workbook.
    sheet1 = wb.add_sheet('Sheet 1')
    
    # For which project you want to estimate a electricity bill

    print("\nEnter the name of project from [HOME], [COMMERCIAL], [SCHOOL], [COLLEGE], [SHOPPING MALL]")

    project = input("\nEnter the name of project: ").lower()
    state = input('Enter the name of state: ').lower()

    if project == 'home':
      bed = input("Enter the total number of BHK: ")
    if project == 'commercial':
      offices = input("Enter the total number of shops: ")
    if project == 'school':
      class_room = input("Enter the total number of classrooms: ")
    if project == 'college':
      class_room1 = input("Enter the total number of classrooms: ") 
    if project == 'Shopping Mall':
      shops = input("Enter the total number of shops: ")

    # Total number os devices
    total_devices = input("Enter Total Number of Devices you want to add: ")

    # Storing heading in excel
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
                                                # consuming device suggested by program



    # Algorithm to calculate optimal devices.
    for i in range(int(total_devices)):
      _name = input("Enter the name of device: ")
      device_name = sheet1.write(i+1, 0, _name )
      b_name = input("Enter the name of brand: ")
      brand = sheet1.write(i+1, 1, b_name )
      if _name in brandsdevices['devices'] and b_name in brandsdevices[_name]:
        v = _name + '_vol'
        current = _name + '_crr'
        index = brandsdevices[_name].index(b_name)
        V = brandsdevices[v][index]
        I = brandsdevices[current][index]
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
      #li = brandsdevices[v]
      #v1 = li.index(min(li))
      #sheet1.write(i+1, 9, brandsdevices[_name][v1])


    # Saving data in excel
    writeToFile(wb, filename)

    # Print data in python emulator.
    printDataFrame(filename)
  except:
    print_exc()

system('cls')
start()
