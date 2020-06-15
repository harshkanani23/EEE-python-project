from traceback import print_exc
# Returns a dictinary of data of devices and brands.
# Note: Data is tentative for project purpose.


def getData(filename):
    try:
        # Import relevant modules.
        from json import load
        data = ""
        with open(filename, 'r') as file:
            data = load(file)
        return data
    except FileNotFoundError:
        print("File not found!")
    except:
        print("Unable to load data!")
        print_exc()


def printDataFrame(filename):
    try:
        # Import relevant modules to read the excel file.
        from pandas import read_excel
        command = 'xlwt '+filename
        print(read_excel(command))
    except ModuleNotFoundError:
        print("Pandas not  Installed!")
    except FileNotFoundError:
        print("File doesn't exist! \nCreate file first!")
    except:
        print_exc()
    finally:
        print("Unable to print data from Excel file.")


def writeToFile(wb, filename):
    try:
        command = 'xlwt '+filename
        from xlwt import Workbook
        wb.save(command)
    except ModuleNotFoundError:
        print("xlwt not installed!")
    except:
        print_exc()
    finally:
        print("Unable to save!")


def createWorkbook():
    try:
        # xlwt package is used for connection of excel and data entered.
        from xlwt import Workbook
        return Workbook()
    except ModuleNotFoundError:
        print("xlwt not installed!")
    except:
        print_exc()
    finally:
        print("Unable to create Workbook")
