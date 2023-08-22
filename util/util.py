import os, sys
import datetime

def create_or_open_folder(fPath):
    if os.path.exists(fPath):
        return fPath
    else:
        os.makedirs(fPath)
        return fPath

def get_custom_template(type):
    if type == 'Mandays':
        return 'customTemplate\\mandaysClaimed\\customTemplateMandays.xlsx'
    elif type == 'Bill':
        return 'customTemplate\\vendorBill\\customTemplateBill.xlsx'
    else:
        print('NO FILE FOUND')

def save_mandays(type, year, month, vName, sName):
    typefolder = f'Mandays\{type}\{year}\{month}'
    typePath = create_or_open_folder(typefolder)
    return f'{typePath}\{vName} - {sName}.xlsx'

def save_bill(year, month, vName, sName):
    billfolder = f'Bills\{year}\{month}'
    billPath = create_or_open_folder(billfolder)
    return f'{billPath}\{vName} - {sName}.xlsx'

def get_mandays(year, month, vName, sName):
    return f'Mandays\\Active\\{year}\\{month}\\{vName} - {sName}.xlsx'

def get_wage_template():
    return f'customTemplate\\wageRate\\customWageRate.xlsx'

def get_estimate_path(year, month, vName, sName):
    try:
        destination_folder = os.path.dirname(f'salaryEstimate\\{year}\\{month}\\{vName} - {sName}.xlsx')
        
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
            return f'salaryEstimate\\{year}\\{month}\\{vName} - {sName}.xlsx'
        else:
            return f'salaryEstimate\\{year}\\{month}\\{vName} - {sName}.xlsx'
    except Exception as e:
            print(f"An error occurred: {e} at line {sys.exc_info()[-1].tb_lineno}")

def get_attendance_path(year, month, vName, sName):
    try:
        destination_folder = os.path.dirname(f'attendanceRecord\\{year}\\{month}\\{vName} - {sName}.xlsx')
        
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
            return f'attendanceRecord\\{year}\\{month}\\{vName} - {sName}.xlsx'
        else:
            return f'attendanceRecord\\{year}\\{month}\\{vName} - {sName}.xlsx'
    except Exception as e:
            print(f"An error occurred: {e} at line {sys.exc_info()[-1].tb_lineno}")

def count_days(year, month):
    weekdays = 0
    sundays = 0

    # Get the first day of the month
    first_day = datetime.date(year, month, 1)

    # Get the last day of the month
    if month == 12:
        last_day = datetime.date(year + 1, 1, 1)
    else:
        last_day = datetime.date(year, month + 1, 1)

    # Iterate through the days of the month
    current_day = first_day
    while current_day < last_day:
        # 0 - Monday, 6 - Sunday
        if current_day.weekday() == 6:
            sundays += 1
        elif current_day.weekday() <= 5:
            weekdays += 1
        current_day += datetime.timedelta(days=1)

    return weekdays, sundays

def get_spi_claim(nonBusSale, busSale, mandays):
    spi = (nonBusSale + (busSale/2))/mandays
    if spi <= 300:
        return 0
    elif spi >= 301 and spi <= 350:
        return 7000
    elif spi >= 351 and spi <= 400:
        return 9000
    elif spi >= 401 and spi <= 450:
        return 11000
    elif spi >= 451 and spi <= 500:
        return 13000
    else:
        return 15000
    
    # if busSale == 0:
    #     result = 0
    # else:
    #     spi = (nonBusSale + (busSale / 2)) / mandays

    #     if spi <= 300:
    #         result = 0
    #     elif 301 <= spi <= 350:
    #         result = 7000
    #     elif 351 <= spi <= 400:
    #         result = 9000
    #     elif 401 <= spi <= 450:
    #         result = 11000
    #     elif 451 <= spi <= 500:
    #         result = 13000
    #     else:
    #         result = 15000

    # print("Result:", result)
