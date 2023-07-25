import os
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

def get_mandays(type, year, month, vName, sName):
    return f'Mandays\\{type}\\{year}\\{month}\\{vName} - {sName}.xlsx'

def get_wage_template():
    return f'customTemplate\\wageRate\\customWageRate.xlsx'

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