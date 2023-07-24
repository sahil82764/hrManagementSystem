import os

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