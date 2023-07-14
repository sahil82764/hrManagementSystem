import os

def create_or_open_folder(fPath):
    if os.path.exists(fPath):
        return fPath
    else:
        os.makedirs(fPath)
        return fPath

def get_custom_mandays():
    return 'customTemplate\\mandaysClaimed\\customTemplateMandays.xlsx'

def save_mandays(type, year, month, vName, sName):
    typefolder = f'Mandays\{type}\{year}\{month}'
    typePath = create_or_open_folder(typefolder)
    return f'{typePath}\{vName} - {sName}.xlsx'