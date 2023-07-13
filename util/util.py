import os

def create_or_open_folder(fPath):
    if os.path.exists(fPath):
        return fPath
    else:
        os.makedirs(fPath)
        return fPath

def get_custom_mandays():
    return 'customTemplate\\mandaysClaimed\\customTemplateMandays.xlsx'

def save_mandays_claimed(year, month, vName, sName):
    claimedfolder = f'Mandays\Claimed\{year}\{month}'
    claimedPath = create_or_open_folder(claimedfolder)
    return f'{claimedPath}\{vName} - {sName}.xlsx'