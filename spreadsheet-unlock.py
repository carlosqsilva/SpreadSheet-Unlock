#!/usr/bin/env python3

from glob import glob
import zipfile
import tempfile
import os
import re

patterns = [
    r'<sheetProtection .*selectLockedCells="1"/>',
    r'<sheetProtection .*scenarios="1"/>',
    r'<sheetProtection .*autoFilter="0"/>',
    r'<sheetProtection .*formatColumns="0"/>',
    r'<sheetProtection .*formatRows="0"/>'
    ]

path = 'xl/worksheets/sheet'


def find_all_spreadsheets():
    sheets = glob('*.xlsx')
    return sheets


def unlock(file):

    for pattern in patterns:
        busca = re.findall(pattern, file)
        if len(busca) == 1:
            break
    if len(busca) == 0:
        return file
    
    file = file.replace(busca[0], '')
    return file


def main(sheetname):
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(sheetname))
    os.close(tmpfd)

    with zipfile.ZipFile(sheetname, 'r') as sheet:
        with zipfile.ZipFile(tmpname, 'w') as sheetunlock:

            for item in sheet.infolist():
                if path not in item.filename:
                    sheetunlock.writestr(item, sheet.read(item.filename))

                if path in item.filename:
                    content = sheet.read(item.filename)
                    content = content.decode('utf-8')
                    unlocked = unlock(content)
                    sheetunlock.writestr(item, unlocked)

    newname = sheetname.split('.')
    newname = newname[:-1] + ['UNLOCK'] + newname[-1:]
    newname = '.'.join(newname)
    os.rename(tmpname, newname)


if __name__ == '__main__':
    sheets = find_all_spreadsheets()
    for sheet in sheets:
        print(sheet, '.....', 'OK')
        main(sheet)

