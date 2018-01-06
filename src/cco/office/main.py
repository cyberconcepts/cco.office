#
#  Copyright (C) 2017 cyberconcepts.org team
#
#  See the file LICENSE for license details
#

"""
Process data using a LibreOffice file.
"""

import csv
from logging import getLogger
import sys
import pyoo

from cco.office import launch, pyoox


logger = getLogger('cco.office')

timeFactor = 24 * 60


def run(dataPath='reportdata.csv', 
        tplPath='template.ods', 
        resultPath='report.ods'):
    desktop = launch.getUnoDesktop()
    logger.info('Desktop: %s', desktop)
    with open(dataPath, newline='', encoding='ISO8859-15') as datafile:
        data = list(csv.reader(datafile, delimiter=';'))[1:]
        result = process(desktop, data, tplPath, resultPath)


def process(desktop, data, tplPath, resultPath):
    doc = desktop.open_spreadsheet(tplPath)
    sheet = doc.sheets[0]
    #dayCells = pyoox.getCellsByName(sheet, 'day')
    #print('***', dayCells.address.col, dayCells.address.col_count)
    #doc.sheets.copy(sheet.name, sheet.name + '-copy', 0)
    startRow = 4
    pyoox.insertRows(sheet, startRow, len(data))
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = sheet[ri + startRow, ci]
            #print('***', repr(cell.number_format))
            if val.isdigit():
                if cell.number_format == 40:
                    value = int(val) / timeFactor
                else:
                    value = int(val)
            else:
                value = val
            cell.value = value
    doc.save(resultPath) #, pyoo.FILTER_EXCEL_2007)
    doc.close()
    return 0


if __name__ == '__main__':
    print(sys.argv)
    run(*sys.argv[1:])
