#
#  Copyright (C) 2017 cyberconcepts.org team
#
#  See the file LICENSE for license details
#

"""
Process data using a LibreOffice file.
"""

from logging import getLogger
import sys
import pyoo

from cco.office import launch, pyoox


logger = getLogger('cco.office')


def run(dataFn='reportdata.csv', resultFn='report.ods', tplFn='template.ods'):
    desktop = launch.getUnoDesktop()
    logger.info('Desktop: %s', desktop)
    data = {}
    result = process(desktop, data, resultFn, tplFn)


def process(desktop, data, resultPath, tplPath):
    doc = desktop.open_spreadsheet(tplPath)
    sheet = doc.sheets[0]
    dayCells = pyoox.getCellsByName(sheet, 'day')
    print('***', dayCells)
    doc.sheets.copy(sheet.name, sheet.name + '-copy', 0)
    doc.save(resultPath) #, pyoo.FILTER_EXCEL_2007)
    doc.close()
    return 0


if __name__ == '__main__':
    run()
