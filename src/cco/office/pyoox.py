#
#  Copyright (C) 2017 cyberconcepts.org team
#
#  See the file LICENSE for license details
#

"""
Some useful extension to the pyoo library.
"""

import pyoo


def getCellsByName(sheet, name):
    rawCells = sheet._target.getCellRangeByName(name)
    addr = rawCells.getRangeAddress()
    return pyoo.TabularCellRange(sheet, pyoo.SheetAddress._from_uno(addr))


def insertRows(sheet, pos, numRows):
    rows = sheet._target.Rows
    rows.insertByIndex(pos, numRows)
