###! /env/bin/python3

"""
Tests for the 'cco.office' package.
"""

import logging
import os
import sys
import unittest, doctest

from cco.office import main


logger = logging.getLogger('cco.office')
logger.level = logging.DEBUG


class Test(unittest.TestCase):
    "Basic tests."

    def testBasicStuff(self):
        main.run()
        #main.run('testdata.csv', 'test.xlsx', 'testtpl.ods')

def test_suite():
    flags = doctest.NORMALIZE_WHITESPACE | doctest.ELLIPSIS
    stream_handler = logging.StreamHandler(sys.stdout)
    logger.addHandler(stream_handler)
    return unittest.TestSuite((
        unittest.makeSuite(Test),
        doctest.DocFileSuite('README.md', optionflags=flags),
        ))


if __name__ == '__main__':
    unittest.main(defaultTest='test_suite')
