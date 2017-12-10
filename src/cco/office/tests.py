###! /env/bin/python3

"""
Tests for the 'cco.office' package.
"""

import os
import unittest, doctest

from cco.office import process


class Test(unittest.TestCase):
    "Basic tests."

    def testBasicStuff(self):
        process.run()


def test_suite():
    flags = doctest.NORMALIZE_WHITESPACE | doctest.ELLIPSIS
    return unittest.TestSuite((
        unittest.makeSuite(Test),
        doctest.DocFileSuite('README.md', optionflags=flags),
        ))


if __name__ == '__main__':
    unittest.main(defaultTest='test_suite')
