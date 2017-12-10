#
#  Copyright (C) 2017 cyberconcepts.org team
#
#  See the file LICENSE for license details
#

"""
Process data using a LibreOffice file.
"""

import pyoo


command = ('soffice --accept="socket,host=localhost,port=2002;urp;" '
           '--norestore --nologo --nodefault --headless')

def run():
    try:
        desktop = pyoo.Desktop('localhost', 2002)
    except OSError:
        print('*** Error: Office connection could not be established.')


if __name__ == '__main__':
    run()
