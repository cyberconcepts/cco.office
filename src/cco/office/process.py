#
#  Copyright (C) 2017 cyberconcepts.org team
#
#  See the file LICENSE for license details
#

"""
Process data using a LibreOffice file.
"""

import os
from subprocess import run, Popen, STDOUT, PIPE
import sys
import time

import pyoo


command = ['soffice', '--accept="socket,host=localhost,port=2002;urp;"',
           '--norestore', '--nologo', '--nodefault', '--headless', '&']
command = ('soffice --accept="socket,host=localhost,port=2002;urp;" '
           '--pidfile=var/pid/office.pid '
           '--norestore --nologo --nodefault --headless &')

def run():
    desktop = None
    try:
        desktop = pyoo.Desktop('localhost', 2002)
    except OSError:
        print('*** Info: Starting office application.')
        with Popen(command, stdout=PIPE, stderr=PIPE, shell=True) as proc:
            #print(proc.stdout.read())
            #print(proc.stderr.read())
            print('*** Info: Office application started:', proc)
            time.sleep(5)
        try:
            desktop = pyoo.Desktop('localhost', 2002)
        except OSError:
            print('*** Error: Office connection could not be established.')
    print('***', desktop)


if __name__ == '__main__':
    run()
