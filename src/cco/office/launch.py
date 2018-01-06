#
#  Copyright (C) 2017 cyberconcepts.org team
#
#  See the file LICENSE for license details
#

"""
Start LibreOffice background process.
"""

import logging
from subprocess import call, Popen
import time

import pyoo


logger = logging.getLogger('cco.office')


command = ' '.join([
        'soffice', '--accept="socket,host=localhost,port=2002;urp;"',
        '--pidfile=var/pid/office.pid',
        '--norestore', '--nologo', '--nodefault', '--headless',
        '&'])


def getUnoDesktop():
    desktop = None
    count = 0
    maxRetry = 5
    while desktop is None and count <= maxRetry:
        try:
            logger.info('Connecting to Office application, count = %i.', count)
            desktop = pyoo.Desktop('localhost', 2002)
        except OSError:
            if count == 0:
                proc = call(command, shell=True)
                logger.info('Office application started: %s.', proc)
            time.sleep(1)
            count += 1
    if desktop is None:
        logger.error('Office connection could not be established.')
    return desktop

