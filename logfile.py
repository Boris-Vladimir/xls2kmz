"""
@python version:
    Python 3.4

@summary:
    Control composed by the LogFilel() class.
    This class has a constructor, __init__() and more two public
    functions:
        - write(msg, level=logging.INFO);
        - def flush();
    

@note:
    I don't understand very well how that class works, only that
    	works.
    See: http://stackoverflow.com/questions/616645/how-do-i-duplicate-
    	sys-stdout-to-a-log-file-in-python/3423392#3423392

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (05/01/2014):
        - Implementation of the LogFile() class.
        - Creation of the function / class attribute __init__(), write(),
        	and flush()

@since:
    05/01/2014
"""
import logging  # to create a logfile

class LogFile(object):
    """
    File-like object to log text using the `logging` module.

    http://stackoverflow.com/questions/616645/how-do-i-duplicate-
    sys-stdout-to-a-log-file-in-python/3423392#3423392
    """

    def __init__(self, name=None):
        self.logger = logging.getLogger(name)

    def write(self, msg, level=logging.INFO):
        self.logger.log(level, msg)

    def flush(self):
        for handler in self.logger.handlers:
            handler.flush()