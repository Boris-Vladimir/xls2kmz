"""
@python version:
    Python 3.4

@summary:
    Control composed by the class MyThread().
    That class has the propose to subclass Thread and create subclass
    instance.
    The control has two public functions:
        - __init__(function, args, name='')
        - get_result()
        - run()

@note:
    function __init__(function, args, name=''):
        Class constructor.
        Take as parameters "function", a string, the name of the
        function to run as a thread, "args", a tuple of strings
        with the function arguments, and optionally a "name" for
        the tread.

    function get_result():
        Retrieves the value of self.res

    function run():
        Applies the arguments to the function running as thread

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0(13/12/2013):
        - Creation of the Class MyThread(), and his functions:
            - __init__
            - get_result
            - run
    1.1(14/12/2013):
        - Changed python version to 3.3
        - Changed the function run():
            self.res = apply(self.func, self.args) to
            self.res = self.func(*self.args)
    1.2(15/12/2013):
        - Added docstrings

@since:
    13/12/2013
"""

import threading


class MyThread(threading.Thread):
    '''Subclass Thread and Create subclass instance'''

    def __init__(self, function, args, name=''):
        '''
        Take as parameters "function", a string, the name of the
        function to run as a thread, "args", a tuple of strings
        with the function arguments, and optionally a "name" for
        the tread.
        '''
        threading.Thread.__init__(self)
        self.name = name
        self.func = function
        self.args = args

    def get_result(self):
        '''Retrieves the value of self.res'''
        return self.res

    def run(self):
        '''Applies the arguments to the function running as thread'''
        self.res = self.func(*self.args)
