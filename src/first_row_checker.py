"""
@python version:
    Python 3.4

@summary:
    Control composed by the class FirstRowChecker().
    That class has the propose to check if the first four columns of the
        first row are 'latitude', 'longitude', 'name' and 'description'.
    The control has two public functions:
        - __init__(firt_row)
        - check()

@note:
    function __init__(first_row):
        Class constructor.
        Take as parameters "first_row", a list, the first row of the EXEL
        file, which we want to make a KMZ file from.

    function check():
        Returns true if the first four columns of the first row are 'latitude'
        'longitude', 'name' and 'description'.

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (12/04/2014):
        - Creation of the Class MotherControl(), and his functions:
            - __init__
            - check

    1.1 (09/12/2014):
        - Added docstrings to the class and functions.

@since:
    12/04/2014
"""


class FirstRowChecker(object):
    """
    Verifica se as primeiras 4 colunas do Excel
    são: latitude, longitude, name e description
    """

    def __init__(self, first_row):
        """
        first_row é uma lista
        """
        self.first_row = first_row

    def check(self):
        low = [x.lower() for x in self.first_row]
        eng = ['latitude', 'longitude', 'name', 'description']
        por = ['latitude', 'longitude', 'nome', 'descricao']
        ptr = ['latitude', 'longitude', 'nome', 'descrição']

        if low[0:4] == eng or low[0:4] == por or low[0:4] == ptr:
            return True
        else:
            return False
