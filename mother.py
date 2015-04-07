"""
@python version:
    Python 3.4

@summary:
    Control composed by the class MotherControl().
    That class has the propose to call the xls.py and kml.py controls,
    making instances of XlsControl() and KmlControl() objects.
    The control has two public functions:
        - __init__(file_name)
        - exel_to_kml()

@note:
    function __init__(file_name):
        Class constructor.
        Take as parameters "file_name", a string, the name of the EXEL
        file, which we want to make a KMZ file from.

    function exel_to_kml():
        Makes an instance of a XlsControl() object and a list with all
        the data from the EXEL cells calling the read_exel() attribute of
        the created object.
        Then, makes an instance of a KmlControl() object, calls his
        attribute build_kml(), which returns another object, a Kml(),
        and, at last, call the save_kmz() attribute of the KmlControl()
        object passing the Kml() object as argument.

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (18/11/2013):
        - Creation of the Class MotherControl(), and his functions:
            - __init__
            - exel_to_kml

    1.1 (22/11/2013):
        - Added docstrings to the class and functions.

    1.2 (06/12/2013):
        - Translation of all comments to English and limitation of the
          maximum line length. Following the rules of the PEP 8,
          Style Guide for Python Code, writed by Guido van Rossum,
          Barry Warsaw and Nick Coghlan:
            - "Python coders from non-English speaking countries: please
              write your comments in English, unless you are 120 per
              cent sure that the code will never be read by people who
              don't speak your language."
            - "Limit all lines to a maximum of 79 characters. For
              flowing long blocks of text with fewer structural
              restrictions (docstrings or comments), the line length
              should be limited to 72 characters."
        - Changed the python version to 3.3

    1.3 (07/02/2014):
        - Added a call to a new class, CreateKMZ(), which helps to 
            rebuild the kmz file with the fotos inside his files folder

    1.4 (13/03/2014):
        - Added a call to the new class, ExcelSplitter(), which divides
            the Excel file if that are to big, to prevent memory leaks
        - Added a call to the new class, KmlJoiner(), to assemble in one
            Kml file the diferent kml files from the divided Excel file
        - Added a call to the new class, TempCleaner, to delete all temp
            files and folder used in the construction of the assembled Kmz

    1.5 (12/04/2014):
        - Added an if statement to confirm that the excel_list and data_list
            are of type list, if not breaks and return the excel_list or
            the data_list to be printed on the GUI interface (xls2kml.py)  

    1.6 (09/12/2014):
        - Commented some piece of code so it don't call rebuild_kmz() twice

@since:
    18/11/2013
"""

import os
from excel_spliter import ExcelSplitter
from xls import XlsControl
from kml import KmlControl
from create_kmz import CreateKMZ
from kml_joiner import KmlJoiner
from temp_cleaner import TempCleaner


class MotherControl(object):
    '''
    Calls the controls xls.py, kml.py and create_kmz.py, and makes
    instances of XlsControl(), KmlControl() and CreateKMZ() objects.
    '''

    def __init__(self, ficheiro, original_working_dir):
        '''
        str -> object MotherControl() object

        file_name is a string, the name of the EXEL file, which we want
        to make a KMZ file from.
        '''
        self.ficheiro = ficheiro
        self.original_working_dir = original_working_dir

    def exel_to_kml(self):
        '''
        None -> None

        Makes an instance of a XlsControl() object, by one attribute of
        that object build a list with all the data from the EXEL cells.
        Then, makes an instance of a KmlControl() object, by one of his
        attributes, makes an Kml() object, which is passed as an
        argument of another KmlControl() attribute to save a KMZ file in
        the drive.
        '''
        excel = ExcelSplitter(self.ficheiro)
        excel_list = excel.splitter()
        # -------------------------------
        if type(excel_list) != list:
            return excel_list
        # -------------------------------
        kmzs_list = []


        for excel in excel_list:
            xls = XlsControl(excel)
            data_list = xls.read_exel()
            # --------------------------
            if type(data_list) != list:
                return data_list
            # ---------------------------
            kml_list = KmlControl(data_list, os.path.abspath(excel))
            kml = kml_list.build_kml()
            kml_list.save_kmz(kml[0])

            path = os.path.abspath(os.path.dirname(excel))
            path_1 = os.path.abspath(os.path.splitext(excel)[0])
            kmzs = [x for x in os.listdir(path) if x[-4:] == '.kmz']
            kmzs.sort()
            kmz_file = path + os.sep + kmzs[-1]
            kmzs_list.append(kmz_file)
            kmz = CreateKMZ(kmz_file, kml[1])
            kmz.rebuild_kmz()

        # Aqui tenho de ver se já lá há kmz e por '_ver-0.n'
        original_dir_path = os.path.abspath(os.path.dirname(self.ficheiro))
        abs_original_file_path = os.path.abspath(os.path.splitext(self.ficheiro)[0])
        original_file_path = abs_original_file_path[abs_original_file_path.rfind('\\')+1:]
        kmzs = [x for x in os.listdir(original_dir_path) if x[-4:] == '.kmz' and x[:-12] ==
                original_file_path]
        kmzs.sort()
        file_name = ''
        if len(kmzs) == 0:
            file_name = self.ficheiro[:self.ficheiro.rfind('.')] + '_ver-0.1.kmz'
        else:
            version = str(round(float(kmzs[-1][-7:-4]) + .1, 2))
            file_name = self.ficheiro[:self.ficheiro.rfind('.')] + '_ver-' + version + '.kmz'

        if os.path.exists(os.path.abspath(os.path.dirname(self.ficheiro)) + '\\Temp'):
            final = KmlJoiner(kmzs_list, file_name)
            final.build_new_kmz()
        else:  #COMENTADO PORQUE ACABAVA POR CHAMAR 2 VEZES O REBUILD_KMZ()
            path = self.ficheiro[:self.ficheiro.rindex('/')]
            path_1 = self.ficheiro[self.ficheiro.rindex('/')+1:self.ficheiro.rfind('.')]
            kmzs = [x for x in os.listdir(path) if x[-4:] == '.kmz' and x[:-12] ==
                    path_1]
            kmzs.sort()
            kmz_file = path + os.sep + kmzs[-1]
            kmz = CreateKMZ(kmz_file, kml[1])
            kmz.rebuild_kmz()

        # LIMPAR A TEMP
        if os.path.exists(os.path.abspath(os.path.dirname(self.ficheiro)) + '\\Temp'):
            os.chdir(os.path.abspath(os.path.dirname(self.ficheiro)) + '\\Temp')
            clean = TempCleaner(os.getcwd())
            clean.clean()

        # Voltar à working directory original
        os.chdir(self.original_working_dir)
