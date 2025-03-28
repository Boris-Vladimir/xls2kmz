"""
@python version:
    Python 3.4

@summary:
    Control composed by the TempCleaner() class.
    This class has a constructor:
        - __init__(temp_folder);
    One public function:
        - clean();
    And, four non public 'helper' fuctions:
        - __file_deleter();
        - __folder_deleter(directory);
        - __sub_folder_searcher();
        - __sub_folder_crawler();

@note:
    function __init__(temp_folder):
        Class constructor.
        temp_folder is the path to the temp folder, a string.
    function clean():
        Handler for the button "Abrir EXEL..." (Open EXEL)
        Opens a new window (filedialog.askopenfilename) to choose the
        EXCEL file that is necessary to make the KMZ file.
    function __file_deleter():
        Deletes the files inside the Temp folder
    fucntion __folder_deleter(directory):
        Deletes the directory
        directory is the path to the directory/folder, a string
    fucntion __sub_folder_searcher():
        Returns a list of folders if folders exists
    function __sub_folder_crawler():
        Crawsl for each sub folder of the parent folder

@author:
    Venancio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (28/03/2014):
        - Creation of the class TempCleaner()

    1.1 (90/12/2014):
        - Added docstrings
    
@since:
    28/03/2014
"""

import os


class TempCleaner(object):

    def __init__(self, temp_folder):
        self.folder = temp_folder

    def clean(self):
        try:
            os.chdir(self.folder)
            self.__sub_folder_crawler()
            self.__file_deleter()
            os.chdir(os.pardir)
            self.__folder_deleter(self.folder)
        except OSError:
            self.clean()

    def __file_deleter(self):
        for files in os.listdir():
            os.remove(files)

    def __folder_deleter(self, directory):
        os.rmdir(directory)

    def __sub_folder_searcher(self):
        return [f for f in os.listdir() if os.path.isdir(f)]

    def __sub_folder_crawler(self):
        sub_folders = self.__sub_folder_searcher()
        if len(sub_folders) > 0:
            for f in sub_folders:
                os.chdir(f)
                if len(os.listdir()) == 0:
                    os.chdir(os.pardir)
                    self.__folder_deleter(f)
                else:
                    return self.__sub_folder_crawler()
        else:
            return os.getcwd()
