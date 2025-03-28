"""
@python version:
    Python 3.4

@summary:
    Control composed by the KmlJoiner() class.
    This class has:
        -  Constructor:
            -  __init__(list_of_kmls, original_path);
        -  One public function:
            - build_new_kmz(); and
        -  Three private/auxiliary functions:
            - __extract_all()
            - __join_docs(doc_list)
            - __kml_parser(doc)

@note:
    function __init__(list_of_kmls, original_path)
        Class constructor
            - list_of_kmls is a list containing strings of paths
            - original_path is a string of the original path

    function build_new_kmz()
        Opens the kmz file(archive file) and add the fotos to
        the archive, then save and close the file

    function __extract_all()
        Extracts all members from the archive to the current
        working directory

    function __join_docs(doc_list)
        Join all kml docs in one only temp_kml and calls __kml_parser(temp_kml)

    function __kml_parser(doc)
        Parses the kml_temp (doc) and creates a new new_kml file

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (07/01/2014):
        - Implementation of the CreateKMZ() class.
        - Creation of the functions / class attributes:
             __init__(), rebuild_kmz(), __extract_all(), __join_docs()
             and __kml_parser()
    1.1 (09/12/2014):
        - Added docstrings

@since:
    07/01/2014
"""
import os
import zipfile  # to zip and unzip files, rebuild the kmz


class KmlJoiner(object):

    def __init__(self, list_of_kmls, original_path):
        self.kmls = list_of_kmls
        self.path = original_path

    def __extract_all(self):
        """
        ZipFile.extractall([path[, members[, pwd]]])

        Extract all members from the archive to the current working directory.
        path specifies a different directory to extract to.
        members must be a subset from list returned by namelist().
        pwd is the password used for encrypted files.
        """
        i = 0

        for k in self.kmls:
            zf = zipfile.ZipFile(k, 'r')
            zf.extractall()
            if 'doc.kml' in os.listdir():
                os.rename('doc.kml', 'doc' + str(i) + '.kml')
                i += 1
            zf.close()  # acho que não é preciso

    def build_new_kmz(self):
        # change working directory p "Temp"
        os.chdir(os.path.dirname(self.kmls[0]))
        self.__extract_all()

        zf = zipfile.ZipFile(self.path, "a")  # make new kmz in parent dir
        doc_list = [doc for doc in os.listdir() if doc[-3:] == 'kml']
        self.__join_docs(doc_list)
        zf.write('doc.kml')

        os.chdir(os.path.dirname(self.kmls[0]) + '\\files')
        for image in os.listdir():
            zf.write(image, arcname='files\\' + image)

        zf.close()

    def __join_docs(self, doc_list):
        os.system("copy *.kml doc_temp.kml")
        self.__kml_parser('doc_temp.kml')

    def __kml_parser(self, doc):
        close_folder = False
        open_folder = True

        lines = open(doc, 'r', encoding='utf-8').readlines()  # list of lines
        new_kml = open('doc.kml', 'w')

        for line in lines[:-4]:
            if '</Folder>' in line:
                close_folder = True
                open_folder = False
            if close_folder is False and open_folder is True:
                new_kml.writelines(line)
            if '<Folder' in line:
                close_folder = False
                open_folder = True
        new_kml.writelines(lines[-4:-1])

        new_kml.close()

