"""
@python version:
    Python 3.4

@summary:
    Control composed by the CreateKMZ() class.
    This class has a constructor, __init__() and one public
    function:
        - rebuild_kmz()    

@note:
    function __init__ (self, kmz_file, images_list)
        Class constructor, kmz_file is a string containing
        the path of the kmz file, and, images_list is a
        list of paths to the images to include in the kmz_file
        file

    function rebuild_kmz()
        Opens the kmz file(archive file) and add the fotos to
        the archive, then save and close the file

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (07/02/2014):
        - Implementation of the CreateKMZ() class.
        - Creation of the functions / class attributes:
             __init__() and rebuild_kmz()

@since:
    07/01/2014
"""

import zipfile

class CreateKMZ(object):
    '''
    Control to create a KMZ file from a KML one
    '''
    def __init__(self, kmz_file, images_list):
        '''
        str, list -> object

        kmz_file is a string containing the path of the
        kmz file
        images_list is a list of paths to the images to
        include in the kmz file
        '''
        self.kmz_file = ''
        for char in kmz_file:
            if char == '/':
                self.kmz_file = self.kmz_file + os.sep
            else:
                self.kmz_file = self.kmz_file + char
        self.images = images_list

    def rebuild_kmz(self):
        '''
        none -> none

        Opens the kmz file(archive file) and add the fotos
        to the archive, then save and close the file
        '''
        zf = zipfile.ZipFile(self.kmz_file, "a")
        for image in self.images:
            zf.write(image, arcname='files/' + image[image.rfind('\\')+1:]) ##Relative Path to the Image
        zf.close()
