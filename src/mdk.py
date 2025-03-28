import xml.etree.ElementTree as et
import os
import random
from coordinates import Coordinates


class MdkControl(object):
    """
    Turns a list into a Mdk() object and saves in the disk as a MKD file.
    """

    def __init__(self, data_list, filename):
        """
        """

        self.data_list = data_list
        self.filename = filename
        self.markupdata = et.Element('MarkupData')
        self.__defaultSubElements()
        self.routes = et.SubElement(self.markupdata, 'Routes')
        self.tree = None

    def build_mkd(self):
        """
        """
        for item in self.data_list:
            route = et.SubElement(self.routes, 'Route')
            et.SubElement(route, 'Name').text = str(item[0][0])  # sheetname
            et.SubElement(route, 'GUID').text = self.__getGUID()
            et.SubElement(route, 'Color A="255" R="0" G="0" B="0"')
            pos = et.SubElement(route, 'Positions')
            if len(item) > 1:
                headers = [x.lower() for x in item[1]]  # column titles
                # Optional column names to build a point-----------------
                col_names = ["icon", "size", "description", "rotation", "appenddatacolumnstodescription", "size"]
                # See what optional names are in the EXEL data ----------
                col_as_name = [True if x in headers else False for x in col_names]
                #  POINT BUILD ##########################################
                for i in range(2, len(item)):
                    rt_pos = et.SubElement(pos, 'RoutePosition')
                    # Coordinates ---------------------------------------
                    lat1 = str(item[i][0])
                    lon1 = str(item[i][1])
                    coords = Coordinates(lon1, lat1).convert()
                    #  Point --------------------------------------------
                    if col_as_name[col_names.index('icon')] and item[i][headers.index('icon')] != '':
                        t = str(item[i][headers.index('icon')])
                        et.SubElement(rt_pos, 'Icon').text = t

                    et.SubElement(rt_pos, 'Name').text = str(item[i][2])

                    if col_as_name[col_names.index('description')] and item[i][headers.index('description')] != '':
                        add = 'appenddatacolumnstodescription'
                        text = self.__point_description(headers, [item[i][headers.index(add)]], item[i])
                        et.SubElement(rt_pos, 'Description').text = text

                    et.SubElement(rt_pos, 'ShowLabel').text = 'true'

                    if col_as_name[col_names.index('rotation')] and item[i][headers.index('rotation')] != '':
                        t = item[i][headers.index('rotation')]
                        et.SubElement(rt_pos, 'Rotation').text = t
                    else:
                        et.SubElement(rt_pos, 'Rotation').text = '0'

                    if col_as_name[col_names.index('size')] and item[i][headers.index('size')] != '':
                        t = str(item[i][headers.index('size')])
                        et.SubElement(rt_pos, 'Size').text = t
                    else:
                        et.SubElement(rt_pos, 'Size').text = '24'

                    geopoint = et.SubElement(rt_pos, 'GeoPoint')
                    et.SubElement(geopoint, 'Lat').text = str(coords[1])
                    et.SubElement(geopoint, 'Lon').text = str(coords[0])

        return et.ElementTree(self.markupdata)

    def save_mkd(self, tree):
        """
        """

        path = self.filename[:self.filename.rindex('\\')]
        path1 = self.filename[self.filename.rindex("\\") + 1:self.filename.rfind('.')]

        mkds = [x for x in os.listdir(path) if x[-4:] == '.mkd' and x[:-12] == path1]

        if len(mkds) > 0:
            mkds.sort()
            version = str(round(float(mkds[-1][-7:-4]) + .1, 2))
            tree.write(self.filename[:self.filename.rfind('.')] + "_ver-" + version + ".mkd")
        else:
            version = "_ver-0.1.mkd"
            tree.write(self.filename[:self.filename.rfind('.')] + version)

    def __defaultSubElements(self):
        """
        """
        et.SubElement(self.markupdata, 'FileVersion').text = '1.1'

        for e in ['Rectangles', 'Circles', 'Arcs', 'Polygons', 'Lines', 'Placemarks']:
            et.SubElement(self.markupdata, e)

    def __setGUID(self):
        """
        Returns a random hexadecimal value
        """

        start = 1000000000000000000
        stop = 9999999999999999999

        return str(hex(random.randrange(start, stop)))[2:]

    def __getGUID(self):
        """
        Returns a random hexadecimal value
        """

        return self.__setGUID()

    def __point_description(self, headers, appenddata, data):
        """
        """

        new_data = data[:]  # to manipukate a data copy
        new_headers = headers[:]  # to manipulate an headers copy
        # the column "AppendDataColumnsTo Description" items
        items = [x.split(',') for x in appenddata][0]
        f_items = [x.lower().strip() for x in items]  # formated items

        # the data indexes of the elements in AppendDataToColToDescr
        indexes = [new_headers.index(x.lower().strip()) for x in f_items if x.lower().strip() in new_headers]
        # add to the indexes the Description column
        indexes.insert(0, new_headers.index('description'))

        pt = [x.capitalize() for x in f_items]  # Capitalize titles

        text = str(new_data[indexes[0]]) + '\n'
        for i in range(1, len(indexes)):
            text += pt[i - 1] + ': ' + str(new_data[indexes[i]]) + '\n'

        return text
