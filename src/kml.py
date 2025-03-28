"""
@python version:
    Python 3.4

@summary:
    Control composed by the KmlControl() class, witch turn a list into
    a Kml() object and saves it as KMZ file.
    This class has three public functions:
        __init__(data_list, file_name),
        build_kml() e,
        save_kmz(kml).
    And six non-public functions (auxiliaries):
        __point_description(headers, append_data_columns_to_description,
            data),
        __formated_date(xldate),
        __formated_time(xldate),
        __icon_heading(coords, next_coords, icon),
        __color_translate(color), and, a auxiliary of __point_description:
        __point_description_foto(data, indexes, titles)

@note:
    function __init__(data_list, file_name):
        Class Constructor.
        Has as parameters a list (values / data to build the Kml() object)
        and a string (the name of the file we want to build).

    function buil_kml():
        Builds a Kml() object from a list passed by the class constructor
        and returns it.

    function __point_description(headers,
            append_data_columns_to_description, data):
        Builds a HTML string to be used in the description/ descriptive
        balloon of a Kml().point
        As parameters are passed three lists:
            headers - has the titles of all columns from the EXEL file.
            append_data_columns_to_description - has the EXEL column
                who has the same name, used to build the description
                balloon.
            data - the EXEL cell values.

    function __formated_date(xldate):
        Formates the EXEL cell values of datetime type.
        Puts a "-" separating days from months and months from years.
        As parameter is passed a tuple, xldate, which has the
        datetime value of the EXEL cell.

    function __formated_time(xldate):
        Formates the EXEL cell values of datetime type.
        Puts a  ":" between hours and minutes, and minutes and seconds.
        As parameter is passed a tuple, xldate, witch has the
        datetime value of the EXEL cell.

    function icon_heading(coords, next_coords, icon):
        Calculates and returns in grades the icon direction of the
        Kml().points.
        As parameters two lists are passed, one (coords) with the
        actual coordinates, and other (next_coords) with the position
        of the next coordinates; and, a int (icon) representing the icon
        number to be used in the Kml().

    function __point_description_foto(data, indexes, titles):
        Is an auxiliary function of the __point_description, to help
        in the building of descriptions / descriptive balloons witch
        have photos.
        As parameters are passed three lists:
            data - the EXEL cell values.
            indexes - the index in data of the titles in titles.
            titles - the names in the column named
                "AppendDataColumnsToDescription".

    function __color_translate(self, color):
        Translates the color names of the icons into the
        correspondent hex value.
        Has as parameter the color name.

    function __polygon(self, lattude, longitude, azimute, radius, altitude):
        Construts a polygon (quasi triangle) given the initial point
        (latitude and longitude), the direction (azimute), the radius,
        and the altitude.

    function __toEart(self, p, altitude)
        Calculates the lat and long of the points to design a perfect circle
        As parameters are passed a point, p, a list with latitude, longitide,
        and altitude, an int.
        Returns a list with the point latitude, longitude and altitude.

    function __toCart(self, longitude, latitude)
        Convert long, lat IN RADIANS to (x,y,z)
        As parameters are passed the longitude and latitude
        Returns a list containing the long and lat in radians

    function __spoints(self, long, lat, meters, altitude, n, offset)
        Get raw list of points in long,lat format
        As parameters are passed the longitude, latitude, meters (radius),
        altitude, n (number of sides), offset (rotation in degrees)
        Returns a list of points comprising the object

    funtion __rotPoints(vet, pt, phi)
        Rotates point pt, around unit vector vec by phi radians
        http://blog.modp.com/2007/09/rotating-point-around-vector.html
        Return the rotated point


@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    DP

@version:
    1.0 (14/11/2013):
        - Creation of the functions:
            - minimalist_xldate_as_datetime
            - formated_datetime
            - color_translate
            - point_description
            - icon_heading
            - construir_kml
    1.1 (24/11/2013):
        - Implementation of the KmlControl() class;
        - Creation of the function save_kmz
        - Modification of the __point_description function so it joins
            the latitude and longitude in a single value, coordenadas;
        - Modification of the function __point_description(), so it allow
            the use of photos in the descriptive balloons, calling a new
            function: __point_description_foto();
        - Added  a new __point_description_foto() function, auxiliary
            of the __point_description() function;
        - Removed the function to read floats of EXEL datetime values
            and turns them in dates and times. This is now a job of the
            XlsControl() class from the xls.py library;
        - Substitution  of the __formated_datetime() function by two new
            functions: __formated_date() and __formated_time(), so they
            can manipulate the new format (a tuple) returned by the
            XlsControl() class from the xls.py library;
        - Alteration of the __color_translate() function so it translate the
            color names by the new ColornameToHex() class from the
            colorname_to_hex.py() library, in substitution of the
            html4_names_to_hex() function of the webcolors() class from
            the webcolors.py library.
    1.2 (05/12/2013):
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
        - Add a trailing number representing the version to the file
            saved in the __save_kmz() function.
        - Alteration of the __point_description() and
            __point_description_foto() functions to add data to the balloon
            only and not to the description.
        - Add more functionality to build_kml() function, now has camera
            positioning (point.lookat), a legend, a call to __polygon()
            and to color that, a call to ColornameToKml(color).get_color()
        - Change the python version to 3.3 a so the xrange to range
        - Modification of the __save_kmz, so it finds the last '.' instead
            of finding the -4 index.
        - creation of the __polygon() function to design polygons in the map.
    1.3 (06/01/2014):
        - Alteration of the __build_kml() function so it manages Excel files
            that have leaves with all cells empty
        - Alteration of the __formated_date(xldate) and __formate_time(xldate)
            so they return the str(xldate) if it is not a tuple
        - Correction of a bug in __point_description, in the "head" variable,
           now, it assegures "head" is a str.
        - Modification of the __build_kml() function, so that doesnt design icons
            descriptions and point names if it designs a polygon.
        - Add functionalities to the __polygon function, designs circles if the
            amplitude was 360, the descriptive balloon apears now in the center
        - Add four more functions to help designing the circle, when the polygon
            has 360 degrees:
                - __toEart(p, altitude)
                - __toCart(longitude, latitude)
                - __spoints(long, lat, meters, altitude, n, offset)
                - __rotPoints(vec, pt, phi) 
            (adapted from kmcircle.py - https://code.google.com/p/kmlcircle/)
    1.4 (01/12/2014):
        - Add functionality to create squares with 2 coordinates
        - Add a log in logfile if the name of the foto in excel does not correspond
            to any foto in the Foto's folder
    1.5 (23/12/2014):
        - Add a call to Coordinates(lat, lon).convert() in the build_kml() method
            so it could use diverse coordinate systems
    1.6 (??/??/????):
        - Corrected bug when the excel file has only one point
    1.7 (24/08/2015):
        - Draw a line per two points instead of a general line to connect every point
        - Change the way the description is added to the point balloon
        - The name of the point cound now be time (Excel time format)
    1.7.1 (09/09/2015):
        - Draw only the point if the field icon is used. Remove it from polygons
    1.7.2 (22/11/2021):
        - Translate coordinates to english (Coordenadas - Coordinates)
    1.8 (28/11/2021):
        - Correct bug in drawing polygons

    TODO:
        - Refactor build_kml to be just a function that calls auxiliary functions
            (one for the line, another for the point, the icon, etc...)

@since:
    14/11/2013
"""

import math  # calculate the inclination of the arows
import os  # work with files and directories
import simplekml  # build KML/KMZ files
from colorname_to_hex import ColornameToHex  # color names to hex-abgr
from colorname_to_kml import ColornameToKml
from coordinates import Coordinates  # convert coordinate types to Decimal Degrees


class KmlControl(object):
    """
    Turns a list into a Kml() object and saves in the disk as a KMZ file.
    """

    def __init__(self, data_list, file_name):
        """
        list, string -> KmlControl() object

        data_list is the EXEL data list.
        file_name is the file name of the KMZ we want to build.
        """
        self.data_list = data_list
        self.file_name = file_name
        self.images_list = []  # list of all icons, photos and images used

    def build_kml(self):
        """
        None -> Kml() object

        Builds a Kml() object from a list passed by the class constructor
        and returns it.

        See: http://simplekml.readthedocs.org/en/latest/index.html
        """
        kml = simplekml.Kml()
        icons = os.getcwd() + os.sep + "icons" + os.sep  # path
        add = "appenddatacolumnstodescription"
        # Build a legend if exist ----------------------------------
        try:
            path = os.getcwd() + os.sep + "fotos" + os.sep + "legenda.png"
            screen = kml.newscreenoverlay(name="Legenda")
            screen.icon.href = path
            screen.overlayxy = simplekml.OverlayXY(x=0, y=1,
                                                   xunits=simplekml.Units.fraction,
                                                   yunits=simplekml.Units.fraction)
            screen.screenxy = simplekml.ScreenXY(x=15, y=15,
                                                 xunits=simplekml.Units.pixels,
                                                 yunits=simplekml.Units.insetpixels)
            screen.size.x = -1
            screen.size.y = -1
            screen.size.xunits = simplekml.Units.fraction
            screen.size.yunits = simplekml.Units.fraction
        except:
            pass

        for item in self.data_list:
            if len(item) > 1:
                folder = kml.newfolder(name=str(item[0][0]))  # sheetname
                headers = [x.lower() for x in item[1]]  # column titles
                # Optional column names to build a point-----------------
                col_names = ["icon", "iconcolor", "iconscale", "description",
                             "appenddatacolumnstodescription", "iconheading",
                             "linestringcolor", "foto", "polygon",
                             "polygoncolor", "polygonaltitude",
                             "polygonazimute", "polygonamplitude",
                             "squarealtitude", "squarelatitude",
                             "squarelongitude", "squarecolor"]
                # See what optional names are in the EXEL data ----------
                next_coords = None
                line = None
                line_color = None
                col_as_name = [True if x in headers else False for x in col_names]
                #  POINT BUILD ##########################################
                for i in range(2, len(item)):
                    # Coordinates ---------------------------------------
                    lat1 = str(item[i][0])
                    lon1 = str(item[i][1])
                    coords = [Coordinates(lon1, lat1).convert()]
                    if i < len(item) - 1:  # Next coordinate (icon heading)
                        lat2 = str(item[i + 1][0])
                        lon2 = str(item[i + 1][1])
                        next_coords = [Coordinates(lon2, lat2).convert()]
                    #  Point --------------------------------------------
                    try:
                        name = self.__formated_time(item[i][2])
                    except:
                        try:
                            name = str(item[i][2])
                        except ValueError:
                            name = item[i][2]
                    #  Icon ---------------------------------------------
                    if col_as_name[col_names.index("icon")] and item[i][headers.index("icon")] != '':
                        point = folder.newpoint(name=name, coords=coords)
                        point.lookat.longitude = coords[0][0]
                        point.lookat.latitude = coords[0][1]
                        point.lookat.altitude = 0
                        point.lookat.heading = 0
                        point.lookat.tilt = 0
                        point.lookat.range = 1230
                        url = icons + str(int(item[i][headers.index("icon")])) + ".png"
                        point.style.iconstyle.icon.href = url
                        point.style.balloonstyle.text = self.__point_description(headers, [item[i][headers.index(add)]],
                                                                                 item[i])
                        # Square Color ---------------------------------------
                        if col_as_name[col_names.index("squarecolor")] and item[i][headers.index("squarecolor")] != '':
                            point.style.iconstyle.scale = 0
                        # Icon color ----------------------------------------
                        if col_as_name[col_names.index("iconcolor")] and item[i][headers.index("iconcolor")] != '':
                            point.style.iconstyle.color = self.__color_translate(item[i][headers.index("iconcolor")])
                        # Icon scale / size ---------------------------------
                        if col_as_name[col_names.index("iconscale")] and item[i][headers.index("iconscale")] != '':
                            point.style.iconstyle.scale = item[i][headers.index("iconscale")]
                        # Ballon Description ----------------------------------
                        if col_as_name[col_names.index("description")]:
                            try:  # hours
                                point.description = self.__formated_time(item[i][headers.index("description")])
                            except ValueError:
                                point.description = str(item[i][headers.index("description")])
                        # Icon heading / inclination ------------------------
                        if col_as_name[col_names.index("iconheading")] and item[i][headers.index("iconheading")] != '':
                            heading = item[i][headers.index("iconheading")]
                            if type(heading) == float or type(heading) == int:
                                point.style.iconstyle.heading = heading
                            else:
                                heading = item[i][headers.index("icon")]
                                if not next_coords:
                                    next_coords = coords
                                point.style.iconstyle.heading = self.__icon_heading(coords, next_coords, heading)
                        # Color and data to build the line ------------------
                        if col_as_name[col_names.index("linestringcolor")] and \
                                item[i][headers.index('linestringcolor')] != '':
                            line = [(coords[0][0], coords[0][1])]
                            line_color = self.__color_translate(item[i][headers.index("linestringcolor")])

                            if i < len(item) - 1:  # Next coordinate (line)
                                if not next_coords:
                                    next_coords = coords
                                line.append((next_coords[0][0], next_coords[0][1]))
                                lin = folder.newlinestring(coords=line)
                                lin.style.linestyle.color = line_color
                                lin.style.linestyle.width = 2  # 2 pixels
                    # Polygon -------------------------------------------
                    if col_as_name[col_names.index("polygon")] and item[i][headers.index("polygon")] != '':
                        try:  # hours
                            description = self.__formated_time(item[i][headers.index("description")])
                        except TypeError:
                            description = str(item[i][headers.index("description")])
                        try:
                            name = self.__formated_time(item[i][2])
                        except TypeError:
                            try:
                                name = str(item[i][2])
                            except TypeError:
                                name = item[i][2]
                        pol = folder.newpolygon(name=name)
                        pol.altitudemode = simplekml.AltitudeMode.relativetoground
                        radius = float(item[i][headers.index("polygon")])
                        color = item[i][headers.index("polygoncolor")]
                        pol_color = ColornameToKml(color).get_color()
                        altitude = float(item[i][headers.index("polygonaltitude")])
                        azimute = float(item[i][headers.index("polygonazimute")])
                        if col_as_name[col_names.index("polygonamplitude")]:
                            amplitude = float(item[i][headers.index("polygonamplitude")])
                        else:
                            amplitude = 60.0
                        pol_points = self.__polygon(float(coords[0][0]),
                                                    float(coords[0][1]), azimute,
                                                    radius, altitude, amplitude)
                        pol.outerboundaryis = pol_points
                        pol.style.balloonstyle.text = self.__point_description(headers,
                                                                               [item[i][headers.index(add)]], item[i])
                        pol.style.linestyle.color = pol_color
                        pol.style.polystyle.color = simplekml.Color.changealphaint(100, pol_color)
                    # Square -------------------------------------------
                    if col_as_name[col_names.index("squarecolor")] and \
                            item[i][headers.index("squarecolor")] != '':
                        description = str(item[i][headers.index("description")])
                        sqr = folder.newpolygon(name=description)
                        sqr.altitudemode = simplekml.AltitudeMode.relativetoground
                        color = item[i][headers.index("squarecolor")]
                        sqr_color = ColornameToKml(color).get_color()
                        altitude = float(item[i][headers.index("squarealtitude")])
                        sqr_lat_original = item[i][headers.index("squarelatitude")]
                        sqr_lon_original = item[i][headers.index("squarelongitude")]
                        sqr_coords = Coordinates(str(sqr_lat_original), str(sqr_lon_original)).convert()
                        sqr_lat = sqr_coords[0]
                        sqr_lon = sqr_coords[1]
                        sqr_points = [[coords[0][0], coords[0][1], altitude],
                                      [sqr_lon, coords[0][1], altitude],
                                      [sqr_lon, sqr_lat, altitude],
                                      [coords[0][0], sqr_lat, altitude],
                                      [coords[0][0], coords[0][1], altitude]]
                        sqr.outerboundaryis = sqr_points
                        sqr.style.balloonstyle.text = self.__point_description(headers,
                                                                               [item[i][headers.index(add)]],
                                                                               item[i])
                        sqr.style.linestyle.color = sqr_color
                        sqr.style.polystyle.color = simplekml.Color.changealphaint(100, sqr_color)

        return kml, self.images_list

    def save_kmz(self, kml):
        """
        kml object -> None

        Turns Kml() object in a KMZ file an saves it in the disk.
        """
        path = self.file_name[:self.file_name.rindex(os.sep)]
        path_1 = self.file_name[self.file_name.rindex(os.sep) + 1:self.file_name.rfind('.')]
        kmzs = [x for x in os.listdir(path) if x[-4:] == '.kmz' and x[:-12] == path_1]

        if len(kmzs) > 0:
            kmzs.sort()
            version = str(round(float(kmzs[-1][-7:-4]) + .1, 2))
            kml.savekmz(self.file_name[:self.file_name.rfind('.')] + "_ver-" + version + ".kmz")
        else:
            version = "_ver-0.1.kmz"
            kml.savekmz(self.file_name[:self.file_name.rfind('.')] + version)

    def __point_description(self, headers, append_data_columns_to_description, data):
        """
        list, list, list -> str

        Builds a HTML string to be used in the description/ descriptive
        balloon of a Kml.point()
        """
        new_data = data[:]  # to manipulate a data copy
        new_headers = headers[:]  # to manipulate a headers copy
        # the column "AppendDataColumnsToDescription" items
        items = [x.split(',') for x in append_data_columns_to_description][0]
        f_items = [x.lower().strip() for x in items]  # formated items

        # Remove latitude and longitude and built a new Coordenadas
        if "latitude" and "longitude" in f_items:
            coordenadas = str(new_data[new_headers.index("latitude")]) + ", " + \
                          str(new_data[new_headers.index("longitude")])
            data_i = min(new_headers.index("latitude"), new_headers.index("longitude"))
            new_data.pop(new_headers.index("latitude"))
            new_headers.pop(new_headers.index("latitude"))
            new_data.pop(new_headers.index("longitude"))
            new_headers.pop(new_headers.index("longitude"))
            new_data.insert(data_i, coordenadas)
            new_headers.insert(data_i, "coordinates")
            f_items_i = min(f_items.index("latitude"), f_items.index("longitude"))
            f_items.pop(f_items.index("latitude"))
            f_items.pop(f_items.index("longitude"))
            f_items.insert(f_items_i, "coordinates")

        # The data indexes of the elements in AppendDataToColToDescr
        indexes = [new_headers.index(x.lower().strip()) for x in f_items if x.lower().strip() in new_headers]
        # Add to the indexes the Description column --------------------
        indexes.insert(0, new_headers.index("description"))

        # Format Dates and Times ---------------------------------------
        if "name" in f_items and type(new_data[new_headers.index("name")]) is tuple:  # Excel time format
            i = new_headers.index("name")
            new_data.insert(i, self.__formated_time(new_data.pop(i)))
        if "data" in f_items:
            i = new_headers.index("data")
            new_data.insert(i, self.__formated_date(new_data.pop(i)))
        if "hora" in f_items:
            i = new_headers.index("hora")
            new_data.insert(i, self.__formated_time(new_data.pop(i)))
        if "duracao" in f_items:
            i = new_headers.index("duracao")
            new_data.insert(i, self.__formated_time(new_data.pop(i)))
        if "duração" in f_items:
            i = new_headers.index("duração")
            new_data.insert(i, self.__formated_time(new_data.pop(i)))
        if "cellfix time" in f_items:
            i = new_headers.index("cellfix time")
            new_data.insert(i, self.__formated_time(new_data.pop(i)))

        # Capitalize titles --------------------------------------------
        pt = [x.capitalize() for x in f_items]
        # Photos -------------------------------------------------------
        if "foto" in f_items:
            return self.__point_description_foto(new_data, indexes, pt)

        # HTML Format --------------------------------------------------
        tags = {0: '<BalloonStyle><text>',
                1: '<table><tr><td colspan="2" align="center">',
                2: '</td></tr>',
                3: '<tr style="background-color:lightgreen"><td align="left">',
                4: '<tr><td>',
                5: '</td></tr>',
                6: '<td>',
                7: '</td>',
                8: '</table>',
                9: '</text></BalloonStyle>'}

        # Return build -------------------------------------------------
        try:
            title = self.__formated_time(new_data[indexes[0]])
        except:
            title = str(new_data[indexes[0]])

        head = tags[0] + tags[1] + str(title) + tags[2]
        body = []
        tail = tags[8] + tags[9]

        for i in range(1, len(indexes)):
            if i % 2 != 0:
                body.append(tags[3] + pt[i - 1] + ": " + tags[7] +
                            tags[6] + str(new_data[indexes[i]]) + tags[5])
            else:
                body.append(tags[4] + pt[i - 1] + ": " + tags[7] +
                            tags[6] + str(new_data[indexes[i]]) + tags[5])

        # Return -------------------------------------------------------
        return head + ''.join(body) + tail

    def __point_description_foto(self, data, indexes, titles):
        '''
        Build the descriptions / descriptive balloons witch have photos.
        '''
        path = os.getcwd() + os.sep + "fotos" + os.sep  # photos path
        original_path = os.getcwd()
        msg = "O nome da foto no Excel difere do nome da foto na pasta 'fotos'"

        # HTML Format --------------------------------------------------
        tags = {0: '<![CDATA[<BalloonStyle><text>',
                1: '<table width="400" border="0" cellspacing="5" \
                    cellpadding="3"><tr><td colspan="2" align="center">\
                    <font color="0000ff"><b>',
                2: '</b></font></td></tr>',
                3: '<tr style="background-color:lightgreen">\
                    <td align="center">',
                4: '<tr><td colspan="2" align="center"></h3>',
                5: '</h3></td></tr>',
                6: '<tr><td colspan="2" align="center">',
                7: '</td></tr>',
                8: '</table>',
                9: '\n<img src="',
                10: '" alt="foto" width="400" height="280">\n</br>',
                11: '<tr><td><hr></td></tr>\
                    <tr style="backgound-color:lightgreen" ><td></td></tr>',
                12: '</text></BalloonStyle>]]>'}

        # Build --------------------------------------------------------
        head = tags[0] + tags[1] + str(data[indexes[0]]) + tags[2] + tags[6] + tags[7]
        body = []
        tail = tags[8] + tags[12]

        for i in range(len(titles)):
            if "foto" in titles[i].lower():
                os.chdir(path)
                if not os.path.isfile(data[indexes[i + 1]]):
                    os.chdir(original_path)
                    logfile = open('error.log', 'a')
                    logfile.write(msg)
                    logfile.close()
                    os.startfile('error.log')
                    try:
                        os.system('taskkill /F /T /IM xls2kmz.exe')
                    except:
                        os.system('taskkill /F /T /IM pythonw.exe')
                # --------------------------------------------------------------
                else:
                    os.chdir(original_path)
                    body.append(tags[6] + tags[9] + 'files/' + (data[indexes[i + 1]]) + tags[10] + tags[7])
                    if path + (data[indexes[i + 1]]) not in self.images_list:
                        self.images_list.append(path + (data[indexes[i + 1]]))
            elif "descrição" and "descricao" in titles[i].lower():
                body.append(tags[6] + (data[indexes[i + 1]]) + tags[7] + tags[11])
            else:
                body.append(tags[3] + (data[indexes[i + 1]]) + tags[7])

        # Return -------------------------------------------------------
        return head + ''.join(body) + tail

    def __formated_date(self, xldate):
        """
        tuple -> str

        xldade is a tuple with the EXEL datetime value.

        Formates the EXEL cell values of datetime type.
        Puts a "-" separating days from months and months from years.
        Has as parameter is passed a tuple, xldate, which has the
        datetime value of the EXEL cell.
        """
        pattern = '{0:02d}'  # two numeric places

        if type(xldate) is not tuple:
            return str(xldate)

        return pattern.format(xldate[2]) + '-' + pattern.format(xldate[1]) + '-' + pattern.format(xldate[0])

    def __formated_time(self, xldate):
        """
        tuple -> str

        xldade is a tuple with the EXEL datetime value.

        Formates the EXEL cell values of datetime type.
        Puts a  ":" between hours and minutes, and minutes and seconds.
        """
        pattern = '{0:02d}'  # two numeric places

        if xldate == (0, 0, 0, 0, 0, 0) or xldate == 0:  # 0 duration
            return "00:00:00"
        if type(xldate) is not tuple:
            return str(xldate)

        return pattern.format(xldate[3]) + ':' + pattern.format(xldate[4]) + ':' + pattern.format(xldate[5])

    def __icon_heading(self, coords, next_coords, icon):
        """
        list, list, int -> float

        Calculates and returns in grades the icon direction of the
        Kml().points.
        Coords are the actual coordinates.
        next_coords are the next coordinates.
        icon is the icon number.

        In a right angled triangle, the hypotenuse is is the side
        opposite to the 90 degrees angle, the opposite is the side
        opposite to the angle we want to find, and, the adjacent, is
        the side who joins the angle we want to find to the 90
        degree angle.
        To find that angle we need to use the Inverse Tangent or
        ArcTangent

        The Tangent of the angle ø is:
            tan(ø) = Opposite / Adjacent
        So, the inverse Tangent is:
            tan^-1(Opposite / Adjacent) = ø

        See: http://www.mathsisfun.com/algebra/trig-inverse-sin-cos-tan.html
        """
        adjac = float(next_coords[0][1]) - float(coords[0][1])
        oppos = float(next_coords[0][0]) - float(coords[0][0])
        angle = 0.0  # If the first adjac is 0

        if adjac != 0.0:  # avoid ZeroDivisionError
            angle = math.atan(oppos / adjac)
        else:
            if oppos < 0:
                return 90.0
            else:
                return - 90.0
        routable_icons = [38, 106, 338, 350, 1000]
        if icon in routable_icons and adjac < 0:  # difference of negative longitude
            return math.degrees(angle)

        if icon in routable_icons and adjac >= 0:  # difference of positive longitude
            return math.degrees(angle) - 180

        return 0.0  # other icons

    def __color_translate(self, color):
        """
        str -> str

        In KML, the values for the color and opacity (alpha) are
        expressed in hexadecimal notation. The range of values for any
        color are 0 to 255 (00 to FF). For the alpha 00 is totally
        transparent and FF is totally opaque.
        The order of the expression are AABBGGRR, where AA=alpha,
        BB=blue, GG=green, and RR=red.
        """
        return ColornameToHex(color).get_abgr()

    def __polygon(self, latitude, longitude, azimute, radius, altitude,
                  amplitude):
        """
        float, float, int, float, float, float -> str

        Construts a polygon (quasi triangle) given the initial point
        (latitude and longitude), the direction (azimute), the radius,
        the altitude and the amplitude (open degree).
        """
        circle_points = self.__spoints(latitude, longitude, radius, altitude, 72, 0)

        if int(amplitude) == 360:
            return circle_points
        else:
            azi_point = int(round(azimute / 10 * 2))
            n_points = int(round(amplitude / 10 * 2))
            origin = (latitude, longitude, altitude)
            triangle_points = [origin]
            reverse_choose_points = []

            for i in range(int(round(n_points / 2))):
                reverse_choose_points.append(circle_points[(azi_point - i) % 73])

            for pt in reversed(reverse_choose_points):
                triangle_points.append(pt)

            for i in range(int(round(n_points / 2))):
                triangle_points.append(circle_points[(azi_point + i) % 73])

            triangle_points.append(origin)

            return triangle_points

    def __toEarth(self, p, altitude):
        if p[0] == 0.0:
            longitude = math.pi / 2.0
        else:
            longitude = math.atan(p[1] / p[0])
        colatitude = math.acos(p[2])
        latitude = (math.pi / 2.0 - colatitude)

        # select correct branch of arctan
        if p[0] < 0.0:
            if p[1] <= 0.0:
                longitude = -(math.pi - longitude)
            else:
                longitude = math.pi + longitude

        DEG = 180.0 / math.pi

        return [longitude * DEG, latitude * DEG, altitude]

    def __toCart(self, longitude, latitude):
        """
        convert long, lat IN RADIANS to (x,y,z)

        spherical coordinate use "co-latitude", not "latitude"
        latiude = [-90, 90] with 0 at equator
        co-latitude = [0, 180] with 0 at north pole
        """
        theta = longitude
        phi = math.pi / 2.0 - latitude

        return [math.cos(theta) * math.sin(phi), math.sin(theta) * math.sin(phi), math.cos(phi)]

    def __spoints(self, lon, lat, meters, altitude, n, offset=0):
        """
        __spoints -- get raw list of points in long,lat format

        meters: radius of polygon
        n: number of sides
        offset: rotate polygon by number of degrees

        Returns a list of points comprising the object
        """
        RAD = math.pi / 180.0  # constant to convert to radians
        MR = 6378.1 * 1000.0  # Mean Radius of Earth, meters
        offsetRadians = offset * RAD
        # compute longitude degrees (in radians) at given latitude
        r = (meters / (MR * math.cos(lat * RAD)))

        vec = self.__toCart(lon * RAD, lat * RAD)
        pt = self.__toCart(lon * RAD + r, lat * RAD)
        pts = []

        for i in range(0, n):
            pts.append(self.__toEarth(self.__rotPoint(vec, pt, offsetRadians + (2.0 * math.pi / n) * i), altitude))

        pts.append(pts[0])  # connect to starting point exactly

        return pts

    def __rotPoint(self, vec, pt, phi):
        '''
        rotate point pt, around unit vector vec by phi radians
        http://blog.modp.com/2007/09/rotating-point-around-vector.html
        '''
        # remap vector for sanity
        (u, v, w, x, y, z) = (vec[0], vec[1], vec[2], pt[0], pt[1], pt[2])

        a = u * x + v * y + w * z
        d = math.cos(phi)
        e = math.sin(phi)

        return [(a * u + (x - a * u) * d + (v * z - w * y) * e),
                (a * v + (y - a * v) * d + (w * x - u * z) * e),
                (a * w + (z - a * w) * d + (u * y - v * x) * e)]

