"""
@python version:
    Python 3.4

@summary:
    Control composed by the class Coordinates().
    That class has the propose to convert diverse coordinates formats
    (degrees, minutes, seconds - dms; and degrees, decimal minutes - h)
    to coordinates in the decimal degrees format.
    The control has a constructor:
        - __init__(latitude, longitude)
    A public method:
        - convert()
    And five private methods:
        - __remove_spaces(string)
        - __remove_chars(string)
        - __get_format(string)
        - __convert_dms()
        - __convert_h()

@note:
    method __init__(latitude, longitude):
        Class constructor.
        Take as parameters "latitude" and "longitude", two strings.
        Builds two class variables lat and lon.
        First removes any white spaces from latitude and longitude,
        calling the __remove_spaces()
        then removes the cardinal direction letter calling __remove_chars.

    method convert():
        Checks for the format of the coordinates and call __convert_dms,
        if the format is degrees, minutes, seconds; call __convert_h,
        if the format is hybrid aka degrees, decimal minutes.
        returns a list with lat and lon in decimal degrees format

    method __remove_spaces(string):
        Returns the string whithout spaces

    method __remove_chars(string):
        Returns the string whithout the cardinal direction letter.
        If the cardinal direction letter is 'S' or 'W' it appends a '-'
        (minus sigh) to the begining of the string.

    method __get_format(string):
        Check the format of the string and returns "dms", "h" or "dd",
        acording to Degrees Minutes Seconds format, Hybrid format, or
        Decimal Degrees format, respectivaly

    method __convert_dms():
        Convert lat and lon from Degrees Minutes Seconds format to
        Decimal Degrees format

    method __convert_h():
        Convert lat and lon Hybrid (Degrees, Decimal Degrees) format
        to Decimal Degrees 

@author:
    Venâncio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (22/12/2014):
        - Creation of the Class Coordinates(), and his methods:
            - __init__
            - convert
            - __remove_spaces
            - __remove_chars
            - __get_format
            - __convert_dms
            - __convert_h
@since:
    22/12/2014
"""


class Coordinates(object):
    """
    Creates a coordinate object with latitude and longitude and convert
    it to the Decimal Degrees format.

    The types of coordinates formats are:
    Decimal Degrees: 32.8303ºN ; 116.7762ºW
    Degrees Minutes Seconds:32º49'49''N ; 116º46'34''W
    Hybrid: 32º49.818'N ; 116º46.574'W
    """

    def __init__(self, latitude, longitude):
        """
        str, str -> None

        Class constructor.
        Take as parameters "latitude" and "longitude", two strings.
        Builds two class variables lat and lon.
        First removes any white spaces from latitude and longitude,
        calling the __remove_spaces()
        then removes the cardinal direction letter calling __remove_chars.
        """
        self.__lat = self.__remove_spaces(latitude)
        self.__lon = self.__remove_spaces(longitude)
        self.lat = self.__remove_chars(self.__lat)
        self.lon = self.__remove_chars(self.__lon)

    def convert(self):
        """
        None -> tuple

        Checks for the format of the coordinates and call __convert_dms,
        if the format is degrees, minutes, seconds; call __convert_h,
        if the format is hybrid aka degrees, decimal minutes.
        returns lat and lon in decimal degrees format
        """
        coords_format = self.__get_format(self.lat)

        if coords_format == "dms":
            self.__convert_dms()
        elif coords_format == "h":
            self.__convert_h()

        return self.lat, self.lon

    def __remove_spaces(self, string):
        """
        str -> str

        Returns the string whithout spaces or tabs
        """
        string = string.replace("\t", "")
        return string.replace(" ", "")

    def __remove_chars(self, string):
        """
        str -> str

        Returns the string whithout the cardinal direction letter.
        If the cardinal direction letter is 'S' or 'W' it appends a '-'
        (minus sigh) to the begining of the string.
        """
        pos_chars = ['N', 'E', 'n', 'e']
        neg_chars = ['S', 'W', 's', 'w']
        new_string = ''

        for char in string:
            if char in pos_chars:
                new_string += ""
            elif char in neg_chars:
                new_string = "-" + new_string
                "-" + string.replace(char, "")
            else:
                new_string += char

        return new_string

    def __get_format(self, string):
        """
        str -> str

        Check the format of the string and returns "dms", "h" or "dd",
        acording to Degrees Minutes Seconds format, Hybrid format, or
        Decimal Degrees format, respectivaly
        """
        m = "'"
        s = "''"
        s2 = '"'
        s3 = "º"
        s4 = "°"
        s5 = ","
        s6 = "."

        if (s in string or s2 in string) or \
                (s3 in string and not s5 in string) or \
                (s3 in string and not s6 in string) or \
                (s4 in string and not s5 in string) or \
                (s4 in string and not s6 in string):
            return 'dms'
        if m in string and s not in string:
            return 'h'
        else:
            return 'dd'

    def __convert_dms(self):
        """
        None -> None

        Convert lat and lon from Degrees Minutes Seconds format to
        Decimal Degrees format
        Decimal Degrees = Degrees + Minutes / 60 + Seconds / 3600
        """
        lat = self.lat.replace('"', "''")
        lon = self.lon.replace('"', "''")
        lat_sign = ""
        lon_sign = ""

        if "º" in lat:
            if "-0" in lat[:lat.index("º")]:
                lat_sign = "-"
            else:
                pass
            lat_deg = int(lat[:lat.index("º")])
            if "-0" in lon[:lon.index("º")]:
                lon_sign = "-"
            else:
                pass
            lon_deg = int(lon[:lon.index("º")])
            lat_min = int(lat[lat.index("º") + 1:lat.index("'")])
            lon_min = int(lon[lon.index("º") + 1:lon.index("'")])
        else:
            if "-0" in lat[:lat.index("°")]:
                lat_sign = "-"
                lat_deg = 0
            else:
                lat_deg = int(lat[:lat.index("°")])

            if "-0" in lon[:lon.index("°")]:
                lon_sign = "-"
                lon_deg = 0
            else:
                lon_deg = int(lon[:lon.index("°")])

            lat_min = int(lat[lat.index("°") + 1:lat.index("'")])
            lon_min = int(lon[lon.index("°") + 1:lon.index("'")])

        if "''" in lat:
            lat_sec = float(lat[lat.index("'") + 1:lat.index("''")])
        else:
            lat_sec = float(lat[lat.index("'") + 1:])
        if "''" in lon:
            lon_sec = float(lon[lon.index("'") + 1:lon.index("''")])
        else:
            lon_sec = float(lon[lon.index("'") + 1:])

        if lat_deg < 0:
            lat_sign = "-"
        if lon_deg < 0:
            lon_sign = "-"

        int_lat = abs(lat_deg) + (lat_min / 60.0) + (lat_sec / 3600.0)
        int_lon = abs(lon_deg) + (lon_min / 60.0) + (lon_sec / 3600.0)
        self.lat = lat_sign + str(int_lat)
        self.lon = lon_sign + str(int_lon)

        self.convert()

    def __convert_h(self):
        """
        None -> None

        Convert lat and lon from Hybrid format to Decimal Degrees format
        Decimal Degrees = Degrees + Minutes / 60
        """
        lat = self.lat
        lon = self.lon
        lat_sign = ""
        lon_sign = ""

        try:
            lat_deg = int(lat[:lat.index("º")])
            lat_min = int(lat[lat.index("º") + 1:lat.index("'")])

            lon_deg = float(lon[:lon.index("º")])
            lon_min = float(lon[lon.index("º") + 1:lon.index("'")])
        except:
            lat_deg = int(lat[:lat.index("°")])
            lat_min = int(lat[lat.index("°") + 1:lat.index("'")])

            lon_deg = float(lon[:lon.index("°")])
            lon_min = float(lon[lon.index("°") + 1:lon.index("'")])

        if lat_deg < 0:
            lat_sign = "-"
        if lon_deg < 0:
            lon_sign = "-"

        int_lat = abs(lat_deg) + (lat_min / 60.0)
        int_lon = abs(lon_deg) + (lon_min / 60.0)
        self.lat = lat_sign + str(int_lat)
        self.lon = lon_sign + str(int_lon)

        self.convert()






