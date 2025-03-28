"""
@python version:
    Python 3.4

@summary:
    ColornameToHex (class)  Control witch the propose to translate
    color names into hexadecimal values.
    It has five functions: a constructor, and four getters of colors
    in hexadecimal format (rgb, rgba, bgr e abgr).

@note:
    Web page with color names and his hexadecimal values:
    http://www.w3schools.com/html/html_colornames.asp

@author:
    Venâncio 2000644

@contact:
    venancio.gne@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (23/12/2013):
        - Creation of the ColornameToKml() class, with the attributes:
            - __init__(colorname)
            - get_color()
    1.1 (14/04/2014):
        - added the strip method to self.colorname to prevent space errors

@since:
    23/12/2013
"""

import simplekml


class ColornameToKml(object):
    """Translates color names into hexadecimal values."""

    def __init__(self, colorname):
        """
        str -> () simplekml.Kml.color() object

        Builds a dictionary where the key is a color name and the value
        is a simplekml.Kml.color() object of that color.
        """
        self.colorname = colorname.lower().strip()
        color = simplekml.Color
        self.colorname_to_kml = {'aliceblue': color.aliceblue,
                                 'antiquewhite': color.antiquewhite,
                                 'aqua': color.aqua,
                                 'aquamarine': color.aquamarine,
                                 'azure': color.azure,
                                 'beige': color.beige,
                                 'bisque': color.bisque,
                                 'black': color.black,
                                 'blanchedalmond ': color.blanchedalmond,
                                 'blue': color.blue,
                                 'blueviolet': color.blueviolet,
                                 'brown': color.brown,
                                 'burlywood': color.burlywood,
                                 'cadetblue': color.cadetblue,
                                 'chartreuse': color.chartreuse,
                                 'chocolate': color.chocolate,
                                 'coral': color.coral,
                                 'cornflowerblue': color.cornflowerblue,
                                 'cornsilk': color.cornsilk,
                                 'crimson': color.crimson,
                                 'cyan': color.cyan,
                                 'darkblue': color.darkblue,
                                 'darkcyan': color.darkcyan,
                                 'darkgoldenrod': color.darkgoldenrod,
                                 'darkgrey': color.darkgrey,
                                 'darkgreen': color.darkgreen,
                                 'darkkhaki': color.darkkhaki,
                                 'darkmagenta': color.darkmagenta,
                                 'darkolivegreen': color.darkolivegreen,
                                 'darkorange': color.darkorange,
                                 'darkorchid': color.darkorchid,
                                 'darkred': color.darkred,
                                 'darksalmon': color.darksalmon,
                                 'darkseagreen': color.darkseagreen,
                                 'darkslateblue': color.darkslateblue,
                                 'darkslategray': color.darkslategray,
                                 'darkturquoise': color.darkturquoise,
                                 'darkviolet': color.darkviolet,
                                 'deeppink': color.deeppink,
                                 'deepskyblue': color.deepskyblue,
                                 'dimgray': color.dimgray,
                                 'dodgerblue': color.dodgerblue,
                                 'firebrick': color.firebrick,
                                 'floralwhite': color.floralwhite,
                                 'forestgreen': color.forestgreen,
                                 'fuchsia': color.fuchsia,
                                 'gainsboro': color.gainsboro,
                                 'ghostwhite': color.ghostwhite,
                                 'gold': color.gold,
                                 'goldenrod': color.goldenrod,
                                 'gray': color.gray,
                                 'green': color.green,
                                 'greenyellow': color.greenyellow,
                                 'honeydew': color.honeydew,
                                 'hotpink': color.hotpink,
                                 'indianred': color.indianred,
                                 'indigo': color.indigo,
                                 'ivory': color.ivory,
                                 'khaki': color.khaki,
                                 'lavender': color.lavender,
                                 'lavenderblush': color.lavenderblush,
                                 'lawngreen': color.lawngreen,
                                 'lemonchiffon': color.lemonchiffon,
                                 'lightblue': color.lightblue,
                                 'lightcoral': color.lightcoral,
                                 'lightcyan': color.lightcyan,
                                 'lightgoldenrodyellow':
                                     color.lightgoldenrodyellow,
                                 'lightgray': color.lightgray,
                                 'lightgreen': color.lightgreen,
                                 'lightpink': color.lightpink,
                                 'lightsalmon': color.lightsalmon,
                                 'lightseagreen': color.lightseagreen,
                                 'lightskyblue': color.lightskyblue,
                                 'lightslategray': color.lightslategray,
                                 'lightsteelblue': color.lightsteelblue,
                                 'lightyellow': color.lightyellow,
                                 'lime': color.lime,
                                 'limegreen': color.limegreen,
                                 'linen': color.linen,
                                 'magenta': color.magenta,
                                 'maroon': color.maroon,
                                 'mediumaquamarine': color.mediumaquamarine,
                                 'mediumblue': color.mediumblue,
                                 'mediumorchid': color.mediumorchid,
                                 'mediumpurple': color.mediumpurple,
                                 'mediumseagreen': color.mediumseagreen,
                                 'mediumslateblue': color.mediumslateblue,
                                 'mediumspringgreen': color.mediumspringgreen,
                                 'mediumturquoise': color.mediumturquoise,
                                 'mediumvioletred': color.mediumvioletred,
                                 'midnightblue': color.midnightblue,
                                 'mintcream': color.mintcream,
                                 'mistyrose': color.mistyrose,
                                 'moccasin': color.moccasin,
                                 'navajowhite': color.navajowhite,
                                 'navy': color.navy,
                                 'oldlace': color.oldlace,
                                 'olive': color.olive,
                                 'olivedrab': color.olivedrab,
                                 'orange': color.orange,
                                 'orangered': color.orangered,
                                 'orchid': color.orchid,
                                 'palegoldenrod': color.palegoldenrod,
                                 'palegreen': color.palegreen,
                                 'paleturquoise': color.paleturquoise,
                                 'palevioletred': color.palevioletred,
                                 'papayawhip': color.papayawhip,
                                 'peachpuff': color.peachpuff,
                                 'peru': color.peru,
                                 'pink': color.pink,
                                 'plum': color.plum,
                                 'powderblue': color.powderblue,
                                 'purple': color.purple,
                                 'red': color.red,
                                 'rosybrown': color.rosybrown,
                                 'royalblue': color.royalblue,
                                 'saddlebrown': color.saddlebrown,
                                 'salmon': color.salmon,
                                 'sandybrown': color.sandybrown,
                                 'seagreen': color.seagreen,
                                 'seashell': color.seashell,
                                 'sienna': color.sienna,
                                 'silver': color.silver,
                                 'skyblue': color.skyblue,
                                 'slateblue': color.slateblue,
                                 'slategray': color.slategray,
                                 'snow': color.snow,
                                 'springgreen': color.springgreen,
                                 'steelblue': color.steelblue,
                                 'tan': color.tan,
                                 'teal': color.teal,
                                 'thistle': color.thistle,
                                 'tomato': color.tomato,
                                 'turquoise': color.turquoise,
                                 'violet': color.violet,
                                 'wheat': color.wheat,
                                 'white': color.white,
                                 'whitesmoke': color.whitesmoke,
                                 'yellow': color.yellow,
                                 'yellowgreen': color.yellowgreen}

    def get_color(self):
        return self.colorname_to_kml[self.colorname]
