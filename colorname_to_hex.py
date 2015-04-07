'''
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
    1.0 (21/11/2013):
        - Creation of the ColornameToHex() class, with the attributes:
            - __init__(colorname)
            - get_rgb()
            - get_rgba()
            - get_bgr()
            - get_abgr()
    1.1 (06/12/2013):
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
    1.2 (13/12/2013):
        - Changed the python version to 3.3
    1.3 (14/04/2014):
        - added the strip method to self.colorname to prevent space errors

@since:
    21/11/2013
'''


class ColornameToHex(object):
    '''
    Translates color names into hexadecimal values.
    '''

    def __init__(self, colorname):
        '''
        str -> Colorname() object

        Builds a dictionary where the key is a color name and the value
        is the hexadecimal value in RGB.
        '''
        self.colorname = colorname.lower().strip()
        self.colorname_to_hex = {'aliceblue':           '#F0F8FF',
                                 'antiquewhite':        '#FAEBD7',
                                 'aqua':                '#00FFFF',
                                 'aquamarine':          '#7FFFD4',
                                 'azure':               '#F0FFFF',
                                 'beige':               '#F5F5DC',
                                 'bisque':              '#FFE4C4',
                                 'black':               '#000000',
                                 'blanchedalmond':       '#FFEBCD',
                                 'blue':                '#0000FF',
                                 'blueviolet':          '#8A2BE2',
                                 'brown':               '#A52A2A',
                                 'burlywood':           '#DEB887',
                                 'cadetblue':           '#5F9EA0',
                                 'chartreuse':          '#7FFF00',
                                 'chocolate':           '#D2691E',
                                 'coral':               '#FF7F50',
                                 'cornflowerblue':      '#6495ED',
                                 'cornsilk':            '#FFF8DC',
                                 'crimson':             '#DC143C',
                                 'cyan':                '#00FFFF',
                                 'darkblue':            '#00008B',
                                 'darkcyan':            '#008B8B',
                                 'darkgoldenrod':       '#B8860B',
                                 'darkgrey':            '#A9A9A9',
                                 'darkgreen':           '#006400',
                                 'darkkhaki':           '#BDB76B',
                                 'darkmagenta':         '#8B008B',
                                 'darkolivegreen':      '#556B2F',
                                 'darkorange':          '#FF8C00',
                                 'darkorchid':         '#9932CC',
                                 'darkred':             '#8B0000',
                                 'darksalmon':          '#E9967A',
                                 'darkseagreen':        '#8FBC8F',
                                 'darkslateblue':       '#483D8B',
                                 'darkslategray':       '#2F4F4F',
                                 'darkturquoise':       '#00CED1',
                                 'darkviolet':          '#9400D3',
                                 'deeppink':            '#FF1493',
                                 'deepskyblue':         '#00BFFF',
                                 'dimgray':             '#696969',
                                 'dodgerblue':          '#1E90FF',
                                 'firebrick':           '#B22222',
                                 'floralwhite':         '#FFFAF0',
                                 'forestgreen':         '#228B22',
                                 'fuchsia':             '#FF00FF',
                                 'gainsboro':           '#DCDCDC',
                                 'ghostwhite':          '#F8F8FF',
                                 'gold':                '#FFD700',
                                 'goldenrod':           '#DAA520',
                                 'gray':                '#808080',
                                 'green':               '#008000',
                                 'greenyellow':         '#ADFF2F',
                                 'honeydew':            '#F0FFF0',
                                 'hotpink':             '#FF69B4',
                                 'indianred':           '#CD5C5C',
                                 'indigo':              '#4B0082',
                                 'ivory':               '#FFFFF0',
                                 'khaki':               '#F0E68C',
                                 'lavender':            '#E6E6FA',
                                 'lavenderblush':       '#FFF0F5',
                                 'lawngreen':           '#7CFC00',
                                 'lemonchiffon':        '#FFFACD',
                                 'lightblue':           '#ADD8E6',
                                 'lightcoral':          '#F08080',
                                 'lightcyan':           '#E0FFFF',
                                 'lightgoldenrodyellow': '#FAFAD2',
                                 'lightgray':           '#D3D3D3',
                                 'lightgreen':          '#90EE90',
                                 'lightpink':           '#FFB6C1',
                                 'lightsalmon':         '#FFA07A',
                                 'lightseagreen':       '#20B2AA',
                                 'lightskyblue':        '#87CEFA',
                                 'lightslategray':      '#778899',
                                 'lightsteelblue':      '#B0C4DE',
                                 'lightyellow':         '#FFFFE0',
                                 'lime':                '#00FF00',
                                 'limegreen':           '#32CD32',
                                 'linen':               '#FAF0E6',
                                 'magenta':             '#FF00FF',
                                 'maroon':              '#800000',
                                 'mediumaquamarine':    '#66CDAA',
                                 'mediumblue':          '#0000CD',
                                 'mediumorchid':        '#BA55D3',
                                 'mediumpurple':        '#9370DB',
                                 'mediumseagreen':      '#3CB371',
                                 'mediumslateblue':     '#7B68EE',
                                 'mediumspringgreen':   '#00FA9A',
                                 'mediumturquoise':     '#48D1CC',
                                 'mediumvioletred':     '#C71585',
                                 'midnightblue':        '#191970',
                                 'mintcream':           '#F5FFFA',
                                 'mistyrose':           '#FFE4E1',
                                 'moccasin':            '#FFE4B5',
                                 'navajowhite':         '#FFDEAD',
                                 'navy':                '#000080',
                                 'oldlace':             '#FDF5E6',
                                 'olive':               '#808000',
                                 'olivedrab':           '#6B8E23',
                                 'orange':              '#FFA500',
                                 'orangered':           '#FF4500',
                                 'orchid':              '#DA70D6',
                                 'palegoldenrod':       '#EEE8AA',
                                 'palegreen':           '#98FB98',
                                 'paleturquoise':       '#AFEEEE',
                                 'palevioletred':       '#DB7093',
                                 'papayawhip':          '#FFEFD5',
                                 'peachpuff':           '#FFDAB9',
                                 'peru':                '#CD853F',
                                 'pink':                '#FFC0CB',
                                 'plum':                '#DDA0DD',
                                 'powderblue':          '#B0E0E6',
                                 'purple':              '#800080',
                                 'red':                 '#FF0000',
                                 'rosybrown':           '#BC8F8F',
                                 'royalblue':           '#4169E1',
                                 'saddlebrown':         '#8B4513',
                                 'salmon':              '#FA8072',
                                 'sandybrown':          '#F4A460',
                                 'seagreen':            '#2E8B57',
                                 'seashell':            '#FFF5EE',
                                 'sienna':              '#A0522D',
                                 'silver':              '#C0C0C0',
                                 'skyblue':             '#87CEEB',
                                 'slateblue':           '#6A5ACD',
                                 'slategray':           '#708090',
                                 'snow':                '#FFFAFA',
                                 'springgreen':         '#00FF7F',
                                 'steelblue':           '#4682B4',
                                 'tan':                 '#D2B48C',
                                 'teal':                '#008080',
                                 'thistle':             '#D8BFD8',
                                 'tomato':              '#FF6347',
                                 'turquoise':           '#40E0D0',
                                 'violet':              '#EE82EE',
                                 'wheat':               '#F5DEB3',
                                 'white':               '#FFFFFF',
                                 'whitesmoke':          '#F5F5F5',
                                 'yellow':              '#FFFF00',
                                 'yellowgreen':         '#9ACD32'}

    def get_rgb(self):
        '''
        None -> str

        Returns a hexadecimal RGB value.
        '''
        return self.colorname_to_hex[self.colorname].lower()

    def get_rgba(self):
        '''
        None -> str

        Returns a hexadecimal RGBA value.
        '''
        return self.colorname_to_hex[self.colorname].lower() + 'ff'

    def get_bgr(self):
        '''
        None -> str

        Returns a hexadecimal BGR value.
        '''
        b = self.colorname_to_hex[self.colorname][-2:].lower()
        g = self.colorname_to_hex[self.colorname][3:5].lower()
        r = self.colorname_to_hex[self.colorname][1:3].lower()

        return "#" + b + g + r

    def get_abgr(self):
        '''
        None -> str

        Returns a hexadecimal ABGR value.
        '''
        bgr = self.get_bgr()
        return bgr[0] + "ff" + bgr[1:]
    
