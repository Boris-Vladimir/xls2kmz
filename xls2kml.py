"""
@python version:
    Python 3.4

@summary:
    Interface composed by the Xls2kml() class.
    This class has a constructor:
        - __init__();
    Three non public functions associated to three buttons:
        - __callback();
        - __callback_2();
        - __callback_3();
    And, five more non public fuctions to help the calls:
        - __threads();
        - __create_kmz();
        - __open_Google_Earth();
        - __progressbar();
        - __about();

@note:
    function __init__():
        Class constructor.
        Creates a Tk() object and designs the elements of the window,
        the icon, the frame, the window name, four labels, three
        separators and three buttons.
    function __callback():
        Handler for the button "Abrir EXEL..." (Open EXEL)
        Opens a new window (filedialog.askopenfilename) to choose the
        EXCEL file that is necessary to make the KMZ file.
    function callback_2():
        Handler for the button "Gravar KMZ" (Save KMZ)
        Calls the function self.__threat()
    function callback_3():
        Handler for the button "Sair" (Exit)
        Kills the program window.
    function __threads():
        Creates two threads to run at the same time the functions:
        self.__create_kmz()
        self.__progressbar()
    fucntion __create_kmz():
        Calls the exel_to_kml() atribute of the MotherControl() class
        And when it returns, calls self.__open_Google_Earth()
    fucntion __open_Google_Earth():
        Opens the maded KMZ file in Google Earth
    function __progressbar():
        Designs a progressbar in the main window
    function __about():
        Associated with the Help Menu.
        Creates a new window with the "About" information.

@author:
    Venancio 2000644

@contact:
    venancio.gnr@gmail.com

@organization:
    SDATO - DP - UAF - GNR

@version:
    1.0 (15/11/2013):
        - Creation of a script who builds a Tkinter window with three
          buttons handled by three functions:
            - function callback();
            - function callback_2();
            - function callback_3();

    1.1 (18/11/2013):
        - Implementation of the Xls2kml() class.
        - Creation of the function / class attribute __init__(), who
          designs the elements of the window
        - Alteration of the program name (window title) to "EXEL to
          KMZ" followed by the version number.
        - Alteration of callback_2() function to create a
          MotherControl() object and call his xls_to_kml() attribute,
          and kill the window in the end if all goes well.

    1.2 (29/11/2013):
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
        - Modification of the python version from 2.7 to 3.3.
        - Creation of the new auxiliary non public functions:
            - __threads();
            - __create_kmz();
            - __open_Google_Earth();
            - __progressbar();
            - __about();
        - Modification of window design, added a new image, a progressbar
            a text area so the user can have a feedback of whats happening
            in the program and a Menu whith buttons to File and Help.
        - Added a initialdir argument to the tkFileDialog.askopenfilename
            so it saves the last opened localization.
    1.3 (06/01/14):
        - Added a log file where the program redirects the stderr
    1.4 (13/04/14):
        - Creation of the new auxiliary non public function:
            - __open_file(doc)
        - Added a new button to the menu, "Opções", where we can choose to
            automatically open or not the Google Earth after the KMZ file
            was builted
        - Added a new button to the menu, "Documentação" whith a cascade menu
            which contains more buttons to "Manual", "KMZ Colors", "KMZ Icons",
            and example files to contruct the Excel Files and KMZ ones.
        - Window redesigned
        - Diferent implementation of the Progress Bar
        - Inform the user in the case that the EXCEL file doesn't have as
            the first four columns:
            - latitude;
            - longitude;
            - name;
            - description;
    1.5 (01/12/2014):
        - Added new submenu "Quadrado" to the submenu "Exemplos" of "Documentação"
            menu.

@since:
    15/11/2013
"""

from mother import MotherControl
from tkinter import HORIZONTAL, BOTTOM, CENTER, FALSE, ALL, E, W, EW, SEPARATOR,\
    Tk, Frame, Label, Button, Message, Menu, Canvas, Toplevel, StringVar, BooleanVar, \
    filedialog, ttk, SUNKEN, TOP, TRUE, BOTH, LEFT, Text, NORMAL, DISABLED
from PIL import ImageTk, Image
import os
import sys
import logging
from time import sleep
from threads import MyThread
from logfile import LogFile


class Xls2kml(object):
    '''
    Interface builted in Tkinter()
    '''

    def __init__(self):
        '''
        None -> None

        Builds the Tkinter window and all his elements.
        '''
        # variables ----------------------------------------------------
        # log file
        open("erros.log", "w").close()  # to open and clean the logfile
        logging.basicConfig(level=logging.DEBUG, filename='erros.log')
        sys.stderr = LogFile('stderr')  # Redirect stderr
        self.original_working_dir = os.getcwd()  # original working dir
        self.master = Tk()  # Tk() object
        self.master.title('EXCEL to KMZ Transformer - ver. 1.6')  # window name
        icons = os.getcwd() + os.sep + "icons" + os.sep  # path to icons
        foto_folder = os.getcwd() + os.sep + "fotos"  # path to fotos
        icon = icons + "compass.ico"
        self.master.iconbitmap(icon)  # window icon
        self.master.resizable(width=FALSE, height=FALSE)
        self.master.geometry("548x314")
        self.file_name = ""  # the name of the EXEL file
        self.last_dir = "C:/"
        # image to decorate the window
        self.img = ImageTk.PhotoImage(Image.open(icons + "excel-kmz.jpg"))
        # to use in frame, message, labels and buttons -----------------
        self.message = StringVar()
        self.message.set("\nSelecciona um ficheiro EXCEL")
        bg = "gray25"
        bg1 = "dark orange"
        fc = "white smoke"
        font = ("Helvetica", "8", "bold")
        text0 = " ----- "  # " ------------------------------------------ "
        text1 = " Boris & Vladimir Software "
        text = text0 + text1 + text0

        # Menu ---------------------------------------------------------
        self.menu = Menu(self.master)
        self.master.config(menu=self.menu)
        filemenu = Menu(self.menu)
        self.menu.add_cascade(label="Ficheiro", menu=filemenu)
        filemenu.add_command(label="Sair", command=self.__callback_3)
        filemenu.add_command(label='Pasta Fotos', command=lambda: (self.__open_folder(foto_folder)))
        # --------------------- NOVO -----------------------------------
        self.openGE = BooleanVar()  # não esquecer de importar BooleanVar
        self.openGE.set(False)
        optionsmenu = Menu(self.menu)
        self.menu.add_cascade(label="Opções", menu=optionsmenu)
        optionsmenu.add_checkbutton(label="Não abrir o Google Earth",
                                    onvalue=True, offvalue=False,
                                    variable=self.openGE)
        docsmenu = Menu(self.menu)
        docs = ["docs\manual.pdf", "docs\icons.pdf", "docs\colors.pdf",
                "docs\GPS.xlsx", "docs\GPS.kmz", "docs\Celulas.xlsx",
                "docs\Celulas.kmz", "docs\Foto.xlsx", "docs\Foto.kmz",
                "docs\Quadrado.xls", "docs\Quadrado.kmz"]
        self.menu.add_cascade(label="Documentação", menu=docsmenu)
        docsmenu.add_command(label="Manual",
                             command=lambda: (self.__open_file(docs[0])))
        docsmenu.add_command(label="Ícones",
                             command=lambda: (self.__open_file(docs[1])))
        docsmenu.add_command(label="Cores",
                             command=lambda: (self.__open_file(docs[2])))

        exemplemenu = Menu(docsmenu)
        docsmenu.add_cascade(label="Exemplos", menu=exemplemenu)

        gpsmenu = Menu(exemplemenu)
        exemplemenu.add_cascade(label="Trajetos", menu=gpsmenu)
        gpsmenu.add_command(label="Excel",
                            command=lambda: (self.__open_file(docs[3])))
        gpsmenu.add_command(label="Google Earth",
                            command=lambda: (self.__open_file(docs[4])))

        cellmenu = Menu(exemplemenu)
        exemplemenu.add_cascade(label="Células Telefónicas", menu=cellmenu)
        cellmenu.add_command(label="Excel",
                             command=lambda: (self.__open_file(docs[5])))
        cellmenu.add_command(label="Google Earth",
                             command=lambda: (self.__open_file(docs[6])))

        fotomenu = Menu(exemplemenu)
        exemplemenu.add_cascade(label="Fotos", menu=fotomenu)
        fotomenu.add_command(label="Excel",
                             command=lambda: (self.__open_file(docs[7])))
        fotomenu.add_command(label="Google Earth",
                             command=lambda: (self.__open_file(docs[8])))

        squaremenu = Menu(exemplemenu)
        exemplemenu.add_cascade(label="Quadrado", menu=squaremenu)
        squaremenu.add_command(label="Excel",
                             command=lambda: (self.__open_file(docs[9])))
        squaremenu.add_command(label="Google Earth",
                             command=lambda: (self.__open_file(docs[10])))

        helpmenu = Menu(self.menu)
        self.menu.add_cascade(label='Ajuda', menu=helpmenu)
        helpmenu.add_command(label="Sobre", command=self.__about)
        helpmenu.add_command(label="Ver erros",
                             command=lambda: (self.__open_file("erros.log")))

        # Frame to suport butons, labels and separators ----------------
        self.f = Frame(self.master, bg=bg)
        self.f.pack_propagate(0)  # don't shrink
        self.f.pack(side=BOTTOM, padx=0, pady=0)

        # Message and Labels -------------------------------------------
        self.l1 = Message(
            self.f, bg=bg1, bd=5, fg=bg, textvariable=self.message,
            font=("Helvetica", "13", "bold italic"), width=500).grid(
            row=0, columnspan=6, sticky=EW, padx=5, pady=5)
        self.l2 = Label(
            self.f, image=self.img, fg=bg
            ).grid(row=1, columnspan=6, padx=5, pady=2)
        self.l6 = Label(
            self.f, text=text, font=("Helvetica", "11", "bold"), bg=bg, fg=bg1
            ).grid(row=3, column=2, columnspan=3, sticky=EW, pady=5)

        # Buttons ------------------------------------------------------
        self.b0 = Button(
            self.f, text="Abrir EXCEL...", command=self.__callback, width=10,
            bg="forest green", fg=fc, font=font
            ).grid(row=3, column=0, padx=5, sticky=W)
        self.b1 = Button(
            self.f, text="Gravar KMZ", command=self.__callback_2, width=10,
            bg="DodgerBlue4", fg=fc, font=font
            ).grid(row=3, column=1, sticky=W)
        self.b2 = Button(
            self.f, text="Sair", command=self.__callback_3, width=10,
            bg="orange red", fg=fc, font=font
            ).grid(row=3, column=5, sticky=E, padx=5)

        # Separator ----------------------------------------------------
        # self.s = ttk.Separator(self.f, orient=HORIZONTAL).grid(
        #    row=4, columnspan=5, sticky=EW, padx=5, pady=5)

        # Progressbar --------------------------------------------------
        # self.pb = Canvas(self.f, width=260, height=10)
        self.s = ttk.Style()
        # themes: winnative, clam, alt, default, classic, vista, xpnative
        self.s.theme_use('winnative')
        self.s.configure("red.Horizontal.TProgressbar", foreground='green',
                         background='forest green')
        self.pb = ttk.Progressbar(self.f, orient='horizontal',
                                  mode='determinate',
                                  style="red.Horizontal.TProgressbar")
        self.pb.grid(row=2, column=0, columnspan=6, padx=5, pady=5, sticky=EW)

        # Mainloop -----------------------------------------------------
        self.master.mainloop()

    def __callback(self):  # "Abrir EXEL..." button handler ------------
        '''
        None -> None

        Opens a new window (filedialog.askopenfilename) to choose the
        EXCEL file that is necessary to make the KMZ file.
        '''
        title = 'Selecciona um ficheiro Excel'
        message = 'Ficheiro EXCEL carregado em memória.\nTransforma-o em KMZ!'
        self.file_name = filedialog.askopenfilename(title=title,
                                                    initialdir=self.last_dir)
        self.last_dir = self.file_name[:self.file_name.rfind('/')]

        if self.file_name[self.file_name.rfind('.')+1:] != 'xls' and \
                self.file_name[self.file_name.rfind('.')+1:] != 'xlsx':
            message = self.file_name + ' não é um ficheiro Excel válido!'
        self.message.set(message)

    def __callback_2(self):  # "Gravar KMZ" button handler ---------------
        '''
        None -> None

        Calls the function self.__threat()
        '''
        sleep(1)
        message = 'Ficheiro EXCEL carregado em memória.\nTransforma-o em KMZ!'
        if self.message.get() != message:
            self.message.set("\nEscolhe um ficheiro EXCEL primeiro")
            self.master.update_idletasks()
        else:
            self.message.set("\nA processar...")
            self.master.update_idletasks()
            sleep(1)
            self.__threads()

    def __callback_3(self):  # "Sair" button handler ---------------------
        '''
        None -> None

        Kills the window
        '''
        self.master.destroy()

    def __threads(self):
        '''
        None -> MyTread() objects

        Creates two threads to run at the same time the functions:
        self.__create_kmz()
        self.__progressbar()
        '''
        funcs = [self.__create_kmz, self.__progressbar]
        threads = []
        nthreads = list(range(len(funcs)))

        for i in nthreads:
            t = MyThread(funcs[i], (), funcs[i].__name__)
            threads.append(t)

        for i in nthreads:
            threads[i].start()

    def __create_kmz(self):
        '''
        None -> None

        Calls the exel_to_kml() atribute of the MotherControl() class
        And when it returns, calls self.__open_Google_Earth()
        '''
        kmz = MotherControl(self.file_name, self.original_working_dir).exel_to_kml()
        if type(kmz) == str:
            self.message.set(kmz)
            self.pb.stop()
            self.master.update_idletasks
        else:
            sleep(2)
            self.pb.stop()
            self.master.update_idletasks()
            self.__open_Google_Earth()

    def __open_Google_Earth(self):
        '''
        None -> None

        Opens the maded KMZ file in Google Earth
        '''
        sleep(1)
        self.master.update_idletasks()
        if not self.openGE.get():
            self.message.set("KMZ gravado com sucesso.\nA abrir o Google Earth...")
        else:
            self.message.set("\nKMZ gravado com sucesso.\n")
        sleep(2)
        self.master.update_idletasks()
        path = self.file_name[:self.file_name.rindex('/')]
        path_1 = self.file_name[self.file_name.rindex('/')+1:self.file_name.rfind('.')]
        kmzs = [x for x in os.listdir(path) if x[-4:] == '.kmz' and x[:-12] ==
                path_1]
        kmzs.sort()
        try:
            if not self.openGE.get():
                os.startfile(path + os.sep + kmzs[-1])
                sleep(2)
            self.message.set("\nSelecciona um ficheiro EXCEL")
        except:
            self.message.set("Instale o Google Earth\nhttp://www.google.com/earth/")
            self.master.update_idletasks()

    def __progressbar(self, ratio=0):
        '''
        None -> None

        Starts the progressbar in the window
        '''
        self.pb.start(50)

    def __about(self):
        '''
        None -> None

        Associated with the Help Menu.
        Creates a new window with the "About" information
        '''
        appversion = "1.6"
        appname = "EXCEL to KML Transformer"
        copyright = 14 * ' ' + '(c) 2013' + 12 * ' ' + \
            'SDATO - DP - UAF - GNR\n' + 34 * ' '\
            + "All Rights Reserved"
        licence = 18 * ' ' + 'http://opensource.org/licenses/GPL-3.0\n'
        contactname = "Nuno Venâncio"
        contactphone = "(00351) 969 564 906"
        contactemail = "venancio.gnr@gmail.com"

        message = "Version: " + appversion + 5 * "\n"
        message0 = "Copyright: " + copyright + "\n" + "Licença: " + licence
        message1 = contactname + '\n' + contactphone + '\n' + contactemail

        icons = os.getcwd() + os.sep + "icons" + os.sep  # path to icons
        icon = icons + "compass.ico"

        tl = Toplevel(self.master)
        tl.configure(borderwidth=5)
        tl.title("Sobre...")
        tl.iconbitmap(icon)
        tl.resizable(width=FALSE, height=FALSE)
        f1 = Frame(tl, borderwidth=2, relief=SUNKEN, bg="gray25")
        f1.pack(side=TOP, expand=TRUE, fill=BOTH)

        l0 = Label(f1, text=appname, fg="white", bg="gray25",
                   font=('courier', 16, 'bold'))
        l0.grid(row=0, column=0, sticky=W, padx=10, pady=5)
        l1 = Label(f1, text=message, justify=CENTER,
                   fg="white", bg="gray25")
        l1.grid(row=2, column=0, sticky=E, columnspan=3, padx=10, pady=0)
        l2 = Label(f1, text=message0,
                   justify=LEFT, fg="white", bg="gray25")
        l2.grid(row=6, column=0, columnspan=2, sticky=W, padx=10, pady=0)
        l3 = Label(f1, text=message1,
                   justify=CENTER, fg="white", bg="gray25")
        l3.grid(row=7, column=0, columnspan=2, padx=10, pady=0)

        button = Button(tl, text="Ok", command=tl.destroy, width=10)
        button.pack(pady=5)

    def __open_file(self, doc):
        try:
            os.startfile(doc)
        except:
            pass
            # os.system(doc)
            # não gosto disto mas os.startfile(doc)
            # faz com que a janela não se desenhe bem

    def __open_folder(self, folder):
        os.system('start explorer "' + folder + '"')
        

if __name__ == '__main__':
    Xls2kml()
