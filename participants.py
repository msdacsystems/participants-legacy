# -*- coding: utf-8 -*-

"""
------------------------------
 MSDAC Systems - Participants
------------------------------

A basic participant organizer and generator for Seventh-day Adventist Church.

This module provides all functionality of organizing, customizing, exporting,
and automating the presentation layout of any church-related service.

Pre-requisites:
    - Microsoft(c) Office 2016 and above

Spacing Gap Formats:
    • Class - 3 to 4 lines
    • Functions - 2 to 3 lines
    • Comments - 1 line

Disclaimer:
    This is an experimental program and we are aware that this software is slow
    due to Python's nature of interpreting code at runtime which affects speed.
    We are planning to switch to a better codebase later on though.


This program is part of MSDAC System's collection of softwares
Made with Qt
(c) 2021-present Ken Verdadero, Reynald Ycong
"""


## Import Modules
import sys


try:
    import os, psutil, winreg, time, datetime, json, shutil, gc
    from PyQt5 import QtCore, QtGui, QtWidgets
    from PyQt5.QtCore import Qt
    from PyQt5.QtGui import QFont
    from pptx import Presentation
    from pptx.util import Inches, Cm, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from kenverdadero.KCore import KPath, KString
    from kenverdadero.KLogging import KLog
    from kenverdadero.KSoftware import KSoftware
    from kenverdadero.KCore.KCore import modHex, p, showLatency
except ImportError as e:
    print(f'Module not found: {e}')
    sys.exit()


        

class System(object):
    """
    System Class Handler
    
    Handles all system related functions for program to work including:
        - System variables
        - Verifying all required files and directory
        - Duplicate instances
        - Application exit event
    """
    def __init__(self):
        """
        All variables have the following format:
            self.TYPE_NAME_ATTRIBUTE_SUBATTR
            
            Example: self.SOFTWARE_HYMNALBROWSER_UI_SETTINGS
        """

        ## External File and Directories
        self.DIR_PARENT =           r'C:\ProgramData\MSDAC Systems'                                                         ## Parent Directory
        self.DIR_PROGRAM =          self.DIR_PARENT + r'\Participants'                                                      ## Program Directory
        self.DIR_LOG =              self.DIR_PROGRAM + r'\Logs'                                                             ## Log Directory
        self.FILE_DATA =            self.DIR_PROGRAM + r'\data.json'                                                        ## Both data and configuration are stored here
        self.FILE_PPT_EXPORTED =    self.DIR_PROGRAM + r'\exported.pptx'                                                    ## Path of the exported file

        ## Resources
        self.RES_LOGO =             './res/images/logo.png'                                                                 ## Resource SDA Logo Image
        self.RES_HEADERLOGO =       './res/images/header.png'                                                               ## Resource Main Header Image
        self.RES_DEFAULT_BG =       'res/images/defBG.png'                                                                  ## Resource Default Background Image
        self.RES_FONT_EXTS =        ['Cameliya.otf', 'HarrietTextBold.otf', 'HarrietTextBoldItalic.otf']                    ## Resource External Fonts for PowerPoint
        
        ## Properties
        self.PROCESS_NAME =         "participants.exe"                                                                      ## Program filename
        self.PROCESS =              psutil.Process(os.getpid())                                                             ## Get PID to detect multiple instances
        self.LOG_FILE_LIMIT =       10                                                                                      ## Maximum threshold for maintaining log files
        self.STARTUP_TIME =         0                                                                                       ## Set initial time for launching the program
        self.GLOBAL_STATE =         1                                                                                       ## 0 - Starting (unused), 1 - Ready, 2 - Reserved, 3 - Shutting Down


    def verifyDirectories(self):
        """
        This method checks for directories and also generate new if the folders does not exist/
        This verifies from parent directory down to subfolders via loop using a dict of directories.

        ..  - Root
        ... - Etc
        <>  - File
        ->  - Directory

        Tree:
            .. MSDAC Systems (Parent/Root)
                -> Participants (Program)
                    -> Logs
                        <> someLogFile.log
                        <> ...
                    <> data.json

            -> POWERPNT.exe (MS Office)
        """
        DIRECTORIES = {
            self.DIR_PARENT: "Parent",
            self.DIR_PROGRAM: "Program",
            self.DIR_LOG: "Logs"
            }
        MISSING = 0

        for DIR, NAME in DIRECTORIES.items():
            if not KPath.exists(DIR, True):
                LOG.warn(f"{NAME} Directory \"{DIR}\" doesn't exist. Creating a new folder.")
                MISSING += 1
        

    def verifyRequisites(self):
        """
        Verifies all required programs to make the software run properly.
        This method prevents the program to run if a valid Microsoft Office
        PowerPoint is not found in both 32-bit and 64-bit installation folders
        """
        try:
            ## Uses Windows Registry to find the current path of MS Office
            self.PPT_EXEC = KPath.upFolder(winreg.EnumValue(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe"), 0)[1])
            LOG.info(f'Using {"64" if "Program Files (x86)" not in self.PPT_EXEC else "32"}-Bit Version of Microsoft Office')
            LOG.info(f'Root Dir: {self.PPT_EXEC}')
        except FileNotFoundError:
            # No MS Office Installed/Detected
            LOG.crit("Microsoft Office is not installed or detected")
            MSG_BOX = QtWidgets.QMessageBox()
            MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok)
            MSG_BOX.setIcon(QtWidgets.QMessageBox.Critical)
            MSG_BOX.setText("Microsoft Office is not installed or detected.\nThis program uses Office PowerPoint to run properly.\n")
            MSG_BOX.setDetailedText("If you think this is a mistake, please contact the developers for further assistance.\n\nhttps://m.me/verdaderoken\nhttps://m.me/reynald.ycong")
            MSG_BOX.setWindowTitle(f"{SW.NAME} - Error")
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            
            MSG_BOX.setStyleSheet(QSS.getStylesheet())
            MSG_BOX.exec_()
            LOG.sys("Program terminated due to an error: Cannot find PowerPoint directory.")
            sys.exit()
        return


    def checkInstances(self):
        """
        Asks the user if they want to run another instance
        """
        self.DUPLICATED = False
        DUPLICATES = list(filter(lambda x: x == SYS.PROCESS_NAME, [i.name() for i in psutil.process_iter()]))
        
        if len(DUPLICATES) > 2:
            MSG_BOX = QtWidgets.QMessageBox()
            MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            MSG_BOX.setIcon(QtWidgets.QMessageBox.Question)
            MSG_BOX.setText("The program is already running.\nDo you want to open another instance?")
            MSG_BOX.setWindowTitle("Duplicate Instance Detected")
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            MSG_BOX.setStyleSheet(QSS.getStylesheet())
            MSG_BOX.setStyleSheet('min-width: 280px; min-height: 35px;')
            
            LOG.info("Duplicate Instance Detected")
            RET = MSG_BOX.exec_()
            if RET == QtWidgets.QMessageBox.Yes:
                return
            else:
                LOG.sys("Program terminated by user.")
                self.DUPLICATED = True
                sys.exit()
                
    
    def closeEvent(self, event):
        """
        Triggers this function when the UI is shut down
        Reports back to the logger.
        """
        
        if self.DUPLICATED:
            LOG.sys(f"Duplicate instance was shut down.")
        else:
            SYS.GLOBAL_STATE = 3
            LOG.sys(f"Shutting down")
            UIA.close()
            UIB.close()
            LOG.sys(f"Program terminated: Duration ({SW.runtime(2)} Seconds)")
            LOG.sys(f"END OF LOG - {datetime.datetime.now().strftime('%B %d, %Y - %I:%M:%S %p')}")




class Data(object):
    """
    Manages the pool database for Participants
    Uses JSON to parse the data.

    DCFG (or Database Configuration) is the dictionary of all recent role-name relations
    along with the configuration settings for users.
    """
    def __init__(self):
        self.check()


    def check(self):
        """
        Checks if the file is existent in directory
        """
        if not os.path.exists(SYS.FILE_DATA):
            LOG.warn("Data is missing. Generating new...")
            self.generateDefault()

        self.DATA = self.load()


    def load(self):
        """
        Retrieves the data from system's data file.
        """
        while True:
            with open(SYS.FILE_DATA, "r") as read:
                try: 
                    DATA = json.load(read)
                except json.decoder.JSONDecodeError:
                    LOG.crit('Failed to load statistic data. Regenerating default...')
                    self.generateDefault()
                else:
                    return DATA
        

    def dump(self, data=None, indent=None, sort_keys=False):
        """
        Saves the passed data to system's data file.
        Uses indention of 4 and sorted keys by default.
        Can be processed with other data if 2nd argument is specified.
        """
        if data is None: data = DCFG
        with open(SYS.FILE_DATA, "w") as write:
            json.dump(data, write, indent=indent, sort_keys=sort_keys)

    
    def generateDefault(self):
        """
        Generates default data
        """
        DATA = {
            "__DATECREATED__": time.time(),
            "__FILETYPE__": f"{SW.NAME} Data",
            "POOL": {
                "ROLES": ['Developed by'],
                "NAMES": ['MSDAC Systems']
            },
            "CONFIG": {
            }
        }
        self.dump(DATA)
        LOG.info(f"Default data was generated successfully. | Hash: {KString.toHashMD5(DATA)}")




class FileManager(object):
    """
    Manages all external files covered by the software
    This includes managing of recent files and temporary folder
    to maintain and prevent building up of unused data.
    """
    def __init__(self):
        self.deleteLogs()
        # self.checkExternalFonts()                                                                             ## Temporarily disabled


    def checkExternalFonts(self):
        """
        Auto installation of external custom fonts
        Needs administrator permissions.
        """
        FONTS = os.listdir("C:/Windows/Fonts")
        EXTS = ([f'res/fonts/{i}' for i in SYS.RES_FONT_EXTS])
        for f in EXTS:
            if f not in FONTS:
                shutil.copy(os.path.join(os.path.dirname(__file__), f), "C:/Windows/Fonts")


    def deleteLogs(self, deleteAll=False):
        """
        Orders the program to remove logs
        """
        self.deleteOldest(self.getLogFiles, SYS.LOG_FILE_LIMIT, deleteAll)


    def getLogFiles(self):
        """
        Returns list of log file paths
        """
        return [f'{SYS.DIR_LOG}\\{i}' for i in os.listdir(SYS.DIR_LOG) if i.endswith('.log')]


    def deleteOldest(self, fileList, threshold, deleteAll):
        """
        Delete older files when the folder items reached the maximum recent files allowed.
        """
        while len(fileList()) > (threshold if not deleteAll else 0):                                            ## Process code until the detected files are below threshold or when the switch is set to "Delete All" (0)
            try: os.remove(min(fileList(), key=os.path.getatime))                                               ## Eliminates the oldest accessed file
            except PermissionError as e: LOG.warn(e)                                                            ## Issue: There is no solution for this one yet. Administrator permissions could be used in future development
            except FileNotFoundError as e: pass                                                                 ## Ignore when file is not there anymore




class Stylesheet(object):
    """
    Handles all interface appearance for this program
    """
    def __init__(self):
        self.toggleMode()                                                                       ## Sets global palette for the application
        self.initStylesheet()                                                                   ## Sets the appearance of the UI's elements
    

    def setupFonts(self):
        """
        Initializes fonts to be used
        """
        FONT_BASE = QFont("Segoe UI", 9)
        FONT_BASE.setStyleStrategy(QFont.PreferAntialias)
        FONT_COLUMN = QFont('Segoe UI', 10)
        FONT_COLUMN.setStyleStrategy(QFont.PreferAntialias)
        FONT_COLUMN.setBold(True)
        APP.setFont(FONT_BASE)
        UIA.LBL_RL.setFont(FONT_COLUMN)
        UIA.LBL_NM.setFont(FONT_COLUMN)


    def initStylesheet(self):
        """
        Updates the stylesheet
        """
        SS = self.getStylesheet() 

        try:
            APP.setStyleSheet(SS)
            UIA.setStyleSheet(SS)
        except (NameError, AttributeError) as e:
            pass


    def QCl(self, c):
        """
        Returns QColor version of a hex value
        """
        if '#' in c: c = c[1:]
        return QtGui.QColor(int(c[0:2],16),int(c[2:4],16),int(c[4:6],16))
    

    def palette2Hex(self, color):
        """
        Returns HEX value of an RGB of a certain palette color
        """
        return self.RGBtoHEX(getattr(APP.palette(), color)().color().getRgb())
    
    
    def RGBtoHEX(self, rgb):
        rgb = list(rgb); del rgb[3]
        return '#%02x%02x%02x' % tuple(rgb)


    def toggleMode(self, mode=0):
        """
        Sets and updates all application color palette
        """

        if not mode:
            """
            For Light Mode
            """
            self.PRIMARY = '#004B74'
            self.SECONDARY = '#008A9A'
            self.TERTIARY = '#F7EBC5'
            self.GOLD = '#FFA92D'
            self.WARN = '#D25900'
            self.ERROR = '#9E1919'
            self.BORDER = '#C0C0C0'
            self.BORDER_HIGHLIGHT = '#C0C0C0'
            self.BTN_DISABLED = '#B6B6B6'
            self.TXT_INV = '#FFFFFF'
            self.TXT_DISABLED = '#999999'
            self.TXT_STATUSBAR = '#707070'
            self.STATUSBAR = '#CFCFCF'
            self.SCROLLBAR = '#A0A0A0'
            self.CARD = '#DDDDDD'
            self.CARDHOVER = '#EEEEEE'
            self.CTX_MENU = '#FFFFFF'
            PLT_LIGHT = QtGui.QPalette()
            PLT_LIGHT.setColor(QtGui.QPalette.Window, self.QCl('#FFFFFF'))
            PLT_LIGHT.setColor(QtGui.QPalette.WindowText, self.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.Base, self.QCl('#DADADA'))
            PLT_LIGHT.setColor(QtGui.QPalette.AlternateBase, self.QCl('#2D2D2D'))
            PLT_LIGHT.setColor(QtGui.QPalette.ToolTipBase, self.QCl('#252525'))
            PLT_LIGHT.setColor(QtGui.QPalette.ToolTipText, self.QCl('#C5C5C5'))
            PLT_LIGHT.setColor(QtGui.QPalette.PlaceholderText, self.QCl('#999999'))
            PLT_LIGHT.setColor(QtGui.QPalette.HighlightedText, self.QCl('#EEEEEE'))
            PLT_LIGHT.setColor(QtGui.QPalette.Highlight, self.QCl(self.PRIMARY))
            PLT_LIGHT.setColor(QtGui.QPalette.Light, self.QCl('#D7D7D7'))
            PLT_LIGHT.setColor(QtGui.QPalette.Text, self.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, self.QCl('#434343')) ## <- Unused
            PLT_LIGHT.setColor(QtGui.QPalette.Midlight, self.QCl('#888888'))
            PLT_LIGHT.setColor(QtGui.QPalette.Mid, self.QCl('#D2D2D2'))
            PLT_LIGHT.setColor(QtGui.QPalette.Dark, self.QCl('#555555'))
            PLT_LIGHT.setColor(QtGui.QPalette.Button, self.QCl('#CCCCCC'))
            PLT_LIGHT.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Button, self.QCl('#252525')) ## <- Unused
            PLT_LIGHT.setColor(QtGui.QPalette.ButtonText, self.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.BrightText, self.QCl('#FFFFFF'))
            PLT_LIGHT.setColor(QtGui.QPalette.Link, self.QCl('#202020'))
            PLT_LIGHT.setColor(QtGui.QPalette.LinkVisited, self.QCl('#151515'))
            APP.setPalette(PLT_LIGHT)

        elif mode == 1:
            """
            For Dark Mode
            """
            self.PRIMARY = '#008A9A' ## #008A9A Original
            self.SECONDARY = '#004B74'
            self.TERTIARY = '#F7EBC5'
            self.GOLD = '#FFA92D'
            self.WARN = '#D25900'
            self.ERROR = '#9E1919'
            self.BORDER = '#303030'
            self.BORDER_HIGHLIGHT = '#505050'
            self.BTN_DISABLED = '#212121'
            self.TXT_INV = '#181818'
            self.TXT_DISABLED = '#434343'
            self.TXT_STATUSBAR = '#434343'
            self.STATUSBAR = '#2A2A2A'
            self.SCROLLBAR = '#3B3B3B'
            self.CARD = '#2A2A2A'
            self.CARDHOVER = '#323232'
            self.CTX_MENU = '#1D1D1D'
            PLT_DARK = QtGui.QPalette()
            PLT_DARK.setColor(QtGui.QPalette.Window, self.QCl('#202020'))
            PLT_DARK.setColor(QtGui.QPalette.WindowText, self.QCl('#D5D5D5'))
            PLT_DARK.setColor(QtGui.QPalette.Base, self.QCl('#191919'))
            PLT_DARK.setColor(QtGui.QPalette.AlternateBase, self.QCl('#2D2D2D'))
            PLT_DARK.setColor(QtGui.QPalette.ToolTipBase, self.QCl('#252525'))
            PLT_DARK.setColor(QtGui.QPalette.ToolTipText, self.QCl('#C5C5C5'))
            PLT_DARK.setColor(QtGui.QPalette.PlaceholderText, self.QCl('#999999'))
            PLT_DARK.setColor(QtGui.QPalette.HighlightedText, self.QCl('#191919'))
            PLT_DARK.setColor(QtGui.QPalette.Highlight, self.QCl(self.PRIMARY))
            PLT_DARK.setColor(QtGui.QPalette.Light, self.QCl('#898989'))
            PLT_DARK.setColor(QtGui.QPalette.Text, self.QCl('#EFEFEF'))
            PLT_DARK.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, self.QCl('#939393')) ## <- Unused
            PLT_DARK.setColor(QtGui.QPalette.Midlight, self.QCl('#888888'))
            PLT_DARK.setColor(QtGui.QPalette.Mid, self.QCl('#424242'))
            PLT_DARK.setColor(QtGui.QPalette.Dark, self.QCl('#555555'))
            PLT_DARK.setColor(QtGui.QPalette.Button, self.QCl('#353535'))
            PLT_DARK.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Button, self.QCl('#252525')) ## <- Unused
            PLT_DARK.setColor(QtGui.QPalette.ButtonText, self.QCl('#EFEFEF'))
            PLT_DARK.setColor(QtGui.QPalette.BrightText, self.QCl('#FFFFFF'))
            PLT_DARK.setColor(QtGui.QPalette.Link, self.QCl('#D0D0D0'))
            PLT_DARK.setColor(QtGui.QPalette.LinkVisited, self.QCl('#CECECE'))
            APP.setPalette(PLT_DARK)


    def getStylesheet(self, objectName=None):
        """
        Returns a string of stylesheet that will be used by QStyleSheet
        Values depends on what is the current theme.
        """
        # self.getThemes()
        RADIUS = "6px" ## Default: 9px
        RADIUS_SML = "4px" ## Default: 5px
        PADDING = "5px"

        if objectName is None:
            STYLESHEET = f"""
                QWidget#WIN_PARTICIPANTS {{
                    image: url('./res/images/bg.png');
                    image-position: bottom;
                }}



                /* Buttons */ 
                QPushButton {{
                    background-color: palette(button);
                    padding: {PADDING};
                    border-radius: {RADIUS};
                    outline: none;
                    border: 1px solid {modHex(self.palette2Hex('button'), 7)}
                }}
                QPushButton::pressed {{
                    background-color: palette(button);
                }}
                QPushButton::disabled {{
                    color: {self.TXT_DISABLED};
                    background-color: {self.BTN_DISABLED};
                }}
                QPushButton::hover {{
                    background-color: palette(light);
                }}
                /*
                QPushButton::focus {{
                    border: 1px solid {self.BORDER};
                }}
                */



                /* Export Powerpoint Button */ 
                QPushButton#BTN_POWERPOINT {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/ppt.png');
                }}

                QPushButton::hover#BTN_POWERPOINT {{
                    background-color: {self.PRIMARY};
                    image: url('./res/icons/ppt_hover.png');
                }}

                QPushButton::disabled#BTN_POWERPOINT {{
                    image: url('./res/icons/ppt_disabled.png');
                }}


                /* Export Plain Text Button */ 
                QPushButton#BTN_PLAINTEXT {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/plaintext.png');
                }}

                QPushButton::hover#BTN_PLAINTEXT {{
                    background-color: {self.PRIMARY};
                    image: url('./res/icons/plaintext_hover.png');
                }}

                QPushButton::disabled#BTN_PLAINTEXT {{
                    image: url('./res/icons/plaintext_disabled.png');
                }}



                /* Set Active Button  */ 
                QPushButton#BTN_ATVS {{
                    background-color: none;
                    border: none;
                    image: none;
                }}

                QPushButton::hover#BTN_ATVS {{
                    image: url('./res/icons/radio_translucent_hover.png');
                }}

                QPushButton::disabled#BTN_ATVS {{
                    image: url('./res/icons/radio_disabled.png');
                }}



                /* Set Active Button SELECTED */ 
                QPushButton#BTN_ATVS_SELECTED {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/radio_selected.png');
                }}

                QPushButton::hover#BTN_ATVS_SELECTED {{
                    image: url('./res/icons/radio_selected_hover.png');
                }}

                QPushButton::disabled#BTN_ATVS_SELECTED {{
                    image: url('./res/icons/select_radio_disabled.png');
                }}



                /* Insert/Add Button */ 
                QPushButton#BTN_INSS {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/add.png');
                }}

                QPushButton::hover#BTN_INSS {{
                    image: url('./res/icons/add_hover.png');
                }}

                QPushButton::disabled#BTN_INSS {{
                    image: url('./res/icons/add_disabled.png');
                }}



                /* Remove Button + Discard BG Img Button (Settings) */ 
                QPushButton#BTN_REMS, QPushButton#BTN_DISCARD {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/xmark.png');
                }}

                QPushButton::hover#BTN_REMS, QPushButton::hover#BTN_DISCARD {{
                    image: url('./res/icons/xmark_hover.png');
                }}

                QPushButton::disabled#BTN_REMS, QPushButton::disabled#BTN_DISCARD {{
                    image: url('./res/icons/xmark_disabled.png');
                }}



                /* Save List Button */
                QPushButton#BTN_SAVELIST {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/save.png');
                }}

                QPushButton::hover#BTN_SAVELIST {{
                    image: url('./res/icons/save_hover.png');
                }}

                QPushButton::disabled#BTN_SAVELIST {{
                    image: url('./res/icons/save_disabled.png');
                }}



                /* Settings Button (Gear) */ 
                QPushButton#BTN_SETTINGS {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/settings.png');
                }}

                QPushButton::hover#BTN_SETTINGS {{
                    image: url('./res/icons/settings_hover.png');
                }}


                /* Browse Button (Settings) */
                QPushButton#BTN_BROWSE {{
                    background-color: none;
                    border: none;
                    image: url('./res/icons/folder.png');
                }}

                QPushButton::hover#BTN_BROWSE {{
                    image: url('./res/icons/folder_hover.png');
                }}

                QPushButton::disabled#BTN_BROWSE {{
                    image: url('./res/icons/folder_disabled.png');
                }}


                /* Dialog Boxes */
                QMessageBox {{
                    background-color: palette(window);
                }}
                


                /* Tooltip */
                QToolTip {{
                    color: palette(text);
                    background-color: palette(base);
                    border: none;
                }}



                /* Status Bar */
                QStatusBar#STATUSBAR {{
                    color: {self.TXT_STATUSBAR};
                    background-color: {self.STATUSBAR};
                }}



                /* Search Bars */
                QLineEdit, QComboBox{{
                    color: palette(text);
                    selection-color: {self.TXT_INV};
                    background-color: {self.CARD};
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                    padding: {PADDING};
                }}
                QLineEdit::focus#LNE_SEARCH {{
                    background-color: #AF{self.BORDER[1:]};
                }}
                QLineEdit::hover#LNE_SEARCH {{
                    border: 1px solid {self.BORDER_HIGHLIGHT};
                }}



                /* Combo Boxes */
                QComboBox {{
                    background-color: transparent;
                    border: 1px solid {self.BORDER};
                }}
                QComboBox::hover {{
                    background-color: transparent;
                    border: 1px solid {self.SECONDARY};
                }}

                QComboBox#CBX_RLS {{
                    margin-bottom: 1px;
                }}
                QComboBox#CBX_NMS {{
                    margin-bottom: 1px;
                }}



                /* Sliders */
                QSlider::handle {{
                    background-color: {self.PRIMARY};
                }}
                QSlider::handle::pressed {{
                    background-color: {self.SECONDARY};
                }}



                /* ScrollBars */
                QScrollBar:vertical {{
                    background-color: palette(base);
                    width: 16px;
                    margin: 0px;
                    border-radius: {RADIUS_SML};
                }}
                QScrollBar:horizontal {{
                    background-color: palette(base);
                    height: 15px;
                    margin: 0px;
                    border-radius: {RADIUS_SML};
                }}
                QScrollBar::handle:vertical {{
                    background-color: {self.SCROLLBAR}; min-height: 20px; margin: 3px; border-radius: {RADIUS_SML}; border: none;
                }}
                QScrollBar::handle:horizontal {{
                    background-color: {self.SCROLLBAR}; min-width: 20px; margin: 3px; border-radius: {RADIUS_SML}; border: none;
                }}
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                    border: none; background: none; height: 0px;
                }}
                QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
                    border: none; background: none; width: 0px;
                }}
                QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical, QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                    border: none; background: none; color: none;
                }}
                QScrollBar::left-arrow:horizontal, QScrollBar::right-arrow:horizontal, QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
                    border: none; background: none; color: none;
                }}
                


                /* Settings Panel (List) */
                QAbstractItemView {{
                    color: palette(text);
                    outline: none;
                    background-color: palette(base);
                    border: none;
                    border-radius: {RADIUS};
                    selection-color: {self.TXT_INV};
                    selection-background-color: {self.SECONDARY};
                    padding: 5px;
                    min-height: 18px;
                }}
                QAbstractItemView::item {{
                    padding: 6px 3px 6px 5px;
                    margin: 2px 0px 2px 0px;
                    border-radius: {RADIUS};
                }}
                QAbstractItemView::item::selected {{
                    background-color: {self.PRIMARY};
                }}
                QAbstractItemView::item::hover {{
                    color: palette(text);
                    background-color: palette(window);
                }}
                QAbstractItemView::item::selected::hover {{
                    color: {self.TXT_INV};
                    background-color: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,stop: 0 {modHex(self.PRIMARY, 50)}, stop: 1 {self.PRIMARY});
                }}
                QListWidget {{
                    font-style: 'Segoe UI Variable Display';
                    font-weight: bold;
                    font-size: 10pt;
                }}

                /* Log Panel */
                QPlainTextEdit {{
                    color: palette(text);
                    background-color: palette(base);
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                }}



                /* Spin

                /* Spin Box */
                QSpinBox {{
                    border: 1px solid {self.BORDER};
                    border-radius: {RADIUS};
                }}
                

                /* Add Queue Button */
                QPushButton#BTN_ADDQUEUE {{
                    background-color: transparent;
                    min-width: 17px;
                    width: 20px;
                    height: 20px;
                    image: url('./res/icons/add_.png');
                }}
                QPushButton::hover#BTN_ADDQUEUE {{
                    image: url('./res/icons/add_hover_.png');
                }}


                /* Queue Button */
                QPushButton#BTN_QUEUES {{
                    background-color: transparent;
                    min-width: 17px;
                    width: 20px;
                    height: 20px;
                    image: url('./res/icons/queue_.png');
                }}
                QPushButton::hover#BTN_QUEUES {{
                    image: url('./res/icons/queue_hover_.png');
                }}



                /* Special */

                QLabel#PIX_HEADER {{
                    color: {self.PRIMARY};
                }}
                QLabel::disabled, QLabel#LBL_BROWSERB {{
                    color: {self.TXT_DISABLED};
                }}
                
                QPushButton::enabled#BTN_LAUNCH, QPushButton::enabled#BTN_OK {{
                    color: {self.TXT_INV};
                    background-color: {self.PRIMARY};
                }}
                QPushButton#BTN_LAUNCH {{
                    min-width: 100px;
                }}
                QPushButton::hover#BTN_LAUNCH, QPushButton::hover#BTN_OK {{
                    color: {self.TXT_INV};
                    background-color: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,stop: 0 {modHex(self.PRIMARY, 50)}, stop: 1 {self.PRIMARY});
                }}
                
                QPushButton::hover#BTN_RESET {{
                    color: palette(highlighted-text);
                    background-color: {self.ERROR};
                }}
                QLabel#LBL_RL, QLabel#LBL_NM {{
                    color: #005278;
                    padding: 3px;
                    border-radius: {RADIUS};
                }}
                """
            return STYLESHEET




class Package(object):
    """
    Represents a single package of configuration for a generated PowerPoint
    """
    def __init__(self):
        self.setupDefaultVariables()
        self.loadConfig()


    def restart(self):
        """
        Re-initiates all the process from the top
        """
        self.__init__()


    def setupDefaultVariables(self):
        """
        Initiates all default variables before loading the user-customized values
        Works as a placeholder incase something went wrong from user
        """
        self.DEF_DIR_EXPORT_RECENT = SW.DIR_CWD
        self.DIR_EXPORT_RECENT = SW.DIR_CWD
        self.DEF_IMG_BACKGROUND = "res/images/defBG.png"
        self.IMG_BACKGROUND = self.DEF_IMG_BACKGROUND
        self.FONT_TITLE = "Harriet Text Bold"
        self.FONT_SUBTITLE = "Cameliya"
        self.FONT_DATE = "Harriet Text Bold"
        self.FONT_CONTENT = "Harriet Text Bold"
        self.TXT_TITLE = "Sabbath Worship Participants"
        self.TXT_SUBTITLE = "Happy Sabbath!"
        self.RGB_TITLE = RGBColor(255, 255, 255)
        self.RGB_SUBTITLE = RGBColor(255, 169, 45)
        self.RGB_DATE = RGBColor(255, 255, 255)
        self.RGB_ROLES = RGBColor(221, 221, 221)
        self.RGB_NAMES = RGBColor(255, 255, 255)
        self.FSZ_TITLE = Pt(40)
        self.FSZ_SUBTITLE = Pt(30)
        self.FSZ_DATE = Pt(15)
        

    def loadConfig(self):
        """
        Retrieves all saved configurations from data.json
        Works by looping through all variables in this program and finding
        the working ones from absent ones.
        """
        LOG.info('Loading configuration')
        LOAD = [
            (self.IMG_BACKGROUND, "IMG_BACKGROUND"),
            (self.FONT_TITLE, "FONT_TITLE"),
            (self.FONT_SUBTITLE, "FONT_SUBTITLE"),
            (self.FONT_DATE, "FONT_DATE"),
            (self.FONT_CONTENT, "FONT_CONTENT"), 
            (self.TXT_TITLE, "TXT_TITLE"),
            (self.TXT_SUBTITLE, "TXT_SUBTITLE"),
            (self.DIR_EXPORT_RECENT, "DIR_EXPORT_RECENT")
        ]
        for obj in LOAD:
            try:
                for var in dict(self.__dict__):
                    if getattr(self, var) == obj[0]:
                        setattr(self, obj[1], DCFG["CONFIG"][obj[1]])
                        break
            except Exception as e:
                pass

        LOG.info('Successfully loaded configuration.')




class Core(object):
    """
    Contains methods that are used for better user experience
    """
    def __init__(self):
        pass


    def centerWindow(self, ui):
        """
        Relocates the specific UI from argument to center of the screen
        """
        FRM_GEOMETRY = ui.frameGeometry()
        SCREEN = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        FRM_GEOMETRY.moveCenter(QtWidgets.QApplication.desktop().screenGeometry(SCREEN).center())
        ui.move(FRM_GEOMETRY.topLeft())


    def splitContents(self, roles:str, names:str):
        """
        Splits the content to distinguish Sabbath school from Divine Service.
        Has fallback for when the service-based categorizing fails
        """
        for i, cmb in enumerate(FLD.CBX_RLS):
            ## Regular Splitter
            if cmb.currentText().upper() in ["CLOSING PRAYER", "CLOSINGPRAYER", "CLOSING"]:
                SS_A = roles.split('\n', i+1); DS_A = SS_A[-1].split('\n'); SS_A.pop(-1)
                SS_B = names.split('\n', i+1); DS_B = SS_B[-1].split('\n'); SS_B.pop(-1)
                return (SS_A, SS_B, DS_A, DS_B)
                
        ## Basic Splitter (Fallback)
        SS_A = roles.split('\n', int(len(roles.split('\n'))/2)); DS_A = SS_A[-1].split('\n', int(len(SS_A[-1].split('\n'))/2)); SS_A.pop(-1)
        SS_B = names.split('\n', int(len(names.split('\n'))/2)); DS_B = SS_B[-1].split('\n', int(len(SS_B[-1].split('\n'))/2)); SS_B.pop(-1)
        return (SS_A, SS_B, DS_A, DS_B)




class Fields(object):
    """
    Handles all field-related events:
        - Add/Remove field(s)
        - Button refreshers
        - Field data access
        - Export field to plain text
        - Export field to Powerpoint
    
    Field - A group of widgets linked to one participant.
            It contains role, name, and the + & x buttons
    """
    def __init__(self):
        """
        Constructor
        """
        self.CBX_RLS, self.CBX_NMS, self.BTN_REMS, self.BTN_INSS, self.BTN_ATVS = [], [], [], [], []
        self.FIELDS = 0
        self.FIELDS_MAX = 20
        self.PREV_ACTIVE = None


    def setup(self):
        """
        Initiates the first run of field generation for setup.
        Usually triggered after setting up the UI
        """
        LENGTH = len(DCFG['POOL']['ROLES'])
        if not LENGTH:
            PDB.generateDefault()
            LENGTH = 1
        self.addFields(LENGTH)
        self.updateItems()
        self.refreshStates()


    def refreshStates(self):
        """
        Refreshes the states of every field.

        Updates the connections of every button.
        This method helps the class handler to reconnect all connections
        of every present button after recent changes
        """
        ## Reconnect Buttons
        for i in range(len(self.BTN_REMS)):                                                 ## Loops through every single button object based on BTN_REMS or BTN_INSS
            try:
                self.BTN_ATVS[i].clicked.disconnect()
                self.BTN_INSS[i].clicked.disconnect()
                self.BTN_REMS[i].clicked.disconnect()
            except Exception:                                                               ## Throws exception when the buttons aren't connected yet which is not valid in the disconnect method
                pass
            finally:                                                                        ## Always connect the Add and Remove button to its main method
                self.BTN_ATVS[i].clicked.connect(lambda: self.setActiveField())
                self.BTN_INSS[i].clicked.connect(lambda: self.redirectFieldInsertion())
                self.BTN_REMS[i].clicked.connect(lambda: self.removeField())
        
        ## Prevent Overflow
        for btn in self.BTN_INSS:
            if self.FIELDS >= self.FIELDS_MAX:
                btn.setEnabled(False)               ## Sets the button to whether Enabled or Disabled depending on how many fields are active
                btn.setToolTip("")
            else:
                btn.setEnabled(True)
                btn.setToolTip("Add new field below")

        ## Prevent Depletion
        try:
            self.BTN_REMS[0].setEnabled(False if self.FIELDS <= 1 else True)                ## Disables the last field's remove button
        except IndexError:
            pass
        
        UIA.updateWindowTitle()
    

    def redirectFieldInsertion(self):
        """
        Redirects function to UIA's Field Insertion function.
        Triggers from the Add button via PyQt signal 
        """
        for i, btn in enumerate(self.BTN_INSS):
            if btn.hasFocus():
                self.insertField(i)
                break


    def setActiveField(self):
        """
        Sets an active field by exporting its current values into a text file 
        to be read by a Text Source from OBS Studio
        """
        for i, btn in enumerate(self.BTN_ATVS):
            if btn.hasFocus():
                if btn.objectName() == 'BTN_ATVS_SELECTED':
                    ## Unset the field from being active
                    EXP.fromActiveField(i, True)
                    btn.setObjectName('BTN_ATVS')
                else:
                    ## Before setting the triggered button to active, determine the previous
                    ## index to unset from being active and return into a normal state.
                    if self.PREV_ACTIVE is not None:
                        POINTER = self.PREV_ACTIVE
                        try:
                            self.BTN_ATVS[POINTER]
                        except IndexError:
                            POINTER = len(self.BTN_ATVS)-1
                        finally:
                            self.BTN_ATVS[POINTER].setObjectName('BTN_ATVS')
                            self.BTN_ATVS[POINTER].setToolTip('Set as active')
                            self.BTN_ATVS[POINTER].setStyleSheet(QSS.getStylesheet())

                    ## Set the focused button to be active and export the file
                    EXP.fromActiveField(i)
                    btn.setObjectName('BTN_ATVS_SELECTED')
                    btn.setToolTip('Unset from being active')
                    self.PREV_ACTIVE = i

                btn.setStyleSheet(QSS.getStylesheet())
                break


    def insertField(self, pos:int=-1):
        """
        Inserts one new set of widgets with items available from both pools (roles, names).
        Creates Role, Name, Add button, Remove button to be placed on a vertical box layout.

        QVBoxLayout was used because of its insert method which you can use for inserting new
        widget to a specific index position.

        There are some adjustments for `pos` argument because of the nature of how Qt handles
        where it puts the next item in layout.

        `CBX_RLS` (Combo Box Roles) - contains all objects for roles
        `CBX_NMS` (Combo Box Names) - contains all objects for names
        `BTN_REMS` (Remove Push Button(s)) - contains all objects for remove/clear button
        `BTN_INSS` (Insert Push Button(s)) - contains all objects for insert/plus/add button
        `pos` - a positional marker used to mark where the next item should be placed

        Code block below will execute the following in order:
        1. Determining the right positional marker to prepare for insertion
        2. Inserting a placeholder (NoneType) in the newly placed index
        3. Compensates `pos` for zero-based indexing
        4. Starts instantiating all four (4) widgets to make a single-line field
        5. Combo boxes are filled with data from the pool for both Roles and Names
        6. Inserting the item in a QVBoxLayout
        7. Refresh states

        This method is not a loop but it can be used multiple times using a different method
        `addFields` which accepts numbers of how many times you want to generate a field.
        """
        # p([str(i)[-6:-1] for i in self.BTN_REMS]) ## Addresses (For debugging)

        ## Index-related adjustments
        pos +=1                                                                                 ## Increase position number by 1 to prevent placing the object in the first index
        if pos == len(self.CBX_RLS): pos += 1                                                   ## When last button is clicked, increment pos by 1 to ensure that the placeholder will be at the very last
        elif pos == 0: pos -= 1                                                                 ## When the pos is at the start, decrease the pos to compensate with insert function
            
        ## Appends a placeholder for an object to be created
        self.CBX_RLS.insert(pos, None)
        self.CBX_NMS.insert(pos, None)
        self.BTN_REMS.insert(pos, None)
        self.BTN_ATVS.insert(pos, None)
        self.BTN_INSS.insert(pos, None)

        if pos == len(self.CBX_RLS): pos -= 1                                                   ## When the pos is at the last index, decrease by 1 to compensate with zero-indexing
        FDS = len(self.CBX_RLS)-1                                                               ## Retrieve number of fields present in the layout
        self.FIELDS = FDS+1

        ## Fills the placeholder with objects (for Role, Name, Add, and Clear/Remove button) 
        self.CBX_RLS[pos] = QtWidgets.QComboBox(UIA.WGT_CENTRAL)
        self.CBX_RLS[pos].setObjectName(f"CBX_RLS")
        self.CBX_RLS[pos].view().window().setWindowFlags(Qt.Popup | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.CBX_RLS[pos].view().window().setAttribute(Qt.WA_TranslucentBackground)
        self.CBX_RLS[pos].setMinimumWidth(140)
        self.CBX_RLS[pos].setMaximumWidth(200)
        self.CBX_RLS[pos].setEditable(True)
        self.CBX_RLS[pos].addItems(RLS)
        self.CBX_RLS[pos].setCurrentText('')

        self.CBX_NMS[pos] = QtWidgets.QComboBox(UIA.WGT_CENTRAL)
        self.CBX_NMS[pos].setObjectName(f"CBX_NMS")
        self.CBX_NMS[pos].view().window().setWindowFlags(Qt.Popup | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.CBX_NMS[pos].view().window().setAttribute(Qt.WA_TranslucentBackground)
        self.CBX_NMS[pos].setMinimumWidth(140)
        self.CBX_NMS[pos].setMaximumWidth(200)
        self.CBX_NMS[pos].setEditable(True)
        self.CBX_NMS[pos].addItems(NMS)
        self.CBX_NMS[pos].setCurrentText('')

        self.BTN_ATVS[pos] = QtWidgets.QPushButton(UIA.WGT_CENTRAL)
        self.BTN_ATVS[pos].setObjectName(f"BTN_ATVS")
        self.BTN_ATVS[pos].setMaximumWidth(25)
        self.BTN_ATVS[pos].setMinimumHeight(30)
        self.BTN_ATVS[pos].setToolTip("Set as active text")

        self.BTN_INSS[pos] = QtWidgets.QPushButton(UIA.WGT_CENTRAL)
        self.BTN_INSS[pos].setObjectName(f"BTN_INSS")
        self.BTN_INSS[pos].setMaximumWidth(26)
        self.BTN_INSS[pos].setMinimumHeight(30)
        self.BTN_INSS[pos].setToolTip("Add new field below")

        self.BTN_REMS[pos] = QtWidgets.QPushButton(UIA.WGT_CENTRAL)
        self.BTN_REMS[pos].setObjectName(f"BTN_REMS")
        self.BTN_REMS[pos].setMaximumWidth(25)
        self.BTN_REMS[pos].setMinimumHeight(30)
        self.BTN_REMS[pos].setToolTip("Remove this field")

        ## Finally adds those widgets into vertical layouts
        UIA.LYT_ROLES.insertWidget(pos, self.CBX_RLS[pos])
        UIA.LYT_NAMES.insertWidget(pos, self.CBX_NMS[pos])
        UIA.LYT_ACTIV.insertWidget(pos, self.BTN_ATVS[pos])
        UIA.LYT_CLEAR.insertWidget(pos, self.BTN_REMS[pos])
        UIA.LYT_INSRT.insertWidget(pos, self.BTN_INSS[pos])

        self.refreshStates()
    

    def addFields(self, fields:int=2):
        """
        Allows to add more field(s) using 2nd argument and redirects to `insertField` method
        """
        for i in range(fields): self.insertField(i)


    def removeField(self):
        """
        Removes a specific field by unlinking the widget from
        the layout and deleting its widgets from memory.
        """
        for i, btn in enumerate(self.BTN_REMS):
            if btn.hasFocus():
                UIA.LYT_ROLES.removeWidget(self.CBX_RLS[i])
                UIA.LYT_NAMES.removeWidget(self.CBX_NMS[i])
                UIA.LYT_INSRT.removeWidget(self.BTN_ATVS[i])
                UIA.LYT_INSRT.removeWidget(self.BTN_INSS[i])
                UIA.LYT_CLEAR.removeWidget(self.BTN_REMS[i])

                del self.CBX_RLS[i]
                del self.CBX_NMS[i]
                del self.BTN_REMS[i]
                del self.BTN_INSS[i]
                del self.BTN_ATVS[i]
                self.FIELDS -= 1
                self.refreshStates()
                gc.collect()

                ## Reset window to shortest possible to remove spaces left by the field.
                UIA.resize(UIA.size().width(), 0)   
                break
    

    def getFieldData(self):
        """
        Returns 2 lists, and a dictionary of all fields' data
        Used for when exporting to a file (Powerpoint, plain text)
        """
        RLS = [role.currentText() for role in self.CBX_RLS]
        NMS = [name.currentText() for name in self.CBX_NMS]
        DCT = {i:[k,v] for i, (k,v) in enumerate(zip(RLS, NMS))}
        
        return RLS, NMS, DCT


    def updateItems(self):
        """
        Syncs the entries for every field.

        This method also removes blank string retrieved from a pool.
        """
        m = time.time()
        R_ITEMS, N_ITEMS = [], []

        ## Sets the index to its exact position
        for i,(r,n) in enumerate(zip(self.CBX_RLS, self.CBX_NMS)):
            r.setCurrentIndex(i)
            n.setCurrentIndex(i)

            if r.currentText() == '': r.setCurrentIndex(-1)
            if n.currentText() == '': n.setCurrentIndex(-1)

        ## Removes blank item in Combo Box by looping through several times
        for i in range(3):
            for r in self.CBX_RLS:
                for i in range(r.count()):
                    R_ITEMS.append(r.itemData(i,0))
                    if r.itemData(i,0)== '':
                        r.removeItem(i)

            for r in self.CBX_NMS:
                for i in range(r.count()):
                    N_ITEMS.append(n.itemData(i,0))
                    if r.itemData(i,0)== '':
                        r.removeItem(i)

        ## Under Construction
        ## 
        ## - Create a algorithm that will filter out all duplicate items in every combo boxes
        ##

        # for r in self.CBX_RLS:
        #     for i in range(r.count()):
        # for r in self.CBX_RLS:
        #     r.clear()
        #     r.addItems(list(set(R_ITEMS)))
        # p()
        # R_FILTERED = list(set(R_ITEMS))
        # # p(R_ITEMS)
        # p(list(set(R_ITEMS) - set(R_FILTERED)))
        # # p([i for i in R_ITEMS if i not in list(set(R_ITEMS))])
        # p(list(set(N_ITEMS)))
        # p(time.time()-m)




class Export(object):
    """
    Handles exporting related functions.
    """
    def __init__(self):
        pass

    
    def fromActiveField(self, i, clear=False):
        """
        Exports file from an active field
        """
        if clear: R, N = '', ''
        else: R, N = FLD.CBX_RLS[i].currentText(), FLD.CBX_NMS[i].currentText() 

        try:
            with open(f"{SYS.DIR_PROGRAM}/Role.txt", "w") as f: f.write(R)
            with open(f"{SYS.DIR_PROGRAM}/Name.txt", "w") as f: f.write(N)
            # if not clear: LOG.info(f"Active field exported: {R} - {N}")
        except Exception as e:
            LOG.error(e)
            


    def toPlainText(self):
        """
        Export current values into a plain text
        """
        RLS, NMS, DCT = FLD.getFieldData()
        
        DIR_TGT = QtWidgets.QFileDialog.getExistingDirectory(None, "Select Folder", PKG.DIR_EXPORT_RECENT)
        if DIR_TGT:
            LOG.info(f"Saving file to {DIR_TGT}")
            try:
                with open(f"{DIR_TGT}/Roles.txt", "w") as f: f.write('\n'.join(RLS))
                with open(f"{DIR_TGT}/Names.txt", "w") as f: f.write('\n'.join(NMS))
                with open(f"{DIR_TGT}/Participants.txt", "w") as f: f.write('\n'.join([f"{r}: {n}" for r,n in zip(RLS,NMS)]))
            except Exception as e:
                LOG.error(e)
            
            ## Save to JSON
            DCFG['POOL'].update({"ROLES": RLS})
            DCFG['POOL'].update({"NAMES": NMS})
            DCFG['CONFIG'].update({"DIR_EXPORT_RECENT": DIR_TGT})
            PKG.DIR_EXPORT_RECENT = DIR_TGT
            PDB.dump()
            LOG.info("File successfully saved.")
            ## 
            ## Add code here that would update the combo boxes connected to FDS.updateItems()
            ## 


    def toPowerpoint(self):
        """
        Export data to a Powerpoint file
        """
        def singleColumn(self):
            RLS, NMS, DCT = FLD.getFieldData()
            self.PRS = Presentation()

            ## 16:9 Ratio
            self.PRS.slide_width, self.PRS.slide_height = Inches(16), Inches(9)

            SLD_MAIN = self.PRS.slides.add_slide(self.PRS.slide_layouts[6])
            SLD_MAIN.shapes.add_picture(PKG.IMG_BACKGROUND, 0, 0, self.PRS.slide_width, self.PRS.slide_height) # Left-Top-Width-Height

            ## Textboxes
            ## Title
            TBX_TITLE = SLD_MAIN.shapes.add_textbox(0, Cm(1), self.PRS.slide_width, Cm(4))
            FRM_TITLE = TBX_TITLE.text_frame

            PARA_TITLE = FRM_TITLE.paragraphs[0]
            PARA_TITLE.text = PKG.TXT_TITLE
            PARA_TITLE.font.size = PKG.FSZ_TITLE
            PARA_TITLE.font.color.rgb = PKG.RGB_TITLE
            PARA_TITLE.alignment = PP_ALIGN.CENTER
            PARA_TITLE.font.name = PKG.FONT_TITLE
            PARA_TITLE.font.italic = True

            ## Subtitle
            TBX_SUBTITLE = SLD_MAIN.shapes.add_textbox(0, Cm(3), self.PRS.slide_width, Cm(4))
            FRM_SUBTITLE = TBX_SUBTITLE.text_frame
            PARA_SUBTITLE = FRM_SUBTITLE.paragraphs[0]
            PARA_SUBTITLE.text = PKG.TXT_SUBTITLE
            PARA_SUBTITLE.font.size = PKG.FSZ_SUBTITLE
            PARA_SUBTITLE.font.color.rgb = PKG.RGB_SUBTITLE
            PARA_SUBTITLE.alignment = PP_ALIGN.CENTER
            PARA_SUBTITLE.font.name = PKG.FONT_SUBTITLE

            ## Date
            TBX_DATE = SLD_MAIN.shapes.add_textbox(0, Cm(4.2), self.PRS.slide_width, Cm(4))
            FRM_DATE = TBX_DATE.text_frame
            PARA_DATE = FRM_DATE.paragraphs[0]
            PARA_DATE.text = f"——— {datetime.datetime.now().strftime('%B %d, %Y')} ———"
            PARA_DATE.font.size = PKG.FSZ_DATE
            PARA_DATE.font.color.rgb = PKG.RGB_DATE
            PARA_DATE.alignment = PP_ALIGN.CENTER
            PARA_DATE.font.name = PKG.FONT_DATE


            ## Paragraph
            left, top, width, height = 0, Cm(5.4), int(self.PRS.slide_width / 2), int(self.PRS.slide_height) - Cm(2)
            PARA_FNTSZ = Pt(39-FLD.FIELDS)

            ## Roles
            TBX_ROLES = SLD_MAIN.shapes.add_textbox(left, top, width, height)
            FRM_ROLES = TBX_ROLES.text_frame

            PARA_RLS = FRM_ROLES.paragraphs[0]
            PARA_RLS.text = '\n'.join([f"{r}:" for r in RLS])
            PARA_RLS.font.size = PARA_FNTSZ
            PARA_RLS.font.color.rgb = PKG.RGB_ROLES
            PARA_RLS.alignment = PP_ALIGN.RIGHT
            PARA_RLS.font.name = PKG.FONT_CONTENT
            PARA_RLS.font.italic = True

            ## Names
            TBX_NAMES = SLD_MAIN.shapes.add_textbox(int(self.PRS.slide_width / 2), top, width, height)
            FRM_NAMES = TBX_NAMES.text_frame

            PARA_NMS = FRM_NAMES.paragraphs[0]
            PARA_NMS.text = '\n'.join(NMS)
            PARA_NMS.font.size = PARA_FNTSZ
            PARA_NMS.font.color.rgb = PKG.RGB_NAMES
            PARA_NMS.alignment = PP_ALIGN.LEFT
            PARA_NMS.font.name = PKG.FONT_CONTENT


        def doubleColumns(self):
            RLS, NMS, DCT = FLD.getFieldData()
            self.PRS = Presentation()

            ## 16:9 Ratio
            self.PRS.slide_width, self.PRS.slide_height = Inches(16), Inches(9)

            SLD_MAIN = self.PRS.slides.add_slide(self.PRS.slide_layouts[6])
            SLD_MAIN.shapes.add_picture(PKG.IMG_BACKGROUND, 0, 0, self.PRS.slide_width, self.PRS.slide_height) # Left-Top-Width-Height

            ## Textboxes
            ## Title
            TBX_TITLE = SLD_MAIN.shapes.add_textbox(0, Cm(1), self.PRS.slide_width, Cm(4))
            FRM_TITLE = TBX_TITLE.text_frame

            PARA_TITLE = FRM_TITLE.paragraphs[0]
            PARA_TITLE.text = PKG.TXT_TITLE
            PARA_TITLE.font.size = PKG.FSZ_TITLE
            PARA_TITLE.font.color.rgb = PKG.RGB_TITLE
            PARA_TITLE.alignment = PP_ALIGN.CENTER
            PARA_TITLE.font.name = PKG.FONT_TITLE
            PARA_TITLE.font.italic = True

            ## Subtitle
            TBX_SUBTITLE = SLD_MAIN.shapes.add_textbox(0, Cm(3), self.PRS.slide_width, Cm(4))
            FRM_SUBTITLE = TBX_SUBTITLE.text_frame
            PARA_SUBTITLE = FRM_SUBTITLE.paragraphs[0]
            PARA_SUBTITLE.text = PKG.TXT_SUBTITLE
            PARA_SUBTITLE.font.size = PKG.FSZ_SUBTITLE
            PARA_SUBTITLE.font.color.rgb = PKG.RGB_SUBTITLE
            PARA_SUBTITLE.alignment = PP_ALIGN.CENTER
            PARA_SUBTITLE.font.name = PKG.FONT_SUBTITLE

            ## Sabbath School Text
            TBX_SUBTITLE = SLD_MAIN.shapes.add_textbox(Inches(4), Inches(3), self.PRS.slide_width, Cm(4))
            FRM_SUBTITLE = TBX_SUBTITLE.text_frame
            PARA_SUBTITLE = FRM_SUBTITLE.paragraphs[0]
            PARA_SUBTITLE.text = "Sabbath School"
            PARA_SUBTITLE.font.size = Pt(40)
            PARA_SUBTITLE.font.color.rgb = PKG.RGB_SUBTITLE
            PARA_SUBTITLE.alignment = PP_ALIGN.CENTER
            PARA_SUBTITLE.font.name = PKG.FONT_SUBTITLE

            ## Divine Service Text
            TBX_SUBTITLE = SLD_MAIN.shapes.add_textbox(Inches(1.25), Inches(6.5), Inches(4), Inches(2))
            FRM_SUBTITLE = TBX_SUBTITLE.text_frame
            PARA_SUBTITLE = FRM_SUBTITLE.paragraphs[0]
            PARA_SUBTITLE.text = "Divine Service"
            PARA_SUBTITLE.font.size = Pt(40)
            PARA_SUBTITLE.font.color.rgb = PKG.RGB_SUBTITLE
            PARA_SUBTITLE.alignment = PP_ALIGN.CENTER
            PARA_SUBTITLE.font.name = PKG.FONT_SUBTITLE


            ## Date
            TBX_DATE = SLD_MAIN.shapes.add_textbox(0, Cm(4.2), self.PRS.slide_width, Cm(4))
            FRM_DATE = TBX_DATE.text_frame
            PARA_DATE = FRM_DATE.paragraphs[0]
            PARA_DATE.text = f"——— {datetime.datetime.now().strftime('%B %d, %Y')} ———"
            PARA_DATE.font.size = PKG.FSZ_DATE
            PARA_DATE.font.color.rgb = PKG.RGB_DATE
            PARA_DATE.alignment = PP_ALIGN.CENTER
            PARA_DATE.font.name = PKG.FONT_DATE


            ## Paragraph
            left, top, width, height = 0, Cm(5.4), int(self.PRS.slide_width/2), int(self.PRS.slide_height) - Cm(2)
            PARA_FNTSZ = Pt(38-FLD.FIELDS)

            ## Sabbath School
            ## Roles
            TBX_ROLES = SLD_MAIN.shapes.add_textbox(0, top, width/2.2, height)
            TBX_NAMES = SLD_MAIN.shapes.add_textbox(width/2.2, top, width/2.2, height)
            FRM_ROLES = TBX_ROLES.text_frame
            FRM_NAMES = TBX_NAMES.text_frame
            TXT_ROLES = '\n'.join([f"{r}:" for r in RLS])
            TXT_NAMES = '\n'.join(NMS)


            ## Service-based Splitter
            SPLITTED = CORE.splitContents(TXT_ROLES, TXT_NAMES)

            PARA_RLS = FRM_ROLES.paragraphs[0]
            PARA_RLS.text = '\n'.join(SPLITTED[0])
            PARA_RLS.font.size = PARA_FNTSZ
            PARA_RLS.font.color.rgb = PKG.RGB_ROLES
            PARA_RLS.alignment = PP_ALIGN.RIGHT
            PARA_RLS.font.name = PKG.FONT_CONTENT
            PARA_RLS.font.italic = True

            ## Names
            PARA_NMS = FRM_NAMES.paragraphs[0]
            PARA_NMS.text = '\n'.join(SPLITTED[1])
            PARA_NMS.font.size = PARA_FNTSZ
            PARA_NMS.font.color.rgb = PKG.RGB_NAMES
            PARA_NMS.alignment = PP_ALIGN.LEFT
            PARA_NMS.font.name = PKG.FONT_CONTENT
            

            ## Divine Worship
            ## Roles
            TBX_ROLES = SLD_MAIN.shapes.add_textbox(Inches(7.25), Inches(5.5), width/2.2, height)
            TBX_NAMES = SLD_MAIN.shapes.add_textbox(Inches(10.9), Inches(5.5), width/2.2, height)
            FRM_ROLES = TBX_ROLES.text_frame
            FRM_NAMES = TBX_NAMES.text_frame

            PARA_RLS = FRM_ROLES.paragraphs[0]
            PARA_RLS.text = '\n'.join(SPLITTED[2])
            PARA_RLS.font.size = PARA_FNTSZ
            PARA_RLS.font.color.rgb = PKG.RGB_ROLES
            PARA_RLS.alignment = PP_ALIGN.RIGHT
            PARA_RLS.font.name = PKG.FONT_CONTENT
            PARA_RLS.font.italic = True

            ## Names
            PARA_NMS = FRM_NAMES.paragraphs[0]
            PARA_NMS.text = '\n'.join(SPLITTED[3])
            PARA_NMS.font.size = PARA_FNTSZ
            PARA_NMS.font.color.rgb = PKG.RGB_NAMES
            PARA_NMS.alignment = PP_ALIGN.LEFT
            PARA_NMS.font.name = PKG.FONT_CONTENT
        
        START = time.time()
        LOG.info("Generating Powerpoint")
        doubleColumns(self) if FLD.FIELDS > FLD.FIELDS_MAX/1.15 else singleColumn(self)

        try:
            self.PRS.save(SYS.FILE_PPT_EXPORTED)
            os.startfile(SYS.FILE_PPT_EXPORTED)
        except PermissionError as e:
            LOG.warn("PermissionError: The file is still open or is already running. Close the file first and try again.")
            MSG_BOX = QtWidgets.QMessageBox()
            MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Ok)
            MSG_BOX.setIcon(QtWidgets.QMessageBox.Warning)
            MSG_BOX.setText("The file is still open or is already running.\nIf you have changes, close the file first and try again.")
            MSG_BOX.setWindowTitle(SW.NAME)
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            MSG_BOX.setStyleSheet(QSS.getStylesheet())
            MSG_BOX.setStyleSheet('min-width: 280px; min-height: 35px;')
            MSG_BOX.exec_()
        except Exception as e:
            LOG.error(f"{e}")
        else:
            RLS, NMS, DCT = FLD.getFieldData()
            DCFG['POOL'].update({"ROLES": RLS})
            DCFG['POOL'].update({"NAMES": NMS})
            PDB.dump()
            LOG.info(f"File successfully saved. ({round((time.time()-START)*1000)} ms)")
            UIB.hide()




class QWGT_PARTICIPANTS(QtWidgets.QMainWindow):
    """
    Main user interface window for this software.
    Uses PyQt5 for handling GUI.

    Alias UIA (User Interface A / Main)
    """
    def __init__(self, parent = None):
        QtWidgets.QMainWindow.__init__(self, parent)
    

    def setupUI(self):
        """
        Initializes all widgets for UIA
        """
        ## Window
        self.setObjectName("WIN_PARTICIPANTS")
        self.resize(0, 0)
        self.setMaximumWidth(500)
        self.setWindowIcon(QtGui.QIcon(SYS.RES_LOGO))

        ## Layouts
        self.WGT_CENTRAL = QtWidgets.QWidget(self); self.WGT_CENTRAL.setObjectName("WGT_CENTRAL")
        self.LYT_MAIN = QtWidgets.QGridLayout(self.WGT_CENTRAL); self.LYT_MAIN.setObjectName("LYT_MAIN")
        self.LYT_HEAD = QtWidgets.QGridLayout(); self.LYT_HEAD.setObjectName("LYT_HEAD")
        self.LYT_BODY = QtWidgets.QGridLayout(); self.LYT_BODY.setObjectName("LYT_BODY")
        self.LYT_FOOTER = QtWidgets.QGridLayout(); self.LYT_FOOTER.setObjectName("LYT_FOOTER")

        self.LYT_ROLES = QtWidgets.QVBoxLayout(); self.LYT_BODY.setObjectName("LYT_ROLES")
        self.LYT_NAMES = QtWidgets.QVBoxLayout(); self.LYT_BODY.setObjectName("LYT_NAMES")
        self.LYT_CLEAR = QtWidgets.QVBoxLayout(); self.LYT_BODY.setObjectName("LYT_CLEAR")
        self.LYT_INSRT = QtWidgets.QVBoxLayout(); self.LYT_BODY.setObjectName("LYT_INSRT")
        self.LYT_ACTIV = QtWidgets.QVBoxLayout(); self.LYT_BODY.setObjectName("LYT_ACTIV")
        
        ## Banner
        self.PIX_HEADER = QtWidgets.QLabel(self.WGT_CENTRAL); self.PIX_HEADER.setObjectName("PIX_HEADER"); self.PIX_HEADER.setAlignment(QtCore.Qt.AlignCenter)        
        self.PIX_HEADER.setPixmap(QtGui.QPixmap(SYS.RES_HEADERLOGO).scaledToHeight(75, QtCore.Qt.KeepAspectRatio & Qt.SmoothTransformation))

        ## Head
        SPC_WINV_TOPP = QtWidgets.QSpacerItem(20, 15, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        SPC_TITLEV2 = QtWidgets.QSpacerItem(20, 15, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.BTN_SETTINGS = QtWidgets.QPushButton(self.WGT_CENTRAL); self.BTN_SETTINGS.setObjectName("BTN_SETTINGS")
        self.BTN_SAVELIST = QtWidgets.QPushButton(self.WGT_CENTRAL); self.BTN_SAVELIST.setObjectName("BTN_SAVELIST")
        SPC_HEADH1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.LBL_EXPORT = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_EXPORT.setObjectName("LBL_EXPORT")
        self.BTN_POWERPOINT = QtWidgets.QPushButton(self.WGT_CENTRAL); self.BTN_POWERPOINT.setObjectName("BTN_POWERPOINT")
        self.BTN_PLAINTEXT = QtWidgets.QPushButton(self.WGT_CENTRAL); self.BTN_PLAINTEXT.setObjectName("BTN_PLAINTEXT")
        
        ## Body
        SPC_BODYV1 = QtWidgets.QSpacerItem(20, 15, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.LBL_RL = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_RL.setObjectName("LBL_RL"); self.LBL_RL.setAlignment(QtCore.Qt.AlignCenter)
        SPC_BODYH = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.LBL_NM = QtWidgets.QLabel(self.WGT_CENTRAL); self.LBL_NM.setObjectName("LBL_NM"); self.LBL_NM.setAlignment(QtCore.Qt.AlignCenter)
        SPC_BODYV2 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        
        ## Window Spacers
        SPC_WINH_LEFT = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        SPC_WINV_BOTM = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        SPC_WINH_RGHT = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)

        ## Layering
        ## Head
        self.LYT_HEAD.addWidget(self.PIX_HEADER, 1, 0, 1, 7)
        self.LYT_HEAD.addItem(SPC_TITLEV2, 2, 0, 1, 3)
        self.LYT_HEAD.addWidget(self.BTN_SETTINGS, 3, 0, 1, 1)
        self.LYT_HEAD.addWidget(self.BTN_SAVELIST, 3, 1, 1, 1)
        self.LYT_HEAD.addItem(SPC_HEADH1, 3, 3, 1, 1)
        self.LYT_HEAD.addWidget(self.LBL_EXPORT, 3, 4, 1, 1)
        self.LYT_HEAD.addWidget(self.BTN_POWERPOINT, 3, 5, 1, 1)
        self.LYT_HEAD.addWidget(self.BTN_PLAINTEXT, 3, 6, 1, 1)
        self.LYT_HEAD.addItem(SPC_BODYV1, 4, 0, 1, 3)
        self.LYT_HEAD.addLayout(self.LYT_BODY, 5, 0, 1, 7)
        self.LYT_HEAD.addItem(SPC_BODYV2, 6, 0, 1, 3)

        ## Body
        self.LYT_BODY.addWidget(self.LBL_RL, 0, 0, 1, 1)
        self.LYT_BODY.addLayout(self.LYT_ROLES, 1, 0, 1, 1)
        self.LYT_BODY.addItem(SPC_BODYH, 0, 1, 1, 1)
        self.LYT_BODY.addWidget(self.LBL_NM, 0, 2, 1, 1)
        self.LYT_BODY.addLayout(self.LYT_NAMES, 1, 2, 1, 1)
        self.LYT_BODY.addLayout(self.LYT_ACTIV, 1, 3, 1, 1)
        self.LYT_BODY.addLayout(self.LYT_INSRT, 1, 4, 1, 1)
        self.LYT_BODY.addLayout(self.LYT_CLEAR, 1, 5, 1, 1)

        ## Window
        self.LYT_MAIN.addLayout(self.LYT_HEAD, 0, 1, 1, 1)
        self.LYT_HEAD.addItem(SPC_WINV_TOPP, 0, 0, 1, 3)
        self.LYT_MAIN.addItem(SPC_WINH_LEFT, 0, 0, 1, 1)
        self.LYT_MAIN.addItem(SPC_WINV_BOTM, 1, 1, 1, 1)
        self.LYT_MAIN.addItem(SPC_WINH_RGHT, 0, 2, 1, 1)
        self.setCentralWidget(self.WGT_CENTRAL)

        ## Stylesheets
        QSS.setupFonts()

        ## Initialization
        self.setupDisplay()
        FLD.setup()
        self.setupConnections()
    

    def setupDisplay(self):
        """
        Handle all UI-based properties including texts
        Identical to `retranslateUi` method.
        """
        ## Main
        self.setWindowTitle("Participants")

        ## Head
        self.PIX_HEADER.setToolTip(f"v{SW.VERSION} {SW.VERSION_NAME}\nMade for Seventh-day Adventist Church\n\n© {SW.PROD_YEAR} {SW.AUTHOR}")
        self.LBL_EXPORT.setText("Export to")

        ## Body
        self.LBL_RL.setText("ROLE")
        self.LBL_NM.setText("NAME")
    

    def setupConnections(self):
        """
        Manages Signals and Slots from user interactions
        """
        self.BTN_PLAINTEXT.clicked.connect(lambda: EXP.toPlainText())
        self.BTN_POWERPOINT.clicked.connect(lambda: EXP.toPowerpoint())
        self.BTN_SETTINGS.clicked.connect(lambda: UIB.enterWindow())
        

    def updateWindowTitle(self):
        """
        Handles UIA's main window title
        """
        self.setWindowTitle(f"{SW.NAME} ({FLD.FIELDS})")
    
    


class QWGT_SETTINGS(QtWidgets.QWidget):
    """
    Settings window

    Alias UIB (User Interface B / Settings)
    """
    def __init__(self, parent = None):
        QtWidgets.QWidget.__init__(self, parent)
        self.CHANGED = False


    def setupUI(self):
        """
        Initializes all widgets for UIB
        """
        ## Window
        self.setObjectName("WIN_SETTINGS")
        self.setFixedSize(QtCore.QSize(380, 320))
        self.setWindowIcon(QtGui.QIcon(SYS.RES_LOGO))


        self.GRID_MAIN = QtWidgets.QGridLayout(self); self.GRID_MAIN.setObjectName("GRID_MAIN")
        self.GRID_BODY = QtWidgets.QGridLayout(); self.GRID_BODY.setObjectName("GRID_BODY")
        self.LYT_HBUTTONS = QtWidgets.QHBoxLayout(); self.LYT_HBUTTONS.setObjectName("LYT_HBUTTONS")

        self.BTN_BROWSE = QtWidgets.QPushButton(self); self.BTN_BROWSE.setObjectName("BTN_BROWSE")
        self.BTN_DISCARD = QtWidgets.QPushButton(self); self.BTN_DISCARD.setObjectName("BTN_DISCARD")
        self.LBL_TITLE = QtWidgets.QLabel(self); self.LBL_TITLE.setObjectName("LBL_TITLE")
        self.LBL_TITLE.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        
        self.GRID_FOOTER = QtWidgets.QGridLayout(); self.GRID_FOOTER.setObjectName("GRID_FOOTER")
        self.BTN_OK = QtWidgets.QPushButton(self); self.BTN_OK.setObjectName("BTN_OK"); self.BTN_OK.setMinimumWidth(70)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        spacerItem1 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.LBL_SELECT = QtWidgets.QLabel(self); self.LBL_SELECT.setObjectName("LBL_SELECT")
        self.LBL_SUBTITLE = QtWidgets.QLabel(self); self.LBL_SUBTITLE.setObjectName("LBL_SUBTITLE"); self.LBL_SUBTITLE.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.LNE_TITLE = QtWidgets.QLineEdit(self); self.LNE_TITLE.setObjectName("LNE_TITLE")
        self.LNE_SUBTITLE = QtWidgets.QLineEdit(self); self.LNE_SUBTITLE.setObjectName("LNE_SUBTITLE")
        self.LBL_PREVIEW = QtWidgets.QLabel(self); self.LBL_PREVIEW.setObjectName("LBL_PREVIEW")
        self.LBL_PREVIEW.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        
        ## Banner
        self.PIX_PREVIEW = QtWidgets.QLabel(self); self.PIX_PREVIEW.setObjectName("PIX_PREVIEW"); self.PIX_PREVIEW.setAlignment(QtCore.Qt.AlignCenter)        
        self.PIX_PREVIEW.setPixmap(QtGui.QPixmap(PKG.IMG_BACKGROUND).scaledToHeight(85, QtCore.Qt.KeepAspectRatio & Qt.SmoothTransformation))

        self.LYT_HBUTTONS.addWidget(self.BTN_BROWSE)
        self.LYT_HBUTTONS.addWidget(self.BTN_DISCARD)
        self.GRID_BODY.addWidget(self.LBL_SELECT, 0, 0, 1, 2)
        self.GRID_BODY.addLayout(self.LYT_HBUTTONS, 0, 3, 1, 1)
        self.GRID_MAIN.addLayout(self.GRID_BODY, 1, 0, 1, 1)
        self.GRID_BODY.addWidget(self.PIX_PREVIEW, 1, 0, 1, 2)
        self.GRID_BODY.addItem(spacerItem1, 1, 2, 1, 1)
        self.GRID_FOOTER.addItem(spacerItem, 2, 0, 1, 1)
        self.GRID_BODY.addWidget(self.LBL_PREVIEW, 2, 0, 1, 2)
        self.GRID_FOOTER.addWidget(self.BTN_OK, 2, 1, 1, 1)
        self.GRID_BODY.addItem(spacerItem2, 3, 0, 2, 2)
        self.GRID_BODY.addWidget(self.LBL_TITLE, 5, 0, 1, 1)
        self.GRID_BODY.addWidget(self.LNE_TITLE, 5, 1, 1, 3)
        self.GRID_BODY.addWidget(self.LBL_SUBTITLE, 6, 0, 1, 1)
        self.GRID_BODY.addWidget(self.LNE_SUBTITLE, 6, 1, 1, 3)
        self.GRID_BODY.addLayout(self.GRID_FOOTER, 7, 0, 1, 5)

        ## Initialization
        self.setupDisplay()
        self.setupConnections()


    def setupDisplay(self):
        """
        Handle all UI-based properties including texts
        Identical to `retranslateUi` method. 
        """
        self.setWindowTitle("Settings")
        self.LBL_TITLE.setText("Title:")
        self.BTN_OK.setText("OK")
        self.LBL_SELECT.setText("Select a background image:")
        self.LBL_SUBTITLE.setText("Subtitle:")
        self.LBL_PREVIEW.setText("Preview")


    def setupConnections(self):
        """
        Manages Signals and Slots from user interactions
        """
        self.BTN_OK.clicked.connect(lambda: self.saveChanges())
        self.BTN_BROWSE.clicked.connect(lambda: UIB.browseForBackgroundImage())
        self.BTN_DISCARD.clicked.connect(lambda: UIB.discardImage())
        self.LNE_TITLE.textChanged.connect(lambda: UIB.updatePackage())
        self.LNE_SUBTITLE.textChanged.connect(lambda: UIB.updatePackage())


    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape:
            self.closeEvent()


    def closeEvent(self, event=None):
        """
        Close event for Settings window.
        Triggers when the user selected "OK" or "x" button.
        """
        if SYS.GLOBAL_STATE == 3: return
        if self.CHANGED:
            MSG_BOX = QtWidgets.QMessageBox()
            MSG_BOX.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            MSG_BOX.setIcon(QtWidgets.QMessageBox.Question)
            MSG_BOX.setText("You have unsaved changes\nDo you want to save your preferences?")
            MSG_BOX.setWindowTitle(SW.NAME)
            MSG_BOX.setWindowFlags(Qt.Drawer | Qt.WindowStaysOnTopHint)
            MSG_BOX.setStyleSheet(QSS.getStylesheet())
            MSG_BOX.setStyleSheet('min-width: 250px; min-height: 40px;')
            
            RET = MSG_BOX.exec_()
            if RET == QtWidgets.QMessageBox.Yes:
                self.saveChanges()
            else:
                PKG.restart()
                self.hide()
        else:
            self.saveChanges()
    

    def saveChanges(self):
        """
        Triggers when user confirms and exits the settings window
        """
        DCFG["CONFIG"].update({"TXT_TITLE": PKG.TXT_TITLE})
        DCFG["CONFIG"].update({"TXT_SUBTITLE": PKG.TXT_SUBTITLE})
        PDB.dump()
        self.hide()


    def enterWindow(self):
        """
        Triggers after UIA's settings button was clicked
        """
        self.ENTERING = True
        self.LNE_TITLE.setText(PKG.TXT_TITLE)
        self.LNE_SUBTITLE.setText(PKG.TXT_SUBTITLE)
        self.show()
        self.raise_()                                   ## Bring to front when clicked again
        self.ENTERING = False
        self.CHANGED = False
    

    def updatePackage(self):
        """
        Updates the current package to what's changed
        """
        if self.ENTERING: return
        self.CHANGED = True                                 ## Indicating that the settings window is modified
        PKG.TXT_TITLE = self.LNE_TITLE.text()
        PKG.TXT_SUBTITLE = self.LNE_SUBTITLE.text()


    def browseForBackgroundImage(self):
        """
        Customizes the background image
        """
        PATH_BGIMG = QtWidgets.QFileDialog.getOpenFileName(None, 'Browse for Hymnal Package', os.getcwd(),
                    f'Images ({" ".join(["*.{}".format(fo.data().decode()) for fo in QtGui.QImageReader.supportedImageFormats()])})'.format())
                    
        if PATH_BGIMG[0] != '':   
            self.PIX_PREVIEW.setPixmap(QtGui.QPixmap(PATH_BGIMG[0]).scaledToHeight(85, QtCore.Qt.KeepAspectRatio & Qt.SmoothTransformation))
            PKG.IMG_BACKGROUND = PATH_BGIMG[0]
            self.updateBackgroundImage(PATH_BGIMG[0])


    def updateBackgroundImage(self, bg):
        """
        Makes copy of the imported bg image to program directory and updating the configuration
        """
        for i in ['png', 'jpg', 'svg', 'bmp', 'tiff', 'jpeg', 'gif']:
            try: os.remove(f"{SYS.DIR_PROGRAM}/BG.{i}")
            except FileNotFoundError: pass

        shutil.copy(bg, SYS.DIR_PROGRAM)
        FILE = bg.split('/')[-1]
        DEST = f"{SYS.DIR_PROGRAM}/BG{os.path.splitext(FILE)[1]}"

        try: os.rename(f"{SYS.DIR_PROGRAM}/{FILE}", DEST)
        except (FileNotFoundError, FileExistsError) as e: LOG.debug(e)
        DCFG["CONFIG"].update({"IMG_BACKGROUND": DEST}); PDB.dump()
    

    def discardImage(self):
        """
        Discards and restores the background image to default
        """
        PKG.IMG_BACKGROUND = PKG.DEF_IMG_BACKGROUND
        self.PIX_PREVIEW.setPixmap(QtGui.QPixmap(PKG.IMG_BACKGROUND).scaledToHeight(85, QtCore.Qt.KeepAspectRatio & Qt.SmoothTransformation))
        self.updateBackgroundImage(PKG.IMG_BACKGROUND)



if __name__ == '__main__':
    INIT_TIME = time.time()
    ## Create Application
    
    APP = QtWidgets.QApplication(sys.argv)

    ## Primary Software Initialization & Logging
    SW = KSoftware("Participants", "1.0.3", "Ken Verdadero, Reynald Ycong", file=__file__, parentName="MSDAC Systems", prodYear=2022, versionName="Alpha")
    LOG = KLog(System().DIR_LOG, __file__, SW.LOG_NAME_DATE(), SW.PY_NAME, SW.AUTHOR, cont=True, tms=True, delete_existing=True, tmsformat="%H:%M:%S.%f %m/%d/%y")

    SYS = System()
    SYS.verifyDirectories()

    PDB = Data()
    global DCFG
    DCFG = PDB.DATA
    RLS = DCFG['POOL']['ROLES']
    NMS = DCFG['POOL']['NAMES']

    QSS = Stylesheet()
    PKG = Package()
    FLD = Fields()
    CORE = Core()
    EXP = Export()
    FMN = FileManager()

    LOG.info('Initializing UIA')
    UIA = QWGT_PARTICIPANTS()
    UIB = QWGT_SETTINGS()
    
    UIA.setupUI()
    UIB.setupUI()

    UIA.closeEvent = lambda event: SYS.closeEvent(event)

    SYS.verifyRequisites()                                                                          ## Check for MS Office PowerPoint availability
    SYS.checkInstances() 

    ## Shows User Interface
    gc.collect()
    UIA.show()
    CORE.centerWindow(UIA)
    UIA.raise_()
    UIA.activateWindow()

    SYS.STARTUP_TIME = time.time()-INIT_TIME
    LOG.info(f'Initialization completed in {round(SYS.STARTUP_TIME, 3)} seconds.')
    sys.exit(APP.exec_())
    


