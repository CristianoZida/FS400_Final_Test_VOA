from GUI.Debug_GUI import Ui_Form
from PyQt5.QtWidgets import *


class DebugGui(QMainWindow, Ui_Form):
    """
    This GUI is a simple interface to config the test parameters and manual power down the control board of FS400
    """

    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        #self.pushButton.clicked.connect(self.save_quit)

    # def save_quit(self):
    #     # pass the settings to main win
    #     pass
