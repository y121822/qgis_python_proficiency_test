# -*- coding: utf-8 -*-

import os

from qgis.PyQt import uic
from qgis.PyQt import QtWidgets

from PyQt5.QtWidgets import QFileDialog

# This loads your .ui file so that PyQt can populate your plugin with the elements from Qt Designer
FORM_CLASS, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'Excel_Data_Exporter_dialog_base.ui'))


class ExcelDataExporterDialog(QtWidgets.QDialog, FORM_CLASS):
    def __init__(self, parent=None):
        """Constructor."""
        super(ExcelDataExporterDialog, self).__init__(parent)
        # Set up the user interface from Designer through FORM_CLASS.
        # After self.setupUi() you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.path = None  # the file path variable
        self.lineEdit.setText("Output file path, click ...")  # Prompt to enter a file path
        self.pushButton.clicked.connect(self.select_output_file)  # the listener calling an action

    def select_output_file(self):
        """Retrieve entered file path, assign it to the proper variable"""
        filename, _filter = QFileDialog.getSaveFileName(self, "", "", '*.xlsm')
        self.lineEdit.setText(filename)
        self.path = filename
