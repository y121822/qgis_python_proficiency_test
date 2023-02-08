# -*- coding: utf-8 -*-

from qgis.PyQt.QtCore import QSettings, QTranslator, QCoreApplication
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction
from qgis.core import QgsProject, QgsMessageLog

# Initialize Qt resources from file resources.py
from .resources import *
# Import the code for the dialog
from .Excel_Data_Exporter_dialog import ExcelDataExporterDialog
import os
import shutil
import re


class ExcelDataExporter:
    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'ExcelDataExporter_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)
            QCoreApplication.installTranslator(self.translator)

        self.actions = []
        self.menu = self.tr(u'&Excel Data Exporter')
        self.first_start = None

    def tr(self, message):
        return QCoreApplication.translate('ExcelDataExporter', message)

    def add_action(
            self,
            icon_path,
            text,
            callback,
            enabled_flag=True,
            add_to_menu=True,
            add_to_toolbar=True,
            status_tip=None,
            whats_this=None,
            parent=None):

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.iface.addToolBarIcon(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        icon_path = ':/plugins/Excel_Data_Exporter/icon.png'
        self.add_action(
            icon_path,
            text=self.tr(u'Excel Data Exporter'),
            callback=self.run,
            parent=self.iface.mainWindow())

        self.first_start = True

    def unload(self):
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'&Excel Data Exporter'),
                action)
            self.iface.removeToolBarIcon(action)

    def run(self):
        if self.first_start:
            self.first_start = False
            self.dlg = ExcelDataExporterDialog()

        self.dlg.show()

        result = self.dlg.exec_()

        if result:
            MyClass(self.dlg.path)  # call the processing class with a file path


class MyClass:
    """Export features and their geolocation from Assessment Splice,
       Assessment Cables, Assessment Strand, Assessment Terminals layers
       within FSA101 boundaries into an Excel file using a template, adding
       new sheets for each layer and saving the file in a user defined location"""

    def __init__(self, path):
        """Initialise required variables, call methods implementing validation
           and processing logic, inform a user about the stage of the process.
           """
        self.path = path
        if self.path:
            self.origin_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Template.xlsm')
            if self.path_check():
                if self.packages_check():
                    self.names, self.fsa101 = ['Splice', 'Cables', 'Strand', 'Terminals'], None
                    self.set_fsa101()
                    if self.fsa101:
                        self.layers_list = [self.get_layer(name) for name in self.names]
                        shutil.copy(self.origin_path, self.path)
                        self.process()
                        QgsMessageLog.logMessage('Process successfully completed!')
                    else:
                        QgsMessageLog.logMessage('Layers are not in a group, no data or wrong data')
            else:
                QgsMessageLog.logMessage('File path is not valid')
        else:
            QgsMessageLog.logMessage('File path is empty')

    def path_check(self):
        """Validate file path. Modified with the more efficient way
           to remove slashes from a file path"""
        filename = os.path.basename(self.path).split('.')
        pattern = re.compile(r'[/\\]')
        new_path = re.sub(pattern, '', self.path)
        old_path = re.sub(pattern, '', self.origin_path)

        if len(filename) == 2 and new_path != old_path:
            name, extension = filename
            if extension == 'xlsm' and re.match("^[A-Za-z0-9_-]*$", name) and len(name) <= 50:
                return 1

    def packages_check(self):
        """Check if required packages have been installed, otherwise,
           following the best practice not to install a software without
           the user`s consent, give the recommendation how to do it"""
        result = 2
        for module in ['pandas', 'openpyxl']:
            try:
                __import__(module)
            except ModuleNotFoundError:
                result -= 1
                QgsMessageLog.logMessage(f'Required package "{module}" not found. In Python console:\n'
                                         f'import pip\n'
                                         f'pip.main(["install", "{module}"])')

        return 1 if result == 2 else 0

    def get_layer(self, name):
        """Get a vector layer by its name"""
        try:
            return QgsProject.instance().mapLayersByName(f'Assessment {name}')[0]
        except IndexError:
            pass

    def set_fsa101(self):
        """Create the general FSA 101 boundary geometry from different
           FSA 101x features"""
        fsa = self.get_layer('FSA')
        if fsa:
            fsa.select([1, 3, 4])
            for f in fsa.selectedFeatures():
                if self.fsa101:
                    self.fsa101 = self.fsa101.combine(f.geometry())
                else:
                    self.fsa101 = f.geometry()
            fsa.removeSelection()

    def process(self):
        """Retrieve all features from the target layers, check if they are
           within the FSA 101 boundary, add geolocation. Create a Pandas
           dataframe for each layer and add it as a new sheet to a resulting
           Excel file"""

        import openpyxl
        import pandas

        with pandas.ExcelWriter(self.path, engine='openpyxl') as writer:
            writer.book = openpyxl.load_workbook(self.path)
            for i in range(len(self.names)):
                cols, data = self.layers_list[i].fields().names(), []
                cols.append('geometry')
                for f in self.layers_list[i].getFeatures():
                    geom = f.geometry()
                    if geom.within(self.fsa101):
                        values = f.attributes()
                        values.append(geom.asWkt())
                        data.append(values)
                df = pandas.DataFrame.from_records(data=data, columns=cols)
                df.to_excel(writer, sheet_name=self.names[i])
