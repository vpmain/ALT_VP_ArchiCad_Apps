import sys
import csv

from PyQt5.QtGui import *
from babel.numbers import format_number, format_decimal, format_percent
from PyQt5.QtCore import Qt
from archicad import ACConnection
from openpyxl import Workbook, load_workbook
from qtpy import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import *
# from PyQt5.QtGui import *
import PDFBase
from ACApps.mainwindow import Ui_MainWindow
from ACApps.SplashScreen import Ui_SplashScreen

from PDFBase import PDF

conn = ACConnection.connect()
# assert conn

acc = conn.commands
act = conn.types
acu = conn.utilities

programmname = "AC25 - V+P-APPS"
programmversion = "1.0"
path_templates_excel = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\Template_Excel"
path_templates_pdf = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\Template_PDF"
company_icon = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\icon\favicon.ico"
splash_icon = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\icon\Klammer.png"
intern_properties = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\intern_properties.csv"
path_layer = r"\\firma.local\dfs\Firmenstandard\VP_ACApps\layer.csv"


def string_to_float(value: str, rounded: int):
    new_value = round(float(value.replace(".", "").replace(",", ".")), rounded)
    return new_value

def float_to_string(value: str):
    new_value = format_decimal(value, format='#,##0.00', locale='de_DE')
    return new_value

# ==> GLOBALS
counter = 1

class SplashScreen(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.ui = Ui_SplashScreen()
        self.ui.setupUi(self)
        self.ui.splashTitlePic.setPixmap(QPixmap(splash_icon))

        ## REMOVE TITLE BAR
        self.setWindowFlags(Qt.WindowStaysOnTopHint | QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        ## DB VALUES


        ## Projektanzahl
        # self.qry1 = """
        #     SELECT Count(DISTINCT GZ)
        #     FROM ProjekteView
        # """
        # self.query = QtSql.QSqlQuery()
        # self.query.exec(self.qry1)
        # self.query.first()
        # self.string = '{:,}'.format(self.query.value(0)).replace(",", ".")
        self.string = ""

        ############################################################################

        ## Positonsanzahl
        # self.qry2 = """
        #     SELECT Count(*)
        #     FROM GewerkeViewOhneBieter
        # """
        # self.query.exec(self.qry2)
        # self.query.first()
        # self.string2 = '{:,}'.format(self.query.value(0)).replace(",", ".")
        self.string2 = ""

        ############################################################################

        ## Bieterpreisanzahl
        # self.qry3 = """
        #     SELECT Count(GZ)
        #     FROM GewerkeView
        # """
        # self.query.exec(self.qry3)
        # self.query.first()
        # self.string3 = '{:,}'.format(self.query.value(0)).replace(",", ".")
        self.string3 = ""

        ############################################################################

        ## QTIMER ==> Start
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.progress)

        ## TIMER IN MILLISECONDS
        self.timer.start(35)

    def progress(self):
        global counter

        # SET VALUE TO PROGRESS BAR
        self.ui.splashProgressBar.setValue(counter)

        # COUNTER = 1
        if counter == 1:
            self.ui.splashText.setText("Apps werden geladen...")

        # COUNTER = 25
        if counter == 25:
                self.ui.splashText.setText(self.string + "Apps werden geladen...")

        # COUNTER = 50
        if counter == 50:
                self.ui.splashText.setText(self.string2 + "Apps werden geladen...")

        # COUNTER = 75
        if counter == 75:
                self.ui.splashText.setText(self.string3 + "Apps werden geladen...")

        # CLOSE SPLASHSCREEN AND OPEN APP
        if counter > 100:
            # STOP TIMER
            self.timer.stop()

            # SHOW MAIN WINDOW
            self.main = MainWindow()
            self.main.show()

            # CLOSE SPLASHSCREEN
            self.close()

        # INCREASE COUNTER
        counter += 1

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.setWindowTitle(programmname + " [Version " + programmversion + "]")
        self.setWindowIcon(QtGui.QIcon(company_icon))
        self.setWindowIconText("logo")

        self.product_info()
        self.show_empty()

        self.attr_raumbuch()

        self.user_attr()
        self.intern_attr()
        self.count_elements()

        self.set_controls_pdf()
        self.attr_raumbuch_pdf()

        # self.ui.txtPathTargetPdf.setText(r"C:\Users\edinger\Desktop")
        # self.ui.txtFilenamePdf.setText("Test")

        self.ui.cmdShowRaumbuchExcel.clicked.connect(self.show_Raumbuch)
        self.ui.cmdShowRaumbuchPdf.clicked.connect(self.show_RaumbuchPDF)
        self.ui.cmdShowAttributuebertragung.clicked.connect(self.show_Attributuebertragung)

        self.ui.cmdAttrAdd.clicked.connect(self.attr_selected)
        self.ui.cmdImportAttr.clicked.connect(self.import_attr)
        self.ui.cmdDeleteRow.clicked.connect(self.row_delete)
        self.ui.cmdSaveCSV.clicked.connect(self.attr_save)

        self.ui.cmdExcel.clicked.connect(self.excel_open)
        self.ui.cmdCSV.clicked.connect(self.csv_open)
        self.ui.cmdTarget.clicked.connect(self.output_directory)

        self.ui.cmdExecRaumbuch.clicked.connect(self.execute_roombook)

        self.ui.gew_user_selected.clicked.connect(self.user_attr_selected)
        self.ui.gew_intern_selected.clicked.connect(self.intern_attr_selected)

        self.ui.execute.clicked.connect(self.execute)
        self.ui.category.currentTextChanged.connect(self.count_elements)

        self.ui.cboSpaltenPdf.currentTextChanged.connect(self.cols_change_pdf)
        self.ui.cmdAddRowPdfUp.clicked.connect(self.add_row_pdf_up)
        self.ui.cmdAddRowPdfDown.clicked.connect(self.add_row_pdf_down)
        self.ui.cmdDeleteRowPdf.clicked.connect(self.delete_row_pdf)
        self.ui.cmdAttrSavePdf.clicked.connect(self.attr_save_pdf)
        self.ui.cmdAttrImportPdf.clicked.connect(self.attr_import_pdf)

        self.ui.cmdTargetPdf.clicked.connect(self.output_directory_pdf)
        self.ui.cmdTargetLogo.clicked.connect(self.open_logo_pdf)
        self.ui.cmdTargetTemplatePdf.clicked.connect(self.open_template_pdf)

        self.ui.cmdExecRaumbuchPdf.clicked.connect(self.execute_roombook_pdf)

    def product_info(self):
        version = acc.GetProductInfo()
        ac_version = str(version[0])
        built_version = str(version[1])
        land = str(version[2])

        self.ui.lblVerbindung.setText("online")
        self.ui.lblACVersion.setText(str(ac_version))
        self.ui.lblACBuilt.setText(str(built_version))
        self.ui.lblACLand.setText(str(land))

    # region Subpages
    def show_empty(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.pageEmpty)

    def show_Raumbuch(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.pageRaumbuchgenerator)

    def show_RaumbuchPDF(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.pageRaumbuchgeneratorPdf)

    def show_Attributuebertragung(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.pageAttributuebertragung)
    # endregion

    # region Raumbuch Excel
    def attr_raumbuch(self):
        all_properties = acc.GetAllPropertyNames()
        tree = self.ui.treeUserAttr
        tree.setHeaderLabels(["Name", "Herkunft"])
        tree.setColumnWidth(0, 200)
        tree.setAlternatingRowColors(True)
        categories = []

        for item in all_properties:
            if item.type == "UserDefined":
                if not item.localizedName[0] in categories:
                    categories.append(item.localizedName[0])

        categories.append("Intern")

        for category in categories:
            item_tree = QtWidgets.QTreeWidgetItem([category])

            for item in all_properties:
                if item.type == "UserDefined":
                    if item.localizedName[0] == category:
                        QtWidgets.QTreeWidgetItem(item_tree, [item.localizedName[1], "benutzerdef."])

            with open(intern_properties, newline='', encoding="utf-8") as csvdatei:
                csv_reader_object = csv.reader(csvdatei, delimiter=';')

                zeilennummer = 0
                for row in csv_reader_object:

                    if zeilennummer == 0:
                        pass
                    elif category == "Intern":
                        QtWidgets.QTreeWidgetItem(item_tree, [row[0], "intern"])
                    zeilennummer += 1

            tree.addTopLevelItem(item_tree)

    def import_attr(self):
        self.ui.tblAttr.setRowCount(0)
        path = self.ui.txtPathCSV.text()
        with open(path, newline='', encoding="utf-8") as csvdatei:
            csv_reader_object = csv.reader(csvdatei, delimiter=';')

            zeilennummer = 0
            for row in csv_reader_object:

                if zeilennummer == 0:
                    pass
                else:
                    row_position = self.ui.tblAttr.rowCount()
                    self.ui.tblAttr.insertRow(row_position)

                    self.ui.tblAttr.setItem(row_position, 0, QtWidgets.QTableWidgetItem(row[0]))
                    self.ui.tblAttr.setItem(row_position, 1, QtWidgets.QTableWidgetItem(row[1]))
                    self.ui.tblAttr.setItem(row_position, 2, QtWidgets.QTableWidgetItem(row[2]))
                zeilennummer += 1

        header = self.ui.tblAttr.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("CSV-Template wurde importiert!")
        msg.setWindowTitle("AC24: Attributübertragung")
        msg.setWindowIcon(QtGui.QIcon(company_icon))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def attr_selected(self):
        get_selected = self.ui.treeUserAttr.selectedItems()
        selected = ""
        if get_selected:
            selected = get_selected[0].text(0)
            selected2 = get_selected[0].text(1)


        all_properties = acc.GetAllPropertyNames()
        categories = []

        for item in all_properties:
            if item.type == "UserDefined":
                if not item.localizedName[0] in categories:
                    categories.append(item.localizedName[0])

        categories.append("Intern")

        if selected in categories:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)

            msg.setText("Fehler Auswahl!")
            msg.setInformativeText("Hauptkategorie kann nicht ausgewählt werden!")
            msg.setWindowTitle("AC24: Attributübertragung")
            msg.setWindowIcon(QtGui.QIcon(company_icon))
            msg.setDetailedText(
                "Für die Übertragung von Attributen muss ein benutzerdefiniertes Attribut ausgewählt werden. Sie "
                "haben eine Gruppenbezeichnung ausgewählt!")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            msg.exec_()

        else:
            if selected2 == "intern":
                with open(r"\\firma.local\dfs\APL-Intern\00_ARCHICAD_PYTHON\AC24_APP_Raumbuch\intern_properties.csv", newline='', encoding="utf-8") as csvdatei:
                    csv_reader_object = csv.reader(csvdatei, delimiter=';')

                    zeilennummer = 0
                    for row in csv_reader_object:
                        if zeilennummer == 0:
                            pass
                        else:
                            if row[0] == selected:
                                col1 = row[0] + "/" + row[1]
                                col2 = "intern"
                                if self.ui.txtCell.text() == "":
                                    col3 = "-"
                                else:
                                    col3 = self.ui.txtCell.text()

                        zeilennummer += 1

            else:

                for item in all_properties:
                    if item.type == "UserDefined":
                        if item.localizedName[1] == selected:
                            col1 = item.localizedName[0] + "/" + item.localizedName[1]
                            col2 = "benutzerdef."
                            if self.ui.txtCell.text() == "":
                                col3 = "-"
                            else:
                                col3 = self.ui.txtCell.text()

            row_position = self.ui.tblAttr.rowCount()
            self.ui.tblAttr.insertRow(row_position)

            self.ui.tblAttr.setItem(row_position, 0, QtWidgets.QTableWidgetItem(col1))
            self.ui.tblAttr.setItem(row_position, 1, QtWidgets.QTableWidgetItem(col2))
            self.ui.tblAttr.setItem(row_position, 2, QtWidgets.QTableWidgetItem(col3))

            header = self.ui.tblAttr.horizontalHeader()
            header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
            header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

    def row_delete(self):
        selected = self.ui.tblAttr.currentRow()
        self.ui.tblAttr.removeRow(selected)

    def attr_save(self):
        path = self.ui.txtPathCSV.text()
        with open(path, "w", newline='', encoding="utf-8") as csvdatei:
            csv_writer_object = csv.writer(csvdatei, delimiter=';', quotechar='"')

            headerline = ["Attribut", "Herkunft", "Zelle"]
            csv_writer_object.writerow(headerline)

            rows = self.ui.tblAttr.rowCount()
            for i in range(0, rows):
                rowContent = [
                    self.ui.tblAttr.item(i, 0).text(),
                    self.ui.tblAttr.item(i, 1).text(),
                    self.ui.tblAttr.item(i, 2).text()
                ]
                csv_writer_object.writerow(rowContent)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("CSV-Template überschrieben und gespeichert!")
        msg.setWindowTitle("AC24: Attributübertragung")
        msg.setWindowIcon(QtGui.QIcon(company_icon))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def excel_open(self):
        fileName = QFileDialog.getOpenFileName(self, "Vorlagendatei wählen", path_templates_excel, "Excel-Dateien (*.xlsm)")
        self.ui.txtPathExcel.setText(fileName[0])
        csv_file = fileName[0][:-4]
        self.ui.txtPathCSV.setText(csv_file + "csv")

    def csv_open(self):
        fileName = QFileDialog.getOpenFileName(self, "CSV-Datei wählen", path_templates_excel, "CSV-Dateien (*.csv)")
        self.ui.txtPathCSV.setText(fileName[0])

    def output_directory(self):
        file = str(QFileDialog.getExistingDirectory(self, "Speicherort wählen", r"\\firma.local\dfs\Projekte"))
        self.ui.txtPathTarget.setText(file)

    def execute_roombook(self):
        self.ui.tblAttr.sortItems(1, QtCore.Qt.DescendingOrder)
        path_tamplate = self.ui.txtPathExcel.text()
        path_target = self.ui.txtPathTarget.text() + "/" + self.ui.txtFilename.text() + ".xlsm"
        workbook_base = load_workbook(filename=path_tamplate, read_only=False, keep_vba=True)
        rooms = acc.GetElementsByType("Zone")

        dict_values_intern = {}
        dict_values_user = {}

        rows = self.ui.tblAttr.rowCount()
        for i in range(0, rows):
            herkunft = self.ui.tblAttr.item(i, 1).text()

            if herkunft == "intern":
                text = self.ui.tblAttr.item(i, 0).text()
                text_split = text.split("/")
                value = text_split[1]
                cell = self.ui.tblAttr.item(i, 2).text()

                dict_values_intern.update({cell:act.BuiltInPropertyUserId(value)})

            elif herkunft == "benutzerdef.":
                text = self.ui.tblAttr.item(i, 0).text()
                text_split = text.split("/")
                value_group = text_split[0]
                value_user = text_split[1]
                cell = self.ui.tblAttr.item(i, 2).text()

                dict_values_user.update({cell: act.UserDefinedPropertyUserId([value_group, value_user])})

        for key in dict_values_intern:
            dict_values_intern[key] = dict_values_intern[key].nonLocalizedName

        for key in dict_values_user:
            dict_values_user[key] = dict_values_user[key].localizedName

        for key, value in dict_values_intern.items():
            if value == "Zone_ZoneName":
                address_name = key
            elif value == "Zone_ZoneNumber":
                address_number = key

        for key in dict_values_intern:
            dict_values_intern[key] = acu.GetBuiltInPropertyId(dict_values_intern[key])

        for key, value in dict_values_user.items():
            #print(value)
            #print(key)
            dict_values_user.update({key: acu.GetUserDefinedPropertyId(value[0], value[1])})

        dict_final = {**dict_values_intern, **dict_values_user}

        zoneNumberPropertyId = acu.GetBuiltInPropertyId('Zone_ZoneNumber')
        zoneNamePropertyId = acu.GetBuiltInPropertyId('Zone_ZoneName')
        propertyValuesDictionary = acu.GetPropertyValuesDictionary(rooms, [zoneNumberPropertyId, zoneNamePropertyId])
        rooms = sorted(rooms, key=lambda r: propertyValuesDictionary[r][zoneNumberPropertyId])

        dict_rooms = {}

        #for i in range(0, 10):
        for room in rooms:
            #print(room.elementId.guid)
            dict_rooms[room.elementId.guid] = {}
            for key, value in dict_final.items():
                value_attr = acc.GetPropertyValuesOfElements([room], [value])
                for a in value_attr:
                    dict_rooms[room.elementId.guid][key] = a.propertyValues[0].propertyValue.value

        for key1, value in dict_rooms.items():

            for key2, value2 in value.items():
                if key2 == address_name:
                    roomName = value2
                elif key2 == address_number:
                    roomId = value2

            #print(roomName + " " + roomId)
            base = workbook_base.worksheets[0]
            worksheet = workbook_base.copy_worksheet(base)
            worksheet.oddHeader.center.text = "Raumbuch"
            worksheet.oddHeader.center.font = "Segoe UI,bold"
            worksheet.oddHeader.center.size = 14
            worksheet.evenHeader.center.text = "Raumbuch"
            worksheet.evenHeader.center.font = "Segoe UI,bold"
            worksheet.evenHeader.center.size = 14
            worksheet.title = str(roomId)
            for key3, value3 in value.items():
                worksheet[key3] = value3

        workbook_base.remove(base)
        workbook_base.save(path_target)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("Raumbuch erfolgreich exportiert!")
        msg.setWindowTitle("AC24: Attributübertragung")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
    # endregion

    # region Raumbuch PDF
    def set_controls_pdf(self):
        self.ui.cboFormatPdf.addItem("Format A4")
        self.ui.cboFormatPdf.addItem("Format A3")
        self.ui.cboFormatPdf.setCurrentIndex(0)

        self.ui.cboAusrichtungPdf.addItem("Hochformat")
        self.ui.cboAusrichtungPdf.addItem("Querformat")
        self.ui.cboAusrichtungPdf.setCurrentIndex(0)

        self.ui.cboSpaltenPdf.addItem("1 Spalte")
        self.ui.cboSpaltenPdf.addItem("2 Spalten")
        self.ui.cboSpaltenPdf.addItem("3 Spalten")
        self.ui.cboSpaltenPdf.addItem("4 Spalten")
        self.ui.cboSpaltenPdf.addItem("5 Spalten")
        self.ui.cboSpaltenPdf.setCurrentIndex(0)

        self.cols_change_pdf()

        with open(path_layer, newline='', encoding="utf-8") as csvdatei:
            csv_reader_object = csv.reader(csvdatei, delimiter=';')

            for row in csv_reader_object:
                self.ui.cboLayerPdf.addItem(row[0])

    def cols_change_pdf(self):
        table = self.ui.tblAttrPdf
        table.setColumnCount(0)
        table.setRowCount(0)
        cols = self.ui.cboSpaltenPdf.currentText()

        header = ["Typ Sp1", "Inhalt Sp1", "Typ Sp2", "Inhalt Sp2", "Typ Sp3", "Inhalt Sp3", "Typ Sp4", "Inhalt Sp4", "Typ Sp5", "Inhalt Sp5"]

        if cols == "1 Spalte":
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels(header[0:2])
            table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)
        elif cols == "2 Spalten":
            table.setColumnCount(4)
            table.setHorizontalHeaderLabels(header[0:4])
            table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)
        elif cols == "3 Spalten":
            table.setColumnCount(6)
            table.setHorizontalHeaderLabels(header[0:6])
            table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)
        elif cols == "4 Spalten":
            table.setColumnCount(8)
            table.setHorizontalHeaderLabels(header[0:8])
            table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)
        elif cols == "5 Spalten":
            table.setColumnCount(10)
            table.setHorizontalHeaderLabels(header)
            table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)

        table.insertRow(0)

    def add_row_pdf_up(self):
        table = self.ui.tblAttrPdf
        if table.rowCount() == 0:
            table.insertRow(0)
        else:
            rowPosition = table.currentRow()
            table.insertRow(rowPosition)

    def add_row_pdf_down(self):
        table = self.ui.tblAttrPdf
        if table.rowCount() == 0:
            table.insertRow(0)
        else:
            rowPosition = table.currentRow()
            table.insertRow(rowPosition+1)

    def delete_row_pdf(self):
        selected = self.ui.tblAttrPdf.currentRow()
        self.ui.tblAttrPdf.removeRow(selected)

    def output_directory_pdf(self):
        file = str(QFileDialog.getExistingDirectory(self, "Speicherort wählen", r"\\firma.local\dfs\Projekte"))
        self.ui.txtPathTargetPdf.setText(file)

    def open_template_pdf(self):
        fileName = QFileDialog.getOpenFileName(self, "CSV-Datei wählen", path_templates_pdf, "CSV-Dateien (*.csv)")
        self.ui.txtPathTemplatePdf.setText(fileName[0])

    def open_logo_pdf(self):
        fileName = QFileDialog.getOpenFileName(self, "Logo wählen", path_templates_pdf, "Bilddatei (*.jpg, *.png)")
        self.ui.txtPathLogoPdf.setText(fileName[0])
        
    def attr_save_pdf(self):
        path = self.ui.txtPathTemplatePdf.text()
        with open(path, "w", newline='', encoding="utf-8") as csvdatei:
            csv_writer_object = csv.writer(csvdatei, delimiter=';', quotechar='"')
    
            csv_writer_object.writerow([self.ui.cboFormatPdf.currentText()])
            csv_writer_object.writerow([self.ui.cboAusrichtungPdf.currentText()])
            csv_writer_object.writerow([self.ui.cboSpaltenPdf.currentText()])
            csv_writer_object.writerow([self.ui.txtProjektnamePdf.text()])

            table = self.ui.tblAttrPdf

            header = ["Typ Sp1", "Inhalt Sp1", "Typ Sp2", "Inhalt Sp2", "Typ Sp3", "Inhalt Sp3", "Typ Sp4", "Inhalt Sp4", "Typ Sp5", "Inhalt Sp5"]
            
            if table.columnCount() == 2:
                csv_writer_object.writerow(header[0:2])

                rows = table.rowCount()
                for i in range(0, rows):
                    rowContent = [
                        table.item(i, 0).text(),
                        table.item(i, 1).text()
                    ]
                    csv_writer_object.writerow(rowContent)
                    
            elif table.columnCount() == 4:
                csv_writer_object.writerow(header[0:4])

                rows = table.rowCount()
                for i in range(0, rows):
                    rowContent = [
                        table.item(i, 0).text(),
                        table.item(i, 1).text(),
                        table.item(i, 2).text(),
                        table.item(i, 3).text()
                    ]
                    csv_writer_object.writerow(rowContent)
                    
            elif table.columnCount() == 6:
                csv_writer_object.writerow(header[0:6])

                rows = table.rowCount()
                for i in range(0, rows):
                    rowContent = [
                        table.item(i, 0).text(),
                        table.item(i, 1).text(),
                        table.item(i, 2).text(),
                        table.item(i, 3).text(),
                        table.item(i, 4).text(),
                        table.item(i, 5).text()
                    ]
                    csv_writer_object.writerow(rowContent)
                    
            elif table.columnCount() == 8:
                csv_writer_object.writerow(header[0:8])

                rows = table.rowCount()
                for i in range(0, rows):
                    rowContent = [
                        table.item(i, 0).text(),
                        table.item(i, 1).text(),
                        table.item(i, 2).text(),
                        table.item(i, 3).text(),
                        table.item(i, 4).text(),
                        table.item(i, 5).text(),
                        table.item(i, 6).text(),
                        table.item(i, 7).text()
                    ]
                    csv_writer_object.writerow(rowContent)
                    
            elif table.columnCount() == 10:
                csv_writer_object.writerow(header)

                rows = table.rowCount()
                for i in range(0, rows):
                    rowContent = [
                        table.item(i, 0).text(),
                        table.item(i, 1).text(),
                        table.item(i, 2).text(),
                        table.item(i, 3).text(),
                        table.item(i, 4).text(),
                        table.item(i, 5).text(),
                        table.item(i, 6).text(),
                        table.item(i, 7).text(),
                        table.item(i, 8).text(),
                        table.item(i, 9).text()
                    ]
                    csv_writer_object.writerow(rowContent)


        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("Template überschrieben und gespeichert!")
        msg.setWindowTitle("AC25: Raumbuch PDF")
        msg.setWindowIcon(QtGui.QIcon(company_icon))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def attr_import_pdf(self):
        table = self.ui.tblAttrPdf
        table.setRowCount(0)
        table.setColumnCount(0)
        path = self.ui.txtPathTemplatePdf.text()
        with open(path, newline='', encoding="utf-8") as csvdatei:
            csv_reader_object = csv.reader(csvdatei, delimiter=';')

            zeilennummer = 0
            for row in csv_reader_object:

                if zeilennummer == 0:
                    self.ui.cboFormatPdf.setCurrentText(row[0])
                elif zeilennummer == 1:
                    self.ui.cboAusrichtungPdf.setCurrentText(row[0])
                elif zeilennummer == 2:
                    self.ui.cboSpaltenPdf.setCurrentText(row[0])
                    self.cols_change_pdf()
                    table.setRowCount(0)
                elif zeilennummer == 3:
                    self.ui.txtProjektnamePdf.setText(row[0])
                elif zeilennummer == 4:
                    table.setHorizontalHeaderLabels(row)
                    table.horizontalHeader().setDefaultAlignment(Qt.AlignHCenter)
                else:
                    row_position = table.rowCount()
                    table.insertRow(row_position)

                    if table.columnCount() == 2:
                        table.setItem(row_position, 0, QtWidgets.QTableWidgetItem(row[0]))
                        table.setItem(row_position, 1, QtWidgets.QTableWidgetItem(row[1]))
                    elif table.columnCount() == 4:
                        table.setItem(row_position, 0, QtWidgets.QTableWidgetItem(row[0]))
                        table.setItem(row_position, 1, QtWidgets.QTableWidgetItem(row[1]))
                        table.setItem(row_position, 2, QtWidgets.QTableWidgetItem(row[2]))
                        table.setItem(row_position, 3, QtWidgets.QTableWidgetItem(row[3]))
                    elif table.columnCount() == 6:
                        table.setItem(row_position, 0, QtWidgets.QTableWidgetItem(row[0]))
                        table.setItem(row_position, 1, QtWidgets.QTableWidgetItem(row[1]))
                        table.setItem(row_position, 2, QtWidgets.QTableWidgetItem(row[2]))
                        table.setItem(row_position, 3, QtWidgets.QTableWidgetItem(row[3]))
                        table.setItem(row_position, 4, QtWidgets.QTableWidgetItem(row[4]))
                        table.setItem(row_position, 5, QtWidgets.QTableWidgetItem(row[5]))
                    elif table.columnCount() == 8:
                        table.setItem(row_position, 0, QtWidgets.QTableWidgetItem(row[0]))
                        table.setItem(row_position, 1, QtWidgets.QTableWidgetItem(row[1]))
                        table.setItem(row_position, 2, QtWidgets.QTableWidgetItem(row[2]))
                        table.setItem(row_position, 3, QtWidgets.QTableWidgetItem(row[3]))
                        table.setItem(row_position, 4, QtWidgets.QTableWidgetItem(row[4]))
                        table.setItem(row_position, 5, QtWidgets.QTableWidgetItem(row[5]))
                        table.setItem(row_position, 6, QtWidgets.QTableWidgetItem(row[6]))
                        table.setItem(row_position, 7, QtWidgets.QTableWidgetItem(row[7]))
                    elif table.columnCount() == 10:
                        table.setItem(row_position, 0, QtWidgets.QTableWidgetItem(row[0]))
                        table.setItem(row_position, 1, QtWidgets.QTableWidgetItem(row[1]))
                        table.setItem(row_position, 2, QtWidgets.QTableWidgetItem(row[2]))
                        table.setItem(row_position, 3, QtWidgets.QTableWidgetItem(row[3]))
                        table.setItem(row_position, 4, QtWidgets.QTableWidgetItem(row[4]))
                        table.setItem(row_position, 5, QtWidgets.QTableWidgetItem(row[5]))
                        table.setItem(row_position, 6, QtWidgets.QTableWidgetItem(row[6]))
                        table.setItem(row_position, 7, QtWidgets.QTableWidgetItem(row[7]))
                        table.setItem(row_position, 8, QtWidgets.QTableWidgetItem(row[8]))
                        table.setItem(row_position, 9, QtWidgets.QTableWidgetItem(row[9]))
                zeilennummer += 1

        table.resizeColumnsToContents()

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("Template wurde importiert!")
        msg.setWindowTitle("AC25: Raumbuch PDF")
        msg.setWindowIcon(QtGui.QIcon(company_icon))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def attr_raumbuch_pdf(self):
        all_properties = acc.GetAllPropertyNames()
        tree = self.ui.treeUserAttrPdf
        tree.setHeaderLabels(["Name"])
        tree.setColumnWidth(0, 200)
        tree.setAlternatingRowColors(True)
        categories = []

        for item in all_properties:
            if item.type == "UserDefined":
                if not item.localizedName[0] in categories:
                    categories.append(item.localizedName[0])

        categories.append("Intern")

        for category in categories:
            item_tree = QtWidgets.QTreeWidgetItem([category])

            for item in all_properties:
                if item.type == "UserDefined":
                    if item.localizedName[0] == category:
                        QtWidgets.QTreeWidgetItem(item_tree, [item.localizedName[1] + "/" + item.localizedName[1] + "/" + category + "/benutzerdef."])

            with open(intern_properties, newline='', encoding="utf-8") as csvdatei:
                csv_reader_object = csv.reader(csvdatei, delimiter=';')

                zeilennummer = 0
                for row in csv_reader_object:

                    if zeilennummer == 0:
                        pass
                    elif category == "Intern":
                        QtWidgets.QTreeWidgetItem(item_tree, [row[0] + "/" + row[0] + "/" + row[1] + "/intern"])
                    zeilennummer += 1

            tree.addTopLevelItem(item_tree)

    def execute_roombook_pdf(self):
        # self.ui.tblAttr.sortItems(1, QtCore.Qt.DescendingOrder)
        path_target = self.ui.txtPathTargetPdf.text() + "/" + self.ui.txtFilenamePdf.text() + ".pdf"
        form = self.ui.cboFormatPdf.currentText()
        ori = self.ui.cboAusrichtungPdf.currentText()
        cols = self.ui.cboSpaltenPdf.currentText()
        proj = self.ui.txtProjektnamePdf.text()

        if form == "Format A4" and ori == "Hochformat":
            pdf = PDFBase.PDF(form=form, ori=ori, cols=cols, proj=proj, orientation='P', unit='mm', format='A4')
        elif form == "Format A4" and ori == "Querformat":
            pdf = PDFBase.PDF(form=form, ori=ori, cols=cols, proj=proj, orientation='L', unit='mm', format='A4')
        elif form == "Format A3" and ori == "Hochformat":
            pdf = PDFBase.PDF(form=form, ori=ori, cols=cols, proj=proj, orientation='P', unit='mm', format='A3')
        elif form == "Format A3" and ori == "Querformat":
            pdf = PDFBase.PDF(form=form, ori=ori, cols=cols, proj=proj, orientation='L', unit='mm', format='A3')

        if self.ui.txtPathLogoPdf.text() != "":
            pdf.logopath = self.ui.txtPathLogoPdf.text()


        table = self.ui.tblAttrPdf
        rows = table.rowCount()
        properties_pdf = {}
        if cols == "1 Spalte":
            for row in range(0, rows):
                typ1 = table.item(row, 0).text()
                if typ1 == "d":
                    value1 = table.item(row, 1).text()
                    value1_split = value1.split("/")
                    bez = value1_split[0]
                    attr = value1_split[1]
                    group = value1_split[2]
                    herkunft = value1_split[3]
                    properties_pdf[attr] = self.get_user_id(group, herkunft, attr)

        elif cols == "2 Spalten":
            for row in range(0, rows):
                for i in [0,2]:
                    typ1 = table.item(row, i).text()
                    if typ1 == "d":
                        value1 = table.item(row, i+1).text()
                        value1_split = value1.split("/")
                        attr = value1_split[1]
                        group = value1_split[2]
                        herkunft = value1_split[3]
                        properties_pdf[attr] = self.get_user_id(group, herkunft, attr)

        elif cols == "3 Spalten":
            for row in range(0, rows):
                for i in [0, 2, 4]:
                    typ1 = table.item(row, i).text()
                    if typ1 == "d":
                        value1 = table.item(row, i+1).text()
                        value1_split = value1.split("/")
                        attr = value1_split[1]
                        group = value1_split[2]
                        herkunft = value1_split[3]
                        properties_pdf[attr] = self.get_user_id(group, herkunft, attr)

        elif cols == "4 Spalten":
            for row in range(0, rows):
                for i in [0, 2, 4, 6]:
                    typ1 = table.item(row, i).text()
                    if typ1 == "d":
                        value1 = table.item(row, i + 1).text()
                        value1_split = value1.split("/")
                        attr = value1_split[1]
                        group = value1_split[2]
                        herkunft = value1_split[3]
                        properties_pdf[attr] = self.get_user_id(group, herkunft, attr)

        elif cols == "5 Spalten":
            for row in range(0, rows):
                for i in [0, 2, 4, 6, 8]:
                    typ1 = table.item(row, i).text()
                    if typ1 == "d":
                        value1 = table.item(row, i + 1).text()
                        value1_split = value1.split("/")
                        attr = value1_split[1]
                        group = value1_split[2]
                        herkunft = value1_split[3]
                        properties_pdf[attr] = self.get_user_id(group, herkunft, attr)

        properties_pdf['Ebene'] = self.get_user_id('ModelView_LayerName', 'intern', 'Ebene')

        rooms_pdf = acc.GetElementsByType("Zone")
        zoneNumberPropertyId = acu.GetBuiltInPropertyId('Zone_ZoneNumber')
        zoneNamePropertyId = acu.GetBuiltInPropertyId('Zone_ZoneName')
        layerPropertyId = acu.GetBuiltInPropertyId('ModelView_LayerName')
        propertyValuesDictionary = acu.GetPropertyValuesDictionary(rooms_pdf, [zoneNumberPropertyId, zoneNamePropertyId, layerPropertyId])
        rooms_pdf = sorted(rooms_pdf, key=lambda r: propertyValuesDictionary[r][zoneNumberPropertyId])
        dict_rooms_pdf = {}

        for room in rooms_pdf:
            dict_rooms_pdf[room.elementId.guid] = {}
            for key, value in properties_pdf.items():
                value_attr = acc.GetPropertyValuesOfElements([room], [value])
                for a in value_attr:
                    dict_rooms_pdf[room.elementId.guid][key] = a.propertyValues[0].propertyValue.value

        dict_rooms_pdf = {key:value for key, value in dict_rooms_pdf.items() if value.get('Ebene') == self.ui.cboLayerPdf.currentText()}

        table = self.ui.tblAttrPdf
        rows = table.rowCount()
        for key1, wert in dict_rooms_pdf.items():
            pdf.add_page()
            if cols == "1 Spalte":
                for row in range(0, rows):
                    typ1 = table.item(row, 0).text()
                    if typ1 == "ü":
                        pdf.chapter(table.item(row, 1).text())
                    elif typ1 == "l":
                        pdf.empty()
                    elif typ1 == "d":
                        value1 = table.item(row, 1).text()
                        value1_split = value1.split("/")
                        bez = value1_split[0]
                        attr = value1_split[1]
                        value = wert.get(attr)
                        if type(value) == int:
                            pdf.data(bez, float_to_string(value))
                        elif type(value) == float:
                            pdf.data(bez, float_to_string(value))
                        else:
                            pdf.data(bez, value)
                pdf.ln(7)

            elif cols == "2 Spalten":
                for row in range(0, rows):
                    for i in [0, 2]:
                        typ1 = table.item(row, i).text()
                        if typ1 == "ü":
                            pdf.chapter(table.item(row, i+1).text())
                        elif typ1 == "l":
                            pdf.empty()
                        elif typ1 == "d":
                            value1 = table.item(row, i+1).text()
                            value1_split = value1.split("/")
                            bez = value1_split[0]
                            attr = value1_split[1]
                            value = wert.get(attr)
                            if type(value) == int:
                                pdf.data(bez, float_to_string(value))
                            elif type(value) == float:
                                pdf.data(bez, float_to_string(value))
                            else:
                                pdf.data(bez, value)
                    pdf.ln(7)

            elif cols == "3 Spalten":
                for row in range(0, rows):
                    for i in [0, 2, 4]:
                        typ1 = table.item(row, i).text()
                        if typ1 == "ü":
                            pdf.chapter(table.item(row, i+1).text())
                        elif typ1 == "l":
                            pdf.empty()
                        elif typ1 == "d":
                            value1 = table.item(row, i+1).text()
                            value1_split = value1.split("/")
                            bez = value1_split[0]
                            attr = value1_split[1]
                            value = wert.get(attr)
                            if type(value) == int:
                                pdf.data(bez, float_to_string(value))
                            elif type(value) == float:
                                pdf.data(bez, float_to_string(value))
                            else:
                                pdf.data(bez, value)
                    pdf.ln(7)

            elif cols == "4 Spalten":
                for row in range(0, rows):
                    for i in [0, 2, 4, 6]:
                        typ1 = table.item(row, i).text()
                        if typ1 == "ü":
                            pdf.chapter(table.item(row, i+1).text())
                        elif typ1 == "l":
                            pdf.empty()
                        elif typ1 == "d":
                            value1 = table.item(row, i+1).text()
                            value1_split = value1.split("/")
                            bez = value1_split[0]
                            attr = value1_split[1]
                            value = wert.get(attr)
                            if type(value) == int:
                                pdf.data(bez, float_to_string(value))
                            elif type(value) == float:
                                pdf.data(bez, float_to_string(value))
                            else:
                                pdf.data(bez, value)
                    pdf.ln(7)

            elif cols == "5 Spalten":
                for row in range(0, rows):
                    for i in [0, 2, 4, 6, 8]:
                        typ1 = table.item(row, i).text()
                        if typ1 == "ü":
                            pdf.chapter(table.item(row, i+1).text())
                        elif typ1 == "l":
                            pdf.empty()
                        elif typ1 == "d":
                            value1 = table.item(row, i+1).text()
                            value1_split = value1.split("/")
                            bez = value1_split[0]
                            attr = value1_split[1]
                            value = wert.get(attr)
                            if type(value) == int:
                                pdf.data(bez, float_to_string(value))
                            elif type(value) == float:
                                pdf.data(bez, float_to_string(value))
                            else:
                                pdf.data(bez, value)
                    pdf.ln(7)

        pdf.output(path_target)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("Raumbuch erfolgreich exportiert!")
        msg.setWindowTitle("AC25: Raumbuch PDF")
        msg.setWindowIcon(QtGui.QIcon(company_icon))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
    # endregion

    # region Raumbuch Attributübertragung
    def user_attr(self):
        all_properties = acc.GetAllPropertyNames()
        tree = self.ui.user_attr
        tree.setHeaderLabels(["Name"])
        tree.setAlternatingRowColors(True)
        categories = []

        for item in all_properties:
            if item.type == "UserDefined":
                if not item.localizedName[0] in categories:
                    categories.append(item.localizedName[0])

        for category in categories:
            item_tree = QtWidgets.QTreeWidgetItem([category])

            for item in all_properties:
                if item.type == "UserDefined":
                    if item.localizedName[0] == category:
                        QtWidgets.QTreeWidgetItem(item_tree, [item.localizedName[1]])

            tree.addTopLevelItem(item_tree)

    def intern_attr(self):
        interns = [["Raumnummer", "Zone_ZoneNumber"]]
        tree = self.ui.intern_attr
        tree.setHeaderLabels(["Name", "Interner Name"])
        tree.setAlternatingRowColors(True)

        for intern in interns:
            item_tree = QtWidgets.QTreeWidgetItem([intern[0], intern[1]])
            tree.addTopLevelItem(item_tree)

    def user_attr_selected(self):
        get_selected = self.ui.user_attr.selectedItems()
        selected = ""
        if get_selected:
            selected = get_selected[0].text(0)

        all_properties = acc.GetAllPropertyNames()
        categories = []

        for item in all_properties:
            if item.type == "UserDefined":
                if not item.localizedName[0] in categories:
                    categories.append(item.localizedName[0])

        if selected in categories:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)

            msg.setText("Fehler Auswahl!")
            msg.setInformativeText("Hauptkategorie kann nicht ausgewählt werden!")
            msg.setWindowTitle("AC25: Attributübertragung")
            msg.setWindowIcon(QtGui.QIcon(company_icon))
            msg.setDetailedText(
                "Für die Übertragung von Attributen muss ein benutzerdefiniertes Attribut ausgewählt werden. Sie "
                "haben eine Gruppenbezeichnung ausgewählt!")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            msg.exec_()

        else:
            for item in all_properties:
                if item.type == "UserDefined":
                    if item.localizedName[1] == selected:
                        self.ui.gew_user_attr_2.setText(item.localizedName[0])

            self.ui.gew_user_attr.setText(selected)
            self.ui.final_user_attr.setText(selected)

    def intern_attr_selected(self):
        get_selected = self.ui.intern_attr.selectedItems()
        if get_selected:
            selected = get_selected[0].text(0)
            selected2 = get_selected[0].text(1)

            self.ui.gew_intern_attr.setText(selected)
            self.ui.gew_intern_attr_2.setText(selected2)
            self.ui.final_intern_attr.setText(selected + " (" + selected2 + ")")

    def execute(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)

        msg.setText("Anweisung wird ausgeführt!")
        msg.setInformativeText(
            "Sind sie sicher das Sie die Anweisung ausführen möchten? Dieser Vorgang kann nicht rückgängig gemacht "
            "werden!")
        msg.setWindowTitle("AC24: Attributübertragung")
        msg.setWindowIcon(QtGui.QIcon(company_icon))
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        return_value = msg.exec_()

        if return_value == QMessageBox.Ok:
            print("Anweisung wird ausgeführt!")

            if self.ui.category.currentText() == "Räume":
                elements = acc.GetElementsByType("Zone")

            user_id = act.UserDefinedPropertyUserId([self.ui.gew_user_attr_2.text(), self.ui.gew_user_attr.text()])
            user_id2 = user_id.localizedName
            user_id3 = acu.GetUserDefinedPropertyId(user_id2[0], user_id2[1])
            interne_id = act.BuiltInPropertyUserId(self.ui.gew_intern_attr_2.text())
            interne_id2 = interne_id.nonLocalizedName
            interne_id3 = acu.GetBuiltInPropertyId(interne_id2)
            #print(interne_id3)
            #print(elements)

            final = []

            for element in elements:
                attr = acc.GetPropertyValuesOfElements([element.elementId], [user_id3])
                #print(attr)
                #print(element)
                #print(act.ElementPropertyValue(element.elementId, interne_id3, act.NormalStringPropertyValue(attr[0].propertyValues[0].propertyValue.value)))
                final.append(act.ElementPropertyValue(element.elementId, interne_id3, act.NormalStringPropertyValue(attr[0].propertyValues[0].propertyValue.value)))

            acc.SetPropertyValuesOfElements(final)
            print("Anweisungen wurden erfolgreich ausgeführt!")

        else:
            print("Prozess abgebrochen!")

    def count_elements(self):
        if self.ui.category.currentText() == "Räume":
            elements = acc.GetElementsByType("Zone")
        elif self.ui.category.currentText() == "Wände":
            elements = acc.GetElementsByType("Wall")

        i = 0
        for element in elements:
            i += 1

        self.ui.count.setText(str(i))
    # endregion

    # region AC-Functions
    def get_user_id(self, name, herkunft, name2=""):
        if herkunft == "intern":
            return acu.GetBuiltInPropertyId(name)
        elif herkunft == "benutzerdef.":
            return acu.GetUserDefinedPropertyId(name, name2)

    # endregion





app = QtWidgets.QApplication(sys.argv)

window = SplashScreen()
window.show()

sys.exit(app.exec_())