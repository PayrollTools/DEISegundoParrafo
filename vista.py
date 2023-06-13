import sys
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QFileDialog
from modelo import Modelo
from modelo import ErrorArchivo
from modelo import ErrorDirectorio
import webbrowser

class MainWindow():
    
    def __init__ (self):
        
        self.new_model = Modelo()

        app = QtWidgets.QApplication(sys.argv)
        self.frm_mainwindow = QtWidgets.QMainWindow()

        self.frm_mainwindow.setObjectName("frm_mainwindow")
        self.frm_mainwindow.setWindowModality(QtCore.Qt.ApplicationModal)
        self.frm_mainwindow.resize(463, 202)
        self.frm_mainwindow.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)
        self.frm_mainwindow.setMinimumSize(463, 202)
        self.frm_mainwindow.setMaximumSize(463, 202)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("img/calculator-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.frm_mainwindow.setWindowIcon(icon)
        self.frm_mainwindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(self.frm_mainwindow)
        self.centralwidget.setObjectName("centralwidget")
        self.lbl_0 = QtWidgets.QLabel(self.centralwidget)
        self.lbl_0.setGeometry(QtCore.QRect(10, 10, 61, 16))
        self.lbl_0.setObjectName("lbl_0")
        self.cmb_resolucion = QtWidgets.QComboBox(self.centralwidget)
        self.cmb_resolucion.setGeometry(QtCore.QRect(120, 10, 321, 22))
        self.cmb_resolucion.setEditable(False)
        self.cmb_resolucion.setPlaceholderText("")
        self.cmb_resolucion.setObjectName("cmb_resolucion")
        self.cmb_resolucion.addItem("Resolución General 5008/2021")
        self.cmb_resolucion.addItem("Resolución General 5076/2021")
        self.cmb_resolucion.addItem("Actualización RIPTE 2022")
        self.cmb_resolucion.addItem("Resolución General 5206/2022")
        self.cmb_resolucion.addItem("Resolución General 5280/2022")
        self.cmb_resolucion.addItem("Actualización RIPTE 2023")
        self.lbl_1 = QtWidgets.QLabel(self.centralwidget)
        self.lbl_1.setGeometry(QtCore.QRect(10, 50, 101, 16))
        self.lbl_1.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.lbl_1.setObjectName("lbl_1")
        self.txt_archivo = QtWidgets.QLineEdit(self.centralwidget)
        self.txt_archivo.setGeometry(QtCore.QRect(120, 50, 241, 20))
        self.txt_archivo.setObjectName("lineEdit")
        self.btn_abrir = QtWidgets.QPushButton(self.centralwidget, clicked=self.abrir_archivo)
        self.btn_abrir.setGeometry(QtCore.QRect(370, 50, 75, 23))
        self.btn_abrir.setObjectName("btn_abrir")
        self.lbl_2 = QtWidgets.QLabel(self.centralwidget)
        self.lbl_2.setGeometry(QtCore.QRect(10, 90, 101, 16))
        self.lbl_2.setObjectName("lbl_2")
        self.txt_directorio = QtWidgets.QLineEdit(self.centralwidget)
        self.txt_directorio.setGeometry(QtCore.QRect(120, 90, 241, 20))
        self.txt_directorio.setObjectName("lineEdit_2")
        self.btn_directorio = QtWidgets.QPushButton(self.centralwidget, clicked = self.directorio_salida)
        self.btn_directorio.setGeometry(QtCore.QRect(370, 90, 75, 23))
        self.btn_directorio.setObjectName("btn_directorio")
        self.btn_procesar = QtWidgets.QPushButton(self.centralwidget, clicked=self.procesar)
        self.btn_procesar.setGeometry(QtCore.QRect(370, 130, 75, 23))
        self.btn_procesar.setObjectName("btn_procesar")
        self.frm_mainwindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self.frm_mainwindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 463, 21))
        self.menubar.setObjectName("menubar")
        self.menuArchivo = QtWidgets.QMenu(self.menubar)
        self.menuArchivo.setObjectName("menuArchivo")
        self.menuAyuda = QtWidgets.QMenu(self.menubar)
        self.menuAyuda.setObjectName("menuAyuda")
        self.frm_mainwindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(self.frm_mainwindow)
        self.statusbar.setObjectName("statusbar")
        self.frm_mainwindow.setStatusBar(self.statusbar)
        self.actionSalir = QtWidgets.QAction(self.frm_mainwindow)
        
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("img/sign-out-alt-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionSalir.setIcon(icon1)
        self.actionSalir.setObjectName("actionSalir")
        self.actionSalir.triggered.connect(lambda:app.quit())
        self.actionFormato_xlsx = QtWidgets.QAction(self.frm_mainwindow)
        
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("img/rss-square-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionFormato_xlsx.setIcon(icon2)
        self.actionFormato_xlsx.setObjectName("actionFormato_xlsx")
        self.actionFormato_xlsx.triggered.connect(lambda: webbrowser.open("help.html"))
        self.actionAcerca_de = QtWidgets.QAction(self.frm_mainwindow)
        
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("img/question-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.actionAcerca_de.setIcon(icon3)
        self.actionAcerca_de.setObjectName("actionAcerca_de")
        self.menuArchivo.addAction(self.actionSalir)
        self.menuAyuda.addAction(self.actionFormato_xlsx)
        self.actionAcerca_de.triggered.connect(self.formulario_acerca)

        self.menuAyuda.addSeparator()
        self.menuAyuda.addAction(self.actionAcerca_de)
        self.menubar.addAction(self.menuArchivo.menuAction())
        self.menubar.addAction(self.menuAyuda.menuAction())    
        
        QtCore.QMetaObject.connectSlotsByName(self.frm_mainwindow)

        _translate = QtCore.QCoreApplication.translate
        self.frm_mainwindow.setWindowTitle(_translate("frm_mainwindow", "Rango deducción ganancias - V 2.0"))
        self.lbl_0.setText(_translate("frm_mainwindow", "Resolución:"))
        self.lbl_1.setText(_translate("frm_mainwindow", "Archivo a procesar:"))
        self.btn_abrir.setText(_translate("frm_mainwindow", "Abrir"))
        self.lbl_2.setText(_translate("frm_mainwindow", "Directorio de salida:"))
        self.btn_directorio.setText(_translate("frm_mainwindow", "Directorio"))
        self.btn_procesar.setText(_translate("frm_mainwindow", "Procesar"))
        self.menuArchivo.setTitle(_translate("frm_mainwindow", "Archivo"))
        self.menuAyuda.setTitle(_translate("frm_mainwindow", "Ayuda"))
        self.actionSalir.setText(_translate("frm_mainwindow", "Salir"))
        self.actionFormato_xlsx.setText(_translate("frm_mainwindow", "Ayuda"))
        self.actionAcerca_de.setText(_translate("frm_mainwindow", "Acerca de..."))
        
        self.frm_mainwindow.show()
        sys.exit(app.exec_())
        

    def abrir_archivo(self):
        opcion = QFileDialog.Options()
        
        file_name = QFileDialog.getOpenFileName(self.frm_mainwindow,"Seleccionar archivo", "",
                                               "Archivos excel (*xlsx);; All Files (*)", options = opcion)        
        self.txt_archivo.clear()
        self.txt_archivo.insert(file_name[0])
        

    def directorio_salida(self):
        opcion = QFileDialog.Options()
        directorio = QFileDialog.getExistingDirectory(self.frm_mainwindow,"Seleccionar directorio","",opcion)

        self.txt_directorio.clear()
        self.txt_directorio.insert(directorio)


    def procesar(self):
        try:
            if self.cmb_resolucion.currentText() == 'Resolución General 5008/2021':
                rg = 'RG_5008'
            elif self.cmb_resolucion.currentText() == 'Resolución General 5076/2021':
                rg = 'RG_5076'
            elif self.cmb_resolucion.currentText() == "Actualización RIPTE 2022":
                rg = 'RIPTE_2022'
            elif self.cmb_resolucion.currentText() == "Resolución General 5206/2022":
                rg = 'RG 5206/2022'
            elif self.cmb_resolucion.currentText() == "Resolución General 5280/2022":
                rg = 'RG 5280/2022'
            elif self.cmb_resolucion.currentText() == "Actualización RIPTE 2023":
                rg = 'RIPTE 2023'
            elif self.cmb_resolucion.currentText() == "Resolución General 5358/2023":
                rg = 'RG 5358/2023'


            if self.txt_archivo.text() == '':
                raise ErrorArchivo
            elif self.txt_directorio.text() == '':
                raise ErrorDirectorio
        
            archivo = self.txt_archivo.text()
            directorio = self.txt_directorio.text()            

            self.new_model.leer_excel(archivo, directorio, rg, self.frm_mainwindow)

            self.txt_directorio.clear()
            self.txt_archivo.clear()

        except ErrorDirectorio:
            bug_icon = QtGui.QIcon()
            bug_icon.addPixmap(QtGui.QPixmap("img/bug-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msg_box = QMessageBox()
            msg_box.setWindowIcon(bug_icon)
            msg_box.setText("Falta seleccionar el directorio de salida.")
            msg_box.setWindowTitle("Error")
            msg_box.setStandardButtons(QMessageBox.Ok)
            #msg_box.buttonClicked.connect(msgButtonClick)
            msg_box.exec()
        
        except ErrorArchivo:
            bug_icon = QtGui.QIcon()
            bug_icon.addPixmap(QtGui.QPixmap("img/bug-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msg_box = QMessageBox()
            msg_box.setWindowIcon(bug_icon)
            msg_box.setText("Falta seleccionar el archivo a procesar o el mismo contiene un formato erroneo.")
            msg_box.setWindowTitle("Error")
            msg_box.setStandardButtons(QMessageBox.Ok)
            #msg_box.buttonClicked.connect(msgButtonClick)
            msg_box.exec()

    def formulario_acerca(self):
        
        frm_acerca = QtWidgets.QDialog()

        frm_acerca.setObjectName("frm_acerca")
        frm_acerca.resize(364, 131)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("img/info-circle-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        frm_acerca.setWindowIcon(icon)
        self.lbl_0 = QtWidgets.QLabel(frm_acerca)
        self.lbl_0.setGeometry(QtCore.QRect(10, 10, 221, 16))
        self.lbl_0.setObjectName("lbl_0")
        self.lbl_1 = QtWidgets.QLabel(frm_acerca)
        self.lbl_1.setGeometry(QtCore.QRect(10, 30, 221, 16))
        self.lbl_1.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.lbl_1.setObjectName("lbl_1")
        self.lbl_2 = QtWidgets.QLabel(frm_acerca)
        self.lbl_2.setGeometry(QtCore.QRect(10, 50, 221, 16))
        self.lbl_2.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.lbl_2.setObjectName("lbl_2")
        self.lbl_3 = QtWidgets.QLabel(frm_acerca)
        self.lbl_3.setGeometry(QtCore.QRect(10, 70, 221, 16))
        self.lbl_3.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.lbl_3.setObjectName("lbl_3")
        self.btn_cerrar = QtWidgets.QPushButton(frm_acerca, clicked = frm_acerca.close)
        self.btn_cerrar.setGeometry(QtCore.QRect(280, 100, 75, 23))
        self.btn_cerrar.setObjectName("btn_cerrar")
        

        #QtCore.QMetaObject.connectSlotsByName(frm_acerca)

        _translate = QtCore.QCoreApplication.translate
        frm_acerca.setWindowTitle(_translate("frm_acerca", "Acerca de..."))
        self.lbl_0.setText(_translate("frm_acerca", "Rango deducciones ganancias - Versión 2.0"))
        self.lbl_1.setText(_translate("frm_acerca", "Escalas actualizadas al 01/01/2023"))
        self.lbl_2.setText(_translate("frm_acerca", "Desarrollado por: Jorge Eduardo Arrieta"))
        self.lbl_3.setText(_translate("frm_acerca", "Año, 2022 - Buenos Aires, Argentina"))
        self.btn_cerrar.setText(_translate("frm_acerca", "Cerrar"))        

        frm_acerca.exec_()        