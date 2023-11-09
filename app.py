#CODIGO PRINCIPAL Y DE LA INTERFAZ

import sys
import os
from PyQt5 import uic
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QDialog
from PyQt5.QtGui import QIntValidator, QPixmap, QIcon
import funciones
import openpyxl
import threading

class Automatizador(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("interfaz.ui", self)
        self.setWindowTitle("Automatizador")
        icon = QIcon("icono.ico")
        self.setWindowIcon(icon)
        self.pdf1_1Button.clicked.connect(self.buscarPdf1_1)
        self.pdf1_2Button.clicked.connect(self.buscarPdf1_2)
        self.pdf2_1Button.clicked.connect(self.buscarPdf2_1)
        self.pdf2_2Button.clicked.connect(self.buscarPdf2_2)
        self.excelButton.clicked.connect(self.buscarExcel)
        self.pathButton.clicked.connect(self.buscarRuta)
        self.autoButton.clicked.connect(self.ejecutarHilo)
        self.helpButton.clicked.connect(self.abrir_ventana)

        self.ruta_pdf1_1 = None
        self.ruta_pdf1_2 = None
        self.ruta_pdf2_1 = None
        self.ruta_pdf2_2 = None
        self.ruta_excel = None
        self.ruta_guardado = None

        int_validator = QIntValidator()
        self.lineEditYear1.setValidator(int_validator)
        self.lineEditYear2.setValidator(int_validator)

        self.hiloFinalizado = False

    def abrir_ventana(self):
        # Crea una instancia de la ventana secundaria
        self.ventana_secundaria = VentanaSecundaria()
        self.ventana_secundaria.show()

    def buscarPdf1_1(self):
        fileName = QFileDialog.getOpenFileName(self, "Abrir archivo", "C:", "Archivos PDF (*.pdf)")
        self.ruta_pdf1_1 = fileName[0]
        self.lineEditPdf1_1.setText(fileName[0])
    def buscarPdf1_2(self):
        fileName = QFileDialog.getOpenFileName(self, "Abrir archivo", "C:", "Archivos PDF (*.pdf)")
        self.ruta_pdf1_2 = fileName[0]
        self.lineEditPdf1_2.setText(fileName[0])
    def buscarPdf2_1(self):
        fileName = QFileDialog.getOpenFileName(self, "Abrir archivo", "C:", "Archivos PDF (*.pdf)")
        self.ruta_pdf2_1 = fileName[0]
        self.lineEditPdf2_1.setText(fileName[0])
    def buscarPdf2_2(self):
        fileName = QFileDialog.getOpenFileName(self, "Abrir archivo", "C:", "Archivos PDF (*.pdf)")
        self.ruta_pdf2_2 = fileName[0]
        self.lineEditPdf2_2.setText(fileName[0])
    def buscarExcel(self):
        fileName = QFileDialog.getOpenFileName(self, "Abrir archivo", "C:", "Archivos de Excel (*.xlsx)")
        self.ruta_excel = fileName[0]
        self.lineEditExcel.setText(fileName[0])
    def buscarRuta(self):
        fileName = QFileDialog.getExistingDirectory(self, "Seleccionar ruta", "C:")
        self.ruta_guardado = fileName
        self.lineEditSave.setText(fileName)
    def ejecutarHilo(self):
        if self.ruta_pdf1_1 and self.ruta_pdf1_2 and self.ruta_pdf2_1 and self.ruta_pdf2_2 and self.ruta_excel and self.ruta_guardado and self.lineEditYear1.text() and self.lineEditYear2.text():
            self.autoButton.setEnabled(False)
            self.pdf1_1Button.setEnabled(False)
            self.pdf1_2Button.setEnabled(False)
            self.pdf2_1Button.setEnabled(False)
            self.pdf2_2Button.setEnabled(False)
            self.excelButton.setEnabled(False)
            self.pathButton.setEnabled(False)
            self.lineEditYear1.setReadOnly(True)
            self.lineEditYear2.setReadOnly(True)

            self.thread = threading.Thread(target=self.automatizar)
            self.thread.start()
        else:
            QMessageBox.warning(self, "Aviso", "Todos los campos son obligatorios, favor de llenar todos.")

    def automatizar(self):
        #main.generarPdfsYExcel(self.ruta_excel, [self.ruta_pdf1_1, self.ruta_pdf1_2], [self.ruta_pdf2_1, self.ruta_pdf2_2], self.ruta_guardado, self.lineEditYear1.text(), self.lineEditYear2.text())

        self.progressBar.setValue(0)

        excel = openpyxl.load_workbook(self.ruta_excel)

        if not os.path.exists(os.path.join(self.ruta_guardado, self.lineEditYear1.text())):
            os.mkdir(os.path.join(self.ruta_guardado, self.lineEditYear1.text()))
        if not os.path.exists(os.path.join(self.ruta_guardado, self.lineEditYear2.text())):
            os.mkdir(os.path.join(self.ruta_guardado, self.lineEditYear2.text()))
    
        nombres = funciones.encontrarNombres(excel)
        contador = 0

        for nombre in nombres:
            celdaTitulo = funciones.encontrarCeldaTitulo(excel, nombre)
            filas, numCelda = funciones.encontrarCeldasTabla(excel, self.lineEditYear1.text(),celdaTitulo.coordinate)
            print(filas)
            funciones.compararConPDF(excel, nombre, filas, [self.ruta_pdf1_1, self.ruta_pdf1_2], self.lineEditYear1.text(), self.ruta_guardado)
            excel.save(self.ruta_excel)

            filas = funciones.encontrarCeldasTabla_2(excel, self.lineEditYear2.text(), numCelda)
            print(filas)
            funciones.compararConPDF(excel, nombre, filas, [self.ruta_pdf2_1, self.ruta_pdf2_2], self.lineEditYear2.text(), self.ruta_guardado)
            excel.save(self.ruta_excel)

            contador += 1
            self.progressBar.setValue(round((contador / len(nombres)) * 100))
            #time.sleep(0.05)

        self.autoButton.setEnabled(True)
        self.pdf1_1Button.setEnabled(True)
        self.pdf1_2Button.setEnabled(True)
        self.pdf2_1Button.setEnabled(True)
        self.pdf2_2Button.setEnabled(True)
        self.excelButton.setEnabled(True)
        self.pathButton.setEnabled(True)
        self.lineEditYear1.setReadOnly(False)
        self.lineEditYear2.setReadOnly(False)

        self.lineEditPdf1_1.clear()
        self.lineEditPdf1_2.clear()
        self.lineEditPdf2_1.clear()
        self.lineEditPdf2_2.clear()
        self.lineEditYear1.clear()
        self.lineEditYear2.clear()
        self.lineEditExcel.clear()
        self.lineEditSave.clear()

class VentanaSecundaria(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi("acercaDe.ui", self)
        self.setWindowTitle("Acerca del programa")
        icon = QIcon("icono.ico")
        self.setWindowIcon(icon)
        # Acceder al objeto gráfico y asignarle la imagen
        pixmapPDFS = QPixmap("./imgs/pdfs.png")
        self.imgPDFS.setPixmap(pixmapPDFS)

        pixmapEXCEL = QPixmap("./imgs/excel.png")
        self.imgEXCEL.setPixmap(pixmapEXCEL)

        pixmapRUTA = QPixmap("./imgs/ruta.png")
        self.imgRUTA.setPixmap(pixmapRUTA)

        pixmapYEARS = QPixmap("./imgs/years.png")
        self.imgYEARS.setPixmap(pixmapYEARS)

        pixmapAUTO = QPixmap("./imgs/automatizarButton.png")
        self.imgAUTO.setPixmap(pixmapAUTO)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    gui = Automatizador()
    gui.show()
    sys.exit(app.exec_())