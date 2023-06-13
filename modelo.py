from PyQt5 import QtGui
from PyQt5.QtWidgets import QMessageBox
from pandas import read_excel
from pandas import DataFrame
from sqlite3 import connect



class Modelo():

    def deduccion_especial(self, reg, promrem):
        
        #Declaración de variables.
        dedesp = 0
        valores_minimo = []
        valores_maximos = []
        deducciones = []

        #Conexión a base de datos.
        conexion = connect('db/BeneficioGanancias.db')

        #Busca la cantidad de registros.
        cursor_count = conexion.cursor()
        valores = cursor_count.execute('SELECT count(id) FROM beneficio_ganancias WHERE resolucion = ?', (reg,))
        cantidad = valores.fetchone()[0]
        
        #Extrae los valores minimos según resolución.
        cursor_vmin = conexion.cursor()
        valores = cursor_vmin.execute('SELECT val_min FROM beneficio_ganancias WHERE resolucion = ?', (reg,))
        v_min = valores.fetchall()
        
        for i in v_min:
            valores_minimo.append(i[0]) 

        cursor_vmin.close()
        
        #Extrae los valores máximos según resolución.
        cursor_vmax = conexion.cursor()
        valores = cursor_vmax.execute('SELECT val_max FROM beneficio_ganancias WHERE resolucion = ?', (reg,))
        v_max = valores.fetchall()
        
        for i in v_max:
            valores_maximos.append(i[0]) 

        cursor_vmin.close()
        
        #Extrae los valores de las deducciones especiales según resolución.
        cursor_ded = conexion.cursor()
        valores = cursor_ded.execute('SELECT deduccion FROM beneficio_ganancias WHERE resolucion = ?', (reg,))
        ded = valores.fetchall()
        
        for i in ded:
            deducciones.append(i[0]) 

        cursor_ded.close()

        conexion.close()
        
        #Compara la información extraida de la BBDD y trae la deducción especial.
        
        i = 0

        while i < cantidad:
            if promrem >= valores_minimo[i] and promrem < valores_maximos[i]: dedesp = deducciones[i] 
            i = i + 1
            
        return dedesp

    def leer_excel(self, archivo, directorio, resolucion, formulario):

        try:

            escala = []

            excel = read_excel(archivo)

            legajo = excel['N° Legajo'].values
            liquidacion = excel['N° Liquidación'].values
            codigo = excel['Código'].values
            importe =excel['Importe'].values
            valor_escala = excel['Importe'].values

            datos = {'N° Legajo' : [],
                    'N° liquidacion' : [],
                    'Código' : [],
                    'Importe' : [],
                    'Escala' : []  }
                               
            info_icon = QtGui.QIcon()
            info_icon.addPixmap(QtGui.QPixmap("img/info-circle-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msg_box = QMessageBox()
            msg_box.setWindowIcon(info_icon)
            msg_box.setText(f"Se va a procesar el archivo {archivo}. Presione 'Ok' para continuar.")
            msg_box.setWindowTitle("Información")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec()
            
            formulario.setEnabled(False)

            for i in valor_escala:
                escala.append(self.deduccion_especial(resolucion, i))                 

            datos['N° Legajo'] = legajo
            datos['N° liquidacion'] = liquidacion
            datos['Código'] = codigo
            datos['Importe'] = importe
            datos['Escala'] = escala

            resultado = DataFrame(datos,columns=['N° Legajo','N° liquidacion','Codigo','Importe','Escala'])

            salida = directorio + '/resultado.xlsx'
            resultado.to_excel(salida, sheet_name='Hoja1')

            x = len(escala)   
            mensaje = f'Se procesaron {x} registros'
            
            info_icon = QtGui.QIcon()
            info_icon.addPixmap(QtGui.QPixmap("img/info-circle-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msg_box = QMessageBox()
            msg_box.setWindowIcon(info_icon)
            msg_box.setText(mensaje)
            msg_box.setWindowTitle("Información")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec()

            formulario.setEnabled(True)

        except KeyError:
            bug_icon = QtGui.QIcon()
            bug_icon.addPixmap(QtGui.QPixmap("img/bug-solid.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msg_box = QMessageBox()
            msg_box.setWindowIcon(bug_icon)
            msg_box.setText("Error en formato de archivo xlxs.")
            msg_box.setWindowTitle("Error")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec()

            formulario.setEnabled(True)

class ErrorDirectorio(Exception):
    
    def __init__(self):
        super().__init__()
        self.mensaje = 'Error en el directorio'
        self.informacion = """Falta definir el directorio de salida o, para el directorio seleccionado,
                              no posee permisos de escritura."""
        print(self.mensaje)
        print(self.informacion)

class ErrorArchivo(Exception):
    
    def __init__(self):
        super().__init__()
        self.mensaje = 'Error en el archivo'
        self.informacion = """Falta seleccionar el archivo a procesar o, 
                              el archivo seleccionado, contiene un formato erroneo."""
        print(self.mensaje)
        print(self.informacion)