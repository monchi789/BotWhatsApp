import os
import pywhatkit
import openpyxl


def bot(nombreExcel, columnaNombre, columnaNumero, filaInicio, filaFinal, mensaje):

    book = openpyxl.load_workbook('./docs/' + nombreExcel + '.xlsx')
    sheet = book.active

    # Convertir tipos de datos
    filaInicio = int(filaInicio)
    filaFinal = int(filaFinal)

    contador = 0

    for i in range(filaInicio, filaFinal):
        numero = str(sheet[columnaNumero + f'{i}'].value)
        if(numero[0] == '9'):
            numero = '+51' + numero
            mensaje = '➡️ Hola, ' + str(sheet[ columnaNombre + f'{i}'].value) + '. ' + mensaje
            pywhatkit.sendwhatmsg_instantly( numero,
                        mensaje,
                        15,
                        tab_close=True)
            os.remove('PyWhatKit_DB.txt')
            mensaje = ''
        else:
            contador += 1
            print('Numero no valido ', contador)
        