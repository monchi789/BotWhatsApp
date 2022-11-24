import os
import pywhatkit
import openpyxl

book = openpyxl.load_workbook('prueba.xlsx')
sheet = book.active



texto = '''
¿Deseas conectarte con nuevos socios y aliados para tu negocio?

La Feria Internacional de Turismo *FITUR CUSCO 2022 “NUEVAS EXPERIENCIAS”* se realizará a partir del 19 al 21 del presente mes y reunirá a los más importantes líderes de la industria del turismo nacional e internacional en la ciudad del Cusco.

Esta feria está dirigida a empresarios, inversionistas y operadores turísticos que deseen conocer, invertir, asociarse o comprar productos, servicios y experiencias turísticas en Cusco.

La *FITUR CUSCO 2022* también contará con una feria de proveedores afines al Sector Turismo, donde se dispondrá de zonas de exposición, rueda de negocios, espacios de networking, Smart Talks especializados y la promoción de nuevos destinos.

*PUEDES ASISTIR ( ¡ÚLTIMOS CUPOS! )* inscribiéndote en la página web www.fiturcusco.com .
*PUEDES SER PARTE DE LA FERIA*, Últimos stands en venta, para más información comunicarse al whatsapp +51 984 674 491. 

Organizado por *EMUFEC*.
Socios estratégicos: Cámara de Comercio de Cusco - Municipalidad del Cusco.
Participan: PromPerú Oficial - Visit Peru - SmArt Tourism & Hospitality Consulting.

https://facebook.com/story.php?story_fbid=pfbid02pjHjcbUH1uR7Dii19dbtzZSgohQQ6HNMFSfNJNs2b3NbccbyZ7XCyC16X87Vj9QCl&id=106091205021196 '''

contador = 0

for i in range(3, 10):
    numero = str(sheet[f'D{i}'].value)
    if(numero[0] == '9'):
        numero = '+51' + numero
        mensaje = '➡️ Hola, ' + str(sheet[f'B{i}'].value) + '.' + texto
        pywhatkit.sendwhatmsg_instantly( numero,
                    mensaje,
                    15,
                    tab_close=True)
        os.remove('PyWhatKit_DB.txt')
        mensaje = ''
    else:
        contador += 1
        print('Numero no valido ', contador)
    