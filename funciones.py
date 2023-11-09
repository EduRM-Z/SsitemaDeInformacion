#CODIGO CON FUNCIONES QUE SE OCUPAN EN EL PROGRAMA

import re
import PyPDF2
import os

def encontrarNombres(excel):
    '''EXTRAE TODOS LOS NOMBRES DEL EXCEL Y REGRESA UN ARRAY DE NOMBRES'''
    nombres = []
    hoja = excel['NOMBRES']
    for celda in hoja["A"]:
        nombres.append(celda.value)
    return nombres

def encontrarCeldaTitulo(excel, nombre):
    '''ENCUENTRA LOS TITULOS DE CADA TABLA Y REGRESA LA CELDA DONDE INICIA DE ACUERDO AL NOMBRE'''
    hoja = excel['TABLAS']
    for celda in hoja["A"]:
        if isinstance(celda.value, str):
            if nombre in celda.value:
                print(celda.value)
                return celda

def encontrarCeldasTabla(excel, year, celdaTitulo):
    '''ENCUENTRA TODAS LAS FILAS DE CELDAS A MODIFICAR DE ACUERDO AL AÑO Y REGRESA UN ARRAY DE ARRAYS CON CADA FILA DE ESE AÑO'''
    hoja = excel['TABLAS']
    numCelda = int((re.findall(r'\d+', celdaTitulo))[0]) + 3
    filas = []
    while hoja["A" + str(numCelda)].value == int(year):
        filas.append([year,
                      hoja["B" + str(numCelda)].value,
                      str(hoja["C" + str(numCelda)].value) if len(str(hoja["C" + str(numCelda)].value)) == 4 else (((4 - len(str(hoja["C" + str(numCelda)].value))) * "0") + str(hoja["C" + str(numCelda)].value)),
                      str(hoja["D" + str(numCelda)].value) if len(str(hoja["D" + str(numCelda)].value)) == 4 else (((4 - len(str(hoja["D" + str(numCelda)].value))) * "0") + str(hoja["D" + str(numCelda)].value)),
                      hoja["H" + str(numCelda)].value,
                      hoja["I" + str(numCelda)].value,
                      hoja["J" + str(numCelda)].value,
                      hoja["L" + str(numCelda)].coordinate,
                      hoja["M" + str(numCelda)].coordinate])

        numCelda += 1
    return filas, numCelda

def encontrarCeldasTabla_2(excel, year, numCelda):
    '''ENCUENTRA TODAS LAS FILAS DE CELDAS A MODIFICAR DE ACUERDO AL AÑO Y REGRESA UN ARRAY DE ARRAYS CON CADA FILA DE ESE AÑO'''
    hoja = excel['TABLAS']
    filas = []
    while hoja["A" + str(numCelda)].value == int(year):
        filas.append([year,
                      hoja["B" + str(numCelda)].value,
                      str(hoja["C" + str(numCelda)].value) if len(str(hoja["C" + str(numCelda)].value)) == 4 else (((4 - len(str(hoja["C" + str(numCelda)].value))) * "0") + str(hoja["C" + str(numCelda)].value)),
                      str(hoja["D" + str(numCelda)].value) if len(str(hoja["D" + str(numCelda)].value)) == 4 else (((4 - len(str(hoja["D" + str(numCelda)].value))) * "0") + str(hoja["D" + str(numCelda)].value)),
                      hoja["H" + str(numCelda)].value,
                      hoja["I" + str(numCelda)].value,
                      hoja["J" + str(numCelda)].value,
                      hoja["L" + str(numCelda)].coordinate,
                      hoja["M" + str(numCelda)].coordinate])

        numCelda += 1
    return filas

def generar_array_paginas(numero_maximo):
    '''GENERA EL NUMERO DE PAGINAS DONDE SE ENCUENTRA LA INFORMACION A COMPARAR'''
    paginasPDF = [num for num in range(1, numero_maximo + 1, 4)]
    return paginasPDF

def alumnosReprobados(contenido):
    '''OBTIENE LAS VARIABLES DE LOS NO APROBADOS Y LOS NO PRESENTADOS DEL PDF'''
    patronNoAprobados = r'Alumnos No Aprobados:\s*(\d+)'
    patronNoPresentados = r'Alumnos No Presentados:\s*(\d+)'
    noAprobados = (re.search(patronNoAprobados, contenido)).group(1)
    noPresentados = (re.search(patronNoPresentados, contenido)).group(1)
    return int(noAprobados), int(noPresentados)

def guardarPaginas(writer, read_pdf, paginaInicio):
    for i in range(4):
        writer.add_page(read_pdf.pages[(paginaInicio - 1) + i])

def compararConPDF(excel, nombre, filas, pdfs, year, ruta_guardado):
    '''BUSCA EN EL PDF DE ACUERDO AL NUMERO DE SEMESTRE Y COMPARA LA INFORMACION, ESCRIBE EN EL EXCEL Y GUARDA LOS PDFS'''

    read_pdf_1 = PyPDF2.PdfReader(open(pdfs[0], 'rb'))
    writer_1 = PyPDF2.PdfWriter()
    paginasPDF_1 = generar_array_paginas(len(read_pdf_1.pages))

    read_pdf_2 = PyPDF2.PdfReader(open(pdfs[1], 'rb'))
    writer_2 = PyPDF2.PdfWriter()
    paginasPDF_2 = generar_array_paginas(len(read_pdf_2.pages))

    hoja = excel['TABLAS']

    for fila in filas:
        contadorCoincidencia = 0
        if fila[1] == 1:
            for pagina in paginasPDF_1:
                contenido = read_pdf_1.pages[pagina - 1].extract_text()
                if f"Nombre: {nombre}" in contenido and f"Asignatura: {fila[3]}" in contenido and f"Grupo: {fila[2]}" in contenido:
                    hoja[fila[8]].value = pagina
                    noAprobados, noPresentados = alumnosReprobados(contenido)
                    if f"Alumnos Inscritos: {fila[4]}" in contenido:
                        contadorCoincidencia += 1
                    if f"Alumnos Aprobados: {fila[5]}":
                        contadorCoincidencia += 1
                    if (noAprobados + noPresentados) == fila[6]:
                        contadorCoincidencia += 1
                    hoja[fila[7]].value = contadorCoincidencia
                    guardarPaginas(writer_1, read_pdf_1, pagina)
        if fila[1] == 2:
            for pagina in paginasPDF_2:
                contenido = read_pdf_2.pages[pagina - 1].extract_text()
                if f"Nombre: {nombre}" in contenido and f"Asignatura: {fila[3]}" in contenido and f"Grupo: {fila[2]}" in contenido:
                    hoja[fila[8]].value = pagina
                    noAprobados, noPresentados = alumnosReprobados(contenido)
                    if f"Alumnos Inscritos: {fila[4]}" in contenido:
                        contadorCoincidencia += 1
                    if f"Alumnos Aprobados: {fila[5]}":
                        contadorCoincidencia += 1
                    if (noAprobados + noPresentados) == fila[6]:
                        contadorCoincidencia += 1
                    hoja[fila[7]].value = contadorCoincidencia
                    guardarPaginas(writer_2, read_pdf_2, pagina)

    with open(os.path.join(ruta_guardado, year, f'{nombre} {year} - 1.pdf'), 'wb') as output_file:
        writer_1.write(output_file)

    with open(os.path.join(ruta_guardado, year, f'{nombre} {year} - 2.pdf'), 'wb') as output_file:
        writer_2.write(output_file)