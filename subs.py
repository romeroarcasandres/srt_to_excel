import re
from openpyxl import Workbook

# Función para procesar el contenido del archivo srt
def procesar_srt(contenido):
    pat = re.compile(r'(\d+)\s+(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})\n(.*?)(?=\n\n\d+|\Z)', re.DOTALL)
    coincidencias = pat.findall(contenido)
    datos = []
    for match in coincidencias:
        _, tiempo_inicio, tiempo_fin, texto = match
        datos.append({
            'Tiempo de entrada': tiempo_inicio,
            'Tiempo de salida': tiempo_fin,
            'Texto de subtítulo': texto.strip().replace('\n', '\n')  # Sustituir \n por salto de línea manual
        })
    return datos

# Función para escribir datos en un archivo Excel
def escribir_excel(datos):
    libro = Workbook()
    hoja = libro.active

    # Escribir encabezados
    encabezados = ['Tiempo de entrada', 'Tiempo de salida', 'Texto de subtítulo']
    hoja.append(encabezados)

    # Escribir datos en el archivo Excel
    for dato in datos:
        fila = [dato[encabezado] for encabezado in encabezados]
        hoja.append(fila)

    # Guardar el archivo Excel
    libro.save('subtitulos.xlsx')

# Leer el contenido del archivo srt
with open('subtitulos.srt', 'r', encoding='utf-8') as archivo_srt:
    contenido_srt = archivo_srt.read()

# Procesar el contenido del archivo srt
datos_subtitulos = procesar_srt(contenido_srt)

# Escribir datos en el archivo Excel
escribir_excel(datos_subtitulos)
