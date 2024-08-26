import os
import re
import extract_msg

# Especifica la ruta de la carpeta
ruta_carpeta = r"C:\\Users\\israel.becerra.ortiz\\Downloads\\Mails\\mails"

# Obtén la lista de todos los archivos en la carpeta
ficheros = [f for f in os.listdir(ruta_carpeta) if os.path.isfile(os.path.join(ruta_carpeta, f))]

# Diccionario para contar las repeticiones de archivos con numeración
conteo_repetidos = {}

# Expresión regular para identificar archivos con numeración (1), (2), etc.
patron_repeticion = re.compile(r' \(\d+\)\.msg$')

# Lista final sin archivos duplicados numerados
ficheros_sin_duplicados = []

for fichero in ficheros:
    if patron_repeticion.search(fichero):
        # Remueve la parte de repetición y la extensión para encontrar el nombre base
        nombre_base = re.sub(patron_repeticion, '', fichero) + ".msg"
        
        # Incrementa el conteo de duplicados
        if nombre_base in conteo_repetidos:
            conteo_repetidos[nombre_base] += 1
        else:
            conteo_repetidos[nombre_base] = 1
    else:
        # Añade a la lista sin duplicados si no tiene numeración
        ficheros_sin_duplicados.append(fichero)

# Función para extraer información de un archivo .msg
def extract_body_from_msg(msg_file_path):
    # Cargar el archivo .msg
    msg = extract_msg.Message(msg_file_path)

    # Extraer el cuerpo del mensaje
    body = msg.body
    bodyParts = body.split('\r\n')
    
    # Extraer la fecha y convertirla al formato M/d/YYYY
    mail_date = f"{int(msg.date.strftime('%m'))}/{int(msg.date.strftime('%d'))}/{msg.date.strftime('%Y')}"


    # Extraer el valor de "Flow"
    flow_line = next((line for line in bodyParts if line.startswith('* Flow:')), None)
    flow_value = flow_line.split(':', 1)[1].strip() if flow_line else ''

    # Extraer el valor de "Error"
    error_line = next((line for line in bodyParts if line.startswith('* Error:')), None)
    error_value = error_line.split(':', 1)[1].strip() if error_line else ''

    return mail_date, flow_value, error_value

# Ruta del archivo de salida .csv
output_csv = r"C:\\Users\\israel.becerra.ortiz\\Downloads\\Mails\\salida.csv"

# Procesa cada fichero y escribe en el archivo de salida
with open(output_csv, 'w', encoding='utf-8') as csv_file:
    for fichero in ficheros_sin_duplicados:
        msg_file_path = os.path.join(ruta_carpeta, fichero)
        mail_date, flow_value, error_value = extract_body_from_msg(msg_file_path)
        
        # Solo escribe en el CSV si flow_value y error_value no están vacíos
        if flow_value and error_value:
            # Obtiene el conteo de duplicados para este archivo
            conteo = conteo_repetidos.get(fichero, 0)
            
            # Escribe la información en el archivo CSV
            csv_file.write(f'{mail_date};;;;{flow_value};;{error_value};;;;{conteo + 1}\n')

print(f"Datos guardados en {output_csv}")
