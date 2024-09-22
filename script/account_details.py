import pandas as pd
import os
import re
import pyautogui
import pygetwindow as gw
import time
import pyperclip
from plyer import notification
import pyodbc
import datetime

def extraer_data_ventana_detalles_cuenta(ventana_detalles_cuenta_titulo, numero):

    data_extraida = []
    error_message = ''
    cuenta = ''
    desembolso = ''
    
    try:
        ventanas = gw.getWindowsWithTitle(ventana_detalles_cuenta_titulo)
        if ventanas:
            ventana_detalles_cuenta = ventanas[0]
            try:
                ventana_detalles_cuenta.activate()
            except Exception as e:
                print(f"Error al activar la ventana: {e}")
                # Fallback: intenta mover el mouse y hacer clic en la ventana para activarla
                x, y, width, height = ventana_detalles_cuenta.left, ventana_detalles_cuenta.top, ventana_detalles_cuenta.width, ventana_detalles_cuenta.height
                pyautogui.click(x + width // 2, y + height // 2)
                time.sleep(1)  # Espera un segundo para asegurarse de que la ventana está activada
            
            cuenta = numero[:11]
            desembolso = numero[-2:]
            
            time.sleep(1)
            pyautogui.press('F1')
            pyautogui.typewrite(cuenta)
            pyautogui.press('enter')
            pyautogui.hotkey('ctrl', 'c')
            texto_pantalla = pyperclip.paste()
            
            desembolso_match = re.search(r'PRESTAM.'+ desembolso , texto_pantalla)
            if desembolso_match:
                lineas = texto_pantalla.splitlines()
                
                # Extracción de datos
                data = {
                    "id_cuenta": cuenta,
                    "numero_desembolso": desembolso,
                    "lugar_emision": lineas[3].split(":")[-1].strip(),
                    "nombre_cliente": lineas[6][:55].strip(),
                    "monto_utilizado": float(lineas[6].split(":")[-1].strip().replace(',', '')),
                    "monto_recuperado": float(lineas[7].split(":")[-1].strip().replace(',', '')),
                    "nombre_sector": lineas[8][13:55].strip(),
                    "doc_cliente": lineas[9][13:30].strip(),
                    "fecha_apertura": lineas[9][42:55].strip(),
                    "cod_cliente": lineas[10][13:30].strip(),
                    "fecha_cierre": lineas[10][42:55].strip(),
                    "fecha_dsb": lineas[11][13:30].strip(),
                    "u_ejecutora": lineas[12][42:55].strip(),
                    "tipo_deudor": lineas[14].split(":")[-1].strip(),
                    "nom_garante": lineas[15][:55].strip(),
                    "magnitud": lineas[15].split(":")[-1].strip(),
                    "tipo_cred": lineas[16].split(":")[-1].strip(),
                    "relacion_BN": lineas[17].split(":")[-1].strip(),
                    "doc_garante": lineas[17][13:30].strip(),
                    "cod_garante": lineas[18][13:30].strip(),
                    "cuenta_garante": lineas[18][42:55].strip(),
                }
                
                data_extraida.append(data)
                print(data_extraida)
            else:
                print("No existe desembolso")
    except Exception as e:
        print(f"Error en extraer datos de la ventana: {e}")
        error_message = 'Último desembolso difiere en Host'
    
    return data_extraida, error_message, cuenta, desembolso

def ejecutar_proceso(titulo ,numero, cursor):
    
    data_correcta = []
    data_incorrecta = []
    data, error_message, cuenta, desembolso = extraer_data_ventana_detalles_cuenta(titulo, numero)
    
    try:
        if not data:
            data_incorrecta.append({
                'numero': cuenta,
                'desembolso': desembolso,
                'observacion': error_message
            })
        else:
            data = data[0]
            # Preparar la llamada al procedimiento almacenado con el parámetro de salida
            sql = """\
            DECLARE @resultado INT;
            EXEC sp_rpa_Actualizar_Detalle_Cuenta @id_cuenta=?, @numero_desembolso=?, @lugar_emision=?, @nombre_cliente=?, 
                                            @monto_utilizado=?, @monto_recuperado=?, @nombre_sector=?, @doc_cliente=?, 
                                            @fecha_apertura=?, @cod_cliente=?, @fecha_cierre=?, @fecha_dsb=?, 
                                            @u_ejecutora=?, @tipo_deudor=?, @nom_garante=?, @magnitud=?, 
                                            @tipo_cred=?, @relacion_BN=?, @doc_garante=?, @cod_garante=?, 
                                            @cuenta_garante=?, @resultado = @resultado OUTPUT;
            SELECT @resultado;
            """
            params = (data['id_cuenta'], data['numero_desembolso'], data['lugar_emision'], data['nombre_cliente'],
            float(data['monto_utilizado']), float(data['monto_recuperado']), data['nombre_sector'], data['doc_cliente'],
            data['fecha_apertura'], data['cod_cliente'], data['fecha_cierre'], data['fecha_dsb'],
            data['u_ejecutora'], data['tipo_deudor'], data['nom_garante'], data['magnitud'],
            data['tipo_cred'], data['relacion_BN'], data['doc_garante'], data['cod_garante'],
            data['cuenta_garante'])

            cursor.execute(sql, params)
            result = cursor.fetchone()[0]
            if result == 1:
                data_correcta.append({
                    'numero': cuenta,
                    'desembolso': desembolso,
                    'observacion': 'Insert SQL OK'
                })
            else:
                data_incorrecta.append({
                    'numero': cuenta,
                    'desembolso': desembolso,
                    'observacion': 'No cumple las condiciones: \n 1. Cuenta diferente \n 2. Revisar el estado del cliente'
                })
            cursor.commit()
            
    except Exception as e:
        # Manejar cualquier error y revertir la transacción si es necesario
        cursor.rollback()
        data_incorrecta.append({
            'numero': cuenta,
            'desembolso': desembolso,
            'observacion': f'Error: {e}'
            })
    return data_correcta, data_incorrecta

def generar_reporte(data_correcta, data_incorrecta, filename):
    
    with open(filename, 'w') as file:
        fecha_actual = datetime.datetime.now().strftime("%d/%m/%Y")
        file.write(f"Reporte de salida - informe de Insert - Fecha {fecha_actual}\n")
        file.write("-------------------------------------------------------------\n")
        file.write("| Número Cuenta   |  Desembolso |  Observacion |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        for item in data_correcta:
            file.write(f"| {item['numero']} | {item['desembolso']} | {item['observacion']} |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        file.write("\n")
        file.write("TX no se enviaron a SQL\n")
        file.write("------------------------------\n")
        file.write("| Número Cuenta   |  Desembolso |  Observacion |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        for item in data_incorrecta:
            file.write(f"| {item['numero']} | {item['desembolso']} | {item['observacion']} |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        
def main():
    
    # Ruta del archivo
    ruta_archivo = r'D:\BN\auxiliar_mejora_sd\Informe-RPA-Extraccion-Calendario-Movimiento-' + datetime.datetime.now().strftime("%d-%m-%Y") + '.txt'
    ventana_detalles_cuenta_titulo = "Prod01 - B - BN328071"
    
    # Leer el archivo y cargar su contenido en un DataFrame
    df_data_insert_sql = pd.read_csv(ruta_archivo, sep='|', skiprows=2, skipfooter=1, engine='python', encoding='ansi', header=0)
    
    # Renombrar columnas para eliminar espacios en blanco
    df_data_insert_sql.columns = df_data_insert_sql.columns.str.strip()
    
    # Filtrar las filas que contienen 'Insert SQL Ok' en la columna 'Observacion'
    df_filtered = df_data_insert_sql[df_data_insert_sql['Observacion'].str.contains('Insert SQL Ok', na=False)]
    
    # Eliminar las columnas en el índice 0 y 6
    df_filtered = df_filtered.drop([df_filtered.columns[0], df_filtered.columns[6]], axis=1)
    
    print(df_filtered)
    
    connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.22.77.35,1433;DATABASE=BN_BD_SoporteSD;UID=sa;PWD=QWEzxc123.'
    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()
    
    data_correcta_total=[]
    data_incorrecta_total=[]
    
    try:
        for index, row in df_filtered.iterrows():
            numero = row['Número Pagare'].strip()
            
            data_correcta, data_incorrecta = ejecutar_proceso(ventana_detalles_cuenta_titulo, numero, cursor)
            data_correcta_total.extend(data_correcta)
            data_incorrecta_total.extend(data_incorrecta)
        
        generar_reporte(data_correcta_total, data_incorrecta_total, "D:\\BN\\auxiliar_mejora_sd\\Informe-Insert-Detalle-Cuenta-" + datetime.datetime.now().strftime("%d-%m-%Y") + ".txt")
                
    except Exception as e:
        print(f"Error en el proceso principal: {e}")
    finally:
        cursor.close()
        conn.close()

#EJECUTAR PROGRAMA
main()
