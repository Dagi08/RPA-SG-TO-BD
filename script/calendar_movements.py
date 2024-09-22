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

pd.set_option('display.max_columns', None)

# Ruta del directorio que contiene el archivo de Excel
directorio = "D:/BN/auxiliar_mejora_sd/prueba/"

# Obtén la lista de archivos en el directorio
archivos = os.listdir(directorio)

# Itera sobre los archivos en el directorio
for archivo in archivos:
    if archivo.endswith(".txt"):
        # Intenta leer el archivo de texto en un DataFrame
        try:
            ruta_completa = os.path.join(directorio, archivo)
            
            # Leer el archivo y procesar las líneas
            data = []
            with open(ruta_completa, 'r') as file:
                for line in file:
                    # Separar los valores de cada línea por el delimitador '|'
                    values = line.strip().split('|')
                    # Verificar si el primer valor es un número
                    if re.match(r'^\d+$', values[0]):
                        data.append(values)

            # Crear un DataFrame con los datos filtrados
            df = pd.DataFrame(data, columns=['Column1', 'Column2', 'Column3'])

            # Mostrar el DataFrame
            print(df)
        except Exception as e:
            # Maneja cualquier excepción que ocurra al leer el archivo
            print(f"No se pudo leer el archivo {archivo}: {e}")

# Título de la ventana que deseas activar
titulo_ventana_calendario = "Prod02 - B - BN326053" #prah6207
titulo_ventana_movimiento = "Prod01 - A - BN32K190" #prah6217

def extraer_datos_ventana_calendario(numero):
    dataPagares = []
    cuotas = []
    desembolso = 0.0
    saldo = 0.0
    fallecido = False
    try:
        ventana_calendario = gw.getWindowsWithTitle(titulo_ventana_calendario)
        if len(ventana_calendario) > 0:
            ventana_calendario[0].activate()
            time.sleep(1)
            pyautogui.press('F1')
            pyautogui.typewrite(numero)
            pyautogui.press('enter')
            pyautogui.press('up', presses=11)
            pyautogui.press('left', presses=4)
            pyautogui.hotkey('ctrl', 'c')
            texto_pantalla = pyperclip.paste()
            desembolso_match = re.search(r'Desembolsado:\s+([\d,]+.\d+)', texto_pantalla)
            recuperado_match = re.search(r'Recuperado\s+:\s+([\d,]+.\d+)', texto_pantalla)
            saldo_match = re.search(r'Saldo\s+:\s+([\d,]+.\d+)', texto_pantalla)
            moneda = 1 if 'NUEVO SOL PERUANO' in texto_pantalla else 0

            desembolso = float(desembolso_match.group(1).replace(',', '')) if desembolso_match else 0.0
            recuperado = float(recuperado_match.group(1).replace(',', '')) if recuperado_match else 0.0
            saldo = float(saldo_match.group(1).replace(',', '')) if saldo_match else 0.0

            dataPagare = {
                "numero": numero,
                "id_moneda": moneda,
                "importe_desembolso": desembolso,
                "importe_recuperado": recuperado,
                "importe_saldo": saldo
            }

            dataPagares.append(dataPagare)

            pasar_mvto = True
            detener = False  # Variable para detener el bucle
            
            while pasar_mvto and not detener:
                lineas = texto_pantalla.splitlines()
                
                 # Verificación adicional para evitar el error 'list index out of range'
                if len(lineas) < 20:
                    print("Advertencia: Menos de 20 líneas encontradas en texto_pantalla")
                    pasar_mvto = False
                    continue
                    
                for i in range(10, 20):
                    if i >= len(lineas):
                        print(f"Advertencia: Intentando acceder a índice {i} en lineas con longitud {len(lineas)}")
                        pasar_mvto = False
                        break
                    
                    linea = lineas[i]
                    datos = linea.split()
                    if len(datos) > 6 and datos[6] in ["FAL"]:
                        fallecido = True
                        #pasar_mvto = False
                        detener = True  # Detener el bucle
                        break
                    #elif len(datos) >= 7:
                    if len(datos) >= 7:
                        cuota = {
                            "fecha_vencimiento": datos[1],
                            "id_pagare": numero,
                            "importe": float(datos[5].replace(',', '')),
                            "importe_amortizacion": float(datos[2].replace(',', '')),
                            "importe_interes": float(datos[3].replace(',', '')),
                            "importe_seguro_desgravem": float(datos[4].replace(',', '')),
                            "numero": datos[0],
                            "situacion": datos[6]
                        }
                        cuotas.append(cuota)
                        
                if not detener:  # Solo continuar si no se ha detenido
                    time.sleep(1)
                    pyautogui.press('F8')
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'c')
                    nuevo_texto_pantalla = pyperclip.paste()

                    if texto_pantalla == nuevo_texto_pantalla:
                        pasar_mvto = False
                    else:
                        texto_pantalla = nuevo_texto_pantalla
        else:
            print(f"No se encontró una ventana con el título '{titulo_ventana_calendario}'")
    except Exception as e:
        print(f"Error en extraer_datos_ventana_calendario: {e}")
    
    return dataPagares, cuotas, desembolso, saldo, fallecido

def extraer_datos_ventana_movimiento(numero):
    movimientos = []
    try:
        ventana_movimiento = gw.getWindowsWithTitle(titulo_ventana_movimiento)
        if len(ventana_movimiento) > 0:
            ventana_movimiento[0].activate()
            time.sleep(1)
            pyautogui.press('F1')
            pyautogui.typewrite(numero)
            pyautogui.press('enter')
            pyautogui.press('up', presses=13)
            pyautogui.press('left', presses=4)
            pyautogui.hotkey('ctrl', 'c')
            texto_pantalla_2 = pyperclip.paste()
            nuevo_texto_pantalla_2 = ""
            
            pasar_mvto = True
            while pasar_mvto:
                lineas_2 = texto_pantalla_2.splitlines()
                coger_fila_siguiente = False

                for i in range(8, 20):
                    linea_2 = lineas_2[i]
                    datos_2 = linea_2.split()
                    fecha_mvto = linea_2[:11].strip()
                    if "DESEMBOLSO" in linea_2:
                        movimiento = {
                            "concepto_operacion": '',
                            "dias_moratorios": 0,
                            "dias_vencidos": 0,
                            "fecha": datos_2[0],
                            "id_pagare": numero,
                            "importe_amortizacion": float(datos_2[2].replace(',', '')),
                            "importe_cuota": 0.0,
                            "importe_interes_moratorio": 0.0,
                            "importe_interes_Vencido_Corridos": 0.0,
                            "importe_seguro_desgravem": 0.0,
                            "importe_saldo_pendiente": float(datos_2[3].replace(',', '')),
                            "numero_cuota": '',
                            "numero_operacion": '',
                            "tipo_operacion": datos_2[1],
                        }
                        movimientos.append(movimiento)
                    else:
                        if fecha_mvto != "":
                            linea_x = linea_2
                            coger_fila_siguiente = True
                        elif fecha_mvto == "" and coger_fila_siguiente:
                            linea_y = linea_2
                            linea_xy = f"{linea_x} {linea_y}"
                            datos_2 = linea_xy.split()
                            if datos_2[1] in ["CUOTA", "ADEL.", "AMPL."]:
                                if "CARGO EN AHORROS AL " in linea_xy:
                                    movimiento = {
                                        "concepto_operacion": ' '.join(datos_2[10:15]),
                                        "dias_moratorios": int(datos_2[15].lstrip('0') or '0'),
                                        "dias_vencidos": int(datos_2[6].lstrip('0') or '0'),
                                        "fecha": datos_2[0],
                                        "id_pagare": numero,
                                        "importe_amortizacion": float(datos_2[3].replace(',', '')),
                                        "importe_cuota": float(datos_2[2].replace(',', '')),
                                        "importe_interes_moratorio": float(datos_2[16].replace(',', '')),
                                        "importe_interes_Vencido_Corridos": float(datos_2[7].replace(',', '')),
                                        "importe_seguro_desgravem": float(datos_2[5].replace(',', '')),
                                        "importe_saldo_pendiente": float(datos_2[4].replace(',', '')),
                                        "numero_cuota": datos_2[8],
                                        "numero_operacion": datos_2[9],
                                        "tipo_operacion": datos_2[1],
                                    }
                                    movimientos.append(movimiento)
                                elif "CARGO VARIAS CUENTAS " in linea_xy or "PAGO ADELANTO DE CUOTA" in linea_xy:
                                    movimiento = {
                                        "concepto_operacion": ' '.join(datos_2[10:14]),
                                        "dias_moratorios": int(datos_2[14].lstrip('0') or '0'),
                                        "dias_vencidos": int(datos_2[6].lstrip('0') or '0'),
                                        "fecha": datos_2[0],
                                        "id_pagare": numero,
                                        "importe_amortizacion": float(datos_2[3].replace(',', '')),
                                        "importe_cuota": float(datos_2[2].replace(',', '')),
                                        "importe_interes_moratorio": float(datos_2[15].replace(',', '')),
                                        "importe_interes_Vencido_Corridos": float(datos_2[7].replace(',', '')),
                                        "importe_seguro_desgravem": float(datos_2[5].replace(',', '')),
                                        "importe_saldo_pendiente": float(datos_2[4].replace(',', '')),
                                        "numero_cuota": datos_2[8],
                                        "numero_operacion": datos_2[9],
                                        "tipo_operacion": datos_2[1],
                                    }
                                    movimientos.append(movimiento)
                                elif "PAGO EN LINEA POR EL " in linea_xy:
                                    movimiento = {
                                        "concepto_operacion": ' '.join(datos_2[10:16]),
                                        "dias_moratorios": int(datos_2[16].lstrip('0') or '0'),
                                        "dias_vencidos": int(datos_2[6].lstrip('0') or '0'),
                                        "fecha": datos_2[0],
                                        "id_pagare": numero,
                                        "importe_amortizacion": float(datos_2[3].replace(',', '')),
                                        "importe_cuota": float(datos_2[2].replace(',', '')),
                                        "importe_interes_moratorio": float(datos_2[17].replace(',', '')),
                                        "importe_interes_Vencido_Corridos": float(datos_2[7].replace(',', '')),
                                        "importe_seguro_desgravem": float(datos_2[5].replace(',', '')),
                                        "importe_saldo_pendiente": float(datos_2[4].replace(',', '')),
                                        "numero_cuota": datos_2[8],
                                        "numero_operacion": datos_2[9],
                                        "tipo_operacion": datos_2[1],
                                    }
                                    movimientos.append(movimiento)
                            elif datos_2[1] in ["LIQUID."]:
                                if "PAGO EN LINEA POR EL " in linea_xy:
                                    movimiento = {
                                        "concepto_operacion": ' '.join(datos_2[10:16]),
                                        "dias_moratorios": int(datos_2[16].lstrip('0') or '0'),
                                        "dias_vencidos": int(datos_2[6].lstrip('0') or '0'),
                                        "fecha": datos_2[0],
                                        "id_pagare": numero,
                                        "importe_amortizacion": float(datos_2[3].replace(',', '')),
                                        "importe_cuota": 0.0,
                                        "importe_interes_moratorio": float(datos_2[17].replace(',', '')),
                                        "importe_interes_Vencido_Corridos": float(datos_2[7].replace(',', '')),
                                        "importe_seguro_desgravem": float(datos_2[5].replace(',', '')),
                                        "importe_saldo_pendiente": float(datos_2[4].replace(',', '')),
                                        "numero_cuota": datos_2[8],
                                        "numero_operacion": datos_2[9],
                                        "tipo_operacion": ' '.join(datos_2[1:3]),
                                    }
                                    movimientos.append(movimiento)
                            elif datos_2[1] in ["PREPAGO"]:
                                if "PREPAGO EN LINEA POR EL " in linea_xy:
                                    movimiento = {
                                        "concepto_operacion": ' '.join(datos_2[8:14]),
                                        "dias_moratorios": 0,
                                        "dias_vencidos": int(datos_2[5].lstrip('0') or '0'),
                                        "fecha": datos_2[0],
                                        "id_pagare": numero,
                                        "importe_amortizacion": float(datos_2[2].replace(',', '')),
                                        "importe_cuota": 0.0,
                                        "importe_interes_moratorio": 0.0,
                                        "importe_interes_Vencido_Corridos": float(datos_2[6].replace(',', '')),
                                        "importe_seguro_desgravem": float(datos_2[4].replace(',', '')),
                                        "importe_saldo_pendiente": float(datos_2[3].replace(',', '')),
                                        "numero_cuota": '',
                                        "numero_operacion": datos_2[7],
                                        "tipo_operacion": datos_2[1],
                                    }
                                    movimientos.append(movimiento)
                            elif datos_2[1] in ["REPROGRAMACION"]:
                                if "REPROGRAMACION ZONA EMERGENCIA" in linea_xy:
                                    movimiento = {
                                        "concepto_operacion": ' '.join(datos_2[7:10]),
                                        "dias_moratorios": 0,
                                        "dias_vencidos": 0,
                                        "fecha": datos_2[0],
                                        "id_pagare": numero,
                                        "importe_amortizacion": float(datos_2[2].replace(',', '')),
                                        "importe_cuota": 0.0,
                                        "importe_interes_moratorio": 0.0,
                                        "importe_interes_Vencido_Corridos": float(datos_2[5].replace(',', '')),
                                        "importe_seguro_desgravem": float(datos_2[4].replace(',', '')),
                                        "importe_saldo_pendiente": float(datos_2[3].replace(',', '')),
                                        "numero_cuota": '',
                                        "numero_operacion": datos_2[6],
                                        "tipo_operacion": datos_2[1],
                                    }
                                    movimientos.append(movimiento)       
                            
                            coger_fila_siguiente = False

                time.sleep(1)
                pyautogui.press('F8')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'c')
                nuevo_texto_pantalla_2 = pyperclip.paste()

                if texto_pantalla_2 == nuevo_texto_pantalla_2:
                    pasar_mvto = False
                else:
                    texto_pantalla_2 = nuevo_texto_pantalla_2
        else:
            print(f"No se encontró una ventana con el título '{titulo_ventana_movimiento}'")
    except Exception as e:
        print(f"Error en extraer_datos_ventana_movimiento: {e}")
        print(linea_xy)
    
    return movimientos

def ejecutar_proceso(numero, cursor):
    dataPagares, cuotas, desembolso, saldo, fallecido = extraer_datos_ventana_calendario(numero)
    
    data_correcta = []
    data_incorrecta = []
    
    if fallecido == False:
        movimientos = extraer_datos_ventana_movimiento(numero)

        total_amortizado = sum(movimiento['importe_amortizacion'] for movimiento in movimientos if movimiento['tipo_operacion'] in ["CUOTA", "ADEL."])
        print(f"Total Desembolso: {round(float(desembolso),2)}")
        print(f"Total Amortizado: {round(float(total_amortizado),2)}")
        print(f"Total Saldo: {round(float(saldo),2)}")

        if round(float(desembolso),2) == round(float(saldo) + float(total_amortizado),2):

            try:
                #creamos con python la tabla temporal de datos para cuotas y pagares
                cursor.execute("CREATE TABLE #dataCuotas (fecha_vencimiento VARCHAR(50), id_pagare VARCHAR(50), importe DECIMAL(18, 2), importe_amortizacion DECIMAL(18, 2), importe_interes DECIMAL(18, 2), importe_seguro_desgravem DECIMAL(18, 2), numero VARCHAR(50), situacion VARCHAR(50))")
                cursor.execute("CREATE TABLE #dataMovimientos (concepto_operacion VARCHAR(255), dias_moratorios INT, dias_vencidos INT, fecha VARCHAR(50), id_pagare VARCHAR(50), importe_amortizacion DECIMAL(18, 2), importe_cuota DECIMAL(18, 2), importe_interes_moratorio DECIMAL(18, 2), importe_interes_Vencido_Corridos DECIMAL(18, 2), importe_seguro_desgravem DECIMAL(18, 2), importe_saldo_pendiente DECIMAL(18, 2), numero_cuota VARCHAR(50), numero_operacion VARCHAR(50), tipo_operacion VARCHAR(50))")

                # Llenar las tablas temporales
                cursor.executemany("INSERT INTO #dataCuotas VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                   [(cuota['fecha_vencimiento'], cuota['id_pagare'], float(cuota['importe']), float(cuota['importe_amortizacion']), float(cuota['importe_interes']), float(cuota['importe_seguro_desgravem']), cuota['numero'], cuota['situacion']) for cuota in cuotas])
                cursor.executemany("INSERT INTO #dataMovimientos VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            [(movimiento['concepto_operacion'], int(movimiento['dias_moratorios']), int(movimiento['dias_vencidos']), movimiento['fecha'], movimiento['id_pagare'], float(movimiento['importe_amortizacion']), float(movimiento['importe_cuota']), float(movimiento['importe_interes_moratorio']), float(movimiento['importe_interes_Vencido_Corridos']), float(movimiento['importe_seguro_desgravem']), float(movimiento['importe_saldo_pendiente']), movimiento['numero_cuota'], movimiento['numero_operacion'], movimiento['tipo_operacion']) for movimiento in movimientos])

                for dataPagare in dataPagares:
                    # Preparar la llamada al procedimiento almacenado con el parámetro de salida
                    sql = """\
                    DECLARE @resultado INT;
                    EXEC sp_rpa_Calendario_Movimiento @numero_pagare = ?, @id_moneda = ?, @importe_desembolso = ?, @importe_recuperado = ?, @importe_saldo = ?, @resultado = @resultado OUTPUT;
                    SELECT @resultado;
                    """
                    params = (dataPagare['numero'], int(dataPagare['id_moneda']), float(dataPagare['importe_desembolso']), float(dataPagare['importe_recuperado']), float(dataPagare['importe_saldo']))

                    cursor.execute(sql, params)
                    result = cursor.fetchone()[0]
                    print(f"Datos enviados a sql pagare número: {numero}")
                                             
                if result == 1:
                    data_correcta.append({
                        'numero': numero,
                        'desembolso': desembolso,
                        'importe_amortizado': total_amortizado,
                        'saldo': saldo
                    })
                else:
                    data_incorrecta.append({
                        'numero': numero,
                        'desembolso': desembolso,
                        'importe_amortizado': total_amortizado,
                        'saldo': saldo,
                        'observacion': 'No se actualiza pagare, calendario de cuotas y movimientos'
                    })

                # Confirmar la transacción
                cursor.commit()
            except Exception as e:
                # Manejar cualquier error y revertir la transacción si es necesario
                cursor.rollback()
                data_incorrecta.append({
                    'numero': numero,
                    'desembolso': desembolso,
                    'importe_amortizado': total_amortizado,
                    'saldo': saldo,
                    'observacion': f'Error: {e}'
                })
        else:
            data_incorrecta.append({
                'numero': numero,
                'desembolso': desembolso,
                'importe_amortizado': total_amortizado,
                'saldo': saldo,
                'observacion': 'Error en Conciliacion'
            })

            print(f"Datos incorrectos para el número: {numero}")
            notification.notify(
                title='Datos Incorrectos',
                message=f'Datos incorrectos para: {numero}',
                timeout=10
            )
        
    else:
            data_incorrecta.append({
                'numero': numero,
                'desembolso': desembolso,
                'importe_amortizado': 0,
                'saldo': saldo,
                'observacion': 'Pagaré en estado fallecido'
            })

            try:
                
                query = """
                    UPDATE pagares 
                    SET id_estado_pagare = ?
                    WHERE numero = ?
                """
                cursor.execute(query, (5, numero))

                cursor.commit()
                
                print(f"Datos fallecidos para el número: {numero}")
                notification.notify(
                    title='Datos Incorrectos',
                    message=f'Se encuentra fallecido: {numero}',
                    timeout=10
                )
            except Exception as e:
                cursor.rollback()
                print(f"Ocurrio un error: {e}")
    
    return data_correcta, data_incorrecta

def generar_reporte(data_correcta, data_incorrecta, file_name):
    with open(file_name, 'w') as file:
        fecha_actual = datetime.datetime.now().strftime("%d/%m/%Y")
        file.write(f"Reporte de salida - informe RPA - Fecha {fecha_actual}\n")
        file.write("-------------------------------------------------------------\n")
        file.write("| Número Pagare   |  Desembolso   |   Importe Amortizado  |  Importe Saldo |  Observacion |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        for item in data_correcta:
            file.write(f"| {item['numero']} | {item['desembolso']} | {item['importe_amortizado']} | {item['saldo']} | Insert SQL Ok |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        file.write("\n")
        file.write("TX no se enviaron a SQL\n")
        file.write("------------------------------\n")
        file.write("| Número Pagare   |  Desembolso |  Importe Amortizado | Importe Saldo |  Observacion |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")
        for item in data_incorrecta:
            file.write(f"| {item['numero']} | {item['desembolso']} | {item['importe_amortizado']} | {item['saldo']} | {item['observacion']} |\n")
        file.write("----------------------------------------------------------------------------------------------------------\n")

def main(df):
    connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.22.77.35,1433;DATABASE=BN_BD_SoporteSD;UID=sa;PWD=QWEzxc123.'
    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    data_correcta_total = []
    data_incorrecta_total = []

    try:
        for index, row in df.iterrows():
            numero = row['Column2']
            print("------------------------------------------------------------------------")
            print(f"Pagaré: {numero}")
            print("------------------------------------------------------------------------")
            data_correcta, data_incorrecta = ejecutar_proceso(numero, cursor)
            data_correcta_total.extend(data_correcta)
            data_incorrecta_total.extend(data_incorrecta)

        generar_reporte(data_correcta_total, data_incorrecta_total, "D:\\BN\\auxiliar_mejora_sd\\Informe-RPA-Extraccion-Calendario-Movimiento-" + datetime.datetime.now().strftime("%d-%m-%Y") + ".txt")
    except Exception as e:
        print(f"Error en el proceso principal: {e}")
    finally:
        cursor.close()
        conn.close()

#ejecutar
main(df)
#try:
    # Aquí debes definir tu DataFrame df
 #   for index, row in df.iterrows():
  #      numero = row['Column2']
   #     ejecutar_proceso(numero)
#except Exception as e:
 #   print(f"Error en el proceso principal: {e}")