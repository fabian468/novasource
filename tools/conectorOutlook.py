import win32com.client
import os
from modificador_excel2 import crearFiltro
# from tools.modificador_excel2 import crearFiltro
import tkinter as tk 
from tkinter import messagebox
from datetime import datetime
import glob
import pandas as pd


root = tk.Tk()
root.withdraw()

ruta_script = os.path.dirname(os.path.abspath(__file__))
print(ruta_script)

attachment_folder = ruta_script 
os.makedirs(attachment_folder, exist_ok=True)

fecha_actual = datetime.now().strftime("%Y-%m-%d")
ahora =datetime.now()

attachment_folder_fecha= os.path.join(attachment_folder, fecha_actual)
os.makedirs(attachment_folder_fecha, exist_ok=True)

# and message.SenderName=="CDC"
i = 0

dia_hoy =str(ahora.day-2)
mes_hoy =str(ahora.month)
print(dia_hoy)

def unir_excels_en_carpeta(carpeta, nombre_salida="excel_unido.xlsx"):
    archivos = glob.glob(os.path.join(carpeta, "*.xlsx")) + glob.glob(os.path.join(carpeta, "*.xls"))

    if not archivos:
        print(" No se encontraron archivos Excel en la carpeta indicada.")
        return None
    
    dfs = []
    for archivo in archivos:
        try:
            df = pd.read_excel(archivo)
            df["Archivo_Origen"] = os.path.basename(archivo)  # opcional: saber de qué archivo viene cada fila
            dfs.append(df)
        except Exception as e:
            print(f" Error leyendo {archivo}: {e}")

    if not dfs:
        print(" No se pudieron leer los archivos.")
        return None

    # Combinar todos los DataFrames
    df_unido = pd.concat(dfs, ignore_index=True)

    # Guardar el archivo combinado
    salida = os.path.join(carpeta, nombre_salida)
    df_unido.to_excel(salida, index=False)
    print(f"Archivos combinados correctamente en: {salida}")

    return salida


def buscar_correo():
    contador_de_mensajes=0
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #inbox = outlook.GetDefaultFolder(6)

    inbox = outlook.Folders('crocc.operator').Folders('Inbox')
    messages = inbox.Items

    print("Empezando analisis...")
    for message in messages:
        # print("buscando")
        # i += 1
        # if(i == 10):
        #     break
        

        try:        
            received_time_pywintype = message.ReceivedTime
            # received_time_datetime = datetime.datetime(
            #         received_time_pywintype.year,
            #         received_time_pywintype.month,
            #         received_time_pywintype.day,
            #     )
            # print(received_time_datetime)            
            if(received_time_pywintype.day == dia_hoy ):
                break

            asuntos=["[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico",
                     "[EXT]: Ajuste Prorrata Generalizada costo SEN 0",
                     "[EXT]: Finaliza Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico",
                     "[EXT]: Ajuste Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico"
                     ]

            if((message.subject in asuntos)and message.SenderName=="CDC" ):
                contador_de_mensajes+=1
                print(message.subject,message.ReceivedTime)
                # print(message.subject,message.subject ==("[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico" or message.subject == "[EXT]: Ajuste Prorrata Generalizada costo SEN 0"))
                # if(i == 2):
                #     break
                # i += 1
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        full_path = os.path.join(attachment_folder_fecha, attachment.FileName)
                        
                        attachment.SaveAsFile(full_path)
                        print(f"Downloaded: {attachment.FileName}")

            path_full = unir_excels_en_carpeta(full_path, nombre_salida="excel_unido.xlsx")
            crearFiltro(attachment_folder_fecha, carpeta_donde_guardar=attachment_folder_fecha)
        
        except AttributeError:
            continue


    messagebox.showinfo("Información", f"Se encontraron {contador_de_mensajes} mensajes.")


