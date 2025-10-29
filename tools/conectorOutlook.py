import win32com.client
import os
from openpyxl import load_workbook
from tools.modificador_excel2 import crearFiltro
import tkinter as tk 
from tkinter import messagebox, ttk
from datetime import datetime
from tools.unirExcel import unir_excels_en_carpeta, eliminar_archivo_unido
from tkinter import filedialog
import pandas as pd
import time
from pages.ventanaDeProgreso import crear_ventana_progreso, actualizar_progreso

root = tk.Tk()
root.withdraw()

ruta_script = os.path.dirname(os.path.abspath(__file__))

attachment_folder = ruta_script 
os.makedirs(attachment_folder, exist_ok=True)

fecha_actual = datetime.now().strftime("%Y-%m-%d")
ahora = datetime.now()
attachment_folder_fecha = filedialog.askdirectory(title="Selecciona la carpeta donde guardar el archivo")
os.makedirs(attachment_folder_fecha, exist_ok=True)

i = 0
dia_hoy = ahora.day-1
mes_hoy = ahora.month


def buscar_correo():
    # Crear ventana de progreso
    ventana_prog, progress_bar, label_estado, label_detalle = crear_ventana_progreso()
    
    try:
        contador_de_mensajes = 0
        
        actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                           10, "Conectando con Outlook...")
        
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.Folders('crocc.operator').Folders('Inbox')
        # inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        
        total_mensajes = sum(1 for msg in messages 
                            if hasattr(msg, 'ReceivedTime') and 
                            msg.ReceivedTime.day == dia_hoy and 
                            msg.ReceivedTime.month == mes_hoy)
        
        actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                           20, "Analizando correos...", 
                           f"Total de correos del día: {total_mensajes}")
        
        archivos_descargados = []

        asuntos = [
            "[ext]: inicia prorrata lt 500 kv nueva pan de azúcar - polpaico",
            "[ext]: ajuste prorrata generalizada costo sen 0",
            "[ext]: finaliza prorrata lt 500 kv nueva pan de azúcar - polpaico"
        ]
        
        mensajes_procesados = 0
        
        for message in messages:
            try:
                if message.ReceivedTime.day == dia_hoy and message.ReceivedTime.month == mes_hoy:
                    mensajes_procesados += 1
                    progreso = 20 + int((mensajes_procesados / max(total_mensajes, 1)) * 50)
                    
                    actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                                       progreso, "Procesando correos...",
                                       f"Revisando: {message.Subject[:50]}...")
                    
                    if any(a in message.Subject.lower() for a in asuntos):
                        contador_de_mensajes += 1
                        
                    
                        for attachment in message.Attachments:
                            full_path = os.path.join(attachment_folder_fecha, attachment.FileName)
                            attachment.SaveAsFile(full_path)
                            archivos_descargados.append(full_path)
                            
                            actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                                               progreso, "Descargando archivos...",
                                               f"Guardado: {attachment.FileName}")
                            
            except AttributeError:
                continue
        
        if archivos_descargados:
            actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                               75, "Uniendo archivos Excel...",
                               f"Total archivos: {len(archivos_descargados)}")
            
            path_full = unir_excels_en_carpeta(attachment_folder_fecha, nombre_salida="excel_unido.xlsx")
            
            if path_full:
                actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                                   85, "Aplicando filtros...", "")
                crearFiltro(path_full, carpeta_donde_guardar=attachment_folder_fecha)
                
                actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                                   95, "Finalizando...", "")
        
        else:
            print("No se descargaron archivos adjuntos.")
        
        actualizar_progreso(ventana_prog, progress_bar, label_estado, label_detalle,
                           100, "¡Proceso completado!", 
                           f"Se encontraron {contador_de_mensajes} mensajes")
        
        time.sleep(1)  
        ventana_prog.destroy()
        
        messagebox.showinfo("Información", 
                           f"Se encontraron {contador_de_mensajes} mensajes con adjuntos.")
        
        wb = load_workbook(path_full)
        wb.close()
        eliminar_archivo_unido(path_full)
    
    except Exception as e:
        ventana_prog.destroy()
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
        raise

