import win32com.client
import os
from modificador_excel2 import crearFiltro
import tkinter as tk 
from tkinter import messagebox
from datetime import datetime


root = tk.Tk()
root.withdraw()
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

print("Empezando analisis...")

contador_de_mensajes=0

ruta_script = os.path.dirname(os.path.abspath(__file__))

attachment_folder = "C:\\adjuntos_prorrata" 
os.makedirs(attachment_folder, exist_ok=True)

fecha_actual = datetime.now().strftime("%Y-%m-%d")

ahora =datetime.now()

attachment_folder_fecha= os.path.join(attachment_folder, fecha_actual)
os.makedirs(attachment_folder_fecha, exist_ok=True)

# and message.SenderName=="CDC"
i = 0

dia_hoy =str(ahora.day)
mes_hoy =str(ahora.month)

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


        
        if(received_time_pywintype.day != dia_hoy and received_time_pywintype.month !=mes_hoy):
            break
        if(((message.subject=="[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico") or (messages.subject == "[EXT]: Ajuste Prorrata Generalizada costo SEN 0"))and message.SenderName=="CDC" ):
            
            contador_de_mensajes+=1
            # print(message.subject,message.subject ==("[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico" or message.subject == "[EXT]: Ajuste Prorrata Generalizada costo SEN 0"))
            # if(i == 2):
            #     break
            # i += 1
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    full_path = os.path.join(attachment_folder_fecha, attachment.FileName)
                    attachment.SaveAsFile(full_path)
                    print(f"Downloaded: {attachment.FileName}")

                    crearFiltro(full_path, carpeta_donde_guardar=attachment_folder_fecha)

    
    except AttributeError:
        continue


messagebox.showinfo("Información", f"Se encontraron {contador_de_mensajes} mensajes.")


