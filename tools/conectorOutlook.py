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


contador_de_mensajes=0

attachment_folder = r"C:\Users\fabia\OneDrive\Escritorio\nova data\novasource\adjuntos_prorrata" 
os.makedirs(attachment_folder, exist_ok=True)

fecha_actual = datetime.now().strftime("%Y-%m-%d")

attachment_folder_fecha= os.path.join(attachment_folder, fecha_actual)
os.makedirs(attachment_folder_fecha, exist_ok=True)

# and message.SenderName=="CDC"
i = 0

for message in messages:
    # i += 1
    # if(i == 10):
    #     break
    try:
        if(((message.subject=="[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico") or (messages.subject == "[EXT]: Ajuste Prorrata Generalizada costo SEN 0")) ):
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


