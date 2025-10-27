import win32com.client
import os
from modificador_excel import crearFiltro
import tkinter as tk 
from tkinter import messagebox

root = tk.Tk()
root.withdraw()
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

contador_de_mensajes=0

output_folder = "C:\\Outlook_Attachments"
os.makedirs(output_folder, exist_ok=True)
 
 
# and message.SenderName=="CDC"
i = 0


for message in messages:
    i += 1
    if(i == 10):
        break
    try:
        if(((message.subject=="[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico") or (messages.subject == "[EXT]: Ajuste Prorrata Generalizada costo SEN 0")) and message.SenderName=="CDC" ):
            contador_de_mensajes+=1
            print(message.subject,message.subject ==("[EXT]: Inicia Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico" or message.subject == "[EXT]: Ajuste Prorrata Generalizada costo SEN 0"))
            if(i == 2):
                break
            i += 1
            # if message.Attachments.Count > 0:
            #     for attachment in message.Attachments:
            #         attachment.SaveAsFile(attachment_path)
            #         print(f"Downloaded: {attachment.FileName}")

            #         crearFiltro(output_folder + "\\" + attachment.FileName)
    
    except AttributeError:
        continue



messagebox.showinfo("Información", f"Se encontraron {contador_de_mensajes} mensajes.")


