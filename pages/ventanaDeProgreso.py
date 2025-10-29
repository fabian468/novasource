from tkinter import ttk
import tkinter as tk
from PIL import Image, ImageTk 

def crear_ventana_progreso():
    ventana = tk.Toplevel()
    ventana.title("Procesando correos")
    ventana.geometry("400x150")
    ventana.resizable(False, False)
    
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (400 // 2)
    y = (ventana.winfo_screenheight() // 2) - (150 // 2)
    ventana.geometry(f"400x150+{x}+{y}")
    
    label_estado = tk.Label(ventana, text="Iniciando...", font=("Arial", 10))
    label_estado.pack(pady=10)
    
    progress_bar = ttk.Progressbar(ventana, length=350, mode='determinate')
    progress_bar.pack(pady=10)
    
    label_detalle = tk.Label(ventana, text="", font=("Arial", 9), fg="gray")
    label_detalle.pack(pady=5)
    
    return ventana, progress_bar, label_estado, label_detalle

def actualizar_progreso(ventana, progress_bar, label_estado, label_detalle, 
                        valor, texto_estado, texto_detalle=""):
    
    progress_bar['value'] = valor
    label_estado.config(text=texto_estado)
    label_detalle.config(text=texto_detalle)
    ventana.update()
