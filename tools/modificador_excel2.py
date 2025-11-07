import pandas as pd
import tkinter as tk 
from tkinter import messagebox, ttk
from datetime import datetime
from estilos_excel import aplicar_formato_con_horas
from unirExcel import eliminar_archivo_unido
# from tools.estilos_excel import aplicar_formato_con_horas
import glob
import os 
from tkinter import filedialog
from openpyxl import load_workbook

root = tk.Tk()
root.withdraw()

prueba = True
prueba_ofi = False

GENERADORAS = ["PFV-ELPELICANO", "PFV-LAHUELLA", "PFV-ELROMERO"]

if prueba:
    if prueba_ofi:
        archivo = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\20241028_1015_InstruccionCDC Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico.xlsx"
        carpeta_donde_guardar = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba"
    else:
        archivo = filedialog.askdirectory(title="Selecciona la carpeta donde están los excels a unir")
        carpeta_donde_guardar = filedialog.askdirectory(title="Selecciona la carpeta donde guardar el archivo")


def crear_ventana_progreso():
    ventana = tk.Toplevel()
    ventana.title("Procesando archivos")
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


def unir_excels_en_carpeta(carpeta, nombre_salida="excel_unido.xlsx", ventana_prog=None):
    archivos = glob.glob(os.path.join(carpeta, "*.xlsx")) + glob.glob(os.path.join(carpeta, "*.xls"))

    if not archivos:
        print("No se encontraron archivos Excel en la carpeta indicada.")
        return None
    
    total_archivos = len(archivos)
    dfs = []
    
    for idx, archivo in enumerate(archivos):
        if ventana_prog:
            ventana, progress_bar, label_estado, label_detalle = ventana_prog
            progreso = int((idx / total_archivos) * 30)  # 30% del total
            actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                              progreso, "Uniendo archivos Excel...", 
                              f"Procesando: {os.path.basename(archivo)}")
        
        try:
            df = pd.read_excel(archivo)
            df["Archivo_Origen"] = os.path.basename(archivo)
            dfs.append(df)
        except Exception as e:
            print(f"Error leyendo {archivo}: {e}")
            messagebox.showerror("Error", f"No se pudo leer el archivo: {archivo}\n{e}")

    if not dfs:
        print("No se pudieron leer los archivos.")
        messagebox.showerror("Error", "No se pudieron leer los archivos Excel en la carpeta indicada.")
        return None

    if ventana_prog:
        ventana, progress_bar, label_estado, label_detalle = ventana_prog
        actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                          30, "Combinando datos...", "")

    df_unido = pd.concat(dfs, ignore_index=True)
    salida = os.path.join(carpeta, nombre_salida)
    df_unido.to_excel(salida, index=False)
    print(f"Archivos combinados correctamente en: {salida}")

    return salida


def eliminar_columnas_innecesarias(filtro):
    columnas_a_eliminar = ["PMAX (MW)", "PMIN (MW)", "SUBE/BAJA", ""]
    filtro = filtro.drop(columns=[col for col in columnas_a_eliminar if col in filtro.columns], errors="ignore")
    return filtro


def ordenar_columnas(filtro):
    columnas_deseadas = [
        'GEN.ACTUAL (MW)',
        'MONTO SUBE/BAJA (MW)',
        'CONSIGNA(MW)'
    ]
    
    if 'HORA' in filtro.columns:
        if 'FECHA' in filtro.columns:
            filtro['FECHA'] = pd.to_datetime(filtro['FECHA']).dt.date
        
        dfs = []
        for columna in columnas_deseadas:
            if columna in filtro.columns:
                pivot = filtro.pivot_table(
                    index=['FECHA', 'GENERADORA'],
                    columns='HORA',
                    values=columna,
                    aggfunc='first'
                )
                pivot.columns = [f'{hora}_{columna}' for hora in pivot.columns]
                dfs.append(pivot)
        
        if dfs:
            resultado = pd.concat(dfs, axis=1).reset_index()
            
            horas_unicas = []
            for col in resultado.columns:
                if col not in ['FECHA', 'GENERADORA']:
                    hora = col.split('_')[0]
                    if hora not in horas_unicas:
                        horas_unicas.append(hora)
            
            try:
                horas_unicas_sorted = sorted([int(h) for h in horas_unicas])
            except:
                horas_unicas_sorted = sorted(horas_unicas)
            
            nuevas_columnas = ['FECHA', 'GENERADORA']
            for hora in horas_unicas_sorted:
                for columna in columnas_deseadas:
                    col_nombre = f'{hora}_{columna}'
                    if col_nombre in resultado.columns:
                        nuevas_columnas.append(col_nombre)
                    for col in resultado.columns:
                if col not in ['FECHA', 'GENERADORA']:
                    partes = col.split('_')
                    hora = partes[0]
                    tipo = partes[1] if len(partes) > 1 else ''
                    
    
            
            resultado = resultado[nuevas_columnas]
            
            rename_dict = {}
            for col in resultado.columns:
                if col not in ['FECHA', 'GENERADORA']:
                    partes = col.split('_')
                    hora = partes[0]
                    tipo = partes[1] if len(partes) > 1 else ''
                
                    if 'GEN.ACTUAL' in tipo:
                        rename_dict[col] = 'GEN.ACTUAL'
                    elif 'MONTO' in tipo:
                        rename_dict[col] = 'MONTO'
                    elif 'CONSIGNA' in tipo:
                        rename_dict[col] = 'CONSIGNA'
                        
                    rename_dict[col] = tipo
                    
            resultado = resultado.rename(columns=rename_dict)
            resultado.attrs['horas_ordenadas'] = horas_unicas_sorted
            
            return resultado, filtro['FECHA'].unique()
    
    else:
        messagebox.showerror("Error", "No existe la columna 'HORA' en el archivo verificar los excels.")
        exit()


def crearFiltro(archivo, carpeta_donde_guardar, ventana_prog=None):
    if not archivo:
        messagebox.showerror("Error", "Error en la ejecucion por favor repetir proceso anterior.")
        return
    
    if ventana_prog:
        ventana, progress_bar, label_estado, label_detalle = ventana_prog
        actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                          35, "Leyendo archivo Excel...", "")
    
    xls = pd.ExcelFile(archivo)
    hoja_origen = xls.sheet_names[0]
    
    df = pd.read_excel(archivo, sheet_name=hoja_origen)

# aca podria colocar el verificador de la finalizacion de la prorrata 
    if "GENERADORA" not in df.columns:
        messagebox.showerror("Error", "La columna 'GENERADORA' no se encontró.")
        print("✗ La columna 'GENERADORA' no se encontró.")
        return

    if ventana_prog:
        actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                          45, "Filtrando datos...", "")

    filtro = df[df["GENERADORA"].isin( GENERADORAS )]
    filtro = eliminar_columnas_innecesarias(filtro)

    if 'FECHA' not in filtro.columns:
        messagebox.showerror("Error", "No existe la columna 'FECHA' en el archivo.")
        return

    filtro['FECHA'] = pd.to_datetime(filtro['FECHA']).dt.date
    fechas_unicas = sorted(filtro['FECHA'].unique())

    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    nuevo_nombre = os.path.join(carpeta_donde_guardar, f"Prorrata_procesada_{fecha_actual}.xlsx")

    if ventana_prog:
        actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                          55, "Creando archivo Excel...", "")

    total_fechas = len(fechas_unicas)

    idx, fecha = enumerate(fechas_unicas)
    
    with pd.ExcelWriter(nuevo_nombre, engine="openpyxl") as writer:
        for idx, fecha in enumerate(fechas_unicas):
            if ventana_prog:
                progreso = 55 + int((idx / total_fechas) * 25)  # 55% a 80%
                actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                                  progreso, "Escribiendo hojas...", 
                                  f"Procesando fecha: {fecha}")
            
            subfiltro = filtro[filtro['FECHA'] == fecha]
            resultado, _ = ordenar_columnas(subfiltro)

            hoja_nombre = str(fecha).replace("-", "_")
            resultado = resultado.round(0)

            resultado.to_excel(writer, sheet_name=hoja_nombre, index=False, startrow=0)
        
        # Aplicar formato DESPUÉS de escribir todos los datos
        for idx, fecha in enumerate(fechas_unicas):
            if ventana_prog:
                progreso = 80 + int((idx / total_fechas) * 20)  # 80% a 100%
                actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                                  progreso, "Aplicando formato...", 
                                  f"Formateando hoja: {fecha}")
            
            subfiltro = filtro[filtro['FECHA'] == fecha]
            resultado, _ = ordenar_columnas(subfiltro)
            hoja_nombre = str(fecha).replace("-", "_")
            aplicar_formato_con_horas(writer, hoja_nombre, resultado)

    if ventana_prog:
        actualizar_progreso(ventana, progress_bar, label_estado, label_detalle,
                          100, "¡Proceso completado!", 
                          f"Archivo guardado: {os.path.basename(nuevo_nombre)}")

    print(f"✓ Archivo guardado: {nuevo_nombre}")


if prueba:
    ventana_prog = crear_ventana_progreso()
    ventana, progress_bar, label_estado, label_detalle = ventana_prog
    
    try:
        archivo_unido = unir_excels_en_carpeta(archivo, ventana_prog=ventana_prog)
        
        crearFiltro(archivo_unido, carpeta_donde_guardar=carpeta_donde_guardar, ventana_prog=ventana_prog)
        
        ventana.after(2000, ventana.destroy)

        wb = load_workbook(archivo_unido)
        wb.close()
        eliminar_archivo_unido(archivo_unido)
    except Exception as e:
        ventana.destroy()
        messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
        raise
