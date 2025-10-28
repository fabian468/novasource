import pandas as pd
import tkinter as tk 
from tkinter import messagebox
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
# from styles.estilos_excel import aplicar_formato_con_horas
from estilos_excel import aplicar_formato_con_horas
import os 

root = tk.Tk()
root.withdraw()

archivo = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\20251027_1328_InstruccionCDC Prorrata Generalizada costo SEN 0.xlsx"
carpeta_donde_guardar = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\"

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
            
            resultado = resultado.rename(columns=rename_dict)
            
            resultado.attrs['horas_ordenadas'] = horas_unicas_sorted
            
            return resultado , filtro['FECHA'].unique()[0]
    
    return filtro



def crearFiltro(archivo , carpeta_donde_guardar):
    # Leer el Excel original
    xls = pd.ExcelFile(archivo)
    hoja_origen = xls.sheet_names[0]
    df = pd.read_excel(archivo, sheet_name=hoja_origen)

    # Verificar columna requerida
    if "GENERADORA" not in df.columns:
        messagebox.showerror("Error", "La columna 'GENERADORA' no se encontró.")
        print("✗ La columna 'GENERADORA' no se encontró.")
        return

    filtro = df[df["GENERADORA"].isin(["PFV-ELPELICANO", "PFV-LAHUELLA"])]
    filtro = eliminar_columnas_innecesarias(filtro)
    filtro3 , fecha = ordenar_columnas(filtro)

    fecha_hoja = str(fecha).replace("-", "_") 

    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    nuevo_nombre = os.path.join(carpeta_donde_guardar, f"Prorrata_procesada_{fecha_actual}.xlsx")

    with pd.ExcelWriter(nuevo_nombre, engine="openpyxl") as writer:
        # filtro.to_excel(writer, sheet_name="Filtro_PFV2", index=False)
        # aplicar_formato_simple(writer, "Filtro_PFV2", filtro)

        filtro3.to_excel(writer, sheet_name=fecha_hoja,index=False)
        aplicar_formato_con_horas(writer,fecha_hoja, filtro3)

    # messagebox.showinfo("Éxito", f"Archivo procesado correctamente.\nGuardado como: {nuevo_nombre}")

    # try:
    #     os.startfile(nuevo_nombre)
    #     print("✓ Archivo abierto automáticamente.")
    # except Exception as e:
    #     print(f"⚠ No se pudo abrir el archivo automáticamente: {e}")



crearFiltro(archivo,carpeta_donde_guardar=carpeta_donde_guardar)