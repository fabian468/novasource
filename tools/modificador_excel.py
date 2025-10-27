import pandas as pd
import tkinter as tk 
from tkinter import messagebox

root = tk.Tk()
root.withdraw()

# archivo = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\20251027_1228_InstruccionCDC Prorrata Generalizada costo SEN 0.xlsx"

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
    
    # Verifica si existe la columna 'Hora'
    if 'HORA' in filtro.columns:
        # Crear pivot tables para cada columna deseada
        dfs = []
        for columna in columnas_deseadas:
            if columna in filtro.columns:
                pivot = filtro.pivot_table(
                    index=['GENERADORA'],
                    columns='HORA',
                    values=columna,
                    aggfunc='first'
                )
                # Renombrar las columnas para incluir el tipo de valor
                pivot.columns = [f'{hora}_{columna}' for hora in pivot.columns]
                dfs.append(pivot)
        
        # Combinar todos los pivot tables
        if dfs:
            resultado = pd.concat(dfs, axis=1).reset_index()
            return resultado
        
    return filtro



def crearFiltro(archivo):
    xls = pd.ExcelFile(archivo)
    hoja_origen = xls.sheet_names[0]
    df = pd.read_excel(archivo, sheet_name=hoja_origen)

    if "GENERADORA" in df.columns:
        
        filtro = df[df["GENERADORA"].isin(["PFV-ELPELICANO", "PFV-LAHUELLA"])]
        filtro = eliminar_columnas_innecesarias(filtro)
        filtro3 = ordenar_columnas(filtro)

        with pd.ExcelWriter(archivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            filtro.to_excel(writer, sheet_name="Filtro_PFV2", index=False)

        with pd.ExcelWriter(archivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            filtro3.to_excel(writer, sheet_name="Filtro_PFV3", index=False)

    else:
        messagebox.showerror("Error", "La columna 'GENERADORA' no se encontró.")
        print("La columna 'GENERADORA' no se encontró.")


# crearFiltro(archivo)