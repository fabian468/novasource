import pandas as pd
import tkinter as tk 
from tkinter import messagebox
from datetime import datetime
# from estilos_excel import aplicar_formato_con_horas
from tools.estilos_excel import aplicar_formato_con_horas
import glob
import os 
from tkinter import filedialog


root = tk.Tk()
root.withdraw()

prueba = False
prueba_ofi = False

if prueba:
    if prueba_ofi:
        archivo = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\20241028_1015_InstruccionCDC Prorrata LT 500 kV Nueva Pan de Azúcar - Polpaico.xlsx"
        carpeta_donde_guardar = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba"
    else:
        archivo = filedialog.askdirectory(title="Selecciona la carpeta donde están los excels a unir")
        carpeta_donde_guardar = filedialog.askdirectory(title="Selecciona la carpeta donde guardar el archivo")


def unir_excels_en_carpeta(carpeta, nombre_salida="excel_unido.xlsx"):
    archivos = glob.glob(os.path.join(carpeta, "*.xlsx")) + glob.glob(os.path.join(carpeta, "*.xls"))

    if not archivos:
        print("No se encontraron archivos Excel en la carpeta indicada.")
        return None
    
    dfs = []
    for archivo in archivos:
        try:
            df = pd.read_excel(archivo)
            df["Archivo_Origen"] = os.path.basename(archivo)
            dfs.append(df)
        except Exception as e:
            print(f"Error leyendo {archivo}: {e}")

    if not dfs:
        print("No se pudieron leer los archivos.")
        return None

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
            
            return resultado, filtro['FECHA'].unique()
    
    return filtro


def crearFiltro(archivo, carpeta_donde_guardar):
    if not archivo:
        messagebox.showerror("Error", "No se encontraron excels para procesar.")
        return
    
    xls = pd.ExcelFile(archivo)
    hoja_origen = xls.sheet_names[0]
    df = pd.read_excel(archivo, sheet_name=hoja_origen)

    if "GENERADORA" not in df.columns:
        messagebox.showerror("Error", "La columna 'GENERADORA' no se encontró.")
        print("✗ La columna 'GENERADORA' no se encontró.")
        return

    filtro = df[df["GENERADORA"].isin(["PFV-ELPELICANO", "PFV-LAHUELLA", "PFV-ELROMERO"])]
    filtro = eliminar_columnas_innecesarias(filtro)

    if 'FECHA' not in filtro.columns:
        messagebox.showerror("Error", "No existe la columna 'FECHA' en el archivo.")
        return

    filtro['FECHA'] = pd.to_datetime(filtro['FECHA']).dt.date
    fechas_unicas = sorted(filtro['FECHA'].unique())

    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    nuevo_nombre = os.path.join(carpeta_donde_guardar, f"Prorrata_procesada_{fecha_actual}.xlsx")

    # CORRECCIÓN CRÍTICA: Usar mode='w' para crear archivo limpio
    with pd.ExcelWriter(nuevo_nombre, engine="openpyxl") as writer:
        for fecha in fechas_unicas:
            subfiltro = filtro[filtro['FECHA'] == fecha]
            resultado, _ = ordenar_columnas(subfiltro)

            hoja_nombre = str(fecha).replace("-", "_")
            resultado = resultado.round(0)

            # NO escribir el header aquí, lo manejará la función de formato
            resultado.to_excel(writer, sheet_name=hoja_nombre, index=False, startrow=0)
            
        # Aplicar formato DESPUÉS de escribir todos los datos
        for fecha in fechas_unicas:
            subfiltro = filtro[filtro['FECHA'] == fecha]
            resultado, _ = ordenar_columnas(subfiltro)
            hoja_nombre = str(fecha).replace("-", "_")
            aplicar_formato_con_horas(writer, hoja_nombre, resultado)

    print(f"✓ Archivo guardado: {nuevo_nombre}")


if prueba:
    crearFiltro(unir_excels_en_carpeta(carpeta_donde_guardar), carpeta_donde_guardar=carpeta_donde_guardar)