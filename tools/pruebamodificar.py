import pandas as pd
import tkinter as tk 
from tkinter import messagebox
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from estilos_excel import aplicar_formato_con_horas
import os 

root = tk.Tk()
root.withdraw()

archivo = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\20251027_1228_InstruccionCDC Prorrata Generalizada costo SEN 0.xlsx"
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
            
            return resultado, filtro['FECHA'].unique()[0]
    
    return filtro, None


def combinar_con_archivo_existente(df_nuevo, fecha_hoja, archivo_salida):
    """
    Combina el DataFrame nuevo con datos existentes si la hoja ya existe.
    Solo añade columnas de horas que no existen previamente.
    """
    if not os.path.exists(archivo_salida):
        return df_nuevo
    
    try:
        # Leer la hoja existente
        df_existente = pd.read_excel(archivo_salida, sheet_name=fecha_hoja)
        
        # Asegurar que ambos DataFrames tengan el mismo tipo de dato para FECHA
        if 'FECHA' in df_existente.columns:
            df_existente['FECHA'] = pd.to_datetime(df_existente['FECHA']).dt.date
        if 'FECHA' in df_nuevo.columns:
            df_nuevo['FECHA'] = pd.to_datetime(df_nuevo['FECHA']).dt.date
        
        # Obtener las horas existentes y las nuevas
        horas_existentes = set()
        horas_nuevas = set()
        
        columnas_fijas = ['FECHA', 'GENERADORA']
        
        for col in df_existente.columns:
            if col not in columnas_fijas:
                # Extraer el número de hora de columnas como "1_GEN.ACTUAL"
                hora = col.split('_')[0] if '_' in str(col) else col
                horas_existentes.add(hora)
        
        for col in df_nuevo.columns:
            if col not in columnas_fijas:
                hora = col.split('_')[0] if '_' in str(col) else col
                horas_nuevas.add(hora)
        
        # Identificar horas que realmente son nuevas
        horas_a_agregar = horas_nuevas - horas_existentes
        
        if not horas_a_agregar:
            print(f"⚠ No hay horas nuevas para agregar en la fecha {fecha_hoja}")
            return df_existente
        
        print(f"✓ Agregando horas nuevas: {sorted(horas_a_agregar)}")
        
        # Filtrar solo las columnas de las horas nuevas del df_nuevo
        columnas_a_agregar = columnas_fijas.copy()
        for col in df_nuevo.columns:
            if col not in columnas_fijas:
                hora = col.split('_')[0] if '_' in str(col) else col
                if hora in horas_a_agregar:
                    columnas_a_agregar.append(col)
        
        df_nuevas_horas = df_nuevo[columnas_a_agregar]
        
        # Hacer merge en FECHA y GENERADORA
        df_combinado = pd.merge(
            df_existente,
            df_nuevas_horas,
            on=['FECHA', 'GENERADORA'],
            how='outer'
        )
        
        # Reordenar columnas: FECHA, GENERADORA, luego por horas ordenadas
        todas_horas = horas_existentes.union(horas_nuevas)
        try:
            horas_ordenadas = sorted([int(h) for h in todas_horas])
        except:
            horas_ordenadas = sorted(todas_horas)
        
        columnas_ordenadas = columnas_fijas.copy()
        columnas_deseadas = ['GEN.ACTUAL', 'MONTO', 'CONSIGNA']
        
        for hora in horas_ordenadas:
            for sufijo in columnas_deseadas:
                # Buscar la columna exacta en el DataFrame
                for col in df_combinado.columns:
                    if col not in columnas_fijas and str(hora) in str(col) and sufijo in str(col):
                        if col not in columnas_ordenadas:
                            columnas_ordenadas.append(col)
        
        df_combinado = df_combinado[columnas_ordenadas]
        df_combinado.attrs['horas_ordenadas'] = horas_ordenadas
        
        return df_combinado
        
    except Exception as e:
        print(f"⚠ Error al combinar con archivo existente: {e}")
        return df_nuevo


def crearFiltro(archivo, carpeta_donde_guardar):
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
    filtro3, fecha = ordenar_columnas(filtro)

    if fecha is None:
        print("✗ No se pudo extraer la fecha del archivo.")
        return

    fecha_hoja = str(fecha).replace("-", "_")
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    nuevo_nombre = os.path.join(carpeta_donde_guardar, f"Prorrata_procesada_{fecha_actual}.xlsx")

    # Combinar con datos existentes si la hoja ya existe
    filtro3 = combinar_con_archivo_existente(filtro3, fecha_hoja, nuevo_nombre)

    # Escribir el archivo
    if os.path.exists(nuevo_nombre):
        # Usar mode='a' para append, que automáticamente carga el workbook existente
        with pd.ExcelWriter(nuevo_nombre, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
            filtro3.to_excel(writer, sheet_name=fecha_hoja, index=False)
            aplicar_formato_con_horas(writer, fecha_hoja, filtro3)
    else:
        # Crear archivo nuevo
        with pd.ExcelWriter(nuevo_nombre, engine="openpyxl") as writer:
            filtro3.to_excel(writer, sheet_name=fecha_hoja, index=False)
            aplicar_formato_con_horas(writer, fecha_hoja, filtro3)

    print(f"✓ Archivo procesado: {nuevo_nombre}")
    print(f"✓ Hoja: {fecha_hoja}")


# Ejemplo de uso procesando múltiples archivos
crearFiltro(archivo, carpeta_donde_guardar=carpeta_donde_guardar)