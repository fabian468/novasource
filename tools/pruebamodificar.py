import pandas as pd
import tkinter as tk 
from tkinter import messagebox
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
# Asegúrate de que 'aplicar_formato_con_horas' esté correctamente definido en 'estilos_excel.py'
from estilos_excel import aplicar_formato_con_horas 
import os 


root = tk.Tk()
root.withdraw()

# Rutas de ejemplo
archivo = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\20251027_1328_InstruccionCDC Prorrata Generalizada costo SEN 0.xlsx"
carpeta_donde_guardar = r"C:\\Users\\GabrielBenelli\\Desktop\\prueba\\"

# --- Funciones Auxiliares ---

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
            # Convierte a objeto date de Python (elimina la hora/timestamp)
            filtro['FECHA'] = pd.to_datetime(filtro['FECHA']).dt.date
            
            # Convierte inmediatamente a str para evitar errores de tipo en pd.merge
            filtro['FECHA'] = filtro['FECHA'].astype(str)
        
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
            
            # Asegurar que FECHA sigue siendo string después del reset_index
            resultado['FECHA'] = resultado['FECHA'].astype(str)
            
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
                    
                    # Renombrado de columna para dejar solo el tipo
                    if 'GEN.ACTUAL' in tipo:
                        rename_dict[col] = f'{hora}_GEN.ACTUAL' # Se mantiene la hora en el nombre para la fusión
                    elif 'MONTO' in tipo:
                        rename_dict[col] = f'{hora}_MONTO'
                    elif 'CONSIGNA' in tipo:
                        rename_dict[col] = f'{hora}_CONSIGNA'
            
            resultado = resultado.rename(columns=rename_dict)
            
            resultado.attrs['horas_ordenadas'] = horas_unicas_sorted
            
            # Devuelve el primer valor de FECHA (que ya es str)
            return resultado , filtro['FECHA'].unique()[0]
    
    return filtro


# --- Función Principal (Modificada) ---

def crearFiltro(archivo, carpeta_donde_guardar):
    xls = pd.ExcelFile(archivo)
    hoja_origen = xls.sheet_names[0]
    df = pd.read_excel(archivo, sheet_name=hoja_origen)

    if "GENERADORA" not in df.columns:
        messagebox.showerror("Error", "La columna 'GENERADORA' no se encontró.")
        print("✗ La columna 'GENERADORA' no se encontró.")
        return

    filtro = df[df["GENERADORA"].isin(["PFV-ELPELICANO", "PFV-LAHUELLA"])]
    filtro = eliminar_columnas_innecesarias(filtro)
    # filtro3 ya tiene la FECHA como STR gracias a ordenar_columnas
    filtro3, fecha = ordenar_columnas(filtro) 

    fecha_hoja = str(fecha).replace("-", "_")

    fecha_actual_str = datetime.now().strftime("%Y-%m-%d")
    nuevo_nombre = os.path.join(carpeta_donde_guardar, f"Prorrata_procesada_{fecha_actual_str}.xlsx")
    
    if os.path.exists(nuevo_nombre):
        try:
            writer = pd.ExcelWriter(nuevo_nombre, engine="openpyxl", mode='a', if_sheet_exists='replace') 
          
            if fecha_hoja in writer.book.sheetnames:
                df_existente = pd.read_excel(nuevo_nombre, sheet_name=fecha_hoja)
            
                df_existente['FECHA'] = df_existente['FECHA'].astype(str)
                
                horas_nuevas = filtro3.attrs.get('horas_ordenadas', [])
                
                horas_existentes = set()
                for col in df_existente.columns:
                    if col not in ['FECHA', 'GENERADORA']:
                        try:
                            hora = col.split('_')[0] 
                            horas_existentes.add(hora)
                        except:
                            pass
                    
                horas_nuevas_a_añadir = [str(h) for h in horas_nuevas if str(h) not in horas_existentes]
                
                if horas_nuevas_a_añadir:
                    columnas_a_mantener = ['FECHA', 'GENERADORA']
                    columnas_nuevas = [col for col in filtro3.columns if col not in df_existente.columns and col not in ['FECHA', 'GENERADORA']]
                    
                    df_nuevas_horas = filtro3[columnas_a_mantener + columnas_nuevas]
                    
                    df_combinado = pd.merge(df_existente, df_nuevas_horas, on=['FECHA', 'GENERADORA'], how='left')
                    
                    columnas_existentes_ordenadas = [col for col in df_existente.columns if col not in ['FECHA', 'GENERADORA']]
                    columnas_finales = ['FECHA', 'GENERADORA'] + columnas_existentes_ordenadas + columnas_nuevas
                    df_combinado = df_combinado[columnas_finales]
                    
                    df_combinado.to_excel(writer, sheet_name=fecha_hoja, index=False)
                    writer.close() 
                    
                    with pd.ExcelWriter(nuevo_nombre, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer_format:
                        aplicar_formato_con_horas(writer_format, fecha_hoja, df_combinado)
                    
                    print(f"✓ Hoja '{fecha_hoja}' actualizada con nuevas horas: {horas_nuevas_a_añadir}")
                    
                else:
                    print(f"⚠ La fecha '{fecha_hoja}' ya existe con las mismas horas. No se realizaron cambios.")
            
            else:
                filtro3.to_excel(writer, sheet_name=fecha_hoja, index=False)
                writer.close() 

                with pd.ExcelWriter(nuevo_nombre, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer_format:
                    aplicar_formato_con_horas(writer_format, fecha_hoja, filtro3)
                    
                print(f"✓ Hoja '{fecha_hoja}' creada.")

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar el archivo existente o crear el nuevo: {e}")
            print(f"✗ Error al procesar: {e}")
            return
            
    else:
        try:
            with pd.ExcelWriter(nuevo_nombre, engine="openpyxl") as writer:
                filtro3.to_excel(writer, sheet_name=fecha_hoja, index=False)
                aplicar_formato_con_horas(writer, fecha_hoja, filtro3)
            print(f"✓ Archivo final creado y hoja '{fecha_hoja}' añadida.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al crear el archivo final: {e}")
            print(f"✗ Error al crear: {e}")
            return
            
crearFiltro(archivo,carpeta_donde_guardar=carpeta_donde_guardar)