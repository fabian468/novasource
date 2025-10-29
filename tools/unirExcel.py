import glob
import os
import pandas as pd


def unir_excels_en_carpeta(carpeta, nombre_salida="excel_unido.xlsx"):
    archivos = glob.glob(os.path.join(carpeta, "*.xlsx")) + glob.glob(os.path.join(carpeta, "*.xls"))

    if not archivos:
        print(" No se encontraron archivos Excel en la carpeta indicada.")
        return None
    
    dfs = []
    for archivo in archivos:
        try:
            df = pd.read_excel(archivo)
            df["Archivo_Origen"] = os.path.basename(archivo)  # opcional: saber de qué archivo viene cada fila
            dfs.append(df)
        except Exception as e:
            print(f" Error leyendo {archivo}: {e}")

    if not dfs:
        print(" No se pudieron leer los archivos.")
        return None

    # Combinar todos los DataFrames
    df_unido = pd.concat(dfs, ignore_index=True)

    # Guardar el archivo combinado
    salida = os.path.join(carpeta, nombre_salida)
    df_unido.to_excel(salida, index=False)
    print(f"Archivos combinados correctamente en: {salida}")

    return salida



def eliminar_archivo_unido(ruta_archivo):
    """
    Elimina el archivo final excel_unido.xlsx si existe.
    """
    if os.path.exists(ruta_archivo):
        try:
            os.remove(ruta_archivo)
            print(f"Archivo eliminado: {ruta_archivo}")
        except Exception as e:
            print(f" No se pudo eliminar {ruta_archivo}: {e}")
    else:
        print(" El archivo final no existe o ya fue eliminado.")