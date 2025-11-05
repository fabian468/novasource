from openpyxl.chart import LineChart, Reference, Series
from openpyxl.utils import get_column_letter
import re
from datetime import datetime

def generarGrafico(column_names, data_values, generadoras, horas_ordenadas, worksheet, fila_actual, DATA_START_ROW):
    try:
        chart = LineChart()
        chart.title = "Consignas por Hora por Generadora"
        chart.style = 12
        chart.y_axis.title = "Consigna (MW)"
        chart.x_axis.title = "Horas"

        # --- columnas de consigna (índices 1-based)
        columnas_consigna = [(i + 1, str(c)) for i, c in enumerate(column_names) if 'CONSIGNA' in str(c).upper()]
        if not columnas_consigna:
            print("No se encontraron columnas con 'CONSIGNA' para graficar.")
            return

        # --- límites: inicio fijo en C (3) según tu hoja, y fin en última columna de consigna
        fila_eje_x = 3
        start_col_index = 3  # columna C
        end_col_index = columnas_consigna[-1][0]

        # --- obtener columnas que tienen valor (no None) en la fila de encabezado (C3..end)
        rango_encabezado = f"{get_column_letter(start_col_index)}{fila_eje_x}:{get_column_letter(end_col_index)}{fila_eje_x}"
        encabezado_row = list(worksheet[rango_encabezado])[0]  # tupla de celdas
        non_empty_cols = []
        encabezados_limpios = []
        for i, cell in enumerate(encabezado_row):
            if cell.value is not None and str(cell.value).strip() != "":
                col_idx = start_col_index + i
                non_empty_cols.append(col_idx)
                # Limpiar texto: extraer HH:MM si existe
                texto = str(cell.value).strip()
                m = re.search(r'(\d{1,2}:\d{2})', texto)
                if m:
                    horas = m.group(1)
                    try:
                        hora_obj = datetime.strptime(horas, "%H:%M")
                        encabezados_limpios.append(hora_obj.strftime("%H:%M"))
                    except:
                        encabezados_limpios.append(horas)
                else:
                    encabezados_limpios.append(texto)

        if not non_empty_cols:
            print("No hay encabezados válidos en la fila de horas.")
            return

        print("Columnas válidas (sin None):", non_empty_cols)
        print("Encabezados limpios:", encabezados_limpios)

        # --- Fila auxiliar para categorías limpias y datos compactados
        fila_aux_cat = fila_actual + 50       # fila donde escribir las categorías limpias
        fila_aux_data_start = fila_aux_cat + 1  # fila donde comienzan las filas de datos por generadora

        # Escribir encabezados limpios en una fila contigua (sin huecos)
        for j, txt in enumerate(encabezados_limpios):
            worksheet.cell(row=fila_aux_cat, column=start_col_index + j, value=txt)

        # --- Para cada generadora: encontrar su fila original y copiar solo las columnas válidas
        filas_aux_generadora = []  # lista de (generadora, fila_aux)
        for g_idx, gen in enumerate(generadoras):
            fila_gen = None
            for i, row_data in enumerate(data_values):
                if row_data[1] == gen:
                    fila_gen = DATA_START_ROW + i
                    break
            if fila_gen is None:
                print(f"Advertencia: no se encontró fila para la generadora '{gen}'")
                continue

            fila_aux = fila_aux_data_start + len(filas_aux_generadora)
            filas_aux_generadora.append((gen, fila_aux))

            # copiar valores de las columnas non_empty_cols en orden contiguo
            for j, orig_col in enumerate(non_empty_cols):
                val = worksheet.cell(row=fila_gen, column=orig_col).value
                worksheet.cell(row=fila_aux, column=start_col_index + j, value=val)

        if not filas_aux_generadora:
            print("No se copiaron filas de generadoras a la zona auxiliar.")
            return

        # --- Crear referencias para categorías y series apuntando a la zona auxiliar
        cats = Reference(
            worksheet,
            min_col=start_col_index,
            min_row=fila_aux_cat,
            max_col=start_col_index + len(encabezados_limpios) - 1,
            max_row=fila_aux_cat
        )

        # Añadir series desde las filas auxiliares (cada fila es una serie)
        for gen, fila_aux in filas_aux_generadora:
            values = Reference(
                worksheet,
                min_col=start_col_index,
                min_row=fila_aux,
                max_col=start_col_index + len(encabezados_limpios) - 1,
                max_row=fila_aux
            )
            serie = Series(values, title=gen)
            chart.series.append(serie)

        chart.set_categories(cats)

        # Estética del gráfico
        chart.height = 8
        chart.width = 20
        chart.legend.position = 'r'
        chart.x_axis.tickLblPos = "low"  # opcional: mejorar posicion etiquetas

        FILA_GRAFICO = fila_actual + 4
        worksheet.add_chart(chart, "C16")

        print("Gráfico creado correctamente (zonas auxiliares usadas para mantener alineación).")

    except Exception as e:
        print(f" Error al crear el gráfico: {e}")
        import traceback
        traceback.print_exc()
