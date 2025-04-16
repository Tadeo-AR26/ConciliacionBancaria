import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def pintar_extracto(df, referencias, nombre_archivo="extracto_coloreado.xlsx"):
    codigos = referencias['codigo']
    df['_color'] = df['Cod. Operativo'].map(codigos).fillna("")

    # Guardamos a Excel primero
    df.to_excel(nombre_archivo, index=False)

    # Aplicamos color a filas en el archivo
    wb = load_workbook(nombre_archivo)
    ws = wb.active

    for idx, color in enumerate(df['_color'], start=2):  # empieza en 2 por el header
        if color:
            fill = PatternFill(start_color=color.replace("#", ""), end_color=color.replace("#", ""), fill_type="solid")
            for cell in ws[idx]:
                cell.fill = fill

    wb.save(nombre_archivo)
    return df


def pintar_contabilidad(df, referencias, nombre_archivo="contabilidad_coloreado.xlsx"):
    observaciones = [(obs.lower(), color) for obs, color in referencias['observacion'].items()]
    colores = []
    for observ in df['Observ.'].astype(str).str.lower():
        color = next((c for o, c in observaciones if o in observ), "")
        colores.append(color)
    df['_color'] = colores

    # Guardamos a Excel
    df.to_excel(nombre_archivo, index=False)

    # Aplicamos color a filas en el archivo
    wb = load_workbook(nombre_archivo)
    ws = wb.active

    for idx, color in enumerate(df['_color'], start=2):  # empieza en 2 por el header
        if color:
            fill = PatternFill(start_color=color.replace("#", ""), end_color=color.replace("#", ""), fill_type="solid")
            for cell in ws[idx]:
                cell.fill = fill

    wb.save(nombre_archivo)
    return df
