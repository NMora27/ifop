from pathlib import Path

import numpy as np
import pandas as pd
from lectura import leer_datos_talla_edad
from metodos_matematicos import filtrar_outliers_iqr_por_sexo_edad
from openpyxl import load_workbook
from openpyxl.styles import Alignment


# ============================================================
# CREAR MATRIZ TALLA-EDAD
def crear_matriz(datos, edad_agrupada):
    matriz = pd.crosstab(index=datos["GRUPO_TALLA"], columns=datos["GRUPO_EDAD"])
    edades = [str(i) for i in range(1, edad_agrupada) if str(i) in matriz.columns]

    if f"{edad_agrupada}+" in matriz.columns:
        edades.append(f"{edad_agrupada}+")

    matriz = matriz[edades]
    matriz["Total general"] = matriz.sum(axis=1)
    matriz.loc["Total general"] = matriz.sum(axis=0)
    return matriz.replace(0, "")


# ============================================================
# PROCESO PRINCIPAL
def generar_matriz_talla_edad(archivo, salida, pesqueria, paso_talla, edad_agrupada):
    archivo = Path(archivo)
    salida = Path(salida)
    columnas = ["LONGITUD_DEL_PEZ_(CM)", "Nº_DE_ANILLOS", "CÓDIGO_DE_PESQUERÍA", "SEXO"]

    print("\n" * 2, "=================== [ CLAVE PESO-TALLA ] =================== \n")

    # --------------------------------------------------
    # LECTURA
    datos = leer_datos_talla_edad(archivo, columnas)

    print("======================[ Datos originales ]======================")
    print(f"Registros leídos: {len(datos)}")

    # --------------------------------------------------
    # FILTROS
    datos = datos[(datos["CÓDIGO_DE_PESQUERÍA"] == pesqueria) & (datos["SEXO"].isin([1, 2])) &
                  (datos["Nº_DE_ANILLOS"] > 0) & (datos["LONGITUD_DEL_PEZ_(CM)"] > 0)].copy()
    datos["Nº_DE_ANILLOS"] = datos["Nº_DE_ANILLOS"].astype(int)
    datos["GRUPO_EDAD"] = datos["Nº_DE_ANILLOS"].astype(str)
    datos.loc[datos["Nº_DE_ANILLOS"] >= edad_agrupada,
              "GRUPO_EDAD"] = f"{edad_agrupada}+"

    print("\n======================[ Después de filtros ]======================")
    print(f"Total   : {len(datos)}")
    print(f"Machos  : {(datos['SEXO'] == 1).sum()}")
    print(f"Hembras : {(datos['SEXO'] == 2).sum()}")

    # --------------------------------------------------
    # APLICAR LIMPIEZA IQR
    print("\n======================[ LIMPIEZA IQR ]======================")
    datos, resumen_iqr = filtrar_outliers_iqr_por_sexo_edad(
        datos, columna="LONGITUD_DEL_PEZ_(CM)")

    for sexo, inicial, eliminados, final in resumen_iqr:
        nombre = "Machos" if sexo == 1 else "Hembras"
        print(f"\n{nombre}")
        print(f"Registros iniciales: {inicial}")
        print(f"Outliers eliminados: {eliminados}")
        print(f"Registros finales  : {final}")

    # --------------------------------------------------
    # AGRUPAR TALLAS DE 5 EN 5
    datos["GRUPO_TALLA"] = (np.floor(datos["LONGITUD_DEL_PEZ_(CM)"] / paso_talla) * paso_talla).astype(int)

    # --------------------------------------------------
    # SEPARAR MACHOS Y HEMBRAS
    machos = datos[datos["SEXO"] == 1].copy()
    hembras = datos[datos["SEXO"] == 2].copy()
    matriz_machos = crear_matriz(machos, edad_agrupada)
    matriz_hembras = crear_matriz(hembras, edad_agrupada)

    print("\n======================[ Datos finales IQR ]======================")
    print(f"Machos  : {len(machos)}")
    print(f"Hembras : {len(hembras)}")
    print(f"Total   : {len(datos)}")

    # --------------------------------------------------
    # EXPORTAR MATRICES A EXCEL
    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        matriz_machos.to_excel(writer, sheet_name="Machos")
        matriz_hembras.to_excel(writer, sheet_name="Hembras")

    # --------------------------------------------------
    # FORMATO DEL EXCEL
    libro = load_workbook(salida)

    for nombre_hoja in ["Machos", "Hembras"]:
        hoja = libro[nombre_hoja]
        hoja.column_dimensions["A"].width = 14

        for columna in hoja.iter_cols(min_col=2, max_col=hoja.max_column):
            letra = columna[0].column_letter
            hoja.column_dimensions[letra].width = 5

        ultima = hoja.cell(1, hoja.max_column).column_letter
        hoja.column_dimensions[ultima].width = 14

        for fila in range(1, hoja.max_row + 1):
            hoja.row_dimensions[fila].height = 18

        for fila in hoja.iter_rows():
            for celda in fila:
                celda.alignment = Alignment(horizontal="center", vertical="center")

    libro.save(salida)
    print(f"\nArchivo generado: {salida}")
    return salida
