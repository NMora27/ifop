"""Genera BPPALS desde cero, sin depender de una plantilla Excel externa."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from lectura import (buscar_desembarque, leer_frecuencias_colocap,
                     leer_matriz_talla_edad, leer_parametros_regresion_iqr)
from metodos_matematicos import calcular_a1_a2, calcular_peso


# ============================================================
# FORMATO
AZUL = "1F4E78"
BORDE = Side(style="thin", color="A6A6A6")

def aplicar_formato(hoja, fila_encabezado, columnas, congelar="A2"):
    for celda in hoja[fila_encabezado][:columnas]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.alignment = Alignment(horizontal="center", vertical="center")

    for fila in hoja.iter_rows(min_row=fila_encabezado, max_col=columnas):
        for celda in fila:
            celda.border = Border(left=BORDE, right=BORDE, top=BORDE, bottom=BORDE)
            celda.alignment = Alignment(horizontal="center", vertical="center")

    hoja.freeze_panes = congelar
    hoja.sheet_view.showGridLines = False

def escribir_titulo(hoja, titulo, columnas):
    hoja.merge_cells(start_row=1, start_column=1, end_row=1, end_column=columnas)
    celda = hoja.cell(1, 1, titulo)
    celda.font = Font(bold=True, color="FFFFFF", size=13)
    celda.fill = PatternFill("solid", fgColor=AZUL)
    celda.alignment = Alignment(horizontal="center")

def escribir_metadata(hoja, especie, sexo, zona, area, anio):
    datos = [("Especie", especie), ("Sexo", sexo), ("Zona", zona),
             ("Área", area), ("Año", anio)]
    for columna, (etiqueta, valor) in enumerate(datos, 1):
        hoja.cell(3, columna * 2 - 1, etiqueta).font = Font(bold=True)
        hoja.cell(3, columna * 2, valor)


# ============================================================
# PROCESO PRINCIPAL
def generar_bppals(especie, sexo, anio, zona, area, flota_desembarque,
                   archivo_matriz, archivo_regresion, archivo_colocap,
                   archivo_desembarque, salida_machos, salida_hembras,
                   talla_min, talla_max, paso_talla, edad_min, edad_max, hoja_colocap):
    if sexo not in ("Machos", "Hembras"):
        raise ValueError('SEXO debe ser "Machos" o "Hembras"')

    archivos = [Path(archivo_matriz), Path(archivo_regresion), Path(archivo_colocap),
                Path(archivo_desembarque)]
    for archivo in archivos:
        if not archivo.exists():
            raise FileNotFoundError(f"No existe el archivo: {archivo}")

    salida = Path(salida_machos if sexo == "Machos" else salida_hembras)
    salida.parent.mkdir(parents=True, exist_ok=True)
    tallas = list(range(talla_min, talla_max + 1, paso_talla))
    edades = list(range(edad_min, edad_max + 1))

    print("===============[ INICIO DEL PROCESAMIENTO ]===============")
    print(f"Especie seleccionada  : {especie}")
    print(f"Sexo seleccionado     : {sexo}")
    print(f"Año                   : {anio}")

    print("\n[1] Leyendo parámetros de regresión IQR...")
    a_machos, b_machos, a_hembras, b_hembras = leer_parametros_regresion_iqr(archivos[1])
    a_sexo, b_sexo = (a_machos, b_machos) if sexo == "Machos" else (a_hembras, b_hembras)

    print("[2] Leyendo matriz talla-edad...")
    datos = leer_matriz_talla_edad(archivos[0], sexo, talla_min, talla_max,
                                   paso_talla, edad_min, edad_max)

    print("[3] Leyendo frecuencias para Colocap...")
    total_ejemplares, n_machos, n_hembras, frecuencias = leer_frecuencias_colocap(
        archivos[2], hoja_colocap, paso_talla)
    desembarque = buscar_desembarque(archivos[3], especie, flota_desembarque)
    a1_machos, a2_machos = calcular_a1_a2(b_machos)
    a1_hembras, a2_hembras = calcular_a1_a2(b_hembras)

    # La composición se calcula en Python: el libro queda independiente de
    # fórmulas y de los BPPALS originales.
    composicion = []
    for talla in tallas:
        marca = talla + paso_talla // 2
        frecuencia = float(frecuencias.get(marca, {}).get(sexo, 0))
        total_clave = sum(datos[talla].values())
        valores = [frecuencia * datos[talla][edad] / total_clave if total_clave else 0
                   for edad in edades]
        composicion.append((talla, talla + paso_talla - 1, marca, frecuencia, valores))

    totales_edad = [sum(fila[4][i] for fila in composicion) for i in range(len(edades))]
    total_comp = sum(totales_edad)
    porcentajes = [100 * valor / total_comp if total_comp else 0 for valor in totales_edad]
    promedios, varianzas, pesos = [], [], []
    for indice in range(len(edades)):
        total = totales_edad[indice]
        promedio = sum(fila[2] * fila[4][indice] for fila in composicion) / total if total else 0
        momento2 = sum(fila[2] ** 2 * fila[4][indice] for fila in composicion) / total if total else 0
        promedios.append(promedio)
        varianzas.append(max(0, momento2 - promedio ** 2))
        pesos.append(float(calcular_peso(promedio, a_sexo, b_sexo)) if promedio else 0)

    print("[4] Creando libro BPPALS desde cero...")
    libro = Workbook()
    libro.remove(libro.active)
    ws_clave = libro.create_sheet("Clave")
    ws_colocap = libro.create_sheet("Colocap")
    ws_comp = libro.create_sheet("Comp.Capt.")
    ws_matrices = libro.create_sheet("Matrices")
    ws_graf = libro.create_sheet("Graf.")
    ws_tabla = libro.create_sheet("Tabla")
    ws_distr = libro.create_sheet("Distr.Capt.Nº")
    ws_print = libro.create_sheet("Tabla Print")

    # ------------------------------------------------------------
    # CLAVE
    escribir_titulo(ws_clave, "Clave edad-talla", 5 + len(edades))
    escribir_metadata(ws_clave, especie, sexo, zona, area, anio)
    edades_txt = [f"{edad}+" if edad == edad_max else edad for edad in edades]
    encabezados = ["Talla inf.", "-", "Talla sup.", "Marca", "Total"] + edades_txt
    for columna, valor in enumerate(encabezados, 1):
        ws_clave.cell(5, columna, valor)
    for fila, talla in enumerate(tallas, 6):
        valores = [talla, "-", talla + paso_talla - 1, talla + paso_talla // 2,
                   sum(datos[talla].values())] + [datos[talla][edad] or None for edad in edades]
        for columna, valor in enumerate(valores, 1):
            ws_clave.cell(fila, columna, valor)
    fila_total = 6 + len(tallas)
    ws_clave.cell(fila_total, 1, "Total")
    for columna in range(5, 6 + len(edades)):
        ws_clave.cell(fila_total, columna,
                      sum(ws_clave.cell(fila, columna).value or 0 for fila in range(6, fila_total)))
    aplicar_formato(ws_clave, 5, len(encabezados), "F6")

    # ------------------------------------------------------------
    # COLOCAP
    escribir_titulo(ws_colocap, "Parámetros y frecuencias Colocap", 5)
    parametros = [("Desembarque (t)", desembarque), ("Total ejemplares", total_ejemplares),
                  ("Nº Machos", n_machos), ("Nº Hembras", n_hembras),
                  ("a Machos", a_machos), ("b Machos", b_machos),
                  ("a Hembras", a_hembras), ("b Hembras", b_hembras),
                  ("a1 Machos", a1_machos), ("a2 Machos", a2_machos),
                  ("a1 Hembras", a1_hembras), ("a2 Hembras", a2_hembras)]
    for fila, (etiqueta, valor) in enumerate(parametros, 3):
        ws_colocap.cell(fila, 1, etiqueta).font = Font(bold=True)
        ws_colocap.cell(fila, 2, valor)
    for columna, valor in enumerate(["Talla inf.", "Talla sup.", "Marca", "Machos", "Hembras"], 1):
        ws_colocap.cell(17, columna, valor)
    for fila, talla in enumerate(tallas, 18):
        marca = talla + paso_talla // 2
        valores = [talla, talla + paso_talla - 1, marca,
                   frecuencias.get(marca, {}).get("Machos", 0),
                   frecuencias.get(marca, {}).get("Hembras", 0)]
        for columna, valor in enumerate(valores, 1):
            ws_colocap.cell(fila, columna, valor)
    aplicar_formato(ws_colocap, 17, 5, "A18")

    # ------------------------------------------------------------
    # COMPOSICIÓN Y MATRICES
    encabezados_comp = ["Talla inf.", "-", "Talla sup.", "Marca", "Frecuencia"] + edades_txt
    escribir_titulo(ws_comp, "Composición de captura talla-edad", len(encabezados_comp))
    escribir_metadata(ws_comp, especie, sexo, zona, area, anio)
    for columna, valor in enumerate(encabezados_comp, 1):
        ws_comp.cell(5, columna, valor)
    for fila, datos_fila in enumerate(composicion, 6):
        valores = [datos_fila[0], "-", datos_fila[1], datos_fila[2], datos_fila[3]] + datos_fila[4]
        for columna, valor in enumerate(valores, 1):
            ws_comp.cell(fila, columna, valor)
    fila_total_comp = 6 + len(tallas)
    ws_comp.cell(fila_total_comp, 1, "Total")
    for columna in range(5, 6 + len(edades)):
        ws_comp.cell(fila_total_comp, columna,
                     sum(ws_comp.cell(fila, columna).value or 0 for fila in range(6, fila_total_comp)))
    aplicar_formato(ws_comp, 5, len(encabezados_comp), "F6")

    encabezados_matriz = ["Marca", "Frecuencia"] + edades_txt
    escribir_titulo(ws_matrices, "Matrices para promedio y varianza", len(encabezados_matriz))
    fila_var_inicio = 3 + len(tallas) + 4
    for columna, valor in enumerate(encabezados_matriz, 1):
        ws_matrices.cell(3, columna, valor)
        ws_matrices.cell(fila_var_inicio, columna, valor)
    for fila, (_, _, marca, _, valores) in enumerate(composicion, 4):
        ws_matrices.cell(fila, 1, marca)
        ws_matrices.cell(fila, 2, sum(valores))
        for columna, valor in enumerate(valores, 3):
            ws_matrices.cell(fila, columna, marca * valor)
        fila_var = fila + len(tallas) + 4
        ws_matrices.cell(fila_var, 1, marca)
        ws_matrices.cell(fila_var, 2, sum(valores))
        for columna, valor in enumerate(valores, 3):
            ws_matrices.cell(fila_var, columna, marca ** 2 * valor)
    aplicar_formato(ws_matrices, 3, len(encabezados_matriz), "C4")
    aplicar_formato(ws_matrices, fila_var_inicio, len(encabezados_matriz), "C4")

    # ------------------------------------------------------------
    # TABLA, GRÁFICO Y DISTRIBUCIÓN
    escribir_titulo(ws_tabla, "Composición del desembarque en número de individuos", len(encabezados_comp))
    ws_tabla.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(encabezados_comp))
    ws_tabla.cell(3, 1, f"{especie}, {sexo.lower()}. {zona}, {area}, {anio}. "
                        f"Desembarque total: {desembarque:,.3f} t.")
    ws_tabla.cell(3, 1).alignment = Alignment(wrap_text=True)
    for columna, valor in enumerate(encabezados_comp, 1):
        ws_tabla.cell(5, columna, valor)
    for fila, datos_fila in enumerate(composicion, 6):
        valores = [datos_fila[0], "-", datos_fila[1], datos_fila[2], datos_fila[3]] + datos_fila[4]
        for columna, valor in enumerate(valores, 1):
            ws_tabla.cell(fila, columna, valor)
    fila_resumen = 7 + len(tallas)
    for etiqueta, valores in [("Total", totales_edad), ("Porcentaje", porcentajes),
                              ("Talla prom. (cm)", promedios), ("Varianza", varianzas),
                              ("Peso prom. (g)", pesos)]:
        ws_tabla.cell(fila_resumen, 1, etiqueta)
        for columna, valor in enumerate(valores, 6):
            ws_tabla.cell(fila_resumen, columna, valor)
        fila_resumen += 1
    aplicar_formato(ws_tabla, 5, len(encabezados_comp), "F6")

    escribir_titulo(ws_graf, "Composición del desembarque por grupo de edad", 3)
    ws_graf.cell(3, 1, "Grupo de edad")
    ws_graf.cell(3, 2, "Porcentaje")
    for fila, (edad, porcentaje) in enumerate(zip(edades_txt, porcentajes), 4):
        ws_graf.cell(fila, 1, edad)
        ws_graf.cell(fila, 2, porcentaje)
    aplicar_formato(ws_graf, 3, 2, "A4")
    grafico = BarChart()
    grafico.title = f"{especie} - {sexo} ({anio})"
    grafico.y_axis.title = "Porcentaje"
    grafico.x_axis.title = "Grupo de edad"
    grafico.add_data(Reference(ws_graf, min_col=2, min_row=3, max_row=3 + len(edades)), titles_from_data=True)
    grafico.set_categories(Reference(ws_graf, min_col=1, min_row=4, max_row=3 + len(edades)))
    ws_graf.add_chart(grafico, "D3")

    escribir_titulo(ws_distr, "Distribución de captura por talla", 6)
    encabezados_distr = ["Talla inf.", "-", "Talla sup.", "Machos", "Hembras", "Total"]
    for columna, valor in enumerate(encabezados_distr, 1):
        ws_distr.cell(3, columna, valor)
    for fila, talla in enumerate(tallas, 4):
        marca = talla + paso_talla // 2
        machos = float(frecuencias.get(marca, {}).get("Machos", 0))
        hembras = float(frecuencias.get(marca, {}).get("Hembras", 0))
        for columna, valor in enumerate([talla, "-", talla + paso_talla - 1, machos, hembras, machos + hembras], 1):
            ws_distr.cell(fila, columna, valor)
    aplicar_formato(ws_distr, 3, 6, "A4")

    escribir_titulo(ws_print, "Tabla para impresión", len(encabezados_comp))
    ws_print.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(encabezados_comp))
    ws_print.cell(3, 1, ws_tabla.cell(3, 1).value)
    ws_print.cell(3, 1).alignment = Alignment(wrap_text=True)
    for columna, valor in enumerate(encabezados_comp, 1):
        ws_print.cell(5, columna, valor)
    for fila, datos_fila in enumerate(composicion, 6):
        valores = [datos_fila[0], "-", datos_fila[1], datos_fila[2], datos_fila[3]] + datos_fila[4]
        for columna, valor in enumerate(valores, 1):
            ws_print.cell(fila, columna, valor)
    aplicar_formato(ws_print, 5, len(encabezados_comp), "F6")
    ws_print.page_setup.orientation = "landscape"
    ws_print.page_setup.fitToWidth = 1
    ws_print.sheet_properties.pageSetUpPr.fitToPage = True

    for hoja in libro.worksheets:
        for columna in range(1, hoja.max_column + 1):
            hoja.column_dimensions[get_column_letter(columna)].width = 13
        for fila in hoja.iter_rows():
            for celda in fila:
                if isinstance(celda.value, float):
                    celda.number_format = "0.000"

    libro.save(salida)
    print("\n=============[ PROCESAMIENTO FINALIZADO ]=============")
    print("Libro creado sin plantilla BPPALS externa.")
    print(f"Archivo: {salida}")
    return salida
