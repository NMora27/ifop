from pathlib import Path

import numpy as np
import pandas as pd
from lectura import leer_excel_palangre
from metodos_matematicos import (
    ajustar_regresion_longpeso,
    aplicar_outliers_longpeso,
    graficar_regresion,
)
from scipy.stats import t


# ============================================================
# FUNCIONES AUXILIARES
def convertir_latitud(valor):
    valor = int(valor)
    grados = valor // 10000
    minutos = (valor // 100) % 100
    segundos = valor % 100
    return -(grados + minutos / 60 + segundos / 3600)

def ajustar_e_informar_regresion(datos, origen):
    a, b, log_a, r = ajustar_regresion_longpeso(datos)
    return a, b, log_a, r

def guardar_dataframe(datos, nombre_archivo):
    if datos.empty:
        raise ValueError(f"No se puede generar {Path(nombre_archivo).name}: no hay registros")

    nombre_archivo = Path(nombre_archivo)
    nombre_archivo.parent.mkdir(parents=True, exist_ok=True)
    datos.to_csv(nombre_archivo, sep=";", decimal=",", index=False)
    print(f"Archivo generado: {nombre_archivo}")


def exportar_tabla_regresiones(resultados, salida):
    """Exporta IC marginales del 95 % del ajuste log10(W) = log10(a) + b log10(L)."""
    salida = Path(salida)
    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for metodo in ("Original", "Cook", "MAD", "IQR"):
            filas = []
            for grupo in ("Machos", "Hembras", "Ambos"):
                resultado = resultados[grupo][metodo]
                datos = resultado["datos"]
                datos = datos.loc[(datos["LONGITUD_ESPECIMEN"] > 0)
                                  & (datos["PESO_ESPECIMEN"] > 0),
                                  ["LONGITUD_ESPECIMEN", "PESO_ESPECIMEN"]]
                datos = datos.loc[np.isfinite(datos).all(axis=1)]
                n = len(datos)
                inferior = [np.nan, np.nan]
                superior = [np.nan, np.nan]
                if n > 2:
                    x = np.log10(datos["LONGITUD_ESPECIMEN"].to_numpy())
                    y = np.log10(datos["PESO_ESPECIMEN"].to_numpy())
                    sxx = np.sum((x - x.mean()) ** 2)
                    if sxx > 0:
                        residuos = y - (resultado["log_a"] + resultado["b"] * x)
                        varianza = np.sum(residuos ** 2) / (n - 2)
                        critico = t.ppf(0.975, df=n - 2)
                        margen_b = critico * np.sqrt(varianza / sxx)
                        margen_log_a = critico * np.sqrt(
                            varianza * (1 / n + x.mean() ** 2 / sxx))
                        inferior = [10 ** (resultado["log_a"] - margen_log_a),
                                    resultado["b"] - margen_b]
                        superior = [10 ** (resultado["log_a"] + margen_log_a),
                                    resultado["b"] + margen_b]
                filas.extend([
                    [grupo, "Área total", resultado["a"], resultado["b"],
                     resultado["r"] ** 2, n],
                    [grupo, "Límite inferior (95 %)", *inferior, None, None],
                    [grupo, "Límite superior (95 %)", *superior, None, None],
                ])
            tabla = pd.DataFrame(filas, columns=["Sexo", "Estimación", "a", "b", "r²", "N"])
            tabla[["a", "b", "r²"]] = tabla[["a", "b", "r²"]].round(7)
            tabla.to_excel(writer, sheet_name=metodo, index=False)
            hoja = writer.sheets[metodo]
            hoja.freeze_panes = "C2"
            hoja.auto_filter.ref = hoja.dimensions
            for columna, ancho in {"A": 14, "B": 28, "C": 22, "D": 22,
                                   "E": 18, "F": 12}.items():
                hoja.column_dimensions[columna].width = ancho
            for fila in hoja.iter_rows(min_row=2):
                fila[2].number_format = "0.#######"
                fila[3].number_format = "0.#######"
                fila[4].number_format = "0.#######"
                fila[5].number_format = "0"
        pd.DataFrame({"Notas": [
            "Modelo: W = a × L^b; ajuste lineal en escala log10.",
            "IC marginales del 95 % con t de Student y N - 2 grados de libertad.",
            "El IC de a se obtiene transformando los límites de log10(a).",
            "r² corresponde al ajuste en escala log10; N es el número de observaciones utilizadas.",
            "a, b y r² se redondean a 7 decimales, igual que en el CSV; r_PEARSON es r, no r².",
            "Los límites se presentan para a y b; r² y N solo en la fila Área total.",
            "IC vacíos si N <= 2 o si no hay variación en longitud.",
            "Área total corresponde al área seleccionada por el filtro geográfico del proceso.",
            "Los IC de Cook, MAD e IQR se calculan sobre los datos retenidos por cada método.",
        ]}).to_excel(writer, sheet_name="Notas", index=False)
        writer.sheets["Notas"].column_dimensions["A"].width = 115
    print(f"Archivo generado: {salida}")


# ============================================================
# PROCESO PRINCIPAL
def generar_regresiones(archivo, anio, especie, latitud_zona_1, latitud_zona_2,
                                directorio_graficos, salida_preprocesamiento,
                                salida_regresiones, salida_iqr_ambos,
                                salida_iqr_machos, salida_iqr_hembras,
                                salida_regresiones_excel=None):
    archivo = Path(archivo)
    directorio_graficos = Path(directorio_graficos)
    salida_preprocesamiento = Path(salida_preprocesamiento)
    salida_regresiones = Path(salida_regresiones)
    salida_iqr_ambos = Path(salida_iqr_ambos)
    salida_iqr_machos = Path(salida_iqr_machos)
    salida_iqr_hembras = Path(salida_iqr_hembras)
    directorio_graficos.mkdir(parents=True, exist_ok=True)

    # --------------------------------------------------
    # PREPROCESAMIENTO
    columnas = ["LATITUD", "ESPECIE_OBJETIVO_LANCE", "SEXO_ESPECIMEN",
                "LONGITUD_ESPECIMEN", "PESO_ESPECIMEN"]
    df, _n_original = leer_excel_palangre(archivo, columnas, especie, anio)

    df["LATITUD"] = df["LATITUD"].apply(convertir_latitud)
    df["ZONA"] = 0
    df.loc[(df["LATITUD"] <= latitud_zona_1) &
           (df["LATITUD"] > latitud_zona_2), "ZONA"] = 1
    df.loc[df["LATITUD"] <= latitud_zona_2, "ZONA"] = 2
    df = df[(df["ESPECIE_OBJETIVO_LANCE"] == especie) &
            (df["SEXO_ESPECIMEN"].isin([1, 2])) &
            (df["LATITUD"] <= latitud_zona_2)]

    salida_preprocesamiento.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(salida_preprocesamiento, sep=";", index=False)

    print(f"Archivo generado: {salida_preprocesamiento}")

    # --------------------------------------------------
    # GRUPOS, MÉTODOS Y GRÁFICOS
    grupos = {"Ambos": {"datos": df.copy(), "origen": "Bacalao",
                         "grafico": "Machos y Hembras", "archivo": "bacalao"},
              "Machos": {"datos": df[df["SEXO_ESPECIMEN"] == 1].copy(), "origen": "Machos",
                          "grafico": "Machos", "archivo": "machos"},
              "Hembras": {"datos": df[df["SEXO_ESPECIMEN"] == 2].copy(), "origen": "Hembras",
                           "grafico": "Hembras", "archivo": "hembras"}}
    metodos = {"cook": "Cook", "mad": "MAD", "iqr": "IQR"}
    metodo_bppals = "IQR"
    resultados = {}

    for grupo, info in grupos.items():
        datos = info["datos"]
        a, b, log_a, r = ajustar_e_informar_regresion(datos, info["origen"])
        resultados[grupo] = {"Original": {"datos": datos, "datos_metodo": None,
                                            "a": a, "b": b, "log_a": log_a,
                                            "r": r, "eliminados": 0}}

        for metodo, etiqueta in metodos.items():
            datos_metodo, datos_limpios, eliminados = aplicar_outliers_longpeso(
                datos, metodo, b, log_a)
            a_nuevo, b_nuevo, log_a_nuevo, r_nuevo = ajustar_e_informar_regresion(
                datos_limpios, f"{info['origen']} sin {etiqueta.upper()}")
            resultados[grupo][etiqueta] = {
                "datos": datos_limpios, "datos_metodo": datos_metodo,
                "a": a_nuevo, "b": b_nuevo, "log_a": log_a_nuevo,
                "r": r_nuevo, "eliminados": eliminados}

    # --------------------------------------------------
    # GUARDAR RESULTADOS Y RESPALDOS
    filas_resumen = []

    for grupo in grupos:
        for metodo in ["Original", *metodos.values()]:
            resultado = resultados[grupo][metodo]
            filas_resumen.append([grupo, metodo, resultado["a"], resultado["b"],
                                  resultado["r"], len(resultado["datos"]),
                                  resultado["eliminados"]])

    resultados_regresion = pd.DataFrame(
        filas_resumen, columns=["GRUPO", "METODO", "a", "b", "r_PEARSON",
                                "REGISTROS_UTILIZADOS", "ELIMINADOS"])
    resultados_regresion["r²"] = resultados_regresion["r_PEARSON"] ** 2
    columnas_numericas = ["a", "b", "r_PEARSON", "r²"]
    resultados_regresion[columnas_numericas] = resultados_regresion[
        columnas_numericas].round(7)
    salida_regresiones.parent.mkdir(parents=True, exist_ok=True)
    resultados_regresion.to_csv(salida_regresiones, sep=";", decimal=",", index=False)

    print(f"Archivo generado: {salida_regresiones}")
    salida_excel = (Path(salida_regresiones_excel) if salida_regresiones_excel is not None
                    else salida_regresiones.with_suffix(".xlsx"))
    exportar_tabla_regresiones(resultados, salida_excel)

    # --------------------------------------------------
    # GRÁFICOS DE REGRESIÓN
    for grupo, info in grupos.items():
        for metodo in ["Original", *metodos.values()]:
            resultado = resultados[grupo][metodo]
            sufijo = "" if metodo == "Original" else f"_{metodo.lower()}"
            nombre = info["grafico"] if metodo == "Original" else f"{info['grafico']} sin {metodo}"
            archivo_grafico = directorio_graficos / f"grafica_regresion_{info['archivo']}{sufijo}.png"
            graficar_regresion(resultado["datos"], resultado["a"], resultado["b"],
                               resultado["r"], nombre, archivo_grafico)
            print(f"Archivo generado: {archivo_grafico}")

    # --------------------------------------------------
    # GUARDAR CONJUNTO UTILIZADO POR BPPALS
    guardar_dataframe(resultados["Ambos"][metodo_bppals]["datos"], salida_iqr_ambos)
    guardar_dataframe(resultados["Machos"][metodo_bppals]["datos"], salida_iqr_machos)
    guardar_dataframe(resultados["Hembras"][metodo_bppals]["datos"], salida_iqr_hembras)

    return salida_regresiones
