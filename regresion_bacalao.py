from pathlib import Path

import pandas as pd
from lectura import leer_csv_palangre
from metodos_matematicos import (
    ajustar_regresion_longpeso,
    aplicar_outliers_longpeso,
    graficar_regresion,
)


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


# ============================================================
# PROCESO PRINCIPAL
def generar_regresiones_bacalao(archivo, especie, latitud_zona_1, latitud_zona_2,
                                directorio_graficos, salida_preprocesamiento,
                                salida_regresiones, salida_iqr_ambos,
                                salida_iqr_machos, salida_iqr_hembras):
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
    df, n_original = leer_csv_palangre(archivo, columnas)

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
    resultados_regresion[["a", "b", "r_PEARSON"]] = resultados_regresion[
        ["a", "b", "r_PEARSON"]].round(5)
    salida_regresiones.parent.mkdir(parents=True, exist_ok=True)
    resultados_regresion.to_csv(salida_regresiones, sep=";", decimal=",", index=False)

    print(f"Archivo generado: {salida_regresiones}")

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
