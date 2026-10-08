import math

import numpy as np
import pandas as pd

# ============================================================
# TABLAS PARA PARÁMETROS a1 Y a2
A1_TABLA = {1.0: 0.00, 1.1: 0.08, 1.2: 0.15, 1.3: 0.23, 1.4: 0.32,
            1.5: 0.42, 1.6: 0.52, 1.7: 0.63, 1.8: 0.74, 1.9: 0.87,
            2.0: 1.00, 2.1: 1.14, 2.2: 1.31, 2.3: 1.48, 2.4: 1.66,
            2.5: 1.86, 2.6: 2.07, 2.7: 2.29, 2.8: 2.52, 2.9: 2.75,
            3.0: 3.00, 3.1: 3.25, 3.2: 3.51, 3.3: 3.77, 3.4: 4.04,
            3.5: 4.32, 3.6: 4.62, 3.7: 4.95, 3.8: 5.29, 3.9: 5.64,
            4.0: 6.00, 4.1: 6.36, 4.2: 6.73, 4.3: 7.11, 4.4: 7.50,
            4.5: 7.90, 4.6: 8.31, 4.7: 8.72, 4.8: 9.14, 4.9: 9.57}

A2_TABLA = {1.0: 0.0, 1.1: 0.0, 1.2: 0.0, 1.3: 0.0, 1.4: 0.0,
            1.5: 0.0, 1.6: 0.0, 1.7: 0.0, 1.8: 0.0, 1.9: 0.0,
            2.0: 0.0, 2.1: 0.0, 2.2: 0.0, 2.3: 0.0, 2.4: 0.0,
            2.5: 0.0, 2.6: 0.0, 2.7: 0.0, 2.8: 0.0, 2.9: 0.0,
            3.0: 0.0, 3.1: 0.2, 3.2: 0.4, 3.3: 0.6, 3.4: 0.9,
            3.5: 1.2, 3.6: 1.5, 3.7: 1.8, 3.8: 2.2, 3.9: 2.6,
            4.0: 3.0, 4.1: 3.5, 4.2: 4.1, 4.3: 4.9, 4.4: 5.9,
            4.5: 7.1, 4.6: 8.4, 4.7: 9.7, 4.8: 11.3, 4.9: 13.0}

# ============================================================
# PARÁMETROS Y RELACIÓN LONGITUD-PESO
def interpolar_tabla(b, tabla):
    b_sup = round(math.ceil(b * 10) / 10, 1)
    b_inf = round(math.floor(b * 10) / 10, 1)

    if b_inf not in tabla or b_sup not in tabla:
        raise ValueError(
            f"b={b:.4f} está fuera del rango de la tabla "
            f"({min(tabla)} a {max(tabla)})")

    if b_sup == b_inf:
        return tabla[b_sup]

    a_sup = tabla[b_sup]
    a_inf = tabla[b_inf]
    return a_sup - ((a_sup - a_inf) * (b_sup - b) / (b_sup - b_inf))

def calcular_a1_a2(b):
    return interpolar_tabla(b, A1_TABLA), interpolar_tabla(b, A2_TABLA)

def calcular_peso(longitud, a, b):
    return a * np.asarray(longitud) ** b

def graficar_regresion(datos, a, b, r, nombre, archivo,columna_longitud="LONGITUD_ESPECIMEN",
                       columna_peso="PESO_ESPECIMEN"):
    import matplotlib.pyplot as plt

    longitudes = np.linspace(datos[columna_longitud].min(),datos[columna_longitud].max(), 500)
    pesos = calcular_peso(longitudes, a, b)

    plt.figure(figsize=(10, 7))
    plt.scatter(datos[columna_longitud], datos[columna_peso],
                alpha=0.4, s=15, label=nombre)
    plt.plot(longitudes, pesos, linewidth=2,
             label=f"W = {a:.5f} L^{b:.4f}\nr = {r:.4f}")
    plt.xlabel("Longitud del espécimen [cm]")
    plt.ylabel("Peso del espécimen [g]")
    plt.title(f"Relación longitud-peso - {nombre}")
    plt.legend()
    plt.grid(alpha=0.3)
    plt.tight_layout()
    plt.savefig(archivo, dpi=300)
    plt.close()
    return archivo

def ajustar_regresion_longpeso(datos, columna_longitud="LONGITUD_ESPECIMEN",
                               columna_peso="PESO_ESPECIMEN"):
    datos = datos[(datos[columna_longitud] > 0) &
                  (datos[columna_peso] > 0)]
    log_longitud = np.log10(datos[columna_longitud])
    log_peso = np.log10(datos[columna_peso])
    b, log_a = np.polyfit(log_longitud, log_peso, 1)
    a = 10 ** log_a
    r = np.corrcoef(log_longitud, log_peso)[0, 1]
    return a, b, log_a, r

def calcular_residuos_longpeso(datos, b, log_a,columna_longitud="LONGITUD_ESPECIMEN",
                               columna_peso="PESO_ESPECIMEN"):
    log_longitud = np.log10(datos[columna_longitud])
    log_peso = np.log10(datos[columna_peso])
    prediccion = log_a + b * log_longitud
    return log_peso - prediccion

# ============================================================
# MÉTODOS PARA OUTLIERS
def calcular_puntaje_mad(valores):
    valores = np.asarray(valores, dtype=float)
    mediana = np.median(valores)
    mad = np.median(np.abs(valores - mediana))

    if mad == 0:
        return np.zeros(len(valores)), mediana, mad

    puntaje = 0.6745 * (valores - mediana) / mad
    return puntaje, mediana, mad

def calcular_limites_iqr(valores, factor=1.5):
    valores = np.asarray(valores, dtype=float)
    q1, q3 = np.percentile(valores, [25, 75])
    iqr = q3 - q1
    limite_inf = q1 - factor * iqr
    limite_sup = q3 + factor * iqr
    return q1, q3, iqr, limite_inf, limite_sup

def calcular_cook_lineal(datos, x="Nº_DE_ANILLOS", y="LONGITUD_DEL_PEZ_(CM)"):
    grupo = datos.copy()
    x_val = grupo[x].to_numpy(dtype=float)
    y_val = grupo[y].to_numpy(dtype=float)
    n = len(grupo)
    p = 2

    if n <= p:
        grupo["COOK"] = 0
        return grupo

    X = np.column_stack((np.ones(n), x_val))
    beta = np.linalg.pinv(X.T @ X) @ X.T @ y_val
    y_pred = X @ beta
    residuos = y_val - y_pred
    sse = np.sum(residuos**2)
    mse = sse / (n - p)

    if mse == 0:
        grupo["COOK"] = 0
        return grupo

    H = X @ np.linalg.pinv(X.T @ X) @ X.T
    h = np.diag(H)
    cook = (residuos**2 / (p * mse)) * (h / np.maximum((1 - h)**2, 1e-12))
    grupo["COOK"] = cook
    return grupo

def calcular_outliers_mad_longpeso(datos, b, log_a,columna_longitud="LONGITUD_ESPECIMEN",
                                   columna_peso="PESO_ESPECIMEN", umbral=3.5):
    datos = datos[(datos[columna_longitud] > 0) & (datos[columna_peso] > 0)].copy()
    residuos = calcular_residuos_longpeso(datos, b, log_a, columna_longitud, columna_peso)
    puntaje, _, mad = calcular_puntaje_mad(residuos)

    if mad == 0:
        datos["MAD"] = 0
        datos["OUTLIER_MAD"] = False
        return datos

    datos["MAD"] = puntaje
    datos["OUTLIER_MAD"] = np.abs(puntaje) > umbral
    return datos

def calcular_outliers_cook_longpeso(datos, b, log_a, columna_longitud="LONGITUD_ESPECIMEN",
                                    columna_peso="PESO_ESPECIMEN"):
    datos = datos[(datos[columna_longitud] > 0) & (datos[columna_peso] > 0)].copy()
    x = np.log10(datos[columna_longitud].to_numpy())
    y = np.log10(datos[columna_peso].to_numpy())
    prediccion = log_a + b * x
    residuos = y - prediccion
    n = len(y)
    p = 2
    mse = np.sum(residuos**2) / (n - p)
    X = np.column_stack((np.ones(n), x))
    H = X @ np.linalg.inv(X.T @ X) @ X.T
    leverage = np.diag(H)
    cook = (residuos**2 / (p * mse)) * (leverage / (1 - leverage)**2)
    datos["COOK"] = cook
    datos["OUTLIER_COOK"] = cook > (4 / n)
    return datos

def calcular_outliers_iqr_longpeso(datos, b, log_a, columna_longitud="LONGITUD_ESPECIMEN",
                                   columna_peso="PESO_ESPECIMEN", factor=3.5):
    datos = datos[(datos[columna_longitud] > 0) & (datos[columna_peso] > 0)].copy()
    residuos = calcular_residuos_longpeso(datos, b, log_a, columna_longitud, columna_peso)
    _, _, _, limite_inf, limite_sup = calcular_limites_iqr(residuos, factor)
    datos["RESIDUO_IQR"] = residuos
    datos["OUTLIER_IQR"] = (residuos < limite_inf) | (residuos > limite_sup)
    return datos

def aplicar_outliers_longpeso(datos, metodo, b, log_a):
    metodo = metodo.lower()

    if metodo == "cook":
        datos_metodo = calcular_outliers_cook_longpeso(datos, b, log_a)
        columna = "OUTLIER_COOK"
    elif metodo == "mad":
        datos_metodo = calcular_outliers_mad_longpeso(datos, b, log_a)
        columna = "OUTLIER_MAD"
    elif metodo == "iqr":
        datos_metodo = calcular_outliers_iqr_longpeso(datos, b, log_a)
        columna = "OUTLIER_IQR"
    else:
        raise ValueError("Método debe ser 'cook', 'mad' o 'iqr'")

    datos_limpios = datos_metodo[~datos_metodo[columna]].copy()
    return datos_metodo, datos_limpios, datos_metodo[columna].sum()

def filtrar_outliers_mad_por_sexo(datos, columna="LONGITUD_DEL_PEZ_(CM)",umbral=2,
                                  columna_sexo="SEXO", sexos=(1, 2)):
    grupos_limpios = []
    resumen = []

    for sexo in sexos:
        grupo = datos[datos[columna_sexo] == sexo].copy()
        puntaje, mediana, mad = calcular_puntaje_mad(grupo[columna])

        if mad == 0:
            grupos_limpios.append(grupo)
            resumen.append((sexo, mediana, mad, len(grupo), 0, len(grupo)))
            continue

        outliers = np.abs(puntaje) > umbral
        grupos_limpios.append(grupo[~outliers].copy())
        resumen.append((sexo, mediana, mad, len(grupo), outliers.sum(),
                        (~outliers).sum()))

    return pd.concat(grupos_limpios, ignore_index=True), resumen

def filtrar_outliers_cook_por_sexo(datos, x="Nº_DE_ANILLOS", y="LONGITUD_DEL_PEZ_(CM)", factor=1.0,
                                   columna_sexo="SEXO", sexos=(1, 2)):
    grupos_limpios = []
    resumen = []

    for sexo in sexos:
        grupo = datos[datos[columna_sexo] == sexo].copy()
        grupo = calcular_cook_lineal(grupo, x=x, y=y)
        n = len(grupo)

        if n == 0:
            continue

        umbral_base = 4 / n
        umbral = umbral_base * factor
        outliers = grupo["COOK"] > umbral
        grupos_limpios.append(grupo[~outliers].drop(columns="COOK"))
        resumen.append((sexo, umbral_base, umbral, n, outliers.sum(),
                        (~outliers).sum()))

    return pd.concat(grupos_limpios, ignore_index=True), resumen

def filtrar_outliers_iqr_por_sexo_edad(datos, columna="LONGITUD_DEL_PEZ_(CM)", columna_sexo="SEXO", columna_edad="GRUPO_EDAD",
                                       factor=3.5, sexos=(1, 2)):
    grupos_limpios = []
    resumen = []

    for sexo in sexos:
        datos_sexo = datos[datos[columna_sexo] == sexo].copy()
        inicial = len(datos_sexo)
        eliminados_total = 0

        for _, grupo in datos_sexo.groupby(columna_edad):
            _, _, _, limite_inf, limite_sup = calcular_limites_iqr(grupo[columna], factor)
            filtro = (grupo[columna] >= limite_inf) & (grupo[columna] <= limite_sup)
            eliminados_total += (~filtro).sum()
            grupos_limpios.append(grupo[filtro].copy())

        resumen.append((sexo, inicial, eliminados_total, inicial - eliminados_total))

    return pd.concat(grupos_limpios, ignore_index=True), resumen
