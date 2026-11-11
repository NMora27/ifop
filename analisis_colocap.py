# Genera el análisis Colocap desde 3 DBF
import math
import warnings
from pathlib import Path

import pandas as pd
from lectura import leer_fuentes, normalizar_texto


class DatosColocapAusentes(ValueError):
    """El dominio no tiene registros o ejemplares para COLOCAP."""


def seleccionar(datos, filtros, nombre):
    faltantes = set(filtros) - set(datos.columns)
    if faltantes:
        raise ValueError(f"{nombre}: faltan columnas {sorted(faltantes)}")
    mascara = pd.Series(True, index=datos.index)
    for columna, esperado in filtros.items():
        valores = esperado if isinstance(esperado, list) else [esperado]
        if columna in ("ESPECIE","PESQUERIA","SEXO","MES"):
            mascara &= pd.to_numeric(datos[columna], errors="coerce").isin(
                pd.to_numeric(pd.Series(valores), errors="raise"))
        else:
            mascara &= datos[columna].map(normalizar_texto).isin(
                [normalizar_texto(v) for v in valores])
    resultado = datos.loc[mascara].copy()
    if resultado.empty:
        raise DatosColocapAusentes(f"{nombre}: selección vacía para {filtros}")
    return resultado


def calcular_colocap(resumen, tallas, sexual, *, especie, pesqueria,
                      paso_talla, filtros_resumen, filtros_tallas, filtros_sexual):
    """Calcula un único dominio anual; no combina proporciones entre dominios."""
    if paso_talla <= 0 or int(paso_talla) != paso_talla:
        raise ValueError("El paso de talla debe ser entero positivo")
    tablas = []
    for nombre, datos, filtros, medidas in [
        ("Resumen", resumen, filtros_resumen, ["EJEMP_LONG"]),
        ("Tallas", tallas, filtros_tallas, ["TALLA","PROPORCION"]),
        ("Sexual", sexual, filtros_sexual, ["PROPORCION"])]:
        requeridas = {"ESPECIE","PESQUERIA","AGRUPACION","ESCALA","PERIODO",
                      "PTO_ZONA","ESTRATO","SEXO", *medidas}
        faltantes = requeridas - set(datos.columns)
        if faltantes:
            raise ValueError(f"{nombre}: faltan columnas {sorted(faltantes)}")
        datos = datos.copy()
        for c in ["ESPECIE","PESQUERIA","SEXO", *medidas]:
            datos[c] = pd.to_numeric(datos[c], errors="coerce")
        if "AGRUPACION" not in filtros:
            raise ValueError(f"{nombre}: declarar AGRUPACION")
        if nombre != "Resumen" and "ESCALA" not in filtros:
            raise ValueError(f"{nombre}: declarar ESCALA anual")
        if any(c in filtros for c in ("ESPECIE","PESQUERIA")):
            raise ValueError("Especie y pesquería se declaran fuera de los filtros")
        filtros = dict(filtros, ESPECIE=especie, PESQUERIA=pesqueria)
        if nombre != "Resumen":
            filtros["SEXO"] = [1,2]
        seleccion = seleccionar(datos, filtros, nombre)
        if nombre != "Resumen":
            if set(seleccion["ESCALA"].map(normalizar_texto)) != {"ANUAL"}:
                raise ValueError("COLOCAP admite proporciones anuales; no promedia meses")
            for campo in ("PERIODO","PTO_ZONA","ESTRATO"):
                if seleccion[campo].fillna("").nunique() != 1:
                    raise ValueError(f"{nombre}: hay varios {campo}; precisar filtros")
        for c in medidas:
            if not seleccion[c].map(lambda v: pd.notna(v) and math.isfinite(v)).all():
                raise ValueError(f"{nombre}: valores inválidos en {c}")
            if seleccion[c].lt(0).any():
                raise ValueError(f"{nombre}: valores negativos en {c}")
        tablas.append(seleccion)
    res, tal, sex = tablas
    flotas = set()
    for tabla in tablas:
        if "FLOTA" in tabla:
            flotas.update(tabla["FLOTA"].dropna().unique())
    if len(flotas) > 1:
        raise ValueError("No se pueden combinar distintas flotas en un proceso")
    if res["ESCALA"].map(normalizar_texto).nunique() != 1:
        raise ValueError("Resumen: no combinar escalas temporales")
    if res["ESTRATO"].fillna("").nunique() != 1:
        raise ValueError("Resumen: seleccionar un único estrato/año")
    if not res["EJEMP_LONG"].mod(1).eq(0).all():
        raise ValueError("Resumen: los conteos deben ser enteros")
    if not res["SEXO"].eq(9).all():
        raise ValueError("Resumen: seleccionar explícitamente el total de sexos (9)")
    llaves = ["ESPECIE","PESQUERIA","AGRUPACION","ESCALA","PERIODO","PTO_ZONA","ESTRATO","SEXO"]
    for nombre, tabla, extra in [("Resumen",res,[]),("Tallas",tal,["TALLA"]),("Sexual",sex,[])]:
        if tabla.duplicated(llaves+extra).any():
            raise ValueError(f"{nombre}: registros duplicados en el dominio seleccionado")
    for c in ("PERIODO","PTO_ZONA","ESTRATO"):
        if normalizar_texto(tal[c].iloc[0]) != normalizar_texto(sex[c].iloc[0]):
            raise ValueError(f"Tallas y Sexual no coinciden en {c}")
    zona = normalizar_texto(tal["PTO_ZONA"].iloc[0])
    if zona and set(res["PTO_ZONA"].map(normalizar_texto)) != {zona}:
        raise ValueError("Resumen y proporciones no corresponden a la misma zona")
    estrato = normalizar_texto(tal["ESTRATO"].iloc[0])
    if estrato and set(res["ESTRATO"].map(normalizar_texto)) != {estrato}:
        raise ValueError("Resumen y proporciones no corresponden al mismo estrato")
    if set(sex["SEXO"]) != {1,2} or len(sex) != 2:
        raise ValueError("Se requiere exactamente una proporción por sexo")
    p = sex.set_index("SEXO")["PROPORCION"]
    if not math.isclose(float(p.sum()),1,abs_tol=0.001):
        raise ValueError("Las proporciones sexuales no suman 1")
    if (sex["PROPORCION"] > 1).any() or (tal["PROPORCION"] > 1).any():
        raise ValueError("Las proporciones deben estar entre 0 y 1")
    for sexo in (1,2):
        total = tal.loc[tal["SEXO"].eq(sexo),"PROPORCION"].sum()
        esperado = 1 if p[sexo] > 0 else 0
        if not math.isclose(float(total),esperado,abs_tol=0.001):
            raise ValueError(f"Proporciones de talla del sexo {sexo}: suma {total}")
    n = float(res["EJEMP_LONG"].sum())
    if n == 0:
        raise DatosColocapAusentes("Resumen: cero ejemplares")
    if n <= 0 or not n.is_integer():
        raise ValueError("El número de ejemplares debe ser entero positivo")
    tal = tal.sort_values(["SEXO","TALLA"]).copy()
    tal["GRUPO_TALLA"] = (tal["TALLA"] // paso_talla)*paso_talla
    agrupadas = tal.groupby(["GRUPO_TALLA","SEXO"])["PROPORCION"].sum().unstack(fill_value=0)
    bins = [(int(t),float(r.get(1,0)),float(r.get(2,0))) for t,r in agrupadas.iterrows()]
    return tal, float(p[1]), float(p[2]), int(n), bins


def generar_analisis_colocap(archivo_resumen, archivo_tallas, archivo_sexual, salida, especie,
                             pesqueria, nombre_hoja, titulo_hoja, titulo_bloque, talla_min,
                             talla_max, paso_talla, fila_totales=None, *,
                             flota="Sin especificar", filtros_resumen=None,
                             filtros_tallas=None, filtros_sexual=None,
                             respaldo_sin_datos=None):
    salida = Path(salida)
    ruta_res, ruta_tal, ruta_sex = map(Path, (archivo_resumen,archivo_tallas,archivo_sexual))
    def leer(ruta):
        return leer_fuentes([{"archivo":str(ruta), "flota":flota}])
    filtros_resumen = filtros_resumen or {"AGRUPACION":"ZONA","SEXO":9}
    filtros_tallas = filtros_tallas or {"AGRUPACION":"Total Periodo","ESCALA":"Anual"}
    filtros_sexual = filtros_sexual or {"AGRUPACION":"Total Periodo","ESCALA":"Anual"}
    fuentes = [leer(ruta_res), leer(ruta_tal), leer(ruta_sex)]
    solicitados = dict(filtros_resumen=filtros_resumen, filtros_tallas=filtros_tallas,
                       filtros_sexual=filtros_sexual)
    efectivos = solicitados
    sustitucion = ""
    try:
        resultado = calcular_colocap(*fuentes, especie=especie, pesqueria=pesqueria,
                                    paso_talla=paso_talla, **efectivos)
    except DatosColocapAusentes:
        if respaldo_sin_datos is None:
            raise
        efectivos = {campo: respaldo_sin_datos[campo] for campo in solicitados}
        resultado = calcular_colocap(*fuentes, especie=especie, pesqueria=pesqueria,
                                    paso_talla=paso_talla, **efectivos)
        sustitucion = respaldo_sin_datos['descripcion']
        warnings.warn(f'{nombre_hoja}: {sustitucion}', stacklevel=2)
    tal, prop_sex1, prop_sex2, n_ejemp, filas_bins = resultado
    if talla_min > talla_max or talla_min % paso_talla or talla_max % paso_talla:
        raise ValueError("Los límites deben ser marcas inferiores alineadas con el paso")
    fuera = [(b,p1,p2) for b,p1,p2 in filas_bins
             if (b<talla_min or b>talla_max) and p1+p2>0]
    if fuera:
        raise ValueError(f"Hay tallas con proporción positiva fuera del rango: {fuera}")
    fila_automatica = 13 + len(range(talla_min,talla_max+1,paso_talla))
    if fila_totales is not None and fila_totales < fila_automatica:
        raise ValueError("La fila de totales se superpone a las tallas")
    fila_totales = fila_totales or fila_automatica
    # ----------------------------
    # 5. ESCRITURA EXCEL
    salida.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_titulo = wb.add_format({"bold": True, "font_size": 13})
        fmt_sub = wb.add_format({"italic": True, "font_size": 10})
        fmt_header = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                                    "font_color": "white", "border": 1,
                                    "align": "center", "valign": "vcenter"})
        fmt_meta_lbl = wb.add_format({"bold": True})
        fmt_num = wb.add_format({"num_format": "0.0000000000"})
        fmt_int = wb.add_format({"num_format": "#,##0"})
        fmt_prop = wb.add_format({"num_format": "0.00000000"})

        ws = wb.add_worksheet(nombre_hoja)
        anchos = {"A": 18, "B": 15, "C": 15, "D": 15, "E": 10, "F": 8,
                  "G": 8, "H": 16, "I": 10, "K": 16, "L": 14, "M": 14,
                  "N": 12, "P": 12, "Q": 14, "R": 14, "S": 14, "T": 14,
                  "U": 14, "V": 14}

        for col, ancho in anchos.items():
            ws.set_column(f"{col}:{col}", ancho)

        # ----------------------------
        # Título y metadatos
        ws.write("A1", titulo_hoja, fmt_titulo)
        ws.write("A2", f"Frecuencia ponderada — {flota}", fmt_sub)
        ws.write("A6", "AGRUPACION", fmt_meta_lbl)
        ws.write("B6", filtros_tallas["AGRUPACION"])
        ws.write("A7", "PESQUERIA", fmt_meta_lbl)
        ws.write("B7", pesqueria)
        ws.write("A8", "ESCALA", fmt_meta_lbl)
        ws.write("B8", "Anual")
        ws.write("A9", "PERIODO", fmt_meta_lbl)
        ws.write("B9", "Ene - Dic")
        ws.write("A10", "ESPECIE", fmt_meta_lbl)
        ws.write("B10", especie)

        # ----------------------------
        # Bloque B: frecuencia de tallas F:I
        ws.write("F1", "SEXO", fmt_header)
        ws.write("G1", "TALLA", fmt_header)
        ws.write("H1", "PROPORCION", fmt_header)
        ws.write("I1", f"TALLA {paso_talla}CM", fmt_header)

        fila = 1
        for _, row in tal.iterrows():
            ws.write_number(fila, 5, int(row["SEXO"]))
            ws.write_number(fila, 6, float(row["TALLA"]))
            ws.write_number(fila, 7, float(row["PROPORCION"]), fmt_num)
            ws.write_formula(fila, 8,
                             f"=+INT(G{fila + 1}/{paso_talla})*{paso_talla}", None, float(row["GRUPO_TALLA"]))
            fila += 1

        # ----------------------------
        # Bloque C: proporción sexual K:N
        ws.write("K2", "PROPORCION", fmt_header)
        ws.write("L2", prop_sex1, fmt_prop)
        ws.write("M2", prop_sex2, fmt_prop)
        ws.write("N2", n_ejemp, fmt_int)
        ws.write("K3", "N° EJEMP_SEXO", fmt_header)
        ws.write_formula("L3", "=+L2*$N$2", fmt_prop, prop_sex1*n_ejemp)
        ws.write_formula("M3", "=+M2*$N$2", fmt_prop, prop_sex2*n_ejemp)

        # ----------------------------
        # Bloque D: bins de 5 cm P:V
        ws.merge_range("P1:V1", titulo_bloque, fmt_header)
        ws.write("P2", f"TALLA ({paso_talla}CM)", fmt_header)
        ws.write("Q2", "sex1", fmt_header)
        ws.write("R2", "sex2", fmt_header)
        ws.write("S2", "TOTAL", fmt_header)
        ws.write("T2", "sex1", fmt_header)
        ws.write("U2", "sex2", fmt_header)
        ws.write("V2", "TOTAL", fmt_header)

        fila_bin = 2
        for b, p1, p2 in filas_bins:
            ws.write_number(fila_bin, 15, b, fmt_int)
            ws.write_number(fila_bin, 16, p1, fmt_num)
            ws.write_number(fila_bin, 17, p2, fmt_num)
            ws.write_formula(fila_bin, 18, f"=+SUM(Q{fila_bin + 1}:R{fila_bin + 1})", fmt_num, p1+p2)
            ws.write_formula(fila_bin, 19, f"=+L$3*Q{fila_bin + 1}", fmt_num, n_ejemp*prop_sex1*p1)
            ws.write_formula(fila_bin, 20, f"=+M$3*R{fila_bin + 1}", fmt_num, n_ejemp*prop_sex2*p2)
            ws.write_formula(fila_bin, 21, f"=+SUM(T{fila_bin + 1}:U{fila_bin + 1})", fmt_num, n_ejemp*(prop_sex1*p1+prop_sex2*p2))
            fila_bin += 1
        ws.write_formula(f"Q{fila_bin + 1}", f"=+SUM(Q3:Q{fila_bin})", None, sum(p1 for _,p1,_ in filas_bins))

        # ----------------------------
        # Bloque E: cantidades por rango de talla A:D
        ws.write("A12", "Etiquetas de fila", fmt_meta_lbl)
        ws.write("B12", 1, fmt_header)
        ws.write("C12", 2, fmt_header)
        ws.write("D11", "Total", fmt_header)

        paso = paso_talla
        fila_ini = 13
        fila_tot = fila_totales
        prop_por_bin = {b: (p1, p2) for b, p1, p2 in filas_bins}

        fila = fila_ini
        for b in range(talla_min, talla_max + 1, paso):
            p1, p2 = prop_por_bin.get(b, (0.0, 0.0))
            ws.write_number(fila - 1, 0, b, fmt_int)
            ws.write_formula(fila - 1, 1, f"=+{p1}*$L$3", fmt_num, p1*prop_sex1*n_ejemp)
            ws.write_formula(fila - 1, 2, f"=+{p2}*$M$3", fmt_num, p2*prop_sex2*n_ejemp)
            ws.write_formula(fila - 1, 3, f"=+B{fila}+C{fila}", fmt_num, n_ejemp*(p1*prop_sex1+p2*prop_sex2))
            fila += 1

        ws.write(f"A{fila_tot}", "TOTAL", fmt_meta_lbl)
        ws.write_formula(f"B{fila_tot}", f"=+SUM(B{fila_ini}:B{fila_tot - 1})", fmt_num, n_ejemp*prop_sex1*sum(p1 for _,p1,_ in filas_bins))
        ws.write_formula(f"C{fila_tot}", f"=+SUM(C{fila_ini}:C{fila_tot - 1})", fmt_num, n_ejemp*prop_sex2*sum(p2 for _,_,p2 in filas_bins))
        ws.write_formula(f"D{fila_tot}", f"=+SUM(D{fila_ini}:D{fila_tot - 1})", fmt_num, n_ejemp*(prop_sex1*sum(p1 for _,p1,_ in filas_bins)+prop_sex2*sum(p2 for _,_,p2 in filas_bins)))

        # ----------------------------
        # Hoja Metadatos
        ws_meta = wb.add_worksheet("Metadatos")
        meta = [("Fuente", f"{ruta_res.resolve()}, {ruta_tal.resolve()}, {ruta_sex.resolve()}"),
                ("Flota",flota),("Especie",especie),("Pesqueria",pesqueria),
                ("Filtro Resumen",str(efectivos['filtros_resumen'])),
                ("Filtro Tallas",str(efectivos['filtros_tallas'])),
                ("Filtro Sexual",str(efectivos['filtros_sexual'])),("Paso talla (cm)",paso_talla),
                ("Filas tallas",len(tal)),("Intervalos",len(filas_bins)),
                ("N° ejemplares",n_ejemp),("Proporción sexo 1",prop_sex1),
                ("Proporción sexo 2",prop_sex2)]
        meta.extend([("Sustitución por falta de datos", sustitucion or "No aplicada"),
                     ("Filtros solicitados", str(solicitados))])
        if sustitucion:
            ws.write("A2", sustitucion, fmt_sub)

        for i, (k, v) in enumerate(meta):
            ws_meta.write(i, 0, k, fmt_meta_lbl)
            ws_meta.write(i, 1, v)

        ws_meta.set_column("A:A", 22)
        ws_meta.set_column("B:B", 70)

    print(f"Archivo generado: {Path(salida).name}")
    return salida
