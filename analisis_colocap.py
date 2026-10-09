# Genera el análisis Colocap desde 3 DBF
from pathlib import Path

import pandas as pd
from lectura import leer_dbf


# ----------------------------
# 0. PROCESO PRINCIPAL
def generar_analisis_colocap(archivo_resumen, archivo_tallas, archivo_sexual, salida,especie,
                             pesqueria, nombre_hoja, titulo_hoja,titulo_bloque, talla_min,
                             talla_max, paso_talla,fila_totales):
    # ----------------------------
    # 1. CONFIGURACIÓN
    salida = Path(salida)

    # ----------------------------
    # 2. LECTURA DBF
    ruta_res = Path(archivo_resumen)
    ruta_tal = Path(archivo_tallas)
    ruta_sex = Path(archivo_sexual)

    df_res = leer_dbf(ruta_res)
    df_tal = leer_dbf(ruta_tal)
    df_sex = leer_dbf(ruta_sex)

    # ----------------------------
    # 3. FILTROS
    df_res_zona = df_res[(df_res["AGRUPACION"] == "ZONA") &
                         (df_res["ESPECIE"] == especie)].copy()
    df_tal_f = df_tal[(df_tal["AGRUPACION"] == "Total Periodo") &
                      (df_tal["ESCALA"] == "Anual") &
                      (df_tal["ESPECIE"] == especie) &
                      (df_tal["SEXO"].isin([1, 2]))].copy()
    df_sex_f = df_sex[(df_sex["AGRUPACION"] == "Total Periodo") &
                      (df_sex["ESCALA"] == "Anual") &
                      (df_sex["ESPECIE"] == especie) &
                      (df_sex["SEXO"].isin([1, 2]))].copy()

    if df_tal_f.empty:
        raise SystemExit("Tallas quedó vacío. Revisa AGRUPACION/ESCALA/SEXO.")
    if df_sex_f.empty:
        raise SystemExit("Sexual quedó vacío. Revisa AGRUPACION/ESCALA/SEXO.")

    # ----------------------------
    # 4. CÁLCULOS
    tal = df_tal_f[["SEXO", "TALLA", "PROPORCION"]].copy()
    tal["TALLA_5CM"] = (tal["TALLA"] // paso_talla) * paso_talla
    tal = tal.sort_values(["SEXO", "TALLA"]).reset_index(drop=True)

    prop_sex1 = float(df_sex_f.loc[df_sex_f["SEXO"] == 1, "PROPORCION"].iloc[0])
    prop_sex2 = float(df_sex_f.loc[df_sex_f["SEXO"] == 2, "PROPORCION"].iloc[0])
    n_ejemp = int(df_res_zona["EJEMP_LONG"].sum())
    bins = sorted(tal["TALLA_5CM"].dropna().unique())
    filas_bins = []

    for b in bins:
        sub = tal[tal["TALLA_5CM"] == b]
        p1 = float(sub.loc[sub["SEXO"] == 1, "PROPORCION"].sum())
        p2 = float(sub.loc[sub["SEXO"] == 2, "PROPORCION"].sum())
        filas_bins.append((int(b), p1, p2))

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
        fmt_pivot_header = wb.add_format({"bold": True, "bg_color": "#D9EAF7",
                                          "border": 1})
        fmt_pivot_total = wb.add_format({"bold": True, "bg_color": "#D9EAF7",
                                         "num_format": "0.0000000000"})

        ws = wb.add_worksheet(nombre_hoja)
        anchos = {"A": 18, "B": 15, "C": 15, "D": 15, "E": 10, "F": 8,
                  "G": 8, "H": 16, "I": 10, "K": 16, "L": 14, "M": 14,
                  "N": 12, "P": 12, "Q": 14, "R": 14, "S": 14, "T": 14,
                  "U": 14, "V": 14, "Y": 20, "Z": 17, "AA": 14, "AB": 16}

        for col, ancho in anchos.items():
            ws.set_column(f"{col}:{col}", ancho)

        # ----------------------------
        # Título y metadatos
        ws.write("A1", titulo_hoja, fmt_titulo)
        ws.write("A2", "Frecuencia pond de tallas Área total", fmt_sub)
        ws.write("A6", "AGRUPACION", fmt_meta_lbl)
        ws.write("B6", "Total periodo")
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
        ws.write("I1", "TALLA 5CM", fmt_header)

        fila = 1
        for _, row in tal.iterrows():
            ws.write_number(fila, 5, int(row["SEXO"]))
            ws.write_number(fila, 6, float(row["TALLA"]))
            ws.write_number(fila, 7, float(row["PROPORCION"]), fmt_num)
            ws.write_formula(fila, 8,
                             f"=+INT(G{fila + 1}/{paso_talla})*{paso_talla}")
            fila += 1

        # ----------------------------
        # Bloque C: proporción sexual K:N
        ws.write("K2", "PROPORCION", fmt_header)
        ws.write("L2", prop_sex1, fmt_prop)
        ws.write("M2", prop_sex2, fmt_prop)
        ws.write("N2", n_ejemp, fmt_int)
        ws.write("K3", "N° EJEMP_SEXO", fmt_header)
        ws.write_formula("L3", "=+L2*$N$2", fmt_prop)
        ws.write_formula("M3", "=+M2*$N$2", fmt_prop)

        # ----------------------------
        # Bloque D: bins de 5 cm P:V
        ws.merge_range("P1:V1", titulo_bloque, fmt_header)
        ws.write("P2", "TALLA (5CM)", fmt_header)
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
            ws.write_formula(fila_bin, 18, f"=+SUM(Q{fila_bin + 1}:R{fila_bin + 1})", fmt_num)
            ws.write_formula(fila_bin, 19, f"=+L$3*Q{fila_bin + 1}", fmt_num)
            ws.write_formula(fila_bin, 20, f"=+M$3*R{fila_bin + 1}", fmt_num)
            ws.write_formula(fila_bin, 21, f"=+SUM(T{fila_bin + 1}:U{fila_bin + 1})", fmt_num)
            fila_bin += 1
        ws.write_formula(f"Q{fila_bin + 1}", f"=+SUM(Q3:Q{fila_bin - 1})")

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
            ws.write_formula(fila - 1, 1, f"=+{p1}*$L$3", fmt_num)
            ws.write_formula(fila - 1, 2, f"=+{p2}*$M$3", fmt_num)
            ws.write_formula(fila - 1, 3, f"=+B{fila}+C{fila}", fmt_num)
            fila += 1

        ws.write(f"A{fila_tot}", "TOTAL", fmt_meta_lbl)
        ws.write_formula(f"B{fila_tot}", f"=+SUM(B{fila_ini}:B{fila_tot - 1})", fmt_num)
        ws.write_formula(f"C{fila_tot}", f"=+SUM(C{fila_ini}:C{fila_tot - 1})", fmt_num)
        ws.write_formula(f"D{fila_tot}", f"=+SUM(D{fila_ini}:D{fila_tot - 1})", fmt_num)

        # ----------------------------
        # Bloque F: reproducción de tabla dinámica Y:AB
        ws.write("Y3", "Suma de PROPORCION", fmt_pivot_header)
        ws.write("Z3", "Etiquetas de columna", fmt_pivot_header)
        ws.write("Y4", "Etiquetas de fila", fmt_pivot_header)
        ws.write("Z4", 1, fmt_pivot_header)
        ws.write("AA4", 2, fmt_pivot_header)
        ws.write("AB4", "Total general", fmt_pivot_header)

        fila_pivot = 4
        for b, p1, p2 in filas_bins:
            ws.write_number(fila_pivot, 24, b, fmt_int)
            ws.write_number(fila_pivot, 25, p1, fmt_num)
            ws.write_number(fila_pivot, 26, p2, fmt_num)
            ws.write_formula(fila_pivot, 27, f"=SUM(Z{fila_pivot + 1}:AA{fila_pivot + 1})", fmt_num)
            fila_pivot += 1

        primera_fila_pivot = 5
        ultima_fila_pivot = fila_pivot
        ws.write(fila_pivot, 24, "Total general", fmt_pivot_total)
        ws.write_formula(fila_pivot, 25,
                         f"=SUM(Z{primera_fila_pivot}:Z{ultima_fila_pivot})", fmt_pivot_total)
        ws.write_formula(fila_pivot, 26,
                         f"=SUM(AA{primera_fila_pivot}:AA{ultima_fila_pivot})", fmt_pivot_total)
        ws.write_formula(fila_pivot, 27,
                         f"=SUM(AB{primera_fila_pivot}:AB{ultima_fila_pivot})", fmt_pivot_total)

        # ----------------------------
        # Hoja Metadatos
        ws_meta = wb.add_worksheet("Metadatos")
        meta = [("Fuente", f"{ruta_res.name}, {ruta_tal.name}, {ruta_sex.name}"),
                ("Filtro Resumen", f"AGRUPACION = ZONA, ESPECIE = {especie}"),
                ("Filtro Tallas", f"AGRUPACION = Total Periodo, ESCALA = Anual, SEXO ∈ {{1,2}}, ESPECIE = {especie}"),
                ("Filtro Sexual", f"AGRUPACION = Total Periodo, ESCALA = Anual, SEXO ∈ {{1,2}}, ESPECIE = {especie}"),
                ("Filas de tallas", len(tal)),("Bins de 5 cm", len(filas_bins)),("N° ejemplares", n_ejemp),
                ("Proporción sexo 1", prop_sex1),("Proporción sexo 2", prop_sex2)]

        for i, (k, v) in enumerate(meta):
            ws_meta.write(i, 0, k, fmt_meta_lbl)
            ws_meta.write(i, 1, v)

        ws_meta.set_column("A:A", 22)
        ws_meta.set_column("B:B", 70)

    print(f"Archivo generado: {salida}")
    return salida
