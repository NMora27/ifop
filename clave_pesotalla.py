"""Matrices talla–edad compartidas por especie; limpieza por sexo y edad agrupada."""
from pathlib import Path
import numpy as np
import pandas as pd
import warnings
from openpyxl.styles import Alignment
from lectura import leer_datos_talla_edad, leer_perfil


def depurar_muestra(datos, *, columna_edad, edad_agrupada=24, metodo="MAD", umbral=3.5):
    """Marca exclusiones por sexo y edad, antes de crear o combinar claves."""
    datos = datos.copy()
    edad = datos[columna_edad]
    datos["GRUPO_EDAD"] = edad.astype(int).astype(str)
    datos.loc[edad.ge(edad_agrupada), "GRUPO_EDAD"] = f"{edad_agrupada}+"
    datos["EXCLUIDO_ESTADISTICO"] = False
    for _, grupo in datos.groupby(["SEXO", "GRUPO_EDAD"]):
        x = grupo["LONGITUD_DEL_PEZ_(CM)"]
        if metodo == "MAD":
            mad = (x - x.median()).abs().median()
            fuera = (0.6745 * (x - x.median()) / mad).abs().gt(umbral) if mad else x.ne(x)
        elif metodo == "IQR":
            q1, q3 = x.quantile([.25, .75])
            fuera = x.lt(q1 - umbral * (q3-q1)) | x.gt(q3 + umbral * (q3-q1))
        else:
            fuera = x.ne(x)
        datos.loc[grupo.index, "EXCLUIDO_ESTADISTICO"] = fuera
    return datos


def crear_matriz(datos, edad_agrupada, talla_min, talla_max, paso_talla, edad_min=1):
    edades = [str(i) for i in range(edad_min, edad_agrupada)] + [f"{edad_agrupada}+"]
    matriz = pd.crosstab(datos["GRUPO_TALLA"], datos["GRUPO_EDAD"])
    matriz = matriz.reindex(index=range(talla_min, talla_max + 1, paso_talla),
                            columns=edades, fill_value=0)
    matriz.index.name = "GRUPO_TALLA"
    matriz["Total general"] = matriz.sum(axis=1)
    matriz.loc["Total general"] = matriz.sum(axis=0)
    return matriz.replace(0, "")


def generar_matriz_talla_edad(archivo, salida, pesqueria, paso_talla, edad_agrupada,
                              talla_min, talla_max, anio=None, especie=None, zona=None,
                              *, edad_min=1, columna_edad=None, hoja=0,
                              metodo="MAD", umbral=3.5, columna_especie="ESPECIE",
                              valor_especie=None, puerto=None, zona_respaldo_sin_datos=None,
                              puerto_respaldo_sin_datos=None, mezclar_arrastre_si_menos_de=None,
                              mezclar_otra_zona_si_menos_de=None,
                              mezcla_secuencial_minimo=None, salida_preliminar=None,
                              depurar_antes_mezcla=True, omitir_si_vacia=False):
    """Los límites de talla son inferiores de clase; la edad final acumula edades mayores."""
    valores = (paso_talla, edad_agrupada, talla_min, talla_max, edad_min)
    if any(not isinstance(v, int) or isinstance(v, bool) for v in valores):
        raise ValueError("Los límites y el paso deben ser enteros")
    if (paso_talla <= 0 or talla_min > talla_max or talla_min % paso_talla
            or not 1 <= edad_min <= edad_agrupada):
        raise ValueError("Rangos de talla o edad inválidos")
    metodo = metodo.upper()
    if metodo not in ("MAD", "IQR", "NINGUNO") or not np.isfinite(umbral) or umbral <= 0:
        raise ValueError("Método o umbral inválido")
    if (anio is None or int(anio) < 2026) and pesqueria is None and puerto is None and especie != 37:
        raise ValueError("Debe indicar pesqueria o puerto para el formato histórico")
    if mezclar_arrastre_si_menos_de is not None:
        if (not isinstance(mezclar_arrastre_si_menos_de, int)
                or isinstance(mezclar_arrastre_si_menos_de, bool)
                or mezclar_arrastre_si_menos_de <= 0):
            raise ValueError("El umbral de mezcla debe ser un entero positivo")
        if puerto not in (70, 71) or pesqueria is not None or zona is not None:
            raise ValueError("La mezcla requiere PUERTO 70 o 71, sin filtros de pesquería ni zona")
    otras_zonas = {61: [64], 64: [61], 70: [71], 71: [70],
                   76: [78, 79], 78: [76], 79: [76]}
    if mezclar_otra_zona_si_menos_de is not None:
        if (not isinstance(mezclar_otra_zona_si_menos_de, int)
                or isinstance(mezclar_otra_zona_si_menos_de, bool)
                or mezclar_otra_zona_si_menos_de <= 0):
            raise ValueError("El umbral de mezcla entre zonas debe ser entero positivo")
        puertos_propios = list(puerto) if isinstance(puerto, (list, tuple)) else [puerto]
        if (any(p not in otras_zonas for p in puertos_propios)
                or pesqueria is not None or zona is not None
                or mezclar_arrastre_si_menos_de is not None
                or puerto_respaldo_sin_datos is not None
                or zona_respaldo_sin_datos is not None):
            raise ValueError("Mezcla entre zonas requiere puertos configurados, sin otras mezclas ni respaldos")
        puertos_donantes = sorted({q for p in puertos_propios for q in otras_zonas[p]})
        if set(puertos_propios) & set(puertos_donantes):
            raise ValueError("Los puertos propios y donantes deben ser distintos")
    if mezcla_secuencial_minimo is not None:
        if (not isinstance(mezcla_secuencial_minimo, int)
                or isinstance(mezcla_secuencial_minimo, bool)
                or mezcla_secuencial_minimo <= 0):
            raise ValueError("El mínimo de mezcla secuencial debe ser entero positivo")
        propios = list(puerto) if isinstance(puerto, (list, tuple)) else [puerto]
        if (any(p not in otras_zonas for p in propios)
                or pesqueria is not None or zona is not None
                or mezclar_arrastre_si_menos_de is not None
                or mezclar_otra_zona_si_menos_de is not None
                or puerto_respaldo_sin_datos is not None
                or zona_respaldo_sin_datos is not None):
            raise ValueError("La mezcla secuencial requiere puertos propios sin otras mezclas ni respaldos")
        etapas_puertos = []
        if any(p in (61, 64, 70, 71) for p in propios):
            otro_arte = {61: 70, 70: 61, 64: 71, 71: 64}
            if len(propios) != 1:
                raise ValueError("La mezcla industrial requiere un único puerto propio")
            etapas_puertos.append(("otro_arte_misma_zona", [otro_arte[propios[0]]]))
        etapas_puertos.append(("mismo_arte_otra_zona",
                               sorted({q for p in propios for q in otras_zonas[p]})))
        if any(p in (61, 64, 70, 71) for p in propios):
            puerto_otra_zona = otras_zonas[propios[0]][0]
            etapas_puertos.append(("otro_arte_otra_zona", [otro_arte[puerto_otra_zona]]))
        usados = set(propios)
        for _, puertos_etapa in etapas_puertos:
            if usados.intersection(puertos_etapa):
                raise ValueError("Los puertos de las etapas de mezcla deben ser distintos")
            usados.update(puertos_etapa)
    if salida_preliminar is not None:
        if Path(salida_preliminar).resolve() == Path(salida).resolve():
            raise ValueError("Las salidas preliminar y definitiva deben ser diferentes")
        # La preliminar usa la muestra propia depurada, sin incorporar donantes.
        generar_matriz_talla_edad(
            archivo, salida_preliminar, pesqueria, paso_talla, edad_agrupada,
            talla_min, talla_max, anio, especie, zona, edad_min=edad_min,
            columna_edad=columna_edad, hoja=hoja, metodo=metodo, umbral=umbral,
            columna_especie=columna_especie, valor_especie=valor_especie, puerto=puerto,
            depurar_antes_mezcla=depurar_antes_mezcla, omitir_si_vacia=True)
    longitud, edad = "LONGITUD_DEL_PEZ_(CM)", "Nº_DE_ANILLOS"
    def depurar_dominio(datos):
        return depurar_muestra(datos, columna_edad=edad, edad_agrupada=edad_agrupada,
                               metodo=metodo, umbral=umbral)

    def contar_retenidos(muestra):
        return int((~muestra["EXCLUIDO_ESTADISTICO"]).sum()) if depurar_antes_mezcla else len(muestra)

    def leer_dominio(zona_fuente, puerto_fuente):
        columnas = [longitud, edad, 'SEXO'] if especie == 37 and pesqueria is None and puerto_fuente is None else None
        datos = leer_datos_talla_edad(archivo, columnas=columnas, anio=anio, especie=especie, zona=zona_fuente,
                                     columna_edad=columna_edad, hoja=hoja,
                                     columna_especie=columna_especie, valor_especie=valor_especie,
                                     puerto=puerto_fuente)
        if (anio is None or int(anio) < 2026) and pesqueria is not None:
            if "CÓDIGO_DE_PESQUERÍA" not in datos:
                raise ValueError("Para filtrar por puerto use pesqueria=None")
            datos = datos[datos["CÓDIGO_DE_PESQUERÍA"].eq(pesqueria)].copy()
        validos = (datos["SEXO"].isin([1, 2]) & np.isfinite(datos[longitud])
                   & np.isfinite(datos[edad]) & datos[longitud].gt(0)
                   & datos[edad].ge(edad_min) & datos[edad].mod(1).eq(0))
        validos_dominio = datos[validos].copy()
        return depurar_dominio(validos_dominio) if depurar_antes_mezcla else validos_dominio
    datos = leer_dominio(zona, puerto)
    puerto_fuente = puerto
    zona_fuente = zona
    sustitucion = "No aplicada"
    if datos.empty and (zona_respaldo_sin_datos is not None or puerto_respaldo_sin_datos is not None):
        zona_fuente = zona_respaldo_sin_datos if zona_respaldo_sin_datos is not None else zona
        puerto_fuente = puerto_respaldo_sin_datos if puerto_respaldo_sin_datos is not None else puerto
        datos = leer_dominio(zona_fuente, puerto_fuente)
        if datos.empty:
            raise ValueError("La zona de respaldo tampoco tiene datos válidos para talla-edad")
        sustitucion = (f"Clave de PUERTO {puerto_fuente}, zona {zona_fuente}, "
                       f"por ausencia de datos válidos de PUERTO {puerto}, zona {zona}")
        warnings.warn(sustitucion, stacklevel=2)
    mezcla = {}
    if mezclar_arrastre_si_menos_de is not None:
        puerto_arrastre = {70: 61, 71: 64}[puerto]
        arrastre = leer_dominio(None, puerto_arrastre)
        partes = [datos]
        for sexo in (1, 2):
            n_propio = contar_retenidos(datos[datos["SEXO"].eq(sexo)])
            agregar = arrastre[arrastre["SEXO"].eq(sexo)].copy()
            aplicado = n_propio < mezclar_arrastre_si_menos_de
            if aplicado:
                partes.append(agregar)
            mezcla[sexo] = dict(palangre=n_propio, puerto_arrastre=puerto_arrastre,
                               agregados=contar_retenidos(agregar) if aplicado else 0, aplicada=aplicado)
        # Índices únicos para depurar cada registro de la muestra combinada una sola vez.
        datos = pd.concat(partes, ignore_index=True)
    if mezclar_otra_zona_si_menos_de is not None:
        otra_zona = leer_dominio(None, puertos_donantes)
        partes = [datos]
        for sexo in (1, 2):
            n_propio = contar_retenidos(datos[datos["SEXO"].eq(sexo)])
            agregar = otra_zona[otra_zona["SEXO"].eq(sexo)].copy()
            aplicado = n_propio < mezclar_otra_zona_si_menos_de
            if aplicado:
                partes.append(agregar)
            mezcla[sexo] = dict(propios=n_propio, puertos_donantes=puertos_donantes,
                               agregados=contar_retenidos(agregar) if aplicado else 0, aplicada=aplicado)
        datos = pd.concat(partes, ignore_index=True)
    if mezcla_secuencial_minimo is not None:
        # Cada etapa lee exclusivamente registros originales de puertos distintos.
        donantes = [(nombre, puertos_etapa, leer_dominio(None, puertos_etapa))
                    for nombre, puertos_etapa in etapas_puertos]
        partes = []
        for sexo in (1, 2):
            propios_sexo = datos[datos["SEXO"].eq(sexo)].copy()
            partes.append(propios_sexo)
            acumulados = contar_retenidos(propios_sexo)
            etapas = []
            for nombre, puertos_etapa, muestra in donantes:
                agregar = muestra[muestra["SEXO"].eq(sexo)].copy()
                aplicada = acumulados < mezcla_secuencial_minimo
                n_agregados = contar_retenidos(agregar) if aplicada else 0
                etapas.append(dict(etapa=nombre, puertos=puertos_etapa,
                                   antes=acumulados, disponibles=contar_retenidos(agregar),
                                   aplicada=aplicada, agregados=n_agregados))
                if aplicada:
                    partes.append(agregar)
                    acumulados += n_agregados
            mezcla[sexo] = dict(propios=contar_retenidos(propios_sexo), etapas=etapas,
                               total_despues_MAD=acumulados,
                               alcanza_minimo=acumulados >= mezcla_secuencial_minimo)
        datos = pd.concat(partes, ignore_index=True)
    if not depurar_antes_mezcla:
        datos = depurar_dominio(datos)
    datos["GRUPO_TALLA"] = (np.floor(datos[longitud] / paso_talla) * paso_talla).astype(int)
    limpios = datos[~datos["EXCLUIDO_ESTADISTICO"]].copy()
    if not limpios["GRUPO_TALLA"].between(talla_min, talla_max).all():
        raise ValueError("Hay tallas retenidas fuera del rango de la matriz")
    salida = Path(salida)
    if omitir_si_vacia and limpios.empty:
        # Evitar que una preliminar de una ejecución anterior aparente tener datos actuales.
        salida.unlink(missing_ok=True)
        return None
    salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for sexo, nombre in [(1, "Machos"), (2, "Hembras")]:
            crear_matriz(limpios[limpios["SEXO"].eq(sexo)], edad_agrupada,
                         talla_min, talla_max, paso_talla, edad_min).to_excel(writer, sheet_name=nombre)
            ws = writer.sheets[nombre]
            for columna in ws.columns:
                ws.column_dimensions[columna[0].column_letter].width = 14 if columna[0].column in (1, ws.max_column) else 6
                for celda in columna:
                    celda.alignment = Alignment(horizontal="center", vertical="center")
        pd.DataFrame(list(dict(archivo=str(Path(archivo).resolve()), anio=anio, especie=especie,
            pesqueria=pesqueria, zona=zona, puerto=puerto, columna_edad=columna_edad, hoja=hoja,
            zona_fuente=zona_fuente, puerto_fuente=puerto_fuente,
            puerto_respaldo_sin_datos=puerto_respaldo_sin_datos,
            sustitucion_sin_datos=sustitucion,
            salida_preliminar=str(salida_preliminar) if salida_preliminar is not None else None,
            mezcla_secuencial_minimo=mezcla_secuencial_minimo,
            mezcla_secuencial_por_sexo=str(mezcla) if mezcla_secuencial_minimo is not None else "{}",
            mezcla_arrastre_umbral=mezclar_arrastre_si_menos_de,
            mezcla_arrastre_por_sexo=str(mezcla) if mezclar_arrastre_si_menos_de is not None else "{}",
            mezcla_otra_zona_umbral=mezclar_otra_zona_si_menos_de,
            mezcla_otra_zona_por_sexo=str(mezcla) if mezclar_otra_zona_si_menos_de is not None else "{}",
            columna_especie=("COS_ESPECIE" if anio is not None and int(anio) >= 2026 else columna_especie),
            valor_especie=(especie if anio is not None and int(anio) >= 2026 else
                           valor_especie if valor_especie is not None else especie),
            depuracion_antes_mezcla=depurar_antes_mezcla,
            metodo=metodo, umbral=umbral, talla_min=talla_min, talla_max=talla_max,
            paso_talla=paso_talla, edad_min=edad_min, edad_agrupada=edad_agrupada,
            registros_validos=len(datos), eliminados=int(datos["EXCLUIDO_ESTADISTICO"].sum()),
            retenidos=len(limpios)).items()), columns=["Parametro", "Valor"]).to_excel(writer, sheet_name="Metadatos", index=False)
        datos.to_excel(writer, sheet_name="Respaldo", index=False)
    return salida


def generar_claves_perfil(archivo_perfil):
    """Genera preliminares y definitivas usando las rutas y reglas del perfil."""
    perfil, base = leer_perfil(archivo_perfil)
    resultados = []
    for proceso in perfil["talla_edad"]:
        config = dict(proceso)
        nombre = config.pop("id")
        for campo in ("archivo", "salida", "salida_preliminar"):
            if config.get(campo) is not None:
                config[campo] = str(base / config[campo])
        config.setdefault("anio", perfil["anio"])
        config.setdefault("especie", perfil["especie_codigo"])
        resultados.append(generar_matriz_talla_edad(**config))
        print(f"Claves generadas: {nombre}", flush=True)
    return resultados


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("perfil", help="Perfil JSON con los procesos talla_edad")
    generar_claves_perfil(parser.parse_args().perfil)
