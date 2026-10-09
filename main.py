"""Ejecuta el flujo completo de automatización para bacalao de profundidad."""
from pathlib import Path

from analisis_colocap import generar_analisis_colocap
from bppals_20XX import generar_bppals
from clave_pesotalla import generar_matriz_talla_edad
from regresion_bacalao import generar_regresiones_bacalao


# ============================================================
# LECTURA DE CONFIGURACIÓN
def leer_configuracion(archivo):
    configuracion = {}

    for numero, linea in enumerate(Path(archivo).read_text(encoding="utf-8").splitlines(), 1):
        linea = linea.strip()

        if not linea or linea.startswith("#"):
            continue
        if "=" not in linea:
            raise ValueError(f"Configuración inválida en línea {numero}: {linea}")

        clave, valor = linea.split("=", 1)
        clave = clave.strip().upper()
        valor = valor.strip()

        if not clave:
            raise ValueError(f"Falta la clave en la línea {numero}")

        configuracion[clave] = valor

    for _ in range(len(configuracion)):
        anterior = configuracion.copy()

        try:
            configuracion = {clave: valor.format_map(configuracion)
                             for clave, valor in configuracion.items()}
        except KeyError as error:
            raise ValueError(
                f"La configuración usa una clave no definida: {error.args[0]}") from error

        if configuracion == anterior:
            break

    return configuracion

def resolver_ruta(base, valor):
    ruta = Path(valor)
    return ruta if ruta.is_absolute() else Path(base) / ruta

def leer_entero(configuracion, clave):
    return int(configuracion[clave])

def leer_booleano(configuracion, clave):
    valor = configuracion[clave].strip().upper()

    if valor in ("1", "SI", "SÍ", "TRUE", "VERDADERO"):
        return True
    if valor in ("0", "NO", "FALSE", "FALSO"):
        return False

    raise ValueError(f"{clave} debe ser SI o NO")


# ============================================================
# PROCESO PRINCIPAL
def main(archivo_configuracion=None):
    base = Path(__file__).resolve().parent
    archivo_configuracion = archivo_configuracion or base / "configuracion.txt"
    configuracion = leer_configuracion(archivo_configuracion)
    anio = leer_entero(configuracion, "ANIO")
    ejecutar_bppals = leer_booleano(configuracion, "EJECUTAR_BPPALS")

    salida_colocap = generar_analisis_colocap(
        archivo_resumen=resolver_ruta(base, configuracion["ARCHIVO_RESUMEN"]),
        archivo_tallas=resolver_ruta(base, configuracion["ARCHIVO_TALLAS"]),
        archivo_sexual=resolver_ruta(base, configuracion["ARCHIVO_SEXUAL"]),
        salida=resolver_ruta(base, configuracion["SALIDA_COLOCAP"]),
        especie=leer_entero(configuracion, "ESPECIE_CODIGO"),
        pesqueria=leer_entero(configuracion, "PESQUERIA_COLOCAP"),
        nombre_hoja=configuracion["NOMBRE_HOJA_COLOCAP"],
        titulo_hoja=configuracion["TITULO_HOJA_COLOCAP"],
        titulo_bloque=configuracion["TITULO_BLOQUE_COLOCAP"],
        talla_min=leer_entero(configuracion, "TALLA_COLOCAP_MIN"),
        talla_max=leer_entero(configuracion, "TALLA_COLOCAP_MAX"),
        paso_talla=leer_entero(configuracion, "PASO_TALLA_COLOCAP"),
        fila_totales=leer_entero(configuracion, "FILA_TOTALES_COLOCAP"))

    salida_regresiones = generar_regresiones_bacalao(
        archivo=resolver_ruta(base, configuracion["ARCHIVO_PALANGRE"]),
        especie=leer_entero(configuracion, "ESPECIE_CODIGO"),
        latitud_zona_1=float(configuracion["LATITUD_ZONA_1"]),
        latitud_zona_2=float(configuracion["LATITUD_ZONA_2"]),
        directorio_graficos=resolver_ruta(base, configuracion["DIRECTORIO_GRAFICOS"]),
        salida_preprocesamiento=resolver_ruta(base, configuracion["SALIDA_PREPROCESAMIENTO"]),
        salida_regresiones=resolver_ruta(base, configuracion["SALIDA_REGRESIONES"]),
        salida_iqr_ambos=resolver_ruta(base, configuracion["SALIDA_IQR_AMBOS"]),
        salida_iqr_machos=resolver_ruta(base, configuracion["SALIDA_IQR_MACHOS"]),
        salida_iqr_hembras=resolver_ruta(base, configuracion["SALIDA_IQR_HEMBRAS"]))

    salida_matriz = generar_matriz_talla_edad(
        archivo=resolver_ruta(base, configuracion["ARCHIVO_TALLA_EDAD"]),
        salida=resolver_ruta(base, configuracion["SALIDA_MATRIZ"]),
        pesqueria=leer_entero(configuracion, "PESQUERIA_TALLA_EDAD"),
        paso_talla=leer_entero(configuracion, "PASO_TALLA_MATRIZ"),
        edad_agrupada=leer_entero(configuracion, "EDAD_AGRUPADA"))

    if ejecutar_bppals:
        generar_bppals(
            especie=configuracion["ESPECIE"], sexo=configuracion["SEXO"], anio=anio,
            zona=configuracion["ZONA"], area=configuracion["AREA"],
            flota_desembarque=configuracion["FLOTA_DESEMBARQUE"],
            archivo_matriz=salida_matriz, archivo_regresion=salida_regresiones,
            archivo_colocap=salida_colocap,
            archivo_desembarque=resolver_ruta(base, configuracion["RUTA_DESEMBARQUE"]),
            salida_machos=resolver_ruta(base, configuracion["SALIDA_BPPALS_MACHOS"]),
            salida_hembras=resolver_ruta(base, configuracion["SALIDA_BPPALS_HEMBRAS"]),
            talla_min=leer_entero(configuracion, "TALLA_MIN"),
            talla_max=leer_entero(configuracion, "TALLA_MAX"),
            paso_talla=leer_entero(configuracion, "PASO_TALLA"),
            edad_min=leer_entero(configuracion, "EDAD_MIN"),
            edad_max=leer_entero(configuracion, "EDAD_MAX"),
            hoja_colocap=configuracion["NOMBRE_HOJA_COLOCAP"])

if __name__ == "__main__":
    main()
