"""Conserva recursos y estilos del diseño Excel; adaptado del módulo definitivo."""
import re
from copy import deepcopy
from io import BytesIO
from xml.etree import ElementTree

ESPACIO_HOJAS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ESPACIO_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
ElementTree.register_namespace("", ESPACIO_HOJAS)
ElementTree.register_namespace("x14ac", ESPACIO_X14AC)


def _obtener_datos_hoja(contenido):
    patron = re.compile(rb"<sheetData(?:\s[^>]*)?>.*?</sheetData>", re.DOTALL)
    return patron.search(contenido)


def _combinar_estilos(original, actualizado):
    """Conserva los recursos originales y remapea los estilos de openpyxl.

    Los índices del XML actualizado sólo tienen sentido dentro de su propia
    tabla. También hay que traducir fuentes, bordes y formatos numéricos.
    """
    base = ElementTree.fromstring(original)
    nuevo = ElementTree.fromstring(actualizado)
    etiqueta = lambda nombre: f"{{{ESPACIO_HOJAS}}}{nombre}"
    mapas = {}
    # ElementTree descarta declaraciones usadas sólo en mc:Ignorable.
    # Conservamos la cabecera y las extensiones originales sin serializarlas.
    for _, (prefijo, uri) in ElementTree.iterparse(BytesIO(original), events=("start-ns",)):
        ElementTree.register_namespace(prefijo, uri)
    resultado = original

    for nombre in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs"):
        origen = nuevo.find(etiqueta(nombre))
        destino = base.find(etiqueta(nombre))
        if origen is None:
            mapas[nombre] = {}
            continue
        if destino is None:
            destino = ElementTree.Element(etiqueta(nombre))
            base.insert(list(nuevo).index(origen), destino)
        mapa = {}
        siguiente_formato = max(
            [163] + [int(n.attrib["numFmtId"]) for n in destino]
        ) + 1 if nombre == "numFmts" else None
        for indice, recurso in enumerate(origen):
            copia = deepcopy(recurso)
            if nombre in ("cellStyleXfs", "cellXfs"):
                for atributo, tabla in (("fontId", "fonts"), ("fillId", "fills"),
                                        ("borderId", "borders"), ("numFmtId", "numFmts"),
                                        ("xfId", "cellStyleXfs")):
                    if atributo in copia.attrib:
                        valor = int(copia.attrib[atributo])
                        copia.attrib[atributo] = str(mapas.get(tabla, {}).get(valor, valor))
            if nombre == "cellXfs":
                # Excel debe aplicar explícitamente los recursos remapeados.
                # openpyxl omite estas marcas aunque existían en la plantilla.
                for atributo in ("applyFont", "applyFill", "applyBorder", "applyNumberFormat"):
                    copia.attrib[atributo] = "1"
            if nombre == "numFmts":
                clave = int(copia.attrib["numFmtId"])
                coincidencia = next((int(n.attrib["numFmtId"]) for n in destino
                                     if n.attrib["formatCode"] == copia.attrib["formatCode"]), None)
                if coincidencia is None:
                    coincidencia = siguiente_formato
                    siguiente_formato += 1
                    copia.attrib["numFmtId"] = str(coincidencia)
                    destino.append(copia)
            else:
                clave = indice
                contenido = ElementTree.tostring(copia)
                coincidencia = next((i for i, n in enumerate(destino)
                                     if ElementTree.tostring(n) == contenido), None)
                if coincidencia is None:
                    coincidencia = len(destino)
                    destino.append(copia)
            mapa[clave] = coincidencia
        destino.attrib["count"] = str(len(destino))
        mapas[nombre] = mapa
        patron = re.compile(
            rb"<" + nombre.encode() + rb"\b[^>]*(?:/>|>.*?</" + nombre.encode() + rb">)",
            re.DOTALL)
        fragmento = ElementTree.tostring(destino, encoding="utf-8")
        resultado, cantidad = patron.subn(lambda _: fragmento, resultado, count=1)
        if cantidad != 1:
            raise ValueError(f"No se encontró la colección de estilos {nombre} en la plantilla")
    return resultado, mapas["cellXfs"], mapas["dxfs"]


def _conservar_estilos_de_plantilla(datos_originales, datos_actualizados, mapa_estilos):
    """Actualiza los valores usando índices de la tabla de estilos combinada."""
    envoltura = (b'<raiz xmlns="' + ESPACIO_HOJAS.encode() +
                 b'" xmlns:x14ac="' + ESPACIO_X14AC.encode() + b'">%s</raiz>')
    raiz_original = ElementTree.fromstring(envoltura % datos_originales)
    raiz_actualizada = ElementTree.fromstring(envoltura % datos_actualizados)
    hoja_original = raiz_original[0]
    hoja_actualizada = raiz_actualizada[0]

    etiqueta_celda = f"{{{ESPACIO_HOJAS}}}c"
    etiqueta_fila = f"{{{ESPACIO_HOJAS}}}row"
    celdas_actualizadas = {
        celda.attrib["r"]: celda
        for celda in hoja_actualizada.iter(etiqueta_celda)
        if "r" in celda.attrib
    }
    for celda in celdas_actualizadas.values():
        # La ausencia de s también referencia un estilo: el número cero.
        celda.attrib["s"] = str(mapa_estilos[int(celda.attrib.get("s", "0"))])
    referencias_originales = set()

    for celda_original in hoja_original.iter(etiqueta_celda):
        referencia = celda_original.attrib.get("r")
        if not referencia:
            continue

        referencias_originales.add(referencia)
        celda_actualizada = celdas_actualizadas.get(referencia)
        if celda_actualizada is None:
            # openpyxl omite celdas vaciadas y sin estilo. No recuperar los
            # valores del ejemplo cuando el proceso las eliminó expresamente.
            celda_original.attrib.clear()
            celda_original.attrib.update(r=referencia, s=str(mapa_estilos[0]))
            celda_original[:] = []
            celda_original.text = None
            continue

        estilo = celda_actualizada.attrib.get("s", celda_original.attrib.get("s"))
        celda_original.attrib.clear()
        celda_original.attrib["r"] = referencia
        if estilo is not None:
            celda_original.attrib["s"] = estilo

        for atributo, valor in celda_actualizada.attrib.items():
            if atributo not in {"r", "s"}:
                celda_original.attrib[atributo] = valor

        celda_original[:] = [deepcopy(hijo) for hijo in celda_actualizada]
        celda_original.text = celda_actualizada.text

    # Las celdas nuevas no estaban en el diseño de referencia, por lo que se
    # añaden sólo si el proceso las requiere. Las existentes conservan siempre
    # diseño con los cambios explícitos del proceso y sus estilos remapeados.
    filas_originales = {
        fila.attrib.get("r"): fila
        for fila in hoja_original.findall(etiqueta_fila)
    }
    for fila_actualizada in hoja_actualizada.findall(etiqueta_fila):
        numero_fila = fila_actualizada.attrib.get("r")
        fila_original = filas_originales.get(numero_fila)
        if fila_original is None:
            hoja_original.append(deepcopy(fila_actualizada))
            continue

        for celda_actualizada in fila_actualizada.findall(etiqueta_celda):
            referencia = celda_actualizada.attrib.get("r")
            if referencia not in referencias_originales:
                fila_original.append(deepcopy(celda_actualizada))

    # Las celdas nuevas deben intercalarse por columna, no quedar al final
    # detrás de P/Q. Excel y Google Drive pueden omitir las celdas desordenadas.
    def numero_columna(celda):
        letras = re.match(r"[A-Z]+", celda.attrib["r"]).group(0)
        numero = 0
        for letra in letras:
            numero = numero * 26 + ord(letra) - ord("A") + 1
        return numero

    hoja_original[:] = sorted(hoja_original, key=lambda fila: int(fila.attrib["r"]))
    for fila in hoja_original:
        fila[:] = sorted(fila, key=numero_columna)
        if len(fila):
            fila.attrib["spans"] = f"{numero_columna(fila[0])}:{numero_columna(fila[-1])}"

    datos = ElementTree.tostring(hoja_original, encoding="utf-8")
    # El fragmento se inserta dentro de una hoja que ya declara este espacio
    # de nombres; se elimina la declaración redundante para mantener el XML
    # de la plantilla lo más intacto posible.
    declaracion = b' xmlns="' + ESPACIO_HOJAS.encode() + b'"'
    return datos.replace(declaracion, b"", 1)


def reemplazar_datos_hoja(original, actualizado, mapa_estilos, mapa_diferenciales):
    datos_originales = _obtener_datos_hoja(original)
    datos_actualizados = _obtener_datos_hoja(actualizado)

    if datos_originales is None or datos_actualizados is None:
        return actualizado

    datos_con_estilos = _conservar_estilos_de_plantilla(
        datos_originales.group(0), datos_actualizados.group(0), mapa_estilos)
    resultado = original[:datos_originales.start()] + datos_con_estilos + original[datos_originales.end():]
    dimension = re.search(rb"<dimension\b[^>]*/>", actualizado)
    if dimension is not None:
        resultado = re.sub(rb"<dimension\b[^>]*/>", lambda _: dimension.group(0),
                           resultado, count=1)
    patron_condicional = re.compile(
        rb"<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>", re.DOTALL)
    reglas = patron_condicional.findall(actualizado)
    if reglas:
        resultado = patron_condicional.sub(b"", resultado)
        reglas = [re.sub(
            rb'dxfId="(\d+)"',
            lambda m: b'dxfId="' + str(mapa_diferenciales[int(m.group(1))]).encode() + b'"',
            regla) for regla in reglas]
        # Orden de elementos de worksheet según OOXML: después de mergeCells,
        # antes de validaciones, hipervínculos y opciones de impresión.
        siguiente = re.search(
            rb"<(?:dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|"
            rb"rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|"
            rb"drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|"
            rb"webPublishItems|tableParts|extLst)\b|</worksheet>", resultado)
        if siguiente is None:
            raise ValueError("No se pudo insertar el formato condicional de la hoja")
        posicion = siguiente.start()
        resultado = resultado[:posicion] + b"".join(reglas) + resultado[posicion:]
    return resultado

def activar_recalculo(libro):
    patron = re.compile(rb"<calcPr[^>]*/>")
    calculo = b'<calcPr calcId="191028" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
    return patron.sub(calculo, libro, count=1)

