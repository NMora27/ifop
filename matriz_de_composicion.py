"""Composición talla–edad y desembarque, con diseños por perfil y cálculo verificable."""
import math
import os
import re
import shutil
import warnings
from copy import copy
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter as letra
from openpyxl.utils import range_boundaries

from diseno_composicion import (
    _combinar_estilos,
    activar_recalculo,
    reemplazar_datos_hoja,
)
from metodos_matematicos import calcular_a1_a2
from lectura import leer_perfil

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
DISENOS = {
    'merluza': {'capacidad_tallas': 61, 'capacidad_edades': 24, 'clave': 84, 'meta': 80,
                    'total_clave': 145, 'total_colocap': 73, 'a1': 77, 'total_comp': 70,
                    'varianza': 79, 'total_varianza': 140, 'total_matriz': 70,
                    'presentacion': 'Tabla Presentación', 'desplazamiento_print': 1, 'resumen_print': 70,
                    'total_distr': 66},
    'bacalao': {'capacidad_tallas': 52, 'capacidad_edades': 37, 'clave': 75, 'meta': 71,
                    'total_clave': 127, 'total_colocap': 64, 'a1': 68, 'total_comp': 61,
                    'varianza': 70, 'total_varianza': 122, 'total_matriz': 61,
                    'presentacion': 'Tabla Print', 'desplazamiento_print': 0, 'resumen_print': 52,
                    'total_distr': 66},}

class DatosBaseNoDisponibles(ValueError):
    """Una fuente aún no aporta datos para el producto solicitado."""


def _codigo_sexo(sexo):
    valor = sexo.strip().casefold() if isinstance(sexo, str) else sexo
    if isinstance(valor, bool):
        raise ValueError('Sexo inválido')
    if valor in (1, 'macho', 'machos'):
        return 1
    if valor in (2, 'hembra', 'hembras'):
        return 2
    raise ValueError('Indique macho/1 o hembra/2')


def _validar_talla_referencia(talla):
    if isinstance(talla, bool):
        raise ValueError('Talla inválida')
    try:
        talla = float(talla)
    except (TypeError, ValueError):
        raise ValueError('Talla inválida') from None
    if not math.isfinite(talla) or talla <= 0:
        raise ValueError('La talla debe ser finita y positiva')
    return talla


def _seleccionar_edad_probable(probabilidades):
    """Selecciona la mayor probabilidad y, en empate, la edad menor."""
    if not probabilidades or any(
        isinstance(p, bool) or not isinstance(p, (int, float))
        or not math.isfinite(p) or not 0 <= p <= 1
        for p in probabilidades.values()
    ) or not math.isclose(sum(probabilidades.values()), 1., rel_tol=1e-9, abs_tol=1e-12):
        raise ValueError('Sin distribución de probabilidades válida para la talla')
    ordenadas = sorted(probabilidades.items(), key=lambda par: (-par[1], par[0]))
    return dict(
        edad_asignada=ordenadas[0][0],
        edades_empatadas_en_primer_lugar=[e for e, p in ordenadas if p == ordenadas[0][1]],
        opciones=[dict(edad=e, probabilidad=p) for e, p in ordenadas[:3]])


def cargar_probabilidades_merluza(sexo, dominio):
    """Consulta las probabilidades guardadas sin leer ni reconstruir históricos."""
    from claves_edad_merluza import CLAVES
    if dominio not in CLAVES:
        raise ValueError(f'Dominio de referencia desconocido: {dominio}')
    return {marca: dict(edades) for marca, edades in CLAVES[dominio][_codigo_sexo(sexo)].items()}


def asignar_edad_merluza(marca_de_clase, probabilidades):
    """Consulta exacta de marca de clase, sin interpolar tallas."""
    marca = _validar_talla_referencia(marca_de_clase)
    return _seleccionar_edad_probable(probabilidades.get(marca, {}))


def asignar_edad_bacalao(talla_cm, sexo):
    """Conserva las reglas fijas de bacalao y consulta las probabilidades por sexo."""
    from claves_edad_bacalao import CLAVES
    talla = _validar_talla_referencia(talla_cm)
    probabilidades = CLAVES[_codigo_sexo(sexo)]
    # Reglas de asignación del módulo original de bacalao, previas a la consulta.
    edad_fija = (2 if talla < 25 else 3 if talla < 35 else
                 4 if talla <= 40 else 37 if talla > 245 else None)
    if edad_fija is not None:
        return _seleccionar_edad_probable({edad_fija: 1.0})
    clase = math.floor(talla / 5) * 5
    return _seleccionar_edad_probable(probabilidades.get(clase, {}))


def leer_regresiones(archivo, *, especie_codigo, anio, flota, dominio, metodo):
    """Selecciona un único ajuste por sexo, con todas las dimensiones explícitas."""
    if Path(archivo).suffix.lower() == '.xlsx':
        with pd.ExcelFile(archivo) as libro:
            hoja = next((s for s in libro.sheet_names if s.casefold() == metodo.casefold()), None)
            if hoja is None:
                raise ValueError(f'No hay hoja de regresión para el método {metodo}')
            d = pd.read_excel(libro, sheet_name=hoja)
    else:
        d = pd.read_csv(archivo, sep=';')
    requeridas = ['COD_ESPECIE', 'ANIO', 'FLOTA', 'DOMINIO', 'METODO', 'GRUPO', 'ESTADO', 'a', 'b']
    if set(requeridas) - set(d.columns):
        raise ValueError('Faltan columnas de contexto en las regresiones')
    d = d[pd.to_numeric(d.COD_ESPECIE, errors='coerce').eq(especie_codigo)
          & pd.to_numeric(d.ANIO, errors='coerce').eq(anio)
          & d.FLOTA.eq(flota) & d.DOMINIO.eq(dominio)
          & d.METODO.str.casefold().eq(metodo.casefold())]
    parametros = {}
    if d.empty:
        raise DatosBaseNoDisponibles(f'Sin regresiones para {flota}, {dominio}, {metodo}')
    for sexo in ['Machos', 'Hembras']:
        grupo = d[d.GRUPO.eq(sexo)]
        if len(grupo) != 1 or grupo.iloc[0].ESTADO != 'OK':
            raise ValueError(f'Se requiere un único ajuste OK para {sexo}, {flota}, {dominio}, {metodo}')
        a, b = (float(str(grupo.iloc[0][campo]).replace(',', '.'))
                for campo in ('a', 'b'))
        if not math.isfinite(a) or not math.isfinite(b) or a <= 0:
            raise ValueError('Parámetros de regresión inválidos')
        parametros[sexo] = (a, b, *calcular_a1_a2(b))
    return parametros


def leer_clave(archivo, sexo, tallas, edades):
    """Lee también la columna final N+; no la descarta al convertir encabezados."""
    d = pd.read_excel(archivo, sheet_name=sexo, index_col=0)
    matriz = np.zeros((len(tallas), len(edades)))
    vistas = set()
    for talla, fila in d.iterrows():
        if talla == 'Total general':
            continue
        try:
            numero = float(talla)
        except (TypeError, ValueError):
            raise ValueError(f'Clase de talla inválida: {talla}') from None
        if not math.isfinite(numero) or numero != int(numero) or numero in vistas:
            raise ValueError(f'Clase de talla inválida o repetida: {talla}')
        vistas.add(numero)
        for encabezado, valor in fila.items():
            if str(encabezado) == 'Total general':
                continue
            texto = str(encabezado).strip()
            try:
                edad = float(texto.rstrip('+'))
            except ValueError:
                raise ValueError(f'Encabezado de edad inválido: {encabezado}') from None
            if not math.isfinite(edad) or edad != int(edad):
                raise ValueError(f'Edad inválida: {encabezado}')
            frecuencia = 0.0 if pd.isna(valor) else float(valor)
            if not math.isfinite(frecuencia) or frecuencia < 0 or frecuencia != int(frecuencia):
                raise ValueError('La clave debe contener conteos enteros no negativos')
            if texto.endswith('+') and edad < edades[-1]:
                raise ValueError('La clase de edad agrupada de la fuente no permite desagregar edades')
            destino = min(int(edad), edades[-1])
            if frecuencia and (numero not in tallas or destino not in edades):
                raise ValueError('La clave contiene conteos fuera del rango configurado')
            if numero in tallas and destino in edades:
                matriz[tallas.index(numero), edades.index(destino)] += frecuencia
    return matriz


def leer_colocap(archivo, hoja, tallas):
    with warnings.catch_warnings():
        warnings.simplefilter('ignore', UserWarning)
        w = load_workbook(archivo, data_only=True)
    try:
        s = w[hoja]
        pm, ph, total = [float(s[c].value) for c in ['L2', 'M2', 'N2']]
        if not all(math.isfinite(v) and v >= 0 for v in [pm, ph, total]) or abs(pm+ph-1) > .001:
            raise ValueError('Proporciones sexuales o total COLOCAP inválidos')
        frecuencias = np.zeros((len(tallas), 2))
        vistas, sumas = set(), np.zeros(2)
        for fila in s.iter_rows(min_row=3, min_col=16, max_col=18, values_only=True):
            if fila[0] is None:
                continue
            try:
                talla = float(fila[0])
            except (TypeError, ValueError):
                continue  # Etiqueta de total del libro, no una clase.
            if not math.isfinite(talla) or talla != int(talla) or talla in vistas:
                raise ValueError('Talla COLOCAP inválida o repetida')
            vistas.add(talla)
            props = np.array([float(v or 0) for v in fila[1:]])
            if not np.isfinite(props).all() or (props < 0).any() or (props > 1).any():
                raise ValueError('Proporción de talla COLOCAP inválida')
            if talla not in tallas:
                if props.sum() > 0:
                    raise ValueError('COLOCAP contiene tallas positivas fuera del rango')
                continue
            sumas += props
            frecuencias[tallas.index(talla)] = props * [total*pm, total*ph]
        for i, proporcion in enumerate([pm, ph]):
            if total * proporcion > 0 and abs(sumas[i]-1) > .001:
                raise ValueError('Las proporciones de talla no suman uno')
        return frecuencias
    finally:
        w.close()


def leer_desembarque(configuracion, *, base='.'):
    """Selecciona celdas explícitas, verificando especie y rótulos de la fuente."""
    from lectura import normalizar_texto
    if 'valor_toneladas' in configuracion:
        if not configuracion.get('fuente'):
            raise ValueError('El desembarque manual requiere fuente explícita')
        valor = float(configuracion['valor_toneladas'])
    else:
        w = load_workbook(Path(base)/configuracion['archivo'], data_only=True)
        try:
            s = w[configuracion['hoja']]
            verificaciones = configuracion.get('verificar', {})
            if not verificaciones:
                raise ValueError('Debe verificar los rótulos de especie y dominio del desembarque')
            for celda, esperado in verificaciones.items():
                if normalizar_texto(s[celda].value) != normalizar_texto(esperado):
                    raise ValueError(f'Rótulo de desembarque inesperado en {celda}')
            vacias_cero = set(configuracion.get('celdas_vacias_cero', []))
            if not vacias_cero.issubset(configuracion['celdas']):
                raise ValueError('Las celdas vacías admitidas deben pertenecer a la suma')
            valores = [0 if s[c].value is None and c in vacias_cero else s[c].value
                       for c in configuracion['celdas']]
            if not valores or any(isinstance(v, bool) or not isinstance(v, (float, int))
                                  or not math.isfinite(v) or v < 0 for v in valores):
                raise ValueError('Celdas de desembarque inválidas o sin valor calculado')
            valor = sum(valores)
        finally:
            w.close()
    if not math.isfinite(valor) or valor < 0:
        raise ValueError('Desembarque inválido')
    return valor


def calcular_composicion(clave, frecuencias, marcas, parametros, desembarque, sexo_indice):
    """Reproduce la expansión de los libros: toneladas a gramos y reparto en peso."""
    pesos = np.column_stack([parametros[s][0]*marcas**parametros[s][1] for s in ['Machos','Hembras']])
    peso_muestra = frecuencias * pesos
    peso_total = peso_muestra.sum()
    if peso_total <= 0 and desembarque > 0:
        raise ValueError('No hay frecuencias positivas para expandir el desembarque')
    # Algebraicamente equivalente al reparto en peso y peso medio por sexo.
    factor = desembarque * 1e6 / peso_total if peso_total else 0
    individuos = frecuencias * factor
    sumas = clave.sum(axis=1)
    prob = np.divide(clave, sumas[:, None], out=np.zeros_like(clave), where=sumas[:, None] > 0)
    captura = individuos[:, sexo_indice]
    composicion = prob * captura[:, None]
    totales = composicion.sum(axis=0)
    momento1 = composicion * marcas[:, None]
    momento2 = composicion * marcas[:, None]**2
    media = np.divide(momento1.sum(axis=0), totales, out=np.zeros_like(totales), where=totales > 0)
    varianza = np.divide(momento2.sum(axis=0)-media*momento1.sum(axis=0), totales-1,
                         out=np.zeros_like(totales), where=totales > 1)
    varianza = np.maximum(varianza, 0)
    a, b, a1, a2 = parametros[['Machos','Hembras'][sexo_indice]]
    peso_edad = np.zeros_like(media)
    validos = media > 0
    m, v = media[validos], varianza[validos]
    peso_edad[validos] = a*(m**b + a1*m**(b-2)*v + a2*m**(b-4)*v**2)
    return {'pesos': pesos, 'peso_muestra': peso_muestra, 'individuos': individuos, 'prob': prob,
                'captura': captura, 'composicion': composicion, 'totales': totales,
                'momento1': momento1, 'momento2': momento2, 'media': media, 'varianza': varianza,
                'peso_edad': peso_edad, 'pendientes': (sumas == 0) & (captura > 0)}


class LibroComposicion:
    def __init__(self, plantilla):
        self.original = Path(plantilla).read_bytes()
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            self.wb = load_workbook(BytesIO(self.original), keep_links=False)
        self.valores = {}
        for s in self.wb:
            for c in list(s._cells.values()):
                if not isinstance(c, MergedCell):
                    c.value = None
                    c.comment = None
            s.conditional_formatting._cf_rules.clear()

    def poner(self, hoja, fila, col, valor, resultado=None):
        s = self.wb[hoja]
        s.cell(fila, col).value = valor
        self.valores[(hoja, f'{letra(col)}{fila}')] = resultado if isinstance(valor, str) and valor.startswith('=') else valor

    def formula(self, hoja, fila, col, expresion, resultado):
        self.poner(hoja, fila, col, '='+expresion, resultado)

    def _datos_cache(self, contenido, hoja):
        raiz = ET.fromstring(contenido)
        for celda in raiz.iter(f'{{{NS}}}c'):
            if celda.find(f'{{{NS}}}f') is None:
                continue
            valor = self.valores.get((hoja, celda.attrib['r']), 0)
            cache = celda.find(f'{{{NS}}}v')
            if cache is None:
                cache = ET.SubElement(celda, f'{{{NS}}}v')
            if isinstance(valor, str):
                celda.set('t', 'str')
                cache.text = valor
            else:
                celda.attrib.pop('t', None)
                cache.text = str(float(valor or 0))
        return ET.tostring(raiz, encoding='utf-8')

    def _grafico(self, contenido, n, k):
        raiz = ET.fromstring(contenido)
        # Las plantillas fijaban 200 miles y saltos de 20, ajenos a cada captura.
        es_distribucion = any('Distr.Capt.Nº' in (f.text or '')
                              for f in raiz.iter(f'{{{CHART}}}f'))
        if es_distribucion:
            for eje in raiz.iter(f'{{{CHART}}}valAx'):
                escala = eje.find(f'{{{CHART}}}scaling')
                if escala is not None:
                    for limite in list(escala):
                        if limite.tag == f'{{{CHART}}}max':
                            escala.remove(limite)
                    minimo = escala.find(f'{{{CHART}}}min')
                    if minimo is None:
                        minimo = ET.SubElement(escala, f'{{{CHART}}}min')
                    minimo.set('val', '0')
                for unidad in list(eje):
                    if unidad.tag in (f'{{{CHART}}}majorUnit', f'{{{CHART}}}minorUnit'):
                        eje.remove(unidad)
                formato = eje.find(f'{{{CHART}}}numFmt')
                if formato is not None:
                    formato.set('formatCode', '0.########')
                    formato.set('sourceLinked', '0')
        for ref in list(raiz.iter(f'{{{CHART}}}numRef')) + list(raiz.iter(f'{{{CHART}}}strRef')):
            f = ref.find(f'{{{CHART}}}f')
            if f is None or not f.text or '!' not in f.text:
                continue
            hoja, rango = f.text.rsplit('!', 1)
            hoja = hoja.strip("'").replace("''", "'")
            if hoja == 'Graf.':
                rango = re.sub(r':\$([BC])\$\d+', lambda m: f':${m[1]}${3+k}', rango)
            elif hoja == 'Distr.Capt.Nº':
                rango = re.sub(r':\$([AGH])\$\d+', lambda m: f':${m[1]}${3+n}', rango)
            f.text = f"'{hoja}'!{rango}"
            for c in list(ref):
                if c.tag.endswith(('numCache', 'strCache')):
                    ref.remove(c)
            numerico = ref.tag.endswith('numRef')
            c1, r1, c2, r2 = range_boundaries(rango)
            valores = [self.valores.get((hoja, f'{letra(c)}{r}'), 0)
                       for r in range(r1, r2+1) for c in range(c1, c2+1)]
            if numerico and any(isinstance(v, str) for v in valores):
                numerico = False
                ref.tag = f'{{{CHART}}}strRef'
            cache = ET.SubElement(ref, f'{{{CHART}}}' + ('numCache' if numerico else 'strCache'))
            if numerico:
                ET.SubElement(cache, f'{{{CHART}}}formatCode').text = 'General'
            ET.SubElement(cache, f'{{{CHART}}}ptCount', val=str(len(valores)))
            for i, valor in enumerate(valores):
                pt = ET.SubElement(cache, f'{{{CHART}}}pt', idx=str(i))
                # Excel representa el último grupo 24+ como categoría textual.
                ET.SubElement(pt, f'{{{CHART}}}v').text = str(valor or 0)
        return ET.tostring(raiz, encoding='utf-8')

    def guardar(self, salida, n, k):
        """Conserva dibujos/impresión originales y actualiza datos y cachés de gráficos."""
        with NamedTemporaryFile(suffix='.xlsx', delete_on_close=False) as temporal:
            self.wb.save(temporal.name)
            with ZipFile(BytesIO(self.original)) as base, ZipFile(temporal.name) as nuevo, \
                    ZipFile(salida, 'w', ZIP_DEFLATED) as destino:
                estilos, mapa, diferenciales = _combinar_estilos(base.read('xl/styles.xml'), nuevo.read('xl/styles.xml'))
                for entrada in base.infolist():
                    nombre = entrada.filename
                    if nombre.startswith('xl/externalLinks/') or nombre in ('xl/calcChain.xml', 'xl/sharedStrings.xml'):
                        continue
                    contenido = base.read(nombre)
                    m = re.fullmatch(r'xl/worksheets/sheet(\d+)\.xml', nombre)
                    if m:
                        hoja = self.wb.sheetnames[int(m[1])-1]
                        datos = self._datos_cache(nuevo.read(nombre), hoja)
                        contenido = reemplazar_datos_hoja(contenido, datos, mapa, diferenciales)
                        # No conservar reglas condicionales del ejemplo.
                        contenido = re.sub(rb'<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>', b'', contenido, flags=re.DOTALL)
                    elif nombre == 'xl/styles.xml':
                        contenido = estilos
                    elif nombre == 'xl/workbook.xml':
                        contenido = activar_recalculo(contenido)
                        contenido = re.sub(rb'<externalReferences\b[^>]*>.*?</externalReferences>', b'', contenido, flags=re.DOTALL)
                    elif nombre in ('xl/_rels/workbook.xml.rels', '[Content_Types].xml'):
                        raiz = ET.fromstring(contenido)
                        for e in list(raiz):
                            if any(t in str(e.attrib) for t in ('calcChain', 'externalLink', 'sharedStrings')):
                                raiz.remove(e)
                        contenido = ET.tostring(raiz, encoding='utf-8')
                    elif re.fullmatch(r'xl/charts/chart\d+\.xml', nombre):
                        contenido = self._grafico(contenido, n, k)
                    destino.writestr(entrada, contenido)


def escribir_libro(libro, diseno, tallas, edades, clave, frecuencias, marcas, parametros,
                    calculo, *, especie, sexo, anio, flota, zona, area, desembarque,
                    asignaciones, paso_talla):
    p, f = libro.poner, libro.formula
    n, k = len(tallas), len(edades)
    idx = 0 if sexo == 'Machos' else 1
    csexo = 'C' if idx == 0 else 'D'
    cnumero = 'E' if idx == 0 else 'F'
    d = diseno
    ck, mk, tc, ts = d['clave'], d['meta'], d['total_colocap'], d['total_comp']
    ultimo_c, ultimo_m = letra(6+k), letra(2+k)
    ultimo_comp = letra(5+k)
    meta = {2: f'Especie: {especie}', 11: f'Sexo: {sexo}',
            20: f'Zona: {zona} - {flota}', 28: f'Fecha: {anio}'}
    for col, texto in meta.items():
        p('Clave', mk, col, texto)
        f('Clave', 5, col, f'{letra(col)}{mk}', texto)
    p('Clave', 2, 5, 'CLAVE: PROBABILIDAD DE PERTENENCIA POR TALLA Y EDAD')
    p('Clave', mk-3, 5, 'CLAVE EDAD - TALLA')
    p('Clave', ck-2, 2, 'Tallas (cm)')
    p('Clave', ck-2, 6, 'Frec.')
    p('Clave', ck-2, 7, 'Grupos de edad')
    p('Clave', ck-1, 1, 'VERIFICACIÓN')
    p('Clave', 7, 2, 'Tallas (cm)')
    p('Clave', 7, 6, 'Frec.')
    p('Clave', 7, 16, 'Grupos de edad')
    for j, edad in enumerate(edades):
        etiqueta = f'{edad}+' if j == k-1 else edad
        p('Clave', ck-1, 7+j, etiqueta)
        f('Clave', 8, 7+j, f'{letra(7+j)}{ck-1}', etiqueta)
    for i, talla in enumerate(tallas):
        r, rp, rc = ck+i, 9+i, 11+i
        estado = 'FALTA CL' if calculo['pendientes'][i] else 'CL OK' if frecuencias[i,idx] > 0 else ' '
        f('Clave', r, 1, f'IF(Colocap!{csexo}{rc}>0,IF(F{r}>0,"CL OK","FALTA CL")," ")', estado)
        for col, valor in [(2,talla),(3,'-'),(4,talla+paso_talla-1)]:
            p('Clave', r, col, valor)
            f('Clave', rp, col, f'{letra(col)}{r}', valor)
        f('Clave', r, 6, f'SUM(G{r}:{ultimo_c}{r})', clave[i].sum())
        f('Clave', rp, 6, f'SUM(G{rp}:{ultimo_c}{rp})', calculo['prob'][i].sum())
        for j in range(k):
            celda = libro.wb['Clave'].cell(r,7+j)
            fuente = copy(celda.font); fuente.color = 'FF000000'; celda.font = fuente
            if clave[i,j]:
                p('Clave', r, 7+j, float(clave[i,j]))
            f('Clave', rp, 7+j, f'IF($F{r}=0,0,{letra(7+j)}{r}/$F{r})', calculo['prob'][i,j])
    for i, j in asignaciones:
        celda = libro.wb['Clave'].cell(ck+i, 7+j)
        fuente = copy(celda.font); fuente.color = 'FFFF0000'; celda.font = fuente
    p('Clave', d['total_clave'], 2, 'Total')
    for j in range(k+1):
        col = 6+j
        valor = clave.sum() if j == 0 else clave[:,j-1].sum()
        f('Clave', d['total_clave'], col, f'SUM({letra(col)}{ck}:{letra(col)}{ck+n-1})', valor)

    p('Colocap', 8, 2, 'DESEMBARQUE (t)')
    p('Colocap', 8, 3, desembarque)
    p('Colocap', 6, 2, 'a:'); p('Colocap', 7, 2, 'b:')
    for col, texto in [(3,'Machos'),(4,'Hembras'),(5,'Peso machos (g)'),(6,'Peso hembras (g)')]:
        p('Colocap', 10, col, texto)
    p('Colocap', 10, 2, 'Marca de clase')
    f('Colocap', 2, 2, f'Clave!B{mk}', meta[2]); f('Colocap', 3, 2, f'Clave!T{mk}', meta[20])
    p('Colocap', 4, 5, 'Machos'); p('Colocap', 4, 6, 'Hembras')
    p('Colocap', 5, 5, 'Desembarque en número')
    p('Colocap', 7, 5, 'Desembarque en peso (t)')
    for j, (nombre, (a,b,a1,a2)) in enumerate((s,parametros[s]) for s in ['Machos','Hembras']):
        col, cp = 3+j, 5+j
        l, lp = letra(col), letra(cp)
        p('Colocap', 6, col, a); p('Colocap', 7, col, b)
        p('Colocap', d['a1'], col, a1); p('Colocap', d['a1']+1, col, a2)
        for i in range(n):
            r = 11+i
            p('Colocap', r, 2, marcas[i]); p('Colocap', r, col, frecuencias[i,j])
            f('Colocap', r, cp, f'{l}{r}*${l}$6*$B{r}^${l}$7', calculo['peso_muestra'][i,j])
        f('Colocap', tc, col, f'SUM({l}11:{l}{10+n})', frecuencias[:,j].sum())
        f('Colocap', tc, cp, f'SUM({lp}11:{lp}{10+n})', calculo['peso_muestra'][:,j].sum())
        total_peso = calculo['peso_muestra'].sum()
        proporcion = calculo['peso_muestra'][:,j].sum()/total_peso if total_peso else 0
        f('Colocap', tc+1, cp, f'IF(SUM($E${tc}:$F${tc})=0,0,{lp}{tc}/SUM($E${tc}:$F${tc}))', proporcion)
        f('Colocap', 8, cp, f'$C$8*{lp}{tc+1}', desembarque*proporcion)
        f('Colocap', 6, cp, f'IF({lp}{tc}=0,0,({l}{tc}/{lp}{tc})*{lp}8*1000000)', calculo['individuos'][:,j].sum())
    p('Colocap', tc, 2, 'Total'); p('Colocap', d['a1'], 2, 'a1:'); p('Colocap', d['a1']+1, 2, 'a2:')
    p('Colocap', tc+1, 2, 'PORCENTAJE EN PESO POR SEXO:')
    p('Colocap', tc+2, 4, 'PESO PROM:')
    frecuencia_sexo = frecuencias[:,idx].sum()
    peso_promedio = calculo['peso_muestra'][:,idx].sum()/frecuencia_sexo if frecuencia_sexo else 0
    f('Colocap', tc+2, 5+idx,
      f'IF({csexo}{tc}=0,0,{cnumero}{tc}/{csexo}{tc})', peso_promedio)
    p('Colocap', tc+3, 3, 'MACHOS'); p('Colocap', tc+3, 4, 'HEMBRAS')
    f('Colocap', 6, 7, 'SUM(E6:F6)', calculo['individuos'].sum())

    # Composición, momentos y tabla de presentación se calculan desde la misma matriz.
    for sn in ['Comp.Capt.', 'Tabla']:
        p(sn, 6, 1, 'Tallas (cm)'); p(sn, 6, 5, 'Frec.'); p(sn, 6, 6, 'Grupos de edad')
        for j, edad in enumerate(edades):
            f(sn, 7, 6+j, f'Clave!{letra(7+j)}{ck-1}', f'{edad}+' if j==k-1 else edad)
        for i, talla in enumerate(tallas):
            r = 8+i
            for col, valor in [(1,talla),(2,'-'),(3,libro.valores[('Clave', f'D{ck+i}')])]:
                f(sn, r, col, f'Clave!{letra(col+1)}{ck+i}', valor)
            if sn == 'Comp.Capt.':
                f(sn, r, 5, f'IF(Colocap!{csexo}${tc}=0,0,Colocap!{csexo}{11+i}/Colocap!{csexo}${tc}*Colocap!${cnumero}$6)', calculo['captura'][i])
            else:
                f(sn, r, 5, f"'Comp.Capt.'!E{r}", calculo['captura'][i])
            for j in range(k):
                expresion = f'Clave!{letra(7+j)}{9+i}*$E{r}' if sn=='Comp.Capt.' else f"'Comp.Capt.'!{letra(6+j)}{r}"
                f(sn, r, 6+j, expresion, calculo['composicion'][i,j])
    p('Comp.Capt.', 2, 4, 'COMPOSICIÓN DEL DESEMBARQUE EN NÚMERO POR GRUPO DE EDAD')
    p('Comp.Capt.', 4, 1, f'{especie} — {flota} — {zona} — {sexo} — {anio}')
    titulo = (f'Composición del desembarque en número por grupo de edad de {especie.lower()}, '
              f'{sexo.lower()}. {flota}, {zona}, {area}, {anio}. Desembarque total: {desembarque:,.3f} t.')
    p('Tabla', 3, 1, titulo)
    # Momentos de talla para media y varianza por edad.
    for inicio, total, nombre in [(9,d['total_matriz'],'momento1'),(d['varianza'],d['total_varianza'],'momento2')]:
        p('Matrices', inicio-2, 1, 'Tallas'); p('Matrices', inicio-2, 2, 'Frec.')
        for j, edad in enumerate(edades):
            p('Matrices', inicio-1, 3+j, f'{edad}+' if j==k-1 else edad)
        for i in range(n):
            r = inicio+i
            p('Matrices', r, 1, marcas[i])
            f('Matrices', r, 2, f'SUM(C{r}:{ultimo_m}{r})', calculo[nombre][i].sum())
            for j in range(k):
                potencia = '^2' if nombre == 'momento2' else ''
                f('Matrices', r, 3+j, f"$A{r}{potencia}*'Comp.Capt.'!{letra(6+j)}{8+i}", calculo[nombre][i,j])
        p('Matrices', total, 1, 'Total')
        for j in range(k+1):
            col = 2+j
            valor = calculo[nombre].sum() if j==0 else calculo[nombre][:,j-1].sum()
            f('Matrices', total, col, f'SUM({letra(col)}{inicio}:{letra(col)}{inicio+n-1})', valor)
    resumenes = [(0,'Total',calculo['totales']),
                 (2,'Porcentaje',calculo['totales']*100/calculo['captura'].sum() if calculo['captura'].sum() else np.zeros(k)),
                 (4,'Talla prom. (cm)',calculo['media']), (6,'Varianza',calculo['varianza']),
                 (8,'Peso prom. (g)',calculo['peso_edad'])]
    a,b,a1,a2 = parametros[sexo]
    for offset, etiqueta, valores in resumenes:
        r = ts+offset
        p('Comp.Capt.', r, 1, etiqueta); p('Tabla', r, 1, etiqueta)
        for j, valor in enumerate(valores):
            l, lm = letra(6+j), letra(3+j)
            if offset==0:
                expresion = f'SUM({l}8:{l}{7+n})'
            elif offset==2:
                expresion = f'IF($E${ts}=0,0,{l}{ts}*100/$E${ts})'
            elif offset==4:
                expresion = f'IF({l}{ts}=0,0,Matrices!{lm}{d["total_matriz"]}/{l}{ts})'
            elif offset==6:
                expresion = f'IF({l}{ts}<=1,0,MAX(0,(Matrices!{lm}{d["total_varianza"]}-{l}{ts+4}*Matrices!{lm}{d["total_matriz"]})/({l}{ts}-1)))'
            else:
                expresion = (f'IF({l}{ts+4}=0,0,Colocap!${csexo}$6*({l}{ts+4}^Colocap!${csexo}$7'
                    f'+Colocap!${csexo}${d["a1"]}*{l}{ts+4}^(Colocap!${csexo}$7-2)*{l}{ts+6}'
                    f'+Colocap!${csexo}${d["a1"]+1}*{l}{ts+4}^(Colocap!${csexo}$7-4)*{l}{ts+6}^2))')
            f('Comp.Capt.', r, 6+j, expresion, valor)
            f('Tabla', r, 6+j, f"'Comp.Capt.'!{l}{r}", valor)
    for sn in ['Comp.Capt.', 'Tabla']:
        f(sn, ts, 5, f'SUM(E8:E{7+n})', calculo['captura'].sum())
    f('Comp.Capt.', ts+1, 5, f'SUM(F{ts}:{ultimo_comp}{ts})', calculo['totales'].sum())

    p('Graf.', 1, 1, 'Composición del desembarque en número (%)')
    p('Graf.', 1, 8, titulo)
    p('Graf.', 3, 2, 'GE'); p('Graf.', 3, 3, '%'); p('Graf.', 3, 4, '>= 5 %')
    porcentajes = resumenes[1][2]
    for j, edad in enumerate(edades):
        r = 4+j
        p('Graf.', r, 2, f'{edad}+' if j==k-1 else edad)
        f('Graf.', r, 3, f"'Comp.Capt.'!{letra(6+j)}{ts+2}", porcentajes[j])
        f('Graf.', r, 4, f'IF(C{r}>=5,C{r},"")', porcentajes[j] if porcentajes[j]>=5 else '')
        if porcentajes[j]>=5:
            libro.wb['Graf.'].cell(r,4).fill = PatternFill('solid', fgColor='FFFFFF00')
    f('Graf.', 4+k, 4, f'SUM(D4:D{3+k})', porcentajes[porcentajes>=5].sum())

    p('Distr.Capt.Nº', 3, 1, 'Intervalo de talla')
    for col, texto in [(5,'Machos'),(6,'Hembras'),(7,'Machos (miles)'),(8,'Hembras (miles)')]:
        p('Distr.Capt.Nº', 3, col, texto)
    for i, talla in enumerate(tallas):
        r = 4+i
        for col, valor in [(1,talla),(2,'-'),(3,libro.valores[('Clave',f'D{ck+i}')])]:
            p('Distr.Capt.Nº', r, col, valor)
        for j in range(2):
            col, l = 5+j, letra(3+j)
            f('Distr.Capt.Nº', r, col, f'IF(Colocap!{l}${tc}=0,0,Colocap!{l}{11+i}/Colocap!{l}${tc}*Colocap!${letra(5+j)}$6)', calculo['individuos'][i,j])
            f('Distr.Capt.Nº', r, col+2, f'{letra(col)}{r}/1000', calculo['individuos'][i,j]/1000)
    total_distr = max(d['total_distr'], n+5)
    p('Distr.Capt.Nº', total_distr, 1, 'Total')
    for col in range(5,9):
        valor = calculo['individuos'][:,(col-5)%2].sum()/(1000 if col>=7 else 1)
        f('Distr.Capt.Nº', total_distr, col, f'SUM({letra(col)}4:{letra(col)}{3+n})', valor)
    sn, shift = d['presentacion'], d['desplazamiento_print']
    p(sn, 3, 1, titulo)
    for col, texto in [(1,'Tallas (cm)'),(5,'Frec.'),(6,'Grupos de edad')]:
        p(sn, 6, col+shift, texto)
    for j, edad in enumerate(edades):
        f(sn, 7, 6+j+shift, f'Tabla!{letra(6+j)}7', f'{edad}+' if j==k-1 else edad)
    for i in range(n):
        for col in [1,2,3]+list(range(5,6+k)):
            valor = libro.valores[('Tabla', f'{letra(col)}{8+i}')]
            f(sn, 8+i, col+shift, f'Tabla!{letra(col)}{8+i}', valor)
    pendientes = float(calculo['captura'][calculo['pendientes']].sum())
    for hoja in ['Comp.Capt.', 'Tabla']:
        p(hoja, ts+10, 1, 'Individuos pendientes de clave')
        p(hoja, ts+10, 5, pendientes)
    p('Graf.', 4, 8, f'Individuos sin clave: {pendientes:.3f}')
    # Evita que el resumen sobrescriba clases, incluso con configuraciones nuevas.
    resumen_print = max(d['resumen_print'], n+9)
    for offset, etiqueta, valores in resumenes:
        r = resumen_print+offset
        p(sn, r, 1+shift, etiqueta)
        if offset==0:
            f(sn, r, 5+shift, f'Tabla!E{ts}', calculo['captura'].sum())
        for j, valor in enumerate(valores):
            f(sn, r, 6+j+shift, f'Tabla!{letra(6+j)}{ts+offset}', valor)


def verificar_metadatos(archivo, esperados):
    """Contrasta el contexto cuando la fuente contiene metadatos del flujo actual."""
    w = load_workbook(archivo, data_only=True, read_only=True)
    try:
        if 'Metadatos' not in w.sheetnames:
            return
        datos = {str(fila[0]).strip().casefold(): fila[1]
                 for fila in w['Metadatos'].iter_rows(min_col=1,max_col=2,values_only=True)
                 if fila[0] is not None}
        for campo, valor in esperados.items():
            encontrado = datos.get(campo.casefold())
            if encontrado is None:
                continue
            if isinstance(valor,(int,float)):
                coincide = pd.to_numeric(encontrado,errors='coerce') == valor
            else:
                coincide = str(encontrado).strip().casefold() == str(valor).strip().casefold()
            if not coincide:
                raise ValueError(f'Metadatos incompatibles en {Path(archivo).name}: {campo}')
    finally:
        w.close()


def generar_matriz_de_composicion(*, especie, especie_codigo, sexo, anio, flota, zona,
        archivo_matriz, archivo_regresion, archivo_colocap, hoja_colocap, salida,
        talla_min, talla_max, paso_talla, edad_min, edad_max, desembarque,
        dominio_regresion, metodo_regresion, diseno, area='', base='.', plantilla=None,
        completar_bacalao=False, nombre_pesqueria=None, dominio_historico_merluza=None):
    """Genera un libro por sexo. Fuentes, dominios y método siempre son explícitos."""
    if sexo not in ('Machos', 'Hembras'):
        raise ValueError('Sexo debe ser Machos o Hembras')
    if diseno not in DISENOS:
        raise ValueError('Diseño no reconocido')
    d = DISENOS[diseno]
    limites = [talla_min,talla_max,paso_talla,edad_min,edad_max]
    if any(not isinstance(v,int) or isinstance(v,bool) for v in limites):
        raise ValueError('Los límites deben ser enteros')
    if paso_talla <= 0 or talla_min <= 0 or talla_max < talla_min or talla_min % paso_talla or not 1 <= edad_min <= edad_max:
        raise ValueError('Rangos de talla o edad inválidos')
    tallas, edades = list(range(talla_min,talla_max+1,paso_talla)), list(range(edad_min,edad_max+1))
    if len(tallas)>d['capacidad_tallas'] or len(edades)>d['capacidad_edades']:
        raise ValueError('Las dimensiones superan la capacidad del diseño seleccionado')
    base = Path(base)
    fuentes = {k: (base/Path(v)).resolve() for k,v in {
        'archivo_matriz': archivo_matriz, 'archivo_regresion': archivo_regresion,
        'archivo_colocap': archivo_colocap}.items()}
    salida = (base/Path(salida)).resolve()
    plantilla = (base/Path(plantilla)).resolve() if plantilla else Path(__file__).parent/'plantillas_composicion'/f'{diseno}_{sexo.lower()}.xlsx'
    if salida in [*fuentes.values(), plantilla.resolve()]:
        raise ValueError('La salida no puede sobrescribir una fuente o plantilla')
    if dominio_historico_merluza is not None and (especie_codigo != 2 or completar_bacalao):
        raise ValueError('La referencia de merluza solo admite especie 2 y no se combina con bacalao')
    parametros = leer_regresiones(fuentes['archivo_regresion'], especie_codigo=especie_codigo,
        anio=anio, flota=flota, dominio=dominio_regresion, metodo=metodo_regresion)
    verificar_metadatos(fuentes['archivo_matriz'],{'especie': especie_codigo,'anio': anio,'paso_talla': paso_talla})
    verificar_metadatos(fuentes['archivo_colocap'],{'Especie':especie_codigo,'Flota':flota,'Paso talla (cm)':paso_talla})
    clave = leer_clave(fuentes['archivo_matriz'], sexo, tallas, edades)
    frecuencias = leer_colocap(fuentes['archivo_colocap'], hoja_colocap, tallas)
    if not clave.sum() and not completar_bacalao and dominio_historico_merluza is None:
        raise DatosBaseNoDisponibles(f'Clave sin datos para {sexo}')
    if not frecuencias.sum():
        raise DatosBaseNoDisponibles('COLOCAP sin frecuencias')
    # Punto medio entre los límites enteros del intervalo: 20–21 -> 20.5; 35–39 -> 37.
    marcas = np.array(tallas, dtype=float)+(paso_talla-1)/2
    idx = 0 if sexo == 'Machos' else 1
    asignaciones = []
    if dominio_historico_merluza is not None:
        probabilidades = cargar_probabilidades_merluza(sexo, dominio_historico_merluza)
        for i, talla in enumerate(tallas):
            if frecuencias[i,idx] <= 0 or clave[i].sum() > 0:
                continue
            try:
                # El histórico etiqueta MARCA_DE_CLASE con valores pares (26, 28...).
                # Se consulta esa etiqueta exacta, sin usar longitud ni el punto medio.
                consulta = asignar_edad_merluza(talla, probabilidades)
            except ValueError:
                continue
            edad = min(consulta['edad_asignada'], edad_max)
            if edad >= edad_min:
                j = edades.index(edad)
                clave[i,j] = 1
                asignaciones.append((i,j))
    if completar_bacalao:
        if especie_codigo != 37:
            raise ValueError('La referencia histórica de bacalao solo admite especie 37')
        for i in range(len(tallas)):
            if frecuencias[i,idx] <= 0 or clave[i].sum() > 0:
                continue
            try:
                # Conserva el punto de consulta original de la referencia de bacalao.
                edad = min(asignar_edad_bacalao(tallas[i]+paso_talla/2, sexo=sexo)['edad_asignada'],edad_max)
            except ValueError:
                continue
            if edad >= edad_min:
                j = edades.index(edad)
                clave[i,j] = 1
                asignaciones.append((i,j))
    toneladas = leer_desembarque(desembarque, base=base)
    calculo = calcular_composicion(clave, frecuencias, marcas, parametros, toneladas, idx)
    libro = LibroComposicion(plantilla)
    escribir_libro(libro,d,tallas,edades,clave,frecuencias,marcas,parametros,calculo,
        especie=especie,sexo=sexo,anio=anio,flota=nombre_pesqueria or flota,zona=zona,area=area,
        desembarque=toneladas,asignaciones=asignaciones,paso_talla=paso_talla)
    salida.parent.mkdir(parents=True, exist_ok=True)
    libro.guardar(salida,len(tallas),len(edades))
    pendientes = float(calculo['captura'][calculo['pendientes']].sum())
    estado = 'PENDIENTE_CLAVE' if pendientes > 0 else 'COMPLETO'
    print(f'{salida.name}: {estado}; individuos sin clave: {pendientes:.3f}',flush=True)
    return salida


def _indice_procesos(procesos):
    resultado = {}
    for p in procesos:
        if p['id'] in resultado:
            raise ValueError(f'Identificador repetido: {p["id"]}')
        resultado[p['id']] = p
    return resultado


def preparar_composiciones(archivo, solo=None, *, verificar_fuentes=True):
    """Resuelve enlaces; el main puede planificar antes de generar las fuentes."""
    perfil, base = leer_perfil(archivo)
    colocap = _indice_procesos(perfil.get('procesos',[]))
    regresiones = _indice_procesos(perfil.get('regresiones',[]))
    claves = _indice_procesos(perfil.get('talla_edad',[]))
    procesos = _indice_procesos(perfil.get('matriz_composicion',[]))
    if not procesos:
        raise ValueError('El perfil no define matriz_composicion')
    if solo and solo not in procesos:
        raise ValueError(f'Proceso desconocido: {solo}')
    preparados, salidas = [], set()
    for id, proceso in procesos.items():
        if solo and id != solo:
            continue
        if proceso.get('pendientes'):
            warnings.warn(f'{id}: omitido por configuración pendiente: ' +
                          '; '.join(proceso['pendientes']), stacklevel=2)
            continue
        try:
            co = colocap[proceso['colocap_id']]
            re = regresiones[proceso['regresion_id']]
            cl = claves[proceso['clave_id']]
        except KeyError as e:
            raise ValueError(f'{id}: referencia de proceso no definida: {e}') from e
        archivo_regresion = Path(re['salida'])/'resultados_regresiones.xlsx'
        if verificar_fuentes and not (base/archivo_regresion).is_file():
            archivo_regresion = Path(re['salida'])/'resultados_regresiones.csv'
        fuentes = [co['salida'], str(archivo_regresion), cl['salida']]
        if 'archivo' in proceso['desembarque']:
            fuentes.append(proceso['desembarque']['archivo'])
        faltantes = [str(base/f) for f in fuentes if not (base/f).is_file()]
        if verificar_fuentes and faltantes:
            warnings.warn(f'{id}: omitido por fuentes ausentes: ' + ', '.join(faltantes), stacklevel=2)
            continue
        for campo in ['talla_min','talla_max','paso_talla']:
            if co[campo] != cl[campo]:  # noqa: SIM102
                # talla_max puede ser un límite no alineado; compara las clases efectivas.
                if campo != 'talla_max' or list(range(co['talla_min'],co['talla_max']+1,co['paso_talla'])) != list(range(cl['talla_min'],cl['talla_max']+1,cl['paso_talla'])):
                    raise ValueError(f'{id}: COLOCAP y clave tienen distinto {campo}')
        for sexo in proceso.get('sexos',['Machos','Hembras']):
            plantilla_nombre = proceso.get('nombre_archivo')
            if not plantilla_nombre:
                raise ValueError(f'{id}: falta el patrón de nombre con la sigla original')
            if sexo not in ('Machos', 'Hembras'):
                raise ValueError('Sexo debe ser Machos o Hembras')
            contexto = {'especie': perfil['especie'].replace(' ','_'), 'flota': co['flota'].replace(' ','_'),
                'zona': proceso['zona'].replace(' ','_'),'sexo': 'Macho' if sexo=='Machos' else 'Hembra',
                'sexo_sigla': 'M' if sexo=='Machos' else 'H','anio': perfil['anio']}
            nombre = plantilla_nombre.format(**contexto)
            if Path(nombre).name != nombre or Path(nombre).suffix.lower() != '.xlsx':
                raise ValueError('El nombre de salida debe ser un nombre de archivo .xlsx')
            salida = base/proceso['directorio_salida']/nombre
            if salida.resolve() in salidas:
                raise ValueError('Dos composiciones tienen la misma salida')
            salidas.add(salida.resolve())
            preparados.append({'base': base,'especie': perfil['especie'],'especie_codigo': perfil['especie_codigo'],
                'sexo': sexo,'anio': perfil['anio'],'flota': co['flota'],'zona': proceso['zona'],'area': proceso.get('area',''),
                'archivo_colocap': co['salida'],'hoja_colocap': co['nombre_hoja'],
                'archivo_regresion': str(archivo_regresion),
                'archivo_matriz': cl['salida'],'salida': salida,
                'talla_min': cl['talla_min'],'talla_max': cl['talla_max'],'paso_talla': cl['paso_talla'],
                'edad_min': proceso.get('edad_min',cl.get('edad_min',1)),'edad_max': cl['edad_agrupada'],
                'dominio_regresion': proceso['dominio_regresion'],'metodo_regresion': proceso['metodo_regresion'],
                'diseno': proceso['diseno'],'desembarque': proceso['desembarque'],
                'nombre_pesqueria': proceso.get('nombre_pesqueria', co['flota']),
                'dominio_historico_merluza': proceso.get('dominio_historico_merluza'),
                'completar_bacalao': proceso.get('completar_bacalao',False)})
    return preparados


def generar_composiciones(archivo, solo=None):
    """Genera los libros del perfil para el main, sin ejecutar módulos anteriores.

    Omite productos sin datos base y vuelve a evaluarlos en cada ejecución.
    Calcula el lote disponible en una carpeta temporal. Si alguna composición
    tiene datos inválidos, informa los errores y conserva las salidas anteriores.
    Solo publica los Excel cuando todos los cálculos del lote han terminado.
    """
    procesos = preparar_composiciones(archivo, solo=solo)
    fuentes = set()
    for proceso in procesos:
        base = Path(proceso['base'])
        for campo in ('archivo_matriz', 'archivo_regresion', 'archivo_colocap'):
            fuentes.add((base / proceso[campo]).resolve())
        if 'archivo' in proceso['desembarque']:
            fuentes.add((base / proceso['desembarque']['archivo']).resolve())
        fuentes.add((Path(__file__).parent / 'plantillas_composicion' /
                     f"{proceso['diseno']}_{proceso['sexo'].lower()}.xlsx").resolve())
    for proceso in procesos:
        if Path(proceso['salida']).resolve() in fuentes:
            raise ValueError('Una salida del lote coincide con una fuente o plantilla')
    errores, resultados = [], []
    with TemporaryDirectory(prefix='matriz_composicion_') as temporal:
        for i, proceso in enumerate(procesos):
            destino = Path(proceso['salida'])
            provisional = Path(temporal) / str(i) / destino.name
            try:
                generar_matriz_de_composicion(**dict(proceso, salida=provisional))
                resultados.append((provisional, destino))
            except (DatosBaseNoDisponibles, FileNotFoundError) as error:
                warnings.warn(f'{destino.name}: omitido: {error}', stacklevel=2)
            except (ValueError, OSError, KeyError, TypeError) as error:
                errores.append(f'{destino.name}: {error}')
        if errores:
            raise ValueError('No se generó el lote de composición:\n' + '\n'.join(errores))
        # Cada reemplazo es atómico y se prepara en el mismo sistema de archivos.
        for provisional, destino in resultados:
            destino.parent.mkdir(parents=True, exist_ok=True)
            with NamedTemporaryFile(dir=destino.parent, suffix='.xlsx', delete=False) as f:
                preparado = Path(f.name)
            try:
                shutil.copyfile(provisional, preparado)
                os.replace(preparado, destino)
            finally:
                preparado.unlink(missing_ok=True)
    return [destino for _, destino in resultados]
