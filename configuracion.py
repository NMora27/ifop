"""Configuración del usuario y reglas fijas del flujo por especie."""
import configparser
import re
from pathlib import Path
ETAPAS = {
    'EJECUTAR_COLOCAP': 'COLOCAP',
    'EJECUTAR_REGRESIONES': 'Regresiones longitud–peso',
    'EJECUTAR_CLAVES': 'Claves talla–edad',
    'EJECUTAR_COMPOSICION': 'Matrices de composición',
}
ESPECIES = {'MERLUZA DEL SUR': ('Merluza del sur', 2, 'merluza'),
            'BACALAO DE PROFUNDIDAD': ('Bacalao de profundidad', 37, 'bacalao')}
FLOTAS = {'arrastre': ('Arrastre', 1, 1), 'palangre': ('Palangre', 1, 5),
          'artesanal': ('Artesanal', 50, 50)}
PUERTOS = {'arrastre': (61, 64), 'palangre': (70, 71), 'artesanal': ([78, 79], 76)}
NOMBRES_COMPOSICION = {'arrastre': ('MSAZN', 'MSAZS'),
                       'palangre': ('MSPALZN', 'MSPALZS'),
                       'artesanal': ('MSEIN', 'MSEIS')}

def booleano(valor):
    if valor.upper() not in ('SI', 'SÍ', 'NO'):
        raise ValueError('se esperaba SI o NO')
    return valor.upper() != 'NO'

def filtros_colocap(zona=None):
    agrupacion = 'Total Periodo' if zona is None else 'Total Pto_Zona'
    resumen = {'AGRUPACION': 'ZONA', 'SEXO': 9}
    proporciones = {'AGRUPACION': agrupacion, 'ESCALA': 'Anual'}
    if zona is not None:
        resumen['PTO_ZONA'] = str(zona)
        proporciones['PTO_ZONA'] = str(zona)
    return dict(filtros_resumen=resumen, filtros_tallas=dict(proporciones),
                filtros_sexual=dict(proporciones))

def construir_procesos(nombre, codigo, especie, anio, tablas, archivos, artes):
    """Construye selecciones, reglas y destinos; no lee perfiles ni históricos."""
    perfil = dict(especie=nombre, especie_codigo=codigo, anio=anio,
                  procesos=[], regresiones=[], talla_edad=[], matriz_composicion=[])
    merluza = codigo == 2
    talla_min, talla_max, paso, edad_max = (20, 140, 2, 24) if merluza else (35, 249, 5, 40)
    directorio = Path(tablas)
    if not directorio.name:
        raise ValueError('DIRECTORIO_TABLAS debe indicar un nombre de carpeta')
    if not directorio.name.casefold().endswith(nombre.casefold()):
        directorio = directorio.with_name(f'{directorio.name} {nombre}')
    carpeta_claves = directorio / 'Clave_talla_edad'
    carpeta_regresiones = directorio / 'Regresiones' / f'{especie}_{anio}'
    def clave(id, pesqueria=None, puerto=None):
        parametros = dict(id=id, archivo=archivos['ARCHIVO_TALLA_EDAD'],
            salida=str(carpeta_claves / ('claves definitivas' if merluza else '') / f'{especie}_{anio}_{id}.xlsx'),
            pesqueria=pesqueria, paso_talla=paso, talla_min=talla_min, talla_max=talla_max,
            edad_min=1, edad_agrupada=edad_max, columna_edad='GEC' if merluza else None,
            metodo='MAD', umbral=3.5, columna_especie='ESPECIE', valor_especie=nombre.upper())
        if merluza:
            parametros.update(zona=None, depurar_antes_mezcla=True,
                salida_preliminar=str(carpeta_claves / 'claves preliminares' / f'{especie}_{anio}_{id}.xlsx'))
        if puerto is not None:
            parametros.update(puerto=puerto, mezcla_secuencial_minimo=500)
        perfil['talla_edad'].append(parametros)
    if merluza:
        clave('industrial_total', pesqueria=25)
        clave('artesanal_total', pesqueria=55)
    for arte, configuracion in artes.items():
        flota, pesqueria_co, pesqueria_reg = FLOTAS[arte]
        for zona in ([None, 1, 2] if merluza else [None]):
            sufijo = 'total' if zona is None else f'zona{zona}'
            proceso = dict(id=f'{arte}_{sufijo}',
                **{campo.lower(): configuracion[campo] for campo in ('ARCHIVO_RESUMEN', 'ARCHIVO_TALLAS', 'ARCHIVO_SEXUAL')},
                salida=str(directorio / 'COLOCAP' / f'{especie}_{anio}_{arte}_{sufijo}.xlsx'),
                pesqueria=pesqueria_co, flota=flota,
                nombre_hoja=f'{flota}_{"total" if zona is None else "Z"+str(zona)}_{anio}',
                titulo_hoja=f'COLOCAP {especie} {anio}' + (f' — {flota}' if merluza else ''),
                titulo_bloque=f'COLOCAP {anio} SP{codigo}', talla_min=talla_min,
                talla_max=talla_max if merluza else 245, paso_talla=paso, **filtros_colocap(zona))
            if merluza and arte == 'artesanal' and zona == 2:
                proceso['respaldo_sin_datos'] = dict(**filtros_colocap(1),
                    descripcion='Artesanal sur: datos de COLOCAP del norte por ausencia de datos del sur')
            perfil['procesos'].append(proceso)
        regresion = dict(id=arte, fuentes=[dict(archivo=archivos['ARCHIVO_PESO_LONGITUD'], flota=flota)],
            salida=str(carpeta_regresiones / flota), codigos_pesqueria=[pesqueria_reg] if merluza else None,
            columna_especie='COD_ESPECIE' if merluza else 'COS_ESPECIE',
            columna_pesqueria='COD_PESQUERIA' if merluza else None,
            zonas_validas=[1, 2] if merluza and arte != 'artesanal' else None,
            factor_longitud=1, factor_peso=1)
        if merluza:
            # En Artesanal, ZONA=1/2 son subdivisiones del mismo dominio Norte.
            # El ajuste Total reúne la muestra y se comparte con la matriz Sur.
            regresion['columna_zona'] = None if arte == 'artesanal' else 'ZONA'
        else:
            regresion.update(columna_zona=None, columna_anio='ANIO_MUESTRO')
        perfil['regresiones'].append(regresion)
        if not merluza:
            clave('palangre_total')
        for zona in ([1, 2] if merluza else [None]):
            norte = zona == 1
            sufijo = ('norte' if norte else 'sur') if merluza else 'total'
            id = f'{arte}_{sufijo}'
            if merluza:
                clave(id, puerto=PUERTOS[arte][zona-1])
            artesanal = arte == 'artesanal'
            tipo_desembarque = 'ARTESANAL' if artesanal else 'INDUSTRIAL'
            desembarque = dict(archivo=archivos[f'ARCHIVO_DESEMBARQUE_{tipo_desembarque}'],
                hoja=archivos[f'HOJA_DESEMBARQUE_{tipo_desembarque}'],
                celdas=configuracion[f'CELDAS_DESEMBARQUE_{sufijo.upper()}' if merluza else 'CELDAS_DESEMBARQUE'])
            if not merluza:
                desembarque['verificar'] = {'C80': 'BACALAO DE PROFUNDIDAD'}
            elif artesanal:
                desembarque['verificar'] = {'A1': 'Pesquería', 'B1': 'Zona', 'C1': 'Desembarque (t)',
                    'A2': 'Artesanal', 'B2' if norte else 'B3': sufijo.capitalize()}
            else:
                desembarque['verificar'] = {'C2': 'MERLUZA DEL SUR', 'E3' if norte else 'F3': 'NOR EXT' if norte else 'SUR EXT'}
                if arte == 'arrastre':
                    desembarque['verificar'].update(C5='Arrastre Fábrica', C6='Arrastresurimero', C8='Arrastre Hielero')
                    if norte:
                        desembarque['verificar']['D3'] = "Al norte 41°28,6'"
                    elif 'F8' in desembarque['celdas']:
                        desembarque['celdas_vacias_cero'] = ['F8']
                else:
                    desembarque['verificar']['C7'] = 'Palangre Fábrica'
            composicion = dict(id=id, colocap_id=f'{arte}_zona{zona}' if merluza else 'palangre_total',
                regresion_id=arte, clave_id=id if merluza else 'palangre_total',
                dominio_regresion=f'Zona_{zona}' if merluza and not artesanal else 'Total',
                metodo_regresion='MAD',
                diseno=especie, zona='Zona '+sufijo.capitalize() if merluza else 'Total',
                directorio_salida=str(directorio / 'Matriz_composicion'),
                nombre_archivo=(NOMBRES_COMPOSICION[arte][zona-1]+'{sexo_sigla}{anio}.xlsx') if merluza else 'BPPALS{sexo_sigla}{anio}_generado.xlsx',
                sexos=['Machos', 'Hembras'], desembarque=desembarque, pendientes=[])
            if merluza:
                composicion.update(area='al norte de 47° S' if norte else 'al sur de 47° S',
                    nombre_pesqueria='Espinel interior' if artesanal else flota, dominio_historico_merluza=id)
            else:
                composicion.update(edad_min=4, completar_bacalao=True)
                if desembarque['hoja'].startswith('REEMPLAZAR'):
                    composicion['pendientes'] = ['Confirmar archivos reales y dominio de las fuentes de bacalao',
                        'Configurar hoja, celda y rótulos de desembarque según fuente real']
            perfil['matriz_composicion'].append(composicion)
    perfil['tabla_regresiones'] = dict(metodo='Original', directorio_salida=str(carpeta_regresiones),
        nombre_archivo=f'peso_longitud_{especie}_{anio}', nombres_zonas={'Zona_1': 'Zona Norte', 'Zona_2': 'Zona Sur'})
    return perfil

def leer_configuracion(archivo):
    archivo = Path(archivo).resolve()
    parser = configparser.ConfigParser(interpolation=None, delimiters=('=',), comment_prefixes=('#',))
    parser.optionxform = str.upper
    try:
        parser.read_string(archivo.read_text(encoding='utf-8-sig'), source=str(archivo))
    except configparser.Error as error:
        raise ValueError(str(error)) from error
    campos = {
        'GENERAL': {'ESPECIE', 'ANIO', 'DIRECTORIO_TABLAS', 'GENERAR_GRAFICOS', *ETAPAS},
        'ARCHIVOS': {'ARCHIVO_TALLA_EDAD', 'ARCHIVO_PESO_LONGITUD', 'ARCHIVO_DESEMBARQUE_INDUSTRIAL',
                     'ARCHIVO_DESEMBARQUE_ARTESANAL', 'HOJA_DESEMBARQUE_INDUSTRIAL', 'HOJA_DESEMBARQUE_ARTESANAL'},
    }
    por_arte = {'ARCHIVO_RESUMEN', 'ARCHIVO_TALLAS', 'ARCHIVO_SEXUAL',
                'CELDAS_DESEMBARQUE_NORTE', 'CELDAS_DESEMBARQUE_SUR', 'CELDAS_DESEMBARQUE'}
    if parser.defaults():
        raise ValueError('No se admite el bloque DEFAULT')
    if not parser.has_section('GENERAL'):
        raise ValueError('Falta el bloque GENERAL')
    especie_elegida = parser['GENERAL'].get('ESPECIE', '').strip().upper()
    if especie_elegida not in ESPECIES:
        raise ValueError('ESPECIE debe ser MERLUZA DEL SUR o BACALAO DE PROFUNDIDAD')
    prefijo = 'MERLUZA' if especie_elegida == 'MERLUZA DEL SUR' else 'BACALAO'
    seleccion = configparser.ConfigParser(interpolation=None)
    seleccion.optionxform = str.upper
    for bloque in parser.sections():
        partes = bloque.split()
        if bloque == 'GENERAL':
            grupo = bloque
        elif len(partes) == 2 and partes[0] in ('MERLUZA', 'BACALAO'):
            grupo = partes[1]
        else:
            raise ValueError(f'Bloque desconocido: [{bloque}]')
        if grupo not in campos and grupo.lower() not in FLOTAS:
            raise ValueError(f'Bloque desconocido: [{bloque}]')
        if grupo == 'GENERAL' and bloque != 'GENERAL':
            raise ValueError('Los parámetros generales se escriben solo en GENERAL')
        desconocidas = set(parser[bloque]) - campos.get(grupo, por_arte)
        if desconocidas:
            raise ValueError(f'[{bloque}] opciones desconocidas: {", ".join(sorted(desconocidas))}')
        if bloque == 'GENERAL' or partes[0] == prefijo:
            seleccion.add_section(grupo)
            for clave, valor in parser[bloque].items():
                seleccion.set(grupo, clave, valor)
    parser = seleccion
    if not parser.has_section('GENERAL') or not parser.has_section('ARCHIVOS'):
        raise ValueError('Se necesitan los bloques GENERAL y ARCHIVOS')
    variables = {**parser['GENERAL'], **parser['ARCHIVOS']}
    def expandir(valor, cadena=()):
        def sustituir(m):
            clave = m[1]
            if clave not in variables:
                raise ValueError(f'Variable no definida: {clave}')
            if clave in cadena:
                raise ValueError('Referencia circular: ' + ' -> '.join((*cadena, clave)))
            return expandir(variables[clave], (*cadena, clave))
        return re.sub(r'\{([A-Za-z_][A-Za-z0-9_]*)\}', sustituir, valor)
    def requerido(bloque, clave):
        valor = expandir(parser[bloque].get(clave, '')).strip()
        if not valor:
            raise ValueError(f'[{bloque}] falta {clave}')
        return valor
    especie = requerido('GENERAL', 'ESPECIE').upper()
    if especie not in ESPECIES:
        raise ValueError('ESPECIE debe ser MERLUZA DEL SUR o BACALAO DE PROFUNDIDAD')
    nombre, codigo, etiqueta = ESPECIES[especie]
    anio = int(requerido('GENERAL', 'ANIO'))
    if anio <= 0:
        raise ValueError('ANIO debe ser un entero positivo')
    config = {'ARCHIVO_CONFIGURACION': archivo}
    for clave in (*ETAPAS, 'GENERAR_GRAFICOS'):
        try:
            config[clave] = booleano(expandir(parser['GENERAL'].get(clave, 'SI')))
        except ValueError as error:
            raise ValueError(f'{clave}: {error}') from error
    if not any(config[k] for k in ETAPAS):
        raise ValueError('Debe activar al menos una etapa')
    requeridos = ['ARCHIVO_TALLA_EDAD', 'ARCHIVO_PESO_LONGITUD', 'ARCHIVO_DESEMBARQUE_INDUSTRIAL', 'HOJA_DESEMBARQUE_INDUSTRIAL']
    if codigo == 2:
        requeridos += ['ARCHIVO_DESEMBARQUE_ARTESANAL', 'HOJA_DESEMBARQUE_ARTESANAL']
    archivos = {k: requerido('ARCHIVOS', k) for k in requeridos}
    artes = {}
    for arte in (FLOTAS if codigo == 2 else ['palangre']):
        bloque = arte.upper()
        if not parser.has_section(bloque):
            raise ValueError(f'Falta el bloque [{bloque}]')
        claves = ['ARCHIVO_RESUMEN', 'ARCHIVO_TALLAS', 'ARCHIVO_SEXUAL']
        claves += ['CELDAS_DESEMBARQUE_NORTE', 'CELDAS_DESEMBARQUE_SUR'] if codigo == 2 else ['CELDAS_DESEMBARQUE']
        artes[arte] = {k: requerido(bloque, k) for k in claves}
        for clave in [c for c in claves if c.startswith('CELDAS_')]:
            celdas = [v.strip().upper() for v in artes[arte][clave].split(';')]
            if any(not re.fullmatch(r'[A-Z]+[1-9]\d*', c) for c in celdas) or len(set(celdas)) != len(celdas):
                raise ValueError(f'[{bloque}] {clave}: indique celdas distintas separadas por ;')
            artes[arte][clave] = celdas
    perfil = construir_procesos(nombre, codigo, etiqueta, anio,
        expandir(parser['GENERAL'].get('DIRECTORIO_TABLAS', 'Tablas')), archivos, artes)
    perfil['base'] = str(archivo.parent)
    config['PERFIL'] = perfil
    return config
