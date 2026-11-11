"""Controlador de COLOCAP, regresiones, claves talla–edad y composición."""
import argparse
import sys
from pathlib import Path

from analisis_colocap import generar_analisis_colocap
from clave_pesotalla import generar_claves_perfil
from configuracion import ETAPAS, leer_configuracion
from lectura import leer_perfil
from matriz_de_composicion import generar_composiciones, preparar_composiciones
from regresion import generar_regresiones_perfil


def preparar_ejecucion(configuracion):
    """Comprueba fuentes y dependencias antes de comenzar a reemplazar salidas.

    Las salidas de etapas activadas pueden no existir todavía. Las de etapas
    desactivadas se exigen cuando composición las necesita.
    """
    archivo = configuracion['PERFIL']
    perfil, base = leer_perfil(archivo)
    for campo in ('especie', 'especie_codigo', 'anio'):
        if campo not in perfil:
            raise ValueError(f'Falta {campo} en el perfil')
    entradas, salidas = set(), set()

    def ruta(valor):
        return (base / valor).resolve()

    def procesos(seccion):
        lista = perfil.get(seccion, [])
        if not lista:
            raise ValueError(f'El perfil no define {seccion}')
        ids = [p['id'] for p in lista]
        if len(set(ids)) != len(ids):
            raise ValueError(f'Hay identificadores repetidos en {seccion}')
        return lista

    colocap = []
    if configuracion['EJECUTAR_COLOCAP']:
        for proceso in procesos('procesos'):
            parametros = {k: v for k, v in proceso.items() if k != 'id'}
            for campo in ('archivo_resumen', 'archivo_tallas', 'archivo_sexual', 'salida'):
                parametros[campo] = ruta(parametros[campo])
            parametros['especie'] = perfil['especie_codigo']
            colocap.append(parametros)
            entradas.update(parametros[c] for c in ('archivo_resumen', 'archivo_tallas', 'archivo_sexual'))
            salidas.add(parametros['salida'])
    if configuracion['EJECUTAR_REGRESIONES']:
        for proceso in procesos('regresiones'):
            entradas.update(ruta(fuente['archivo']) for fuente in proceso['fuentes'])
            salidas.add(ruta(proceso['salida']) / 'resultados_regresiones.xlsx')
    if configuracion['EJECUTAR_CLAVES']:
        for proceso in procesos('talla_edad'):
            entradas.add(ruta(proceso['archivo']))
            salidas.add(ruta(proceso['salida']))
            if proceso.get('salida_preliminar'):
                salidas.add(ruta(proceso['salida_preliminar']))
    composiciones = []
    if configuracion['EJECUTAR_COMPOSICION']:
        procesos('matriz_composicion')
        composiciones = preparar_composiciones(archivo, verificar_fuentes=False)
        if not composiciones:
            raise ValueError('No hay composiciones habilitadas en el perfil')
        for proceso in composiciones:
            for campo in ('archivo_colocap', 'archivo_regresion', 'archivo_matriz'):
                fuente = ruta(proceso[campo])
                if fuente not in salidas:
                    entradas.add(fuente)
            if 'archivo' in proceso['desembarque']:
                entradas.add(ruta(proceso['desembarque']['archivo']))
            entradas.add(Path(__file__).resolve().parent / 'plantillas_composicion' /
                         f"{proceso['diseno']}_{proceso['sexo'].lower()}.xlsx")
            if proceso['salida'].resolve() in entradas | salidas:
                raise ValueError(f"La composición sobrescribiría una fuente: {proceso['salida']}")
        destinos = {p['salida'].resolve() for p in composiciones}
        if destinos & entradas:
            raise ValueError('Una composición sobrescribiría una fuente de otra composición')
    if entradas & salidas:
        raise ValueError('Una salida coincide con una fuente de entrada: ' +
                         ', '.join(str(p) for p in sorted(entradas & salidas)))
    faltantes = sorted(p for p in entradas if not p.is_file())
    if faltantes:
        raise FileNotFoundError('Faltan archivos necesarios:\n' + '\n'.join(f'  {p}' for p in faltantes))
    return perfil, colocap, composiciones


def main(argumentos=None):
    base = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=base / 'configuracion.txt',
                        help='Archivo de configuración (por defecto, junto a main.py)')
    parser.add_argument('--validar', action='store_true',
                        help='Comprueba configuración y existencia de archivos sin generar salidas')
    args = parser.parse_args(argumentos)
    etapa = 'Validación inicial'
    try:
        config = leer_configuracion(args.config)
        perfil, colocap, composiciones = preparar_ejecucion(config)
        print(f"{perfil['especie']} — {perfil['anio']}\nConfiguración: {config['ARCHIVO_CONFIGURACION'].name}", flush=True)
        print('Etapas: ' + ', '.join(nombre for clave, nombre in ETAPAS.items() if config[clave]), flush=True)
        if args.validar:
            print('Configuración y archivos comprobados. No se generaron salidas.', flush=True)
            return 0
        for clave, etapa in ETAPAS.items():
            if not config[clave]:
                print(f'Etapa omitida por configuración: {etapa}', flush=True)
                continue
            print(f'\n{etapa}', flush=True)
            if clave == 'EJECUTAR_COLOCAP':
                for parametros in colocap:
                    generar_analisis_colocap(**parametros)
                    print(f"COLOCAP generado: {parametros['salida'].name}", flush=True)
            elif clave == 'EJECUTAR_REGRESIONES':
                generar_regresiones_perfil(config['PERFIL'], graficos=config['GENERAR_GRAFICOS'])
            elif clave == 'EJECUTAR_CLAVES':
                generar_claves_perfil(config['PERFIL'])
            else:
                resultados = generar_composiciones(config['PERFIL'])
                print(f'Composiciones generadas: {len(resultados)} de {len(composiciones)} previstas.', flush=True)
        print('\nFlujo finalizado. Revise los avisos y el estado de cada composición.', flush=True)
        return 0
    except (OSError, ValueError, KeyError, TypeError, ImportError) as error:
        print(f'\nError en {etapa}: {error}', file=sys.stderr, flush=True)
        return 1


if __name__ == '__main__':
    sys.exit(main())
