"""Regresiones longitud–peso por especie, flota y zona."""
from pathlib import Path
import json
import numpy as np
import pandas as pd
from scipy.stats import t
from lectura import leer_fuentes, leer_perfil
from metodos_matematicos import ajustar_regresion_longpeso, aplicar_outliers_longpeso, graficar_regresion
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

METODOS = ("Original", "Cook", "MAD", "IQR")
GRUPOS = {"Ambos": None, "Machos": 1, "Hembras": 2}


def asignar_zonas(datos, configuracion):
    """Zonificación explícita para fuentes sin ZONA; no detecta coordenadas."""
    columna=configuracion["columna"]
    if columna not in datos:
        raise ValueError(f"No existe la coordenada {columna}")
    d=datos.copy()
    valores=pd.to_numeric(d[columna],errors="coerce")
    formato=configuracion["formato"]
    if formato=="ddmmss_sur":
        magnitud=valores.abs()
        grados=np.floor(magnitud/10000)
        minutos=np.floor(magnitud/100)%100
        segundos=magnitud%100
        valida=(np.isfinite(valores)&magnitud.mod(1).eq(0)&grados.le(90)&minutos.lt(60)&segundos.lt(60)&
                ((grados<90)|((minutos==0)&(segundos==0))))
        latitud=-(grados+minutos/60+segundos/3600)
    elif formato=="decimal":
        latitud=valores
        valida=np.isfinite(valores)&valores.between(-90,90)
    else:
        raise ValueError("Formato de latitud no admitido; usar ddmmss_sur o decimal")
    norte=float(configuracion["limite_norte"])
    sur=float(configuracion["limite_sur"])
    if not (-90<=sur<norte<=90):
        raise ValueError("Límites de latitud inválidos")
    d["LATITUD_DECIMAL_ANALISIS"]=latitud.where(valida)
    zona=pd.Series(np.nan,index=d.index)
    zona.loc[valida & latitud.gt(norte)]=0
    zona.loc[valida & latitud.le(norte)&latitud.gt(sur)]=1
    zona.loc[valida & latitud.le(sur)]=2
    d["ZONA_CONFIGURADA"]=zona
    return d

def preparar_datos(datos, *, especie, codigos_pesqueria, columna_especie="COD_ESPECIE",
                    columna_pesqueria="COD_PESQUERIA", columna_zona="ZONA",
                    zonas_validas=(1,2), factor_longitud=1, factor_peso=1):
    requeridas = [columna_especie, "SEXO_ESPECIMEN", "LONGITUD_ESPECIMEN", "PESO_ESPECIMEN"]
    if codigos_pesqueria is not None:
        requeridas.append(columna_pesqueria)
    if zonas_validas is not None:
        requeridas.append(columna_zona)
    faltantes = set(requeridas)-set(datos.columns)
    if faltantes:
        raise ValueError(f"Faltan columnas para regresión: {sorted(faltantes)}")
    if codigos_pesqueria is not None and not codigos_pesqueria or zonas_validas is not None and not zonas_validas:
        raise ValueError("Los filtros de pesquería y zona no pueden ser listas vacías")
    if not all(np.isfinite(f) and f>0 for f in [factor_longitud,factor_peso]):
        raise ValueError("Los factores de conversión deben ser finitos positivos")
    d = datos.copy()
    for c in dict.fromkeys(requeridas):
        d[c+"_ORIGINAL"] = d[c]
        d[c] = pd.to_numeric(d[c],errors="coerce")
    d["LONGITUD_ESPECIMEN"] *= factor_longitud
    d["PESO_ESPECIMEN"] *= factor_peso
    d["ZONA_ANALISIS"] = d[columna_zona] if zonas_validas is not None else "Total"
    razones = pd.DataFrame({
        "ESPECIE_NO_SELECCIONADA": ~d[columna_especie].eq(especie),
        "SEXO_INVALIDO": ~d["SEXO_ESPECIMEN"].isin([1,2]),
        "LONGITUD_INVALIDA": ~np.isfinite(d["LONGITUD_ESPECIMEN"]) | d["LONGITUD_ESPECIMEN"].le(0),
        "PESO_INVALIDO": ~np.isfinite(d["PESO_ESPECIMEN"]) | d["PESO_ESPECIMEN"].le(0),
    },index=d.index)
    if codigos_pesqueria is not None:
        razones["PESQUERIA_NO_SELECCIONADA"] = ~d[columna_pesqueria].isin(codigos_pesqueria)
    if zonas_validas is not None:
        razones["ZONA_INVALIDA"] = ~d["ZONA_ANALISIS"].isin(zonas_validas)
    d["MOTIVO_EXCLUSION"] = razones.apply(lambda r: ";".join(r.index[r]),axis=1)
    return d.loc[~razones.any(axis=1)].copy(), d.loc[razones.any(axis=1)].copy()

def ajustar(datos):
    if len(datos)<3:
        raise ValueError("Se requieren al menos tres observaciones")
    if datos["LONGITUD_ESPECIMEN"].nunique()<2:
        raise ValueError("No hay variación en longitud")
    if datos["PESO_ESPECIMEN"].nunique()<2:
        raise ValueError("No hay variación en peso para estimar r")
    a,b,log_a,r = ajustar_regresion_longpeso(datos)
    if not np.isfinite([a,b,log_a,r]).all():
        raise ValueError("El ajuste produjo parámetros no finitos")
    x=np.log10(datos["LONGITUD_ESPECIMEN"].to_numpy())
    y=np.log10(datos["PESO_ESPECIMEN"].to_numpy())
    n=len(x)
    sxx=np.sum((x-x.mean())**2)
    varianza=np.sum((y-log_a-b*x)**2)/(n-2)
    critico=t.ppf(.975,n-2)
    margen_b=critico*np.sqrt(varianza/sxx)
    margen_a=critico*np.sqrt(varianza*(1/n+x.mean()**2/sxx))
    return dict(a=a,b=b,log_a=log_a,r_PEARSON=r,**{"r²":r*r},
                a_IC95_INF=10**(log_a-margen_a),a_IC95_SUP=10**(log_a+margen_a),
                b_IC95_INF=b-margen_b,b_IC95_SUP=b+margen_b)

def analizar_dominio(datos, contexto):
    filas, resultados, diagnosticos = [], {}, []
    for grupo,sexo in GRUPOS.items():
        original=datos.copy() if sexo is None else datos.loc[datos["SEXO_ESPECIMEN"].eq(sexo)].copy()
        resultados[grupo]={}
        try:
            inicial=ajustar(original)
            fallo=None
        except ValueError as error:
            inicial=None
            fallo=str(error)
        for metodo in METODOS:
            fila=dict(contexto,GRUPO=grupo,METODO=metodo,N_ORIGINAL=len(original),
                      REGISTROS_UTILIZADOS=0,ELIMINADOS=0,ESTADO="SIN_AJUSTE",DETALLE="")
            if fallo:
                fila["DETALLE"]=fallo
                filas.append(fila)
                continue
            if metodo=="Original":
                retenidos=original.copy()
                eliminados=0
            else:
                diagnostico,retenidos,eliminados=aplicar_outliers_longpeso(
                    original,metodo.lower(),inicial["b"],inicial["log_a"])
                diagnostico=diagnostico.assign(GRUPO_ANALISIS=grupo,METODO=metodo,**contexto)
                diagnosticos.append(diagnostico)
            fila.update(REGISTROS_UTILIZADOS=len(retenidos),ELIMINADOS=int(eliminados))
            try:
                parametros=inicial if metodo=="Original" else ajustar(retenidos)
                fila.update(parametros,ESTADO="OK")
                resultados[grupo][metodo]=dict(datos=retenidos,**parametros)
            except ValueError as error:
                fila["DETALLE"]=str(error)
            filas.append(fila)
    return pd.DataFrame(filas),resultados,diagnosticos

NUMERICAS = ['a', 'b', 'r²', 'a_IC95_INF', 'a_IC95_SUP', 'b_IC95_INF', 'b_IC95_SUP']


def preparar_resumen(tabla, metodo='Original', nombres_zonas=None):
    """Devuelve bloques por flota con estimaciones e IC marginales del 95%."""
    d = tabla[tabla.METODO.str.casefold().eq(metodo.casefold())].copy()
    if d.empty:
        raise ValueError(f'No hay resultados del método {metodo}')
    if d[['COD_ESPECIE', 'ANIO']].drop_duplicates().shape[0] != 1:
        raise ValueError('La tabla debe corresponder a una sola especie y año')
    if d.duplicated(['FLOTA', 'DOMINIO', 'GRUPO']).any():
        raise ValueError('Hay ajustes duplicados por flota, dominio y sexo')
    for c in NUMERICAS:
        d[c] = pd.to_numeric(d[c].astype(str).str.replace(',', '.', regex=False), errors='coerce') if c in d else float('nan')
    zonas = nombres_zonas or {'Zona_1': 'Zona Norte', 'Zona_2': 'Zona Sur'}
    bloques = []
    for flota, datos in d.groupby('FLOTA', sort=False):
        dominios = [v for v in datos.DOMINIO.unique() if v != 'Total'] or ['Total']
        filas = []
        for dominio in dominios:
            for grupo in ('Machos', 'Hembras', 'Ambos'):
                seleccion = datos[datos.DOMINIO.eq(dominio) & datos.GRUPO.eq(grupo)]
                if seleccion.empty:
                    continue
                r = seleccion.iloc[0]
                etiqueta = f'{zonas.get(dominio, dominio)} {grupo}'
                n = int(r.REGISTROS_UTILIZADOS)
                if r.ESTADO != 'OK':
                    filas.extend([[etiqueta, 'Sin ajuste', None, None, n],
                                  ['Lím. inferior', None, None, None, None],
                                  ['Lím. superior', None, None, None, None]])
                else:
                    if r[NUMERICAS].isna().any():
                        raise ValueError(f'Parámetros incompletos: {flota}, {dominio}, {grupo}')
                    filas.extend([[etiqueta, r.a, r.b, r['r²'], n],
                                  ['Lím. inferior', r.a_IC95_INF, r.b_IC95_INF, None, None],
                                  ['Lím. superior', r.a_IC95_SUP, r.b_IC95_SUP, None, None]])
        bloques.append((flota, filas))
    especie, anio = d.iloc[0][['ESPECIE', 'ANIO']]
    titulo = f'Relaciones peso–longitud — {especie}, {int(anio)}'
    nota = f'Ajuste {metodo}. W(g) = a × L(cm)^b. IC marginales del 95%. r² en escala log10. N = registros utilizados.'
    return titulo, nota, bloques


def escribir_resumen_excel(libro, tabla, metodo='Original', nombres_zonas=None):
    titulo, nota, bloques = preparar_resumen(tabla, metodo, nombres_zonas)
    hoja = libro.create_sheet('Resumen', 0)
    ultima = len(bloques)*6-1
    hoja.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ultima)
    hoja.cell(1, 1, titulo).font = Font(name='Calibri', size=14, bold=True)
    hoja.merge_cells(start_row=2, start_column=1, end_row=3, end_column=ultima)
    hoja.cell(2, 1, nota).alignment = Alignment(wrap_text=True, vertical='center')
    borde = Border(bottom=Side(style='thin', color='000000'))
    max_filas = 0
    for indice, (flota, filas) in enumerate(bloques):
        inicio = indice*6+1
        for j, encabezado in enumerate([flota.upper(), 'a', 'b', 'r²', 'N']):
            c = hoja.cell(5, inicio+j, encabezado)
            c.font = Font(name='Calibri', size=11, bold=True)
            c.border = borde
        for i, valores in enumerate(filas):
            for j, valor in enumerate(valores):
                c = hoja.cell(6+i, inicio+j, valor)
                c.font = Font(name='Calibri', size=11, bold=i%3 == 0)
                c.alignment = Alignment(horizontal='left' if j == 0 else 'right')
                if j in (1, 2): c.number_format = '0.0000000'
                elif j == 3: c.number_format = '0.00'
                elif j == 4: c.number_format = '#,##0'
                if i%9 == 8: c.border = borde
            hoja.row_dimensions[6+i].height = 20
        for j, ancho in enumerate([28, 17, 17, 10, 12, 3]):
            hoja.column_dimensions[get_column_letter(inicio+j)].width = ancho
        max_filas = max(max_filas, len(filas))
    hoja.freeze_panes = 'B6'
    hoja.sheet_view.showGridLines = False
    hoja.sheet_properties.pageSetUpPr.fitToPage = True
    hoja.page_setup.orientation = 'landscape'
    hoja.page_setup.paperSize = hoja.PAPERSIZE_A3 if len(bloques)>2 else hoja.PAPERSIZE_A4
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 1
    hoja.print_area = f'A1:{get_column_letter(ultima)}{5+max_filas}'
    return hoja


def _texto(valor, columna):
    if valor is None: return ''
    if isinstance(valor, str): return valor
    if columna == 4: return f'{int(valor):,}'.replace(',', '.')
    return f'{valor:.{2 if columna == 3 else 7}f}'.replace('.', ',')


def exportar_resumen_jpg(tabla, archivo, metodo='Original', nombres_zonas=None):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    titulo, nota, bloques = preparar_resumen(tabla, metodo, nombres_zonas)
    alto = max(3.2, 1.8+.32*max(len(f) for _, f in bloques))
    with plt.rc_context({'font.family': 'DejaVu Sans'}):
        fig, ejes = plt.subplots(1, len(bloques), figsize=(7.6*len(bloques), alto), squeeze=False)
        try:
            fig.suptitle(titulo, fontsize=14, fontweight='bold', y=.98)
            for ax, (flota, filas) in zip(ejes[0], bloques):
                ax.axis('off')
                valores = [[_texto(v, j) for j,v in enumerate(f)] for f in filas]
                tabla_graf = ax.table(cellText=valores, colLabels=[flota.upper(), 'a', 'b', 'r²', 'N'],
                                     colWidths=[.36, .21, .21, .10, .12], cellLoc='right', bbox=[0, 0, 1, 1])
                tabla_graf.auto_set_font_size(False)
                tabla_graf.set_fontsize(10)
                for (fila, col), celda in tabla_graf.get_celld().items():
                    celda.visible_edges = 'B' if fila == 0 or fila%9 == 0 else ''
                    celda.set_linewidth(.6)
                    celda.set_facecolor('white')
                    celda.get_text().set_ha('left' if col == 0 else 'right')
                    if fila == 0 or (fila-1)%3 == 0: celda.get_text().set_fontweight('bold')
            fig.text(.025, .025, nota, fontsize=9, ha='left')
            fig.subplots_adjust(left=.025, right=.985, top=.90, bottom=.10, wspace=.04)
            Path(archivo).parent.mkdir(parents=True, exist_ok=True)
            fig.savefig(archivo, format='jpg', dpi=220, facecolor='white')
        finally:
            plt.close(fig)


def exportar_tabla_resumen(tabla, archivo_excel, archivo_jpg, metodo='Original', nombres_zonas=None):
    Path(archivo_excel).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(archivo_excel, engine='openpyxl') as writer:
        escribir_resumen_excel(writer.book, tabla, metodo, nombres_zonas)
    exportar_resumen_jpg(tabla, archivo_jpg, metodo, nombres_zonas)


def exportar_resultados(tabla, archivo, metadatos):
    with pd.ExcelWriter(archivo,engine="openpyxl") as writer:
        escribir_resumen_excel(writer.book, tabla, metadatos.get('metodo_tabla', 'Original'))
        for metodo in METODOS:
            tabla.loc[tabla.METODO.eq(metodo)].to_excel(writer,sheet_name=metodo,index=False)
            hoja=writer.sheets[metodo]
            hoja.freeze_panes="A2"
            hoja.auto_filter.ref=hoja.dimensions
            for columna in hoja.columns:
                hoja.column_dimensions[columna[0].column_letter].width=22
                if columna[0].value in ["a","b","r_PEARSON","r²","a_IC95_INF","a_IC95_SUP","b_IC95_INF","b_IC95_SUP"]:
                    for celda in columna[1:]:
                        celda.number_format="0.#######"
        pd.DataFrame({"Configuración":[json.dumps(metadatos,ensure_ascii=False,indent=2)]}).to_excel(
            writer,sheet_name="Metadatos",index=False)
        pd.DataFrame({"Notas":[
            "Modelo W(g) = a × L(cm)^b; mínimos cuadrados en escala log10.",
            "r_PEARSON y r² corresponden al ajuste log10. IC marginales del 95% con t de Student.",
            "Cada método parte del conjunto original de su sexo y dominio; no se aplican en cadena.",
            "Cook: D > 4/N. MAD: puntaje modificado absoluto > 3.5. IQR: residuos fuera de Q1/Q3 ± 3.5 IQR.",
            "Mediana de desviaciones absolutas nula en MAD: se retienen todos los registros (criterio del código anterior).",
            "Ambos se ajusta y depura independientemente de Machos y Hembras.",
            "SIN_AJUSTE informa grupos insuficientes o degenerados. No se asigna un método a la matriz de composición.",
            "El año identifica la fuente; solo filtra registros si se configura columna_anio.",
        ]}).to_excel(writer,sheet_name="Notas",index=False)

def limpiar_salidas_anteriores(salida):
    """Retira únicamente archivos auxiliares conocidos del módulo anterior."""
    salida = Path(salida)
    for nombre in ("registros_validos.csv", "registros_excluidos.csv",
                   "resultados_regresiones.csv", "metadatos.json"):
        (salida / nombre).unlink(missing_ok=True)
    for carpeta in salida.iterdir():
        if not carpeta.is_dir() or not (carpeta.name == "Total" or carpeta.name.startswith("Zona_")):
            continue
        nombres = {"diagnosticos.csv"}
        nombres.update(f"{grupo.lower()}_{metodo.lower()}.csv"
                       for grupo in GRUPOS for metodo in METODOS)
        for nombre in nombres:
            (carpeta / nombre).unlink(missing_ok=True)
        if not any(carpeta.iterdir()):
            carpeta.rmdir()


def generar_regresiones(*, fuentes, base, salida, especie, nombre_especie, anio,
                         codigos_pesqueria=None, columna_especie="COD_ESPECIE",
                         columna_pesqueria="COD_PESQUERIA", columna_zona="ZONA",
                         zonas_validas=(1,2), factor_longitud=1, factor_peso=1,
                         columna_anio=None, graficos=True, zonificacion=None,
                         metodo_tabla='Original'):
    if metodo_tabla not in METODOS:
        raise ValueError(f'Método de tabla inválido: {metodo_tabla}')
    datos=leer_fuentes(fuentes,base=base,especie=especie,anio=anio)
    if especie == 37:
        datos = datos.rename(columns={c: str(c).strip().upper() for c in datos.columns})
        for destino, alternativas in {'ANIO_MUESTRO': ('AÑO', 'ANIO'),
                'COS_ESPECIE': ('COD_ESPECIE',), 'SEXO_ESPECIMEN': ('SEXO',)}.items():
            if destino not in datos:
                for origen in alternativas:
                    if origen in datos:
                        datos = datos.rename(columns={origen: destino})
                        break
    if zonificacion:
        datos=asignar_zonas(datos,zonificacion)
        columna_zona="ZONA_CONFIGURADA"
    flotas=datos["FLOTA"].unique()
    if len(flotas)!=1:
        raise ValueError("Cada proceso de regresión debe contener una sola flota")
    duplicados=datos.duplicated(["ORIGEN_ARCHIVO","ORIGEN_HOJA","ORIGEN_FILA"])
    if duplicados.any():
        raise ValueError("Hay registros de origen repetidos")
    validos,excluidos=preparar_datos(datos,especie=especie,codigos_pesqueria=codigos_pesqueria,
        columna_especie=columna_especie,columna_pesqueria=columna_pesqueria,columna_zona=columna_zona,
        zonas_validas=zonas_validas,factor_longitud=factor_longitud,factor_peso=factor_peso)
    if columna_anio:
        if columna_anio not in validos:
            raise ValueError(f"No existe la columna de año {columna_anio}")
        mascara=pd.to_numeric(validos[columna_anio],errors="coerce").eq(anio)
        fuera=validos.loc[~mascara].assign(MOTIVO_EXCLUSION="ANIO_NO_SELECCIONADO")
        excluidos=pd.concat([excluidos,fuera],ignore_index=True)
        validos=validos.loc[mascara].copy()
    if validos.empty:
        raise ValueError("No hay pares longitud–peso válidos con los filtros configurados")
    salida=Path(salida)
    salida.mkdir(parents=True,exist_ok=True)
    columnas_csv = ["ZONA_ANALISIS", "SEXO_ESPECIMEN", "LONGITUD_ESPECIMEN", "PESO_ESPECIMEN"]
    validos[columnas_csv].to_csv(salida/"datos_filtrados.csv",sep=";",index=False,encoding="utf-8-sig")
    tablas=[]
    dominios=[("Total",validos)]+[(f"Zona_{z}",validos.loc[validos.ZONA_ANALISIS.eq(z)]) for z in (zonas_validas or ())]
    for dominio,subgrupo in dominios:
        contexto=dict(ESPECIE=nombre_especie,COD_ESPECIE=especie,ANIO=anio,FLOTA=flotas[0],DOMINIO=dominio)
        tabla,resultados,_=analizar_dominio(subgrupo,contexto)
        tablas.append(tabla)
        if graficos:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            carpeta = salida / dominio
            carpeta.mkdir(exist_ok=True)
            for grupo,por_metodo in resultados.items():
                fig,ejes=plt.subplots(2,2,figsize=(16,12),sharex=True,sharey=True)
                try:
                    fig.suptitle(f"{nombre_especie} — {flotas[0]} — {dominio} — {grupo}")
                    for ax,metodo in zip(ejes.flat,METODOS):
                        if metodo not in por_metodo:
                            ax.text(.5,.5,"Sin ajuste disponible",ha="center",transform=ax.transAxes)
                            ax.set_title(metodo)
                            continue
                        r=por_metodo[metodo]
                        graficar_regresion(r["datos"],r["a"],r["b"],r["r_PEARSON"],metodo,ax=ax,grupo=grupo)
                    fig.tight_layout(rect=(0,0,1,.96))
                    fig.savefig(carpeta/f"regresion_{grupo.lower()}.png",dpi=200)
                finally:
                    plt.close(fig)
    tabla=pd.concat(tablas,ignore_index=True)
    metadatos=dict(fuentes=fuentes,base=str(Path(base).resolve()),anio=anio,
        columna_anio=columna_anio,especie=especie,codigos_pesqueria=codigos_pesqueria,
        columna_especie=columna_especie,columna_pesqueria=columna_pesqueria,
        columna_zona=columna_zona,zonas_validas=zonas_validas,
        factor_longitud=factor_longitud,factor_peso=factor_peso,zonificacion=zonificacion,
        recibidos=len(datos),validos=len(validos),excluidos=len(excluidos),metodo_tabla=metodo_tabla)
    exportar_resultados(tabla,salida/"resultados_regresiones.xlsx",metadatos)
    exportar_resumen_jpg(tabla, salida/'tabla_peso_longitud.jpg', metodo_tabla)
    limpiar_salidas_anteriores(salida)
    if zonas_validas is None:
        for carpeta in salida.glob('Zona_*'):
            if carpeta.is_dir():
                for grupo in GRUPOS:
                    (carpeta / f'regresion_{grupo.lower()}.png').unlink(missing_ok=True)
                if not any(carpeta.iterdir()):
                    carpeta.rmdir()
    print(f"{flotas[0]}: {len(validos)} válidos; {len(excluidos)} excluidos. Archivo generado: resultados_regresiones.xlsx",flush=True)
    for motivo, cantidad in excluidos['MOTIVO_EXCLUSION'].value_counts().items():
        print(f'  Antes del ajuste: {motivo} = {cantidad}', flush=True)
    return tabla


def generar_regresiones_perfil(archivo, *, graficos=True):
    """Ejecuta las flotas y exporta una tabla conjunta de la especie en Excel/JPG."""
    perfil, base = leer_perfil(archivo)
    configuracion = perfil.get('tabla_regresiones', {})
    metodo = configuracion.get('metodo', 'Original')
    tablas = []
    for proceso in perfil.get('regresiones', []):
        parametros = {k: v for k,v in proceso.items() if k not in ('id', 'salida')}
        tablas.append(generar_regresiones(**parametros, base=base, salida=base/proceso['salida'],
            especie=perfil['especie_codigo'], nombre_especie=perfil['especie'], anio=perfil['anio'],
            graficos=graficos, metodo_tabla=metodo))
    if not tablas:
        raise ValueError('El perfil no define regresiones')
    tabla = pd.concat(tablas, ignore_index=True)
    salida = base/configuracion.get('directorio_salida', 'Tablas/Regresiones')
    nombre = configuracion.get('nombre_archivo', f"tabla_peso_longitud_sp{perfil['especie_codigo']}_{perfil['anio']}")
    if Path(nombre).name != nombre or Path(nombre).suffix:
        raise ValueError('El nombre de la tabla debe ser un nombre base sin extensión')
    exportar_tabla_resumen(tabla, salida/f'{nombre}.xlsx', salida/f'{nombre}.jpg', metodo,
                          configuracion.get('nombres_zonas'))
    return tabla
