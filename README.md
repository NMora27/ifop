# Automatización de Edad y Crecimiento — IFOP

Herramienta en Python para automatizar el procesamiento de datos de bacalao de profundidad. El programa genera el análisis COLOCAP, regresiones longitud–peso, matrices talla–edad y libros BPPALS, sin depender de una plantilla BPPALS externa.

La versión actual está configurada para bacalao de profundidad, pero los parámetros principales se modifican desde `configuracion.txt`.

## Flujo de trabajo

1. Lee los archivos DBF, CSV y Excel de entrada.
2. Genera el análisis COLOCAP a partir de los archivos DBF.
3. Calcula regresiones longitud–peso para machos, hembras y ambos sexos.
4. Aplica los métodos Cook, MAD e IQR como comparación de limpieza.
5. Genera gráficos de regresión y archivos de respaldo.
6. Crea matrices talla–edad depuradas mediante IQR.
7. Construye el libro BPPALS para el sexo seleccionado.

## Requisitos

- Python 3.
- Los archivos de entrada indicados en este manual.
- Paquetes de Python:

```bash
python -m pip install numpy pandas openpyxl matplotlib dbfread xlsxwriter
```

## Estructura del proyecto

Los scripts y los archivos de entrada deben mantenerse en la misma carpeta:

```text
ifop/
├── main.py
├── configuracion.txt
├── analisis_colocap.py
├── regresion_bacalao.py
├── clave_pesotalla.py
├── bppals_20XX.py
├── lectura.py
├── metodos_matematicos.py
├── E_Resumen.DBF
├── ET_Estructura_de_Tallas.DBF
├── ET_Proporción_Sexual.DBF
├── Palangre_2025.csv
├── Unido_Bacalao_de_profundidad_2025.xlsx
└── Desembarque 2025 industr para inf final EDAD.xlsx
```

Las carpetas `Tablas/` y `Graficos/` se crean automáticamente durante la ejecución.

## Archivos de entrada

| Archivo | Uso |
|---|---|
| `E_Resumen.DBF` | Obtiene el número de ejemplares para el análisis COLOCAP. |
| `ET_Estructura_de_Tallas.DBF` | Obtiene proporciones de talla por sexo. |
| `ET_Proporción_Sexual.DBF` | Obtiene la proporción de machos y hembras. |
| `Palangre_{AÑO}.csv` | Datos para la regresión longitud–peso. |
| `Unido_Bacalao_de_profundidad_{AÑO}.xlsx` | Datos para construir la matriz talla–edad. |
| `Desembarque {AÑO} industr para inf final EDAD.xlsx` | Obtiene el desembarque utilizado en BPPALS. |

Los nombres de los archivos deben coincidir con los definidos en `configuracion.txt`.

### Columnas requeridas

El archivo `Palangre_{AÑO}.csv` debe contener, al menos, las siguientes columnas:

```text
LATITUD
ESPECIE_OBJETIVO_LANCE
SEXO_ESPECIMEN
LONGITUD_ESPECIMEN
PESO_ESPECIMEN
```

El archivo `Unido_Bacalao_de_profundidad_{AÑO}.xlsx` debe contener:

```text
LONGITUD_DEL_PEZ_(CM)
Nº_DE_ANILLOS
CÓDIGO_DE_PESQUERÍA
SEXO
```

Los archivos DBF deben conservar las columnas utilizadas por el proceso, tales como `AGRUPACION`, `ESCALA`, `ESPECIE`, `SEXO`, `TALLA`, `PROPORCION` y `EJEMP_LONG`, según corresponda.

## Configuración

Antes de ejecutar el programa, abra `configuracion.txt` y modifique los parámetros de acuerdo con el año y análisis requerido.

Ejemplo:

```text
ANIO=2025
ESPECIE=BACALAO DE PROFUNDIDAD
ESPECIE_CODIGO=37
SEXO=Machos
ZONA=Palangre
AREA=área sur austral
FLOTA_DESEMBARQUE=Palangre Fábrica
EJECUTAR_BPPALS=SI
```

### Parámetros principales

| Parámetro | Descripción |
|---|---|
| `ANIO` | Año de los datos procesados. |
| `ESPECIE` | Nombre de la especie utilizada en BPPALS. |
| `ESPECIE_CODIGO` | Código numérico de la especie en los datos. |
| `SEXO` | Sexo del libro BPPALS a generar: `Machos` o `Hembras`. |
| `ZONA` | Zona de estudio mostrada en BPPALS. |
| `AREA` | Área de estudio mostrada en BPPALS. |
| `FLOTA_DESEMBARQUE` | Nombre de la flota buscada en el archivo de desembarque. |
| `EJECUTAR_BPPALS` | Use `SI` para generar BPPALS o `NO` para ejecutar solamente los análisis previos. |
| `DIRECTORIO_TABLAS` | Carpeta de salida para tablas, CSV y archivos Excel. |
| `DIRECTORIO_GRAFICOS` | Carpeta de salida para los gráficos de regresión. |

> `PESQUERIA_COLOCAP` y `PESQUERIA_TALLA_EDAD` corresponden a procesos distintos. No deben modificarse sin confirmar el código de pesquería que corresponde a cada fuente de datos.

### Generar BPPALS para ambos sexos

El programa genera un libro BPPALS por ejecución. Para obtener ambos resultados:

1. Configure `SEXO=Machos` y ejecute el programa.
2. Cambie a `SEXO=Hembras`.
3. Ejecute nuevamente el programa.

Se generarán los archivos correspondientes para cada sexo.

## Ejecución

Abra una terminal en la carpeta del proyecto y ejecute:

```bash
python main.py
```

En sistemas donde Python se ejecuta como `python3`:

```bash
python3 main.py
```

Al finalizar correctamente se mostrará el mensaje:

```text
[ PROCESO FINALIZADO ]
```

## Resultados generados

### Carpeta `Tablas/`

| Archivo | Descripción |
|---|---|
| `{AÑO}_Bp_Analisis_Colocap.xlsx` | Análisis COLOCAP generado desde los archivos DBF. |
| `resultados_regresiones.csv` | Parámetros `a`, `b`, correlación de Pearson, registros utilizados y eliminados para cada método. |
| `bacalao_filtrado.csv` | Datos de palangre filtrados durante el preprocesamiento. |
| `bacalao_filtrado_iqr.csv` | Datos de ambos sexos filtrados mediante IQR. |
| `bacalao_machos_filtrado_iqr.csv` | Datos de machos filtrados mediante IQR. |
| `bacalao_hembras_filtrado_iqr.csv` | Datos de hembras filtrados mediante IQR. |
| `matriz_talla_edad_pesqueria25.xlsx` | Matrices talla–edad depuradas, separadas en hojas de Machos y Hembras. |
| `BPPALSM{AÑO}_generado.xlsx` | Libro BPPALS generado cuando `SEXO=Machos`. |
| `BPPALSH{AÑO}_generado.xlsx` | Libro BPPALS generado cuando `SEXO=Hembras`. |

### Carpeta `Graficos/`

Se generan gráficos de dispersión y curvas longitud–peso para:

- Ambos sexos.
- Machos.
- Hembras.
- Datos originales.
- Datos después de Cook.
- Datos después de MAD.
- Datos después de IQR.

## Criterios aplicados

### Regresión longitud–peso

La relación se ajusta mediante el modelo:

```text
W = a × Lᵇ
```

donde:

- `W` corresponde al peso.
- `L` corresponde a la longitud.
- `a` y `b` son los parámetros estimados.

Los métodos Cook y MAD se generan como respaldo y comparación. El conjunto utilizado por BPPALS corresponde al filtrado mediante IQR.

### Matriz talla–edad

La matriz talla–edad utiliza los siguientes criterios:

- Pesquería definida en `PESQUERIA_TALLA_EDAD`.
- Sexo 1 para machos y 2 para hembras.
- Se excluyen registros con cero anillos.
- Se excluyen longitudes iguales o menores que cero.
- Las tallas se agrupan según `PASO_TALLA_MATRIZ`.
- Las edades iguales o superiores a `EDAD_AGRUPADA` se agrupan como `40+`.
- Se aplica limpieza IQR por sexo y grupo etario.

## Problemas frecuentes

### No existe el archivo

Verifique que el archivo se encuentre en la carpeta del proyecto y que su nombre coincida exactamente con el indicado en `configuracion.txt`.

### `SEXO debe ser "Machos" o "Hembras"`

Revise el valor de `SEXO` en `configuracion.txt`. Debe escribirse exactamente como:

```text
SEXO=Machos
```

o:

```text
SEXO=Hembras
```

### Tallas o proporción sexual quedaron vacías

Revise que los archivos DBF contengan registros compatibles con los filtros definidos:

- Especie configurada.
- Agrupación correspondiente.
- Escala anual.
- Sexos 1 y 2.

### No se encontró desembarque

Verifique que el nombre configurado en `ESPECIE` y `FLOTA_DESEMBARQUE` exista en el archivo de desembarque.

### BPPALS no se genera

Confirme que:

1. `EJECUTAR_BPPALS=SI`.
2. Los procesos de COLOCAP, regresión y matriz talla–edad hayan finalizado correctamente.
3. Existan los archivos generados en la carpeta `Tablas/`.

## Consideraciones

- Los datos de entrada no se incluyen en este repositorio.
- No modifique los nombres de las columnas requeridas sin actualizar el código.
- Antes de procesar un nuevo año, revise todas las rutas, nombres de archivos y parámetros de `configuracion.txt`.
- Se recomienda conservar las carpetas `Tablas/` y `Graficos/` como respaldo de cada ejecución.
- Si los datos son de uso restringido, no los suba al repositorio público.

## Estado del proyecto

Versión inicial integrada de la automatización para análisis de edad y crecimiento.

---