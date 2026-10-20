# Automatización Procesamiento de Datos de Edad y Crecimiento — IFOP
# USO EXCLUSIVO PARA LOS DATOS DEL 2024

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
python -m pip install numpy pandas scipy openpyxl matplotlib dbfread xlsxwriter
```

## Estructura del proyecto

Los scripts y los archivos de entrada deben mantenerse en la misma carpeta:

```text
ifop/
├── main.py
├── configuracion.txt
├── analisis_colocap.py
├── regresion.py
├── clave_pesotalla.py
├── bppals_20XX.py
├── lectura.py
├── metodos_matematicos.py
├── E_Resumen.DBF
├── ET_Estructura_de_Tallas.DBF
├── ET_Proporción_Sexual.DBF
├── Biologico_Palangre2024.xlsx
├── Unido_Bacalao_de_Pofundidad_2024.xlsx
└── Desembarque 2024 industr para inf final EDAD.xlsx
```

Las carpetas `Tablas/` y `Graficos/` se crean automáticamente durante la ejecución.

## Archivos de entrada

| Archivo | Uso |
|---|---|
| `E_Resumen.DBF` | Obtiene el número de ejemplares para el análisis COLOCAP. |
| `ET_Estructura_de_Tallas.DBF` | Obtiene proporciones de talla por sexo. |
| `ET_Proporción_Sexual.DBF` | Obtiene la proporción de machos y hembras. |
| `Biologico_Palangre2024.xlsx` | Datos para la regresión longitud–peso. |
| `Unido_Bacalao_de_Pofundidad_2024.xlsx` | Datos para construir la matriz talla–edad. |
| `Desembarque {AÑO} industr para inf final EDAD.xlsx` | Obtiene el desembarque utilizado en BPPALS. |

Los nombres de los archivos deben coincidir con los definidos en `configuracion.txt`.

### Columnas requeridas

La hoja `Sp37-Pal-24` de `Biologico_Palangre2024.xlsx` debe contener, al menos, las siguientes columnas:

```text
LATITUD
COD_ESPECIE
SEXO_ESPECIMEN
LONGITUD_ESPECIMEN
PESO_ESPECIMEN
```

La hoja `Clave Edad-Talla` de `Unido_Bacalao_de_Pofundidad_2024.xlsx` debe contener:

```text
Longitud
N°Anillos
Pesqueria
Sexo
```

Los archivos DBF deben conservar las columnas utilizadas por el proceso, tales como `AGRUPACION`, `ESCALA`, `ESPECIE`, `SEXO`, `TALLA`, `PROPORCION` y `EJEMP_LONG`, según corresponda.

## Configuración

Antes de ejecutar el programa, abra `configuracion.txt` y modifique los parámetros de acuerdo con el año y análisis requerido.

Ejemplo:

```text
ANIO=2024
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
La terminal mostrará `Archivo generado: ...` por cada salida creada.

## Resultados generados

### Carpeta `Tablas/`

Los nombres del CSV y del Excel de regresiones se configuran por separado con
`SALIDA_REGRESIONES` y `SALIDA_REGRESIONES_EXCEL` en `configuracion.txt`.
Ambos exportan `a`, `b` y `r²` con 7 decimales como máximo. El Excel muestra
`a` sin notación científica. `r_PEARSON` en el CSV corresponde a `r`, no a `r²`.

| Archivo | Descripción |
|---|---|
| `{AÑO}_Bp_Analisis_Colocap.xlsx` | Análisis COLOCAP generado desde los archivos DBF. |
| `resultados_regresiones.csv` | Parámetros `a`, `b`, `r_PEARSON`, `r²`, registros utilizados y eliminados para cada método. |
| `resultados_regresiones.xlsx` | Tabla con `a`, `b`, `r²` y `N` para machos, hembras y ambos, e intervalos de confianza del 95 % de `a` y `b`. Hojas Original, Cook, MAD e IQR. |
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

La matriz usa la hoja `Clave Edad-Talla`, filtra `Pesqueria=Industrial` y
separa machos (`Sexo=1`) y hembras (`Sexo=2`).

La matriz talla–edad utiliza los siguientes criterios:

- Sexo 1 para machos y 2 para hembras.
- Se excluyen edades iguales o menores que cero, o datos de edad/talla no válidos.
- Se excluyen longitudes iguales o menores que cero.
- Las tallas se agrupan según `PASO_TALLA_MATRIZ`.
- Las edades iguales o superiores a `EDAD_AGRUPADA` se agrupan como `40+`.
- Se aplica limpieza IQR por sexo y grupo etario.

### Consulta de edad por talla

Las funciones `cargar_clave()` y `asignar_edad()` están en
`metodos_matematicos.py`. Consultan `claves_edad.py`, una clave estática
con frecuencias y probabilidades históricas de `UNIDO_2007-2024.xlsx`,
hoja `2007-2024`, pesquería 25 / Industrial y sexos conocidos 1 y 2.
La consulta no requiere los Excel ni ejecutar `main.py` previamente.
Esta clave histórica conserva edades individuales de 1 a 100 años,
sin agrupación `40+`, suavizado ni limpieza IQR adicional.

```python
from metodos_matematicos import cargar_clave, asignar_edad

resultado = asignar_edad(123, sexo="hembra")
print(resultado["edad_asignada"], resultado["opciones"])

clave = cargar_clave()  # suma frecuencias de ambos sexos
resultado_ambos = asignar_edad(123, clave=clave)
```

No combinar `clave` y `sexo`. Se entregan hasta tres opciones con
probabilidades calculadas sobre el total de la clase; los empates se
informan y se resuelven por edad ascendente. Las reglas fijas son
`(0,25) → 2`, `[25,35) → 3`, `[35,40] → 4` y `>245 → 37` años.
Estas reglas tienen `origen="regla_fija"`, probabilidad 1 por decisión
explícita y frecuencias `None`. Fuera de las reglas, una clase sin
observaciones produce `ValueError`.

### BPPALS

El libro conserva las ocho hojas y el diseño original de 2024. En la hoja
`Clave`, la columna `VERIFICACIÓN` informa `CL OK` cuando una clase con captura
tiene datos de clave, o `FALTA CL` cuando no los tiene.

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
