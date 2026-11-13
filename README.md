# Procesamiento de edad y crecimiento — IFOP

Herramienta para **merluza del sur** y **bacalao de profundidad**. Ejecuta
COLOCAP, regresiones longitud–peso, claves talla–edad y composición del
desembarque por edad. El usuario trabaja con un único `configuracion.txt`.

## Descargar ejecutable para Windows

[Descargar Consolidado_IFOP.zip](https://github.com/NMora27/ifop/releases/download/v2.0.0/Consolidado_IFOP.zip)

Incluye el ejecutable, `configuracion.txt` y `Ejecutar_IFOP.bat`. Descomprima el ZIP, ajuste la configuración y coloque los archivos de entrada en la misma carpeta antes de ejecutar `Ejecutar_IFOP.bat`.

## Uso desde Python

Instalar las dependencias:

```bash
python -m pip install numpy pandas scipy openpyxl matplotlib dbfread
```

Mantener los once módulos Python, `configuracion.txt` y las cuatro plantillas
de `plantillas_composicion/` en la distribución del programa. Colocar los datos
en las rutas del TXT; los inputs no se incluyen en GitHub.

En `[GENERAL]`, seleccionar una de estas opciones:

```ini
ESPECIE=MERLUZA DEL SUR
```

```ini
ESPECIE=BACALAO DE PROFUNDIDAD
```

Después ejecutar desde la carpeta del programa:

```bash
python main.py --validar
python main.py
```

`--validar` comprueba configuración, enlaces entre etapas y existencia de
archivos; no calcula productos ni valida todo el contenido de las fuentes.
La ejecución normal valida los datos dentro de cada módulo.

## Qué se configura

| Bloque del TXT | Contenido editable |
|---|---|
| `[GENERAL]` | Especie, año, carpeta de salida y etapas/gráficos SI o NO. |
| `[MERLUZA ARCHIVOS]` | Excel de edad, peso–longitud y desembarque de merluza; hojas de desembarque. |
| `[MERLUZA ARRASTRE]`, `[MERLUZA PALANGRE]`, `[MERLUZA ARTESANAL]` | DBF y celdas de desembarque norte/sur. |
| `[BACALAO ARCHIVOS]` | Excel de edad, peso–longitud y desembarque de bacalao; hoja de desembarque. |
| `[BACALAO PALANGRE]` | DBF y `CELDAS_DESEMBARQUE`, sin separación zonal. |

Las rutas relativas parten de la carpeta del TXT. `{ANIO}` se sustituye por
el año seleccionado. Solo se exigen las fuentes de la especie activa; los
bloques y nombres de opciones del archivo completo deben ser válidos.

Las reglas fijas están en `configuracion.py`: códigos, columnas, dimensiones,
MAD 3,5, mezclas de merluza, filtros anuales y nombres de productos. No se
repiten en la configuración del usuario. Véase [MAIN.md](MAIN.md).

## Diferencias entre especies

| Etapa | Merluza del sur | Bacalao de profundidad |
|---|---|---|
| COLOCAP | Tres artes, total y dos zonas: nueve libros. | Palangre total: un libro. |
| Regresión | Total y zonas norte/sur por arte. | Únicamente Total; sin COD_PESQUERIA, zona ni latitud. |
| Edad | GEC en la fuente histórica 2025; edades 1–24+. | EDAD en el formato estándar o Nº_DE_ANILLOS en el histórico 2025; clave 1–40+. |
| Claves | Preliminares y definitivas; mezclas por sexo cuando hay menos de 500 retenidos. | Una clave total por sexo dentro del mismo libro; sin mezcla por pesquería o zona. |
| Composición | Seis dominios y dos sexos: doce libros. | Total de Palangre y dos sexos: dos libros; composición 4–40+. |

En ambas especies, los sexos se identifican como 1 = machos y 2 = hembras.
La regresión utiliza longitud y peso del ejemplar, y la composición consulta
los ajustes MAD de cada sexo. Las tablas de informe muestran el ajuste Original.

## Archivos del programa

```text
main.py
configuracion.py
configuracion.txt
lectura.py
analisis_colocap.py
regresion.py
clave_pesotalla.py
matriz_de_composicion.py
metodos_matematicos.py
diseno_composicion.py
claves_edad_merluza.py
claves_edad_bacalao.py
plantillas_composicion/
    merluza_machos.xlsx
    merluza_hembras.xlsx
    bacalao_machos.xlsx
    bacalao_hembras.xlsx
.github/workflows/compilar-windows.yml
```

`inputs_MS/` y `inputs_BP/` son fuentes locales. Los JSON antiguos de
`perfiles/`, los ZIP y `analisis_historico_merluza/` no participan en la
operación de main. Las probabilidades de edad activas están en los dos módulos
`claves_edad_*.py`; no hace falta reconstruir una base histórica.

## Resultados

Con `DIRECTORIO_TABLAS=Tablas`, las carpetas son:

```text
Tablas Merluza del sur/
Tablas Bacalao de profundidad/
```

Contienen `COLOCAP/`, `Regresiones/`, `Clave_talla_edad/` y
`Matriz_composicion/`. El programa reemplaza los productos de las etapas
solicitadas. Las etapas omitidas pueden dejar archivos de ejecuciones anteriores.

Las celdas completadas por referencia de edad se muestran con un **1 rojo**
en la clave del libro de composición: indican una asignación estimada, no un
pez observado. `PENDIENTE_CLAVE` identifica captura expandida aún sin edad.
No se renormalizan los porcentajes para ocultarla.

## Ejecutable Windows desde GitHub

El workflow `.github/workflows/compilar-windows.yml` se ejecuta manualmente en
**Actions → Compilar IFOP para Windows → Run workflow**. Seleccionar la rama
que contiene la versión que se quiere compilar, normalmente `main`.

Usa Windows y Python 3.12; PyInstaller genera `Consolidado_IFOP.exe` e incluye
las cuatro plantillas con `--add-data`. El paquete descargable contiene:

```text
Consolidado_IFOP.exe
configuracion.txt
Ejecutar_IFOP.bat
```

Los inputs se agregan localmente según las rutas del TXT. Las plantillas y
probabilidades quedan dentro del ejecutable. No se necesita instalar Python
en el equipo que usa el paquete.

El workflow comprueba los archivos necesarios y ejecuta el `.exe --help` antes
de publicar el artefacto. Esa comprobación valida el arranque sin inputs; la
prueba con datos propios se realiza después de descargar el paquete. La
ejecución de Python con datos de 2025 y la compilación Windows son
verificaciones distintas.

Para usar el paquete, extraerlo, revisar el TXT y abrir `Ejecutar_IFOP.bat`.
También se puede comprobar desde PowerShell:

```powershell
.\Consolidado_IFOP.exe --validar
```

La suite de trabajo está en `tests/`. Su prueba de equivalencia utiliza los
JSON de referencia de `perfiles/`; las pruebas de rutas usan también los inputs
locales. Esos recursos son necesarios para ejecutar esa suite, pero no para
distribuir ni iniciar el programa sin datos.
