# Madecentro Autofill Agent

Aplicacion de escritorio en Python para administrar datos maestros, aprender
formularios PDF, Excel y Word, y generar copias diligenciadas.

## Requisitos

- Python 3.9 o superior
- Windows, macOS o Linux con interfaz grafica

## Instalacion

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
python main.py
```

`DATABASE_PATH` puede ser absoluto o relativo a la raiz del proyecto. Las
claves de OpenAI y Gemini quedan reservadas para una integracion futura.

## Flujo

1. Importe un CSV/XLSX con columnas `clave`, `valor` y `categoria`, o capture
   los datos manualmente.
2. Registre el formulario vacio y una copia del mismo formulario diligenciada.
   Ambos archivos deben ser del mismo formato: PDF, XLSX o DOCX.
3. El sistema compara los archivos, detecta los campos y crea un mapeo JSON.
4. Revise el mapeo y guardelo. Cada mapeo queda asociado a su plantilla.
5. Seleccione la plantilla y genere el formulario. La salida conserva el
   formato original, recibe una marca de tiempo y no sobrescribe la plantilla.

Cada plantilla nueva queda en una carpeta independiente y genera un manifiesto
versionado. Su huella estructural permite reconocer posteriormente una copia
diligenciada sin depender del nombre del archivo.

## Identidad visual

La interfaz usa la identidad de Madecentro:

- Primario: `#F08419`
- Secundario: `#2D2D2D`
- Fondo: `#FFFFFF`
- Texto: `#333333`

Coloque el logo en `assets/logo`, preferiblemente como PNG transparente. La
aplicación carga automáticamente la primera imagen encontrada.

Coloque la firma en `assets/firma`, preferiblemente como `firma.png`. La firma
se inserta automáticamente en PDF, Excel y Word. El sistema busca una zona con
la palabra `firma` y, si no existe, utiliza una ubicación de respaldo. En Excel
también reconoce la línea de firma, calcula el espacio disponible con los
anchos de columnas y altos de filas, y redimensiona el PNG conservando su
proporción.

La ubicación puede configurarse en el manifiesto. Por ejemplo, para Excel:

```json
{
  "signature": {
    "enabled": true,
    "sheet": "FORMULARIO",
    "cell": "G114",
    "width_columns": 8,
    "height_rows": 4
  }
}
```

Los checkboxes se activan con `si`, `sí`, `true`, `1`, `yes`, `x`, `on`,
`checked` o `marcado`.

## Apariencia de los PDF

Los campos de texto se normalizan al generar la salida: Helvetica, tamaño base
9 y color negro. Los tamaños automaticos y configuraciones `/DA` distintas se
reemplazan por una apariencia consistente. Cuando el contenido no cabe, solo
ese campo reduce proporcionalmente su tamaño, con un minimo legible de 6.

## Formularios Excel

El proyecto puede aprender un formulario Excel comparando una plantilla vacia
con una copia diligenciada. El proceso detecta celdas modificadas y controles
checkbox, y escribe las salidas directamente sobre el paquete XLSX para
conservar formulas, estilos, validaciones, dibujos y configuracion de impresion.

```powershell
.\.venv\Scripts\python.exe tools\excel_workflow.py learn
.\.venv\Scripts\python.exe tools\excel_workflow.py fill
```

`learn` genera `plantillas/mapeos/mapeo_formulario_excel.json` y una copia de
verificacion. `fill` usa las claves maestras asignadas en ese JSON y genera una
copia diligenciada desde `Sarlaft-Somer-Incare vacio.xlsx`. Los campos ambiguos y los
checkboxes quedan sin `master_key` hasta que se defina su fuente de datos.

Las casillas sin una clave maestra conservan el estado de la copia diligenciada
usada como referencia. La fecha principal se actualiza automáticamente con el
día, mes en español y año en que se genera el formulario.

Las selecciones representadas mediante una `X` dentro de una celda, como las de
la sección 8, también se conservan desde el formulario de referencia.

## Aprender un PDF AcroForm

En `Plantillas`, use `Aprender plantilla` para registrar un PDF vacío sin
necesitar una copia diligenciada:

1. Seleccione el PDF AcroForm.
2. El sistema extrae todos sus campos editables.
3. Relacione cada `Campo PDF` con una `Clave maestra`.
4. Pulse `Guardar plantilla`.
5. La aplicación copia el PDF, registra su versión y guarda automáticamente el
   manifiesto JSON.

El mapeo queda asociado a la plantilla y se reutiliza tanto al diligenciar como
al importar futuras copias del mismo formulario.

## Datos maestros

`data/maestros/datos_maestros.csv` es el archivo general de claves y valores
que usa el diligenciamiento de formularios. La aplicacion lo actualiza como
respaldo visible cada vez que se importan, editan, eliminan o aprueban cambios
de datos maestros.

El archivo debe tener estas columnas:

- `clave`
- `valor`
- `categoria`

`data/maestros/clientes_certificados.xlsx` es independiente: solo alimenta el
buscador de clientes para certificados y no reemplaza los datos maestros
generales.

`data/maestros/diccionario_madecentro.json` y
`data/maestros/diccionario_madecentro.csv` son otra capa: ayudan a reconocer
etiquetas, aliases y nombres de campos cuando se aprende un formulario vacio.
No reemplazan ni se mezclan con `datos_maestros.csv`; el JSON es la fuente
principal y el CSV queda como respaldo editable o revisable.

Para actualizar el diccionario, agregue nuevas relaciones entre etiquetas de
formulario y claves maestras usando columnas como `clave`, `master_key`,
`campo_maestro`, `alias`, `etiqueta`, `label` o `texto`. Si ambos archivos
existen, las coincidencias del JSON tienen prioridad y el CSV solo complementa
aliases faltantes. El diccionario mejora la sugerencia automatica de mapeos,
pero los valores oficiales reutilizables siguen viviendo solo en
`datos_maestros.csv` y en la base operativa.

## Certificados de clientes

Desde `Datos maestros`, use `Importar clientes certificados` para cargar un
XLSX con estas columnas:

- `cliente_razon_social`
- `cliente_nit`
- `cliente_anio_vinculacion`

Cuando una plantilla contiene claves `cliente_*`, al diligenciar se abre una
búsqueda por razón social o NIT. Al seleccionar el cliente se completan
automáticamente la razón social, el NIT, el año de vinculación y la fecha de
expedición. El sistema admite tanto `cliente_anio_vinculacion` como
`cliente_ano_vinculacion`.

Los certificados tienen un flujo independiente:

1. Abra la pestaña `Crear certificado`.
2. Busque al cliente por razón social o NIT.
3. Seleccione una fila de resultados.
4. Pulse `Crear certificado`.

La plantilla marcada como certificado no aparece en `Diligenciar formulario`.
El PDF generado se guarda en `data/salidas` con el nombre del cliente y una
marca de tiempo.

## Importar formularios diligenciados

En `Datos maestros`, use `Importar formulario diligenciado`. La aplicacion:

1. Detecta si el archivo es PDF, XLSX o DOCX.
2. Identifica la plantilla mediante su huella y manifiesto.
3. Guarda todas las respuestas extraídas en el historial de formularios.
4. Muestra el valor maestro actual y el valor encontrado.
5. Permite desmarcar cambios antes de confirmar.
6. Actualiza los datos maestros en una transacción independiente.

PDF usa los nombres AcroForm. Excel usa las hojas y coordenadas de su propio
manifiesto. Word usa los controles o celdas aprendidas. Los mapeos globales
antiguos continúan disponibles solamente por compatibilidad.

Si un archivo coincide con varios manifiestos, la importación se detiene para
evitar actualizar información con un mapeo ambiguo.

En el formulario Excel de Madecentro, los miembros de junta directiva se
importan como `junta_1_*` hasta `junta_9_*`, con nombre, tipo de identificación
y número de identificación. Se guardan en la categoría `junta_directiva`; una
posición sin nombre se considera incompleta y no se importa.

## Limites actuales

- Los PDF deben contener campos AcroForm.
- Los Word deben usar controles de contenido o tablas donde el archivo vacio y
  el diligenciado conserven la misma estructura.
- Los Excel deben conservar hojas y estructura entre la copia vacia y la
  diligenciada.
- No hay OCR ni firma digital criptográfica. La firma insertada es una imagen.
- La sugerencia de mapeo usa alias y similitud textual local; no llama a una IA.

## Verificacion

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m compileall -q main.py app tests
.\.venv\Scripts\python.exe -m pip check
```

## Estructura

```text
app/core          Configuracion y rutas
app/database      Persistencia SQLite
app/models        Modelos y validacion
app/services      Campos, mapeos y llenado PDF
app/ui            Interfaz PySide6
tests             Pruebas automatizadas
data              Entrada, salidas y base local
plantillas        PDFs y mapeos JSON
```
