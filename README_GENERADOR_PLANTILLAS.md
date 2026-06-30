# Generador de Plantillas

Modulo para tomar documentos diligenciados de clientes o proveedores, ubicar
datos propios de Madecentro definidos en un diccionario y generar copias
limpias en PDF, DOCX, XLSX o XLSM.

## Instalacion

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Estructura

```text
entrada_documentos/       Documentos diligenciados originales
plantillas_generadas/     Copias limpias generadas y reporte
diccionario_madecentro.xlsx
```

El diccionario debe incluir estas columnas:

- `categoria`
- `valor`
- `reemplazo`
- `confianza`

Tambien se acepta JSON con una lista de registros equivalentes.

## Uso

Reemplazar datos por marcadores:

```powershell
python scripts\generar_plantillas.py --input entrada_documentos --dictionary diccionario_madecentro.xlsx --output plantillas_generadas --mode markers
```

Borrar datos y dejar el espacio vacio:

```powershell
python scripts\generar_plantillas.py --input entrada_documentos --dictionary diccionario_madecentro.xlsx --output plantillas_generadas --mode blank
```

Si `reemplazo` tiene valor y el modo es `markers`, se usa ese valor. Si esta
vacio, el sistema usa marcadores por categoria, por ejemplo `<<NIT>>`,
`<<RAZON_SOCIAL>>` o `<<REPRESENTANTE_LEGAL>>`.

## Salida

Cada archivo original se conserva intacto. La copia procesada queda en
`plantillas_generadas` manteniendo la ruta relativa. El reporte se guarda como:

```text
plantillas_generadas/reporte_generacion_plantillas.xlsx
```

El reporte incluye archivo, tipo, estado, cantidad de reemplazos, valores
reemplazados y errores. Si un archivo falla, el proceso continua con los demas.

## Limitaciones

Los PDF escaneados no se procesan todavia porque no hay OCR en este modulo. En
PDF con texto digital se intenta cubrir cada coincidencia con un rectangulo
blanco e insertar el marcador sobre el area encontrada. Si la insercion precisa
no es posible, el dato queda cubierto.
