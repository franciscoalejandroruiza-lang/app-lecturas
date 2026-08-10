# Generador de Formatos de Lectura ISSSTE — Delegación Estatal Chihuahua

App de Streamlit para generar automáticamente los formatos de lectura mensual
(en PDF, listos para firma) a partir del Excel de consumo (`ISSSTE_DETALLE_CONSUMO`)
y el catálogo maestro de equipos.

## Cómo funciona

1. Los 25 formatos Word originales (uno por ubicación) se usan como **plantilla
   viva**: la app solo edita las celdas de "Lectura inicial", "Lectura final" y
   "Volumen mensual" de cada equipo (por número de serie), dejando el resto del
   documento exactamente igual (logos, CLUE, folios, domicilio, firmas, etc.).
2. El catálogo maestro (`data/catalog_final.json`) fue extraído automáticamente
   de esos 25 Word y contiene los datos fijos de 180 equipos (ver
   `data/CATALOGO_MAESTRO_ISSSTE.xlsx` para revisarlo en Excel, con los 7 casos
   sin match resaltados en amarillo y explicados en la hoja "NOTAS").
3. Cada mes, subes el Excel de consumo actualizado. La app detecta los bloques
   de columnas mensuales (lee el CONTENIDO de los encabezados, no la posición,
   porque el archivo real tiene inconsistencias: bloques de 4 o 5 columnas,
   meses repetidos/mal copiados, columnas sueltas de "FECHA Y HORA").
4. Eliges el modo:
   - **Lectura inicial**: genera el PDF con la lectura inicial ya llena
     (= lectura final del bloque elegido) y "Lectura final" en blanco, listo
     para imprimir y mandar a campo a que la operadora capture la lectura final.
   - **Lectura final**: usa el bloque ANTERIOR como inicial y el bloque ACTUAL
     como final, calcula el volumen mensual, y genera el PDF completo listo
     para firma y sello.
5. Descargas un .zip con un PDF por ubicación (ej. `FORMATOS DE LECTURA DE
   ISSSTE ALDAMA.pdf`).

## Instalación

Requiere Python 3.9+ y LibreOffice instalado en el sistema (para convertir a
PDF). En Windows/Mac se instala LibreOffice normal; en Linux:
`sudo apt install libreoffice`.

```bash
cd app
pip install -r requirements.txt
streamlit run app.py
```

## Estructura de carpetas

```
app/
  app.py              <- interfaz Streamlit
  excel_reader.py      <- lee y detecta bloques mensuales del Excel "OK"
  fill_formats.py       <- llena las plantillas Word y convierte a PDF
  data/
    catalog_final.json          <- catálogo maestro (usado por la app)
    CATALOGO_MAESTRO_ISSSTE.xlsx <- mismo catálogo, para revisar/editar en Excel
  templates/
    FORMATOS DE LECTURA DE ISSSTE *.docx  <- los 25 formatos originales (plantilla viva)
```

## Pendientes / decisiones que quedaron abiertas (revisar contigo)

- **7 equipos sin match** entre Word y Excel de consumo (ver hoja NOTAS del
  catálogo): 2 parecen typos de un dígito en el número de serie, y 5 son altas
  de CAF que aún no están en el Excel mensual. Mientras no se agreguen al
  Excel, esos 5 equipos no van a tener lecturas para generar.
- **CLUE y Folio Unidad Médica en blanco**: se dejaron tal cual vienen en cerca
  de 30% de los formatos (así lo pediste), no se inventan valores.
- **Bloque de Noviembre partido en dos** (columnas P:S y T sueltas) en el
  primer año del Excel: es un error de captura real, ya documentado en la
  hoja `Hoja2` del Excel original. La app lo muestra como dos bloques
  separados en vez de adivinar cuál es el correcto — al generar ese mes en
  particular, revisa manualmente cuál bloque (o si hay que sumar ambos) es
  el correcto antes de imprimir para firma.
- Si se dan de alta o de baja equipos, hay que actualizar
  `data/catalog_final.json` (o pedirme que vuelva a extraer el catálogo si
  mandas formatos Word nuevos/actualizados).
- Si quieres, en una siguiente vuelta puedo agregar a la app: edición del
  catálogo desde la misma interfaz (sin tocar el JSON a mano), o integrarla
  como pestaña de tu dashboard de KPIs existente.

