# Sincronizar gráficos del PPTX con el Excel fuente

`Laminas 11112026.pptx` no trae los datos de sus gráficos incrustados: cada
`ppt/charts/chartN.xml` guarda una **caché** de los valores que se dibujan
(`<c:numCache>`) más un vínculo a un Excel externo
(`Graficos 11112026.xlsx`, hoy enlazado a una ruta de red que PowerPoint sólo
resuelve si abres el archivo conectado a esa red). Para actualizar un gráfico
hoy hay que abrir el Excel, actualizar los datos, abrir el PPTX conectado a
esa red y refrescar los vínculos a mano.

`pptx_chart_sync.py` hace lo mismo sin PowerPoint ni esa red: lee
`Graficos 11112026.xlsx`, ubica para cada serie el rango de celdas que la
alimenta (la misma fórmula `<c:f>` que ya trae el gráfico, p. ej.
`Combustibles!$B$3:$B$48`), toma los valores actuales de esas celdas
—extendiendo el rango si hay filas o columnas nuevas con datos— y reescribe
**solo esa caché** dentro del XML del gráfico. Todo lo demás del .pptx
(diseño, colores, fuentes, textos, logo, capas) queda intacto byte a byte.

## Uso

```bash
python3 tools/pptx_chart_sync.py \
  --xlsx "Graficos 11112026.xlsx" \
  --pptx "Laminas 11112026.pptx" \
  --out  "Laminas 11112026.pptx"
```

Esto actualiza **dos cosas** en cada corrida:

1. Los gráficos del `--pptx` (en `--out`), leyendo los datos del `--xlsx`.
2. Los gráficos nativos del propio `--xlsx` (los que ves al abrirlo en Excel),
   **en el mismo archivo**, para que ambos documentos queden consistentes.
   Usa `--skip-excel-chart` si por algún motivo sólo quieres tocar el pptx.

El Excel trae sus propios `chart1.xml`…`chart8.xml` (son los gráficos que se
copiaron hacia el pptx en su momento) con la misma estructura y las mismas
fórmulas de rango que los del pptx, así que se sincronizan con el mismo
mecanismo: seguir la fórmula de la celda, releer el rango (extendiéndolo si
hay datos nuevos) y reescribir la caché.

- Sin `--out`, escribe `<pptx>.synced.pptx` junto al original (no pisa nada).
  El `--xlsx` sí se actualiza en el mismo archivo (a menos que uses
  `--skip-excel-chart`), porque normalmente quieres que quede como la fuente
  de verdad al día.
- `--chart chart1 --chart chart7` limita la sincronización a gráficos puntuales
  (los nombres de archivo, no el número de lámina — ver mapa abajo; aplica a
  ambos documentos).
- `--quiet` silencia el detalle por serie.

Mapa de `chartN.xml` → lámina y gráfico:

| Archivo | Lámina | Gráfico |
|---|---|---|
| chart1 | 1 | Gasolinas |
| chart2 | 2 | Balance efectivo acumulado (% PIB) |
| chart3 | 2 | Crecimiento anual del gasto |
| chart4 | 3 | Ingresos y Gastos del Gobierno Central |
| chart5 | 3 | Balance Efectivo del Gobierno Central |
| chart6 | 4 | Crecimiento del PIB en 2027 (por institución) |
| chart7 | 5 | Creación de empleo |
| chart8 | 7 | IPC efectivo y futuros |

(La lámina 6 es una tabla, no un gráfico; no la toca este script.)

## Validado

- Sincronizar el Excel actual contra el PPTX actual (caso "sin cambios")
  reproduce exactamente los mismos gráficos — ver `scripts/office/validate.py`
  del skill de pptx, que da `All validations PASSED!`.
- Agregar una fila nueva en el Excel (p. ej. una semana más de precios de
  combustible, un mes más de gasto fiscal) y volver a correr el script
  extiende el gráfico correctamente: nueva fecha en el eje, nuevo punto/barra
  con el valor correcto, sin romper nada más.

## Aviso sobre fórmulas arrastradas (y por qué `update_from_apis.py` no usa openpyxl para escribir)

Varias columnas del Excel —no sólo fechas— son fórmulas, incluyendo
**fórmulas compartidas** de Excel (`<f t="shared" si="N"/>`, el mismo
mecanismo detrás de "arrastrar" una fórmula por muchas filas). `Balance 2`
en particular tiene columnas de apoyo (H, I, J, K, Y, Z:AK) con decenas de
estas fórmulas compartidas que ningún gráfico usa directamente.

`pptx_chart_sync.py` sólo *lee* el Excel (nunca lo guarda), así que no tiene
problema con esto. Pero se comprobó en pruebas que si otro script abre el
libro completo con `openpyxl` y lo vuelve a **guardar** — aunque sólo
modifique 2 o 3 celdas — openpyxl no reevalúa las fórmulas compartidas que
no tocó, y su valor cacheado se pierde **en cualquier hoja del archivo**, se
haya tocado o no (se confirmó: un guardado así dejó en blanco la columna de
fechas de "Combustibles", una hoja completamente ajena al cambio). Por eso
`xlsx_row_writer.py` nunca abre el libro con `openpyxl` en modo escritura:
inserta la fila nueva directamente como XML dentro de `xl/worksheets/sheetN.xml`
(mismo espíritu que `pptx_chart_sync.py` con los gráficos), así que el resto
del archivo queda intacto byte a byte. (`xlsx_utils.flatten_formulas_to_values`
sigue disponible pero ya no se usa en este flujo — aplanar todo el libro
también borraría fórmulas que sí queremos conservar.)

## Traer los datos directamente desde las APIs (BCCh + INE)

`update_from_apis.py` busca datos nuevos en el Banco Central y el INE y, si
los hay, agrega una fila al Excel (conservando sus fórmulas) y vuelve a
correr `pptx_chart_sync.py`.

```bash
python3 tools/update_from_apis.py \
  --xlsx "Graficos 11112026.xlsx" \
  --pptx "Laminas 11112026.pptx" \
  --apply       # sin --apply, sólo muestra qué haría (dry run)
```

Fuentes confirmadas (los valores que devuelve la API calzan exacto, al
decimal, con lo que ya había en el Excel — ver `tools/bcch_client.py` y
`tools/ine_client.py`):

| Fuente | Serie / dataflow | Hoja → columnas | Alimenta |
|---|---|---|---|
| BCCh | `F051.BT02.FLU.C.D.CLP.M` (ingresos) y `F051.BT03.FLU.C.D.CLP.M` (gastos) | `Balance 2` → A,B,C,D,M | chart4, chart5 |
| BCCh | `F089.PIB.V12.2027.M` (EEE, expectativa PIB 2027) | `PIB 2027` → A,B,C,F | chart6 (columna EEE) |
| INE (SDMX, sin credenciales) | `DF_OCU_SEXO` y `DF_OCUINF_SEXO`, total nacional (`AREA_REF=_T`, `SEXO=AS`) | `Empleo` → A,B,D,F,U,V,W,X | chart7 |

Autenticación: BCCh usa `BCCH_USER`/`BCCH_PASS` (ya configuradas en este
entorno); el modo moderno de la API (librería oficial `bcchapi`) usa un
"API Key Token" en vez de usuario/clave, pero el endpoint REST con
usuario/clave —el que ya teníamos— sigue funcionando. El INE no pide
credenciales.

Lo que **no** se auto-completa, y por qué:

- **`Balance 2`, columnas E/G/N/O/P (PIB nominal y %PIB)**: la columna E no
  es un dato plano — en el archivo ya es, ella misma, una fórmula que
  proyecta el trimestre (p. ej. `+E163*(1+3%)`, un supuesto de crecimiento
  t/t) hasta que alguien la reemplaza a mano por el PIB nominal real. No hay
  una fuente pública que calce con ese supuesto, así que no se inventa: las
  filas nuevas llegan con A-D y M, pero sin E/G/N/O/P, y chart4/chart5 no se
  extienden hasta completarlas (a mano, o replicando esa fórmula si
  confirman que es el criterio correcto).
- **`PIB 2027`, columnas Bloomberg y FMI**: sin fuente pública gratuita
  (Bloomberg) o sin conectar todavía (FMI/WEO). Quedan en blanco en la fila
  nueva, igual que ya aparecen en blanco en meses existentes sin ese dato.
- **Gasolinas, "Balance" (fan chart mensual) y los forwards de inflación**:
  fuera de alcance — ver el resumen que se le dio al usuario en el chat:
  ENAP no publica el $/lt final (incluye impuesto específico y MEPCO, no
  reproducible desde la tabla de paridad), y los forwards son de Bloomberg.
- **Una fila nueva como barra en `chart6`** (p. ej. un vintage WEO nuevo en
  abril/julio/octubre): el script sólo agrega la fila a la hoja `PIB 2027`;
  decidir si esa fecha debe aparecer como una barra nueva en el gráfico
  es un cambio de estructura del chart (agregar un `<c:ser>`), no un dato,
  así que queda para hacerlo a mano cuando corresponda.

Probado con un "retroceso" controlado: se borró la última fila real de
`Balance 2`, `PIB 2027` y `Empleo` en una copia del archivo (nunca el
original), se corrió `update_from_apis.py --apply`, y reconstruyó exactamente
los mismos valores que tenía el archivo real — validado con
`scripts/office/validate.py` del skill de pptx y comparando el .pptx
renderizado.
