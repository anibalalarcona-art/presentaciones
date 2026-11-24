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
- `--no-grow-sheet "Nombre hoja"` (repetible): para esa hoja, nunca crece un
  rango más allá de lo ya declarado (sí lo sigue recortando si el final
  tiene celdas vacías). Pensado para hojas de uso mixto donde la celda de
  al lado de un rango puede pertenecer a otro bloque de datos sin relación
  con el gráfico — ver el caso real más abajo.
- `--quiet` silencia el detalle por serie.

### Cuando "crecer el rango" no es lo que quieres: `--no-grow-sheet`

Por defecto, si la celda siguiente a un rango ya tiene un dato, se asume
que es el próximo punto real de la serie y se agrega al gráfico. Eso es
exactamente lo que quieres en una hoja de serie de tiempo que se va
extendiendo (una fila nueva cada semana/mes). Pero en una hoja de uso
mixto —un "Resumen" con una tabla chica para el gráfico y, un poco más
allá en la misma fila, celdas de apoyo para otro cálculo— esa celda
vecina puede no tener nada que ver.

Pasó con `Gráficos_TC.xlsx`, hoja `Resumen`: el gráfico de "Proyecciones
de Tipo de Cambio" usa `Resumen!$M$3:$Q$3` (spot + 4 trimestres), y la
celda siguiente (`$R$3`) sí tiene un dato real — "2027", la proyección
anual — pero el autor del gráfico decidió no incluirla (no calza con el
patrón trimestral del eje). Sincronizar sin avisar la habría vuelto a
agregar. Por eso, para ese archivo:

```bash
python3 tools/pptx_chart_sync.py \
  --xlsx "Gráficos_TC.xlsx" --pptx "Láminas_TC.pptx" --out "Láminas_TC.pptx" \
  --no-grow-sheet Resumen
```

Si alguna vez SÍ quieres que el "2027" entre al gráfico, corre sin ese
flag una vez (o agrega la columna a mano en el rango del gráfico, en
Excel).

### Ejes de categoría de dos niveles (`<c:multiLvlStrCache>`)

Un gráfico cuyo eje de categorías usa un rango de **dos filas** (p.ej.
`Hoja!$C$25:$AW$26`, con el año en una fila dispersa y la fecha en la
otra, como el histórico de revisión del catastro CBC) guarda su caché
como `<c:multiLvlStrCache>`, no como `<c:strCache>`. Antes de que esto se
manejara explícitamente (`process_multilvl_ref_block`), el sync
reconocía el rango de dos filas, ensanchaba el `<c:f>` correctamente,
pero al no encontrar un `<c:strCache>`/`<c:numCache>` para reemplazar
**insertaba uno nuevo de sobra** (leído mal, como si fuera una celda
suelta) dejando el `<c:multiLvlStrCache>` original sin actualizar — dos
cachés distintas conviviendo en el mismo `<c:cat>`, y la que de verdad
dibuja el eje con el `<c:ptCount>` viejo. PowerPoint abre ese archivo
pidiendo reparar, y tras "reparar" puede mostrar la lámina casi vacía o
con contenido irreconocible en vez del gráfico. Si ves `(multinivel)` en
la salida del sync, es esta rama la que corrió; si en cambio ves
`! rango multinivel con forma inesperada`, es un rango de dos niveles
con una forma distinta a "dos filas, muchas columnas" — revisa ese
gráfico a mano, no lo toca.

## Tipo de Cambio Real (BIS): `update_tc_from_apis.py`

Mismo patrón que `update_from_apis.py`, pero para `Gráficos_TC.xlsx` /
`Láminas_TC.pptx`, con el BIS (Bank for International Settlements) como
fuente. El enlace que ya traía la hoja "TCR" (`data.bis.org/topics/EER`)
apunta a la serie `WS_EER` (REER real, canasta amplia, mensual) de los 15
países/áreas del gráfico — confirmado: los valores de la API calzan exacto
con los que ya había en el Excel (`tools/bis_client.py`).

```bash
python3 tools/update_tc_from_apis.py \
  --xlsx "Gráficos_TC.xlsx" --pptx "Láminas_TC.pptx" --apply
```

Qué hace cuando el BIS publica un mes nuevo (ver el docstring del script
para el detalle completo):

1. Llena la fila ya pre-armada de "TCR" (trae la fecha puesta, esperando
   los 15 valores) y extiende los "Prom. 5a"/"Prom. 10a" de esa hoja.
2. En "TCR (ajustado)": esa fila tenía una fórmula que PROYECTABA el valor
   con el forward de tipo de cambio (p. ej. `+B145*TCN!B208/TCN!B207`,
   mientras no existía el dato real). La reemplaza por la misma fórmula
   que ya usan los meses reales (`(100/TCR!B146)*100` o `+TCR!G146` según
   la columna), tomada de la fila anterior y desplazada una fila — no
   inventa el patrón, copia el que el propio archivo ya usa para esa
   columna.
3. Recalcula en cascada las filas de proyección siguientes (que ahora
   parten de un valor real), los promedios 5a/10a, y "TCR estimado <mes>"
   (que pasa a apuntar al próximo mes, el que todavía no es real).
4. Reordena la tabla de ranking (15 países de mayor a menor TCR, la que
   alimenta los gráficos de las láminas 8 y 9) — en el archivo original
   esa tabla está pegada a mano, no por fórmula; el script la recalcula
   entera cada vez a partir de los promedios del paso 3.
5. Sincroniza pptx + gráficos nativos del Excel, con `--no-grow-sheet
   Resumen` (ver arriba) para no reintroducir la columna "2027".

**Validado con un "retroceso" controlado** (el mismo método que para el
otro archivo): en una copia descartable, se borró el último mes real
(sep-26) simulando que el BIS todavía no lo publicaba, y se revirtieron a
mano las fórmulas de "TCR (ajustado)" a cómo se verían antes de esa
publicación. Se corrió el script con los datos reales del BIS para ese
mes, y reconstruyó — al dato, incluida la tabla de ranking reordenada —
exactamente los mismos valores que ya tenía el archivo real.

Hoy (nov-2026) el BIS todavía no publica octubre, así que correrlo no
cambia nada — pero queda listo para cuando lo publique.

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

### PIB nominal (columna E de "Balance 2"): se auto-corrige solo cuando el Banco Central publica el dato real

La columna E sí tiene fuente pública real: `F032.PIB.FLU.N.CLP.EP18.Z.Z.0.T`
(PIB nominal trimestral, Cuentas Nacionales del Banco Central — viene en
miles de millones de pesos, la hoja usa millones, de ahí el ×1000). El
archivo ya trae esa serie para todos los trimestres publicados. El
problema es sólo el trimestre MÁS RECIENTE, mientras el Banco Central
todavía no lo publica (tiene ~2 meses de rezago): para no dejar el gráfico
cortado, esa celda trae un supuesto propio puesto a mano (p. ej.
`+E163*(1+3%)`, asumiendo X% de crecimiento t/t) que luego alguien
reemplaza por el dato real cuando sale.

`update_from_apis.py` automatiza justo ese reemplazo: en cada corrida
revisa si la celda E del trimestre más reciente sigue siendo un supuesto
(fórmula con un `%` en vez de un número) y, si el Banco Central ya publicó
ese trimestre, la reemplaza por el valor real y recalcula en cascada
G/N/O/P de los 3 meses de ese trimestre — usando la MISMA fórmula que ya
tenía cada celda (`SUM(B...)/SUM(G...)*100`), sólo actualiza el resultado
cacheado. No toca ninguna fórmula ni ningún otro trimestre.

Probado con una Cuentas Nacionales ya publicada (se simuló que el
trimestre abr-jun 2026 todavía estaba "supuesto" y se dejó que el script lo
resolviera con el dato real del Banco Central): reconstruyó E, G, N, O y P
prácticamente idénticos —hasta el décimo decimal— a los que ya tenía el
archivo real para ese trimestre.

Hoy (nov-2026) el trimestre jul-sep 2026 sigue sin publicarse, así que esto
no cambia nada todavía — corre solo cuando el Banco Central lo publique.

Lo que **no** se auto-completa, y por qué:
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
