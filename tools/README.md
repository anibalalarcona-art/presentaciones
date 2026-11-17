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

## Aviso sobre fórmulas arrastradas

Varias columnas de fecha del Excel (p. ej. `Combustibles!A`) son fórmulas
arrastradas (`=A3+7`), no valores literales. `pptx_chart_sync.py` sólo *lee*
el Excel (nunca lo guarda), así que no tiene problema con esto: Excel ya dejó
el valor calculado en caché y basta leerlo.

El problema aparece si en el futuro otro script (no Excel) abre y vuelve a
**guardar** ese mismo archivo con `openpyxl` sin tocar esas celdas de fecha:
`openpyxl` no reevalúa fórmulas, así que el valor cacheado se pierde y la
celda queda en blanco. Si vas a automatizar también la escritura hacia el
Excel (por ejemplo, insertar filas nuevas con datos que vengan del Banco
Central), usa primero `xlsx_utils.flatten_formulas_to_values(...)` para
convertir el archivo a valores literales, y recién ahí agrega filas y guarda.

## Pendiente: traer los datos directamente desde el Banco Central

Ya se probó la conexión a la API SIETE del Banco Central
(`si3.bcentral.cl`, habilitada en este entorno) con las credenciales
guardadas como `BCCH_USER` / `BCCH_PASS`: responde y las credenciales son
válidas. Falta decidir, por serie, el código de la serie en el Banco Central
(p. ej. el balance efectivo del Gobierno Central, o la fila "BCCh" de
proyecciones de PIB) para completar un `fetch_bcch.py` que escriba esas filas
directamente en `Graficos 11112026.xlsx` (usando `flatten_formulas_to_values`
antes de guardar) y luego llame a `pptx_chart_sync.py`. Eso queda para
cuando se definan qué gráficos y qué series se alimentan así.
