# Estructura del reporte Power BI — LUMBRA FW27

Este documento es la especificación exacta para replicar las tres pestañas del
dashboard HTML en un reporte de Power BI de tres páginas. Úsalo junto con
`tema_lumbra.json` (Ver > Temas > Explorar temas > Examinar... para importarlo).

## Tablas a importar (Obtener datos > Texto/CSV)

| Archivo | Filas | Uso |
|---|---|---|
| `data/processed/lumbra_datos_limpios.csv` | 83 | Páginas 1 y 3 |
| `data/processed/lumbra_scatter_oportunidad.csv` | 8 | Página 1 (mapa de oportunidad) |
| `data/processed/simulacion_montecarlo.csv` | 10.000 | Página 2, vista Monte Carlo — **estático, ver aviso abajo** |

No hace falta crear relaciones entre estas tres tablas — cada una alimenta
visuales distintos y no comparten una clave común útil para el modelo.

**Aviso importante sobre `simulacion_montecarlo.csv` (corregido tras auditoría):** a diferencia
de la vista "Valor fijo" (que sí usa los parámetros "Y si" en vivo, editables desde el propio
informe), este archivo es una **foto estática de un solo escenario**, generada una vez desde la
terminal con:
```
python src/montecarlo.py --pais Suecia --precio 200 --coste 80 --tasa-min 10 --tasa-max 50 --coste-devolucion 10
```
Si quieres ver la distribución Monte Carlo de otro país, precio o costo, hay que volver a correr
ese comando con los valores nuevos y reimportar el CSV en Power BI — **no se actualiza solo** al
mover los parámetros "Y si" de la vista "Valor fijo", porque son dos mecanismos distintos (Power
BI no puede ejecutar Python en vivo). Si algún día se quiere que ambas vistas respondan al mismo
control, la vía es armar la simulación completa en DAX con `GENERATESERIES` + `RAND()`, no CSV.

---

## Página 1 — Resumen

**Slicers arriba:** Segmento (columna `Segmento` de `lumbra_datos_limpios`),
Field parameter "Moneda" con dos campos: `Valor USD` y `Valor EUR`.

| Visual | Tipo | Filtro | Ejes | Color |
|---|---|---|---|---|
| Poder de compra | Barras horizontales | `Indicador = "PIB per capita nominal"` | Eje: `Pais`, Valor: field parameter Moneda | Por `Segmento` (usa el tema) |
| Mapa de oportunidad | Dispersión (Scatter) | tabla `lumbra_scatter_oportunidad` completa | X: `TamanoUSD`, Y: `CrecimientoPct`, Tamaño: `PibPerCapitaUSD` | Leyenda: `Segmento` |
| Cuota de lujo | Barras horizontales | `Indicador` contiene "Cuota de lujo" | Eje: `Pais`, Valor: `Valor Porcentaje` | Por `Segmento` |
| Riesgo de devolución | Barras horizontales | `Indicador = "Tasa de devolucion eCommerce moda"` | Eje: `Pais`, Valor: `Valor Porcentaje` | Por `Segmento` |

---

## Página 2 — Simulador de producto

### Parámetros "Y si" (Modelado > Nuevo parámetro > Numérico)

| Parámetro | Mín | Máx | Incremento | Default |
|---|---|---|---|---|
| Precio | 0 | 1000 | 1 | 155 |
| Coste | 0 | 500 | 1 | 52 |
| CosteDevolucion | 0 | 200 | 1 | 13 |
| Unidades (modo manual) | 0 | 50000 | 10 | 500 |

### Datos del caso incorporados (ficticios)

- **Costes por categoría:** Zueco 27€, Bamba/Zapatilla 32€, Zapato 52€, Barefoot 38€, Botín plano 56€, Bailarina 22€.
- **PVP vigente por categoría:** Zueco 129€, Bamba 149€, Zapato 189€, Barefoot 159€, Botín plano 209€, Bailarina 139€.
- **Precio wholesale (canal Presencial/B2B)** = Coste × 2,5. Si usas un parámetro "Y si" para este multiplicador (como se describe más abajo), solo hay que cambiar su valor por defecto — no hace falta tocar ninguna fórmula DAX.
- **Coste adicional de aprovisionamiento** = 10€ por unidad, universal en ambos canales,
  SE SUMA AL COSTE (`CosteTotalUnitario = Coste + CosteAdicionalAprovisionamiento`, ver
  las medidas más abajo) — supuesto del caso, tratado como coste de producción, no como
  comisión de venta. 
- **IVA por país** (para el canal E-commerce, PVP incluye IVA): España 21%, Francia 20%,
  Alemania 19%, Países Bajos 21%, Suecia 25% (del Excel de LUMBRA); Dinamarca 25%, Noruega 25%,
  Finlandia 25,5%, Austria 20%, Suiza 8,1%, Bélgica 21%, Luxemburgo 17% (verificado con
  fuentes fiscales independientes, ver notas).
- **Dos canales confirmados**: Presencial (B2B, showrooms/agentes) y E-commerce (B2C,
  tienda online propia). LUMBRA asume el coste de devolución en ambos, bajo regulación UE, salvo
  incumplimiento de condiciones por el cliente final.

### Medidas DAX (vista "Valor fijo") — actualizadas para reflejar el canal

**Nota de modelado:** la primera versión de esta fórmula no restaba
el ingreso perdido por devolución — solo el costo logístico del envío de vuelta,
como si un producto devuelto igual generara la venta completa. Eso hacía que el
margen casi no variara entre países con distinta tasa de devolución real (62-63%
sin importar si la tasa era 36% o 45%). La versión corregida sí resta esa venta
perdida, igual que ya hacía la vista Monte Carlo — con esto, Austria/Alemania/Suiza
(36%/44%/45% reales) muestran 27,4%/18,8%/17,7% de margen respectivamente: una
diferencia de negocio real, no un redondeo.

```
Ingreso = [Unidades] * [Precio]
DeduccionCanal = IF([Canal] = "ecommerce", [Ingreso] * [IVA]/100 / (1 + [IVA]/100), 0)
IngresoNeto = [Ingreso] - [DeduccionCanal]
CosteTotalUnitario = [Coste] + [CosteAdicionalAprovisionamiento]
CosteProduccionTotal = [Unidades] * [CosteTotalUnitario]
MargenBruto = [IngresoNeto] - [CosteProduccionTotal]
DevolucionesEstimadas = [Unidades] * ([TasaDevolucion] / 100)
PrecioNetoUnitario = DIVIDE([IngresoNeto], [Unidades])
IngresoPerdido = [DevolucionesEstimadas] * [PrecioNetoUnitario]
CosteDevolucionTotal = [DevolucionesEstimadas] * [CosteDevolucion]
ComisionAgenteTotal = IF([Canal] = "presencial", [IngresoNeto] * [ComisionAgentePct]/100, 0)
CosteTransporteTotal = IF([Canal] = "presencial", [Unidades] * [CosteTransporteUnitario], 0)
MargenNeto = [MargenBruto] - [IngresoPerdido] - [CosteDevolucionTotal] - [ComisionAgenteTotal] - [CosteTransporteTotal]
MargenNetoPct = DIVIDE([MargenNeto], [Ingreso])
```

Necesitas cinco parámetros nuevos: `IVA` (numérico, 0-30, default 19), `CosteAdicionalAprovisionamiento`
(numérico, 0-50, default 10 — se suma al coste base, aplica en los dos canales, no está en el
Excel de costes de LUMBRA), `Canal` (parámetro de campo con dos valores: "ecommerce" /
"presencial", o dos medidas separadas si prefieres empezar más simple), y dos nuevos
**solo para canal Presencial/B2B** :

- `ComisionAgentePct` (numérico, 0-100, default 12) — comisión del agente de ventas B2B, se
  descuenta del Ingreso Neto, no del Ingreso Bruto.
- `CosteTransporteUnitario` (numérico, 0-20, default 2,40) — supuesto ficticio de coste medio de transporte por par; en un caso real saldría de las tarifas del transportista, promediadas por grupo de países y categoría.
  para el detalle de los 6 valores que componen este promedio. También se revisaron las tarifas
  UPS por kg/zona (tarifa individual, no por lote) y los extracargos de UPS/FedEx/TNT 2026 — estos
  últimos son recargos por casos excepcionales (entrega en sábado, aduana Reino Unido, paquete
  fuera de medida, dirección residencial), no aplican a un envío estándar de zapatos dentro de la
  UE, así que no se incluyeron en el promedio.

### Medidas DAX (vista "Monte Carlo", sobre `simulacion_montecarlo`)

```
MC_MargenMedio = AVERAGE(simulacion_montecarlo[margen_neto_unidad])
MC_P10 = PERCENTILE.INC(simulacion_montecarlo[margen_neto_unidad], 0.1)
MC_P90 = PERCENTILE.INC(simulacion_montecarlo[margen_neto_unidad], 0.9)
MC_DesvEst = STDEV.P(simulacion_montecarlo[margen_neto_unidad])
MC_MargenTotal = [MC_MargenMedio] * [Unidades]
```

### Cómo alternar entre las dos vistas

Power BI no recalcula 10.000 filas en vivo como hace el navegador (ver
conversación previa) — por eso las dos vistas son estáticas y se alternan con
un **bookmark**:
1. Arma la vista "Valor fijo" (tarjetas de arriba) y guárdala como bookmark
   `Vista_Fijo` (oculta los visuales de Monte Carlo con "Ocultar este visual"
   en las opciones del bookmark).
2. Arma la vista "Monte Carlo" (histograma + tarjetas MC_*) y guárdala como
   `Vista_MonteCarlo`.
3. Agrega dos botones de texto ("Valor fijo" / "Monte Carlo (rango)") y
   asígnales la acción "Bookmark" apuntando a cada uno.

### Histograma

Visual de histograma nativo (o barras agrupando `margen_neto_unidad` en bins):
panel de formato > "Bins" > tamaño de contenedor, por ejemplo cada 5.

---

## Página 3 — Explorar datos

**Slicers arriba:** Segmento, Indicador (slicer de lista sobre la columna
`Indicador`), Field parameter "Moneda" (mismo patrón que Página 1).

| Visual | Tipo | Campos |
|---|---|---|
| Gráfico principal | Barras horizontales | Eje: `Pais`, Valor: field parameter Moneda (o `Valor Porcentaje` si el indicador filtrado es de tipo Porcentaje — usar dos visuales superpuestos con "Editar interacciones" si quieres que cambie de tipo automáticamente) |
| Tabla | Tabla nativa | `Pais`, Valor (USD/EUR según parámetro), `Anio`, `Fuente` |

**Para que la URL de la fuente sea clicable:** selecciona la columna `URL` en
la vista de Datos > pestaña "Herramientas de columna" > Categoría de datos >
"URL web".

**Exportar para presentación:** cada visual tiene, en su menú `...`,
"Exportar datos" (CSV/Excel). Para el reporte completo: Archivo > Exportar >
PowerPoint o PDF.

---

## Página 4 — Competencia (actualizada — canal online/físico, filtros de país y ciudad)

**Siete tablas a importar en total.** Las tres del Excel de siempre, más cuatro CSV nuevos que generé desde la misma lógica de la web (no hace falta teclear nada a mano):

| Tabla | Origen | Para qué sirve |
|---|---|---|
| `Categorias FW27` | Excel, pestaña del mismo nombre | Coste y PVP vigente por categoría |
| `Benchmark PVP` (queda como `Benchmark_PVP` al importar) | Excel, pestaña del mismo nombre | LUMBRA + cada competidor, con PVP y Modelo |
| `Analisis Prospeccion` | Excel, pestaña del mismo nombre | Mínimo/media/máximo de mercado, Posición/Lectura (del planner) |
| `presencia_marcas` | **Nuevo**: `data/processed/presencia_marcas.csv` | 153 filas: una por cada Marca+País+Canal (Online/Físico) confirmado, con su fuente |
| `Dim_Marca` | **Nuevo**: `data/processed/dim_marca.csv` | Un atributo por marca: país de origen, cuota de mercado (solo Birkenstock la tiene), y nota de "sin presencia física" (Satorisan, Alohas) |
| `Sel_Ciudad` | **Nuevo**: `data/processed/sel_ciudad.csv` | Las 18 ciudades del estudio + su país — tabla desconectada, solo para el segmentador |
| `Sel_Pais` | **Nuevo**: `data/processed/sel_pais.csv` | Los 10 países del estudio — tabla desconectada, solo para el segmentador |

**Por qué `Sel_Ciudad` y `Sel_Pais` van desconectadas (sin relación en el Modelo):** igual que hicimos con `Selector Moneda` hace tiempo — si las relacionas directamente, Power BI intenta cruzar automáticamente y se pierde el control fino de "mostrar con menos opacidad los que solo venden online". Las dejas sueltas, y las medidas leen la selección con `SELECTEDVALUE`.

**Relaciones a crear** (vista de Modelo): solo `Dim_Marca[Marca]` ↔ `Benchmark_PVP[Marca]` (uno a varios, `Dim_Marca` es el uno). **`presencia_marcas` no necesita ninguna relación** — queda suelta, igual que `Sel_Ciudad` y `Sel_Pais` — porque todas las medidas de abajo la filtran a mano con `CALCULATE(..., presencia_marcas[Marca] = MarcaActual, ...)`, sin depender de que exista una línea en el Modelo. Si ya creaste esa relación, bórrala: causa un error de "rutas de acceso ambiguas" al combinarse con la de `Dim_Marca`.

**Dos relaciones más, imprescindibles, que se quedaron fuera en la primera versión de este documento** (sin ellas, `SELECTEDVALUE` en `PVP_Vigente` y `Factibilidad` siempre devuelve vacío, porque el segmentador de Categoría vive en `Benchmark_PVP` y nunca llega a las otras dos tablas):

- `Categorias_FW27[Categoría]` ↔ `Benchmark_PVP[Categoría]` — Cardinalidad **Uno a varios**, dirección de filtro cruzado **Ambas** (no "Única" — ese es el detalle que hace que el segmentador sí llegue hasta `Categorias_FW27`).
- `Analisis_Prospeccion[Categoría]` ↔ `Benchmark_PVP[Categoría]` — mismos ajustes: Uno a varios, dirección **Ambas**.

**Segmentadores:** `Categoría` (de `Benchmark_PVP`), `Sel_Pais[Pais]`, `Sel_Ciudad[Ciudad]` — los tres en selección única.

### Medidas DAX — presencia por canal

```dax
EstadoOnline = 
VAR MarcaActual = SELECTEDVALUE('Benchmark_PVP'[Marca])
VAR PaisSel = SELECTEDVALUE(Sel_Pais[Pais])
VAR NPaisesOnline = CALCULATE(DISTINCTCOUNT(presencia_marcas[Pais]), presencia_marcas[Marca] = MarcaActual, presencia_marcas[Canal] = "Online")
RETURN
    IF(NPaisesOnline = 0, "Pendiente de verificar",
        IF(ISBLANK(PaisSel) || CALCULATE(COUNTROWS(presencia_marcas), presencia_marcas[Marca]=MarcaActual, presencia_marcas[Canal]="Online", presencia_marcas[Pais]=PaisSel) > 0,
            "Confirmada (" & NPaisesOnline & "/10)", "Sin venta online confirmada en " & PaisSel))

EstadoFisico = 
VAR MarcaActual = SELECTEDVALUE('Benchmark_PVP'[Marca])
VAR NCiudades = CALCULATE(DISTINCTCOUNT(presencia_marcas[Ciudad]), presencia_marcas[Marca] = MarcaActual, presencia_marcas[Canal] = "Físico")
VAR TieneNota = CALCULATE(VALUES(Dim_Marca[NotaSinPresenciaFisica]), Dim_Marca[Marca] = MarcaActual)
RETURN
    IF(NCiudades > 0, "Confirmada",
        IF(NOT(ISBLANK(TieneNota)), "Sin indicios de presencia internacional", "Pendiente de verificar"))
```

### Medidas DAX — el mensaje combinado de país/ciudad (equivalente al aviso de la web)

```dax
TieneTiendaEnCiudadSel = 
VAR MarcaActual = SELECTEDVALUE('Benchmark_PVP'[Marca])
VAR CiudadSel = SELECTEDVALUE(Sel_Ciudad[Ciudad])
RETURN CALCULATE(COUNTROWS(presencia_marcas), presencia_marcas[Marca]=MarcaActual, presencia_marcas[Canal]="Físico", presencia_marcas[Ciudad]=CiudadSel) > 0

TieneOnlineEnPaisDeCiudad = 
VAR MarcaActual = SELECTEDVALUE('Benchmark_PVP'[Marca])
VAR PaisDeCiudad = SELECTEDVALUE(Sel_Ciudad[Pais])
RETURN CALCULATE(COUNTROWS(presencia_marcas), presencia_marcas[Marca]=MarcaActual, presencia_marcas[Canal]="Online", presencia_marcas[Pais]=PaisDeCiudad) > 0

MensajeFiltro = 
VAR CiudadSel = SELECTEDVALUE(Sel_Ciudad[Ciudad])
VAR PaisDeCiudad = SELECTEDVALUE(Sel_Ciudad[Pais])
VAR ConTienda = CALCULATE(DISTINCTCOUNT('Benchmark_PVP'[Marca]), FILTER('Benchmark_PVP', [TieneTiendaEnCiudadSel]), 'Benchmark_PVP'[Tipo] <> "LUMBRA")
VAR SoloOnline = CALCULATE(DISTINCTCOUNT('Benchmark_PVP'[Marca]), FILTER('Benchmark_PVP', [TieneOnlineEnPaisDeCiudad] && NOT([TieneTiendaEnCiudadSel])), 'Benchmark_PVP'[Tipo] <> "LUMBRA")
RETURN
    IF(ISBLANK(CiudadSel), BLANK(),
        ConTienda & " competidor(es) con tienda física en " & CiudadSel & ", y " & SoloOnline &
        " más con venta online confirmada en " & PaisDeCiudad & " aunque sin tienda ahí.")
```

Pon `MensajeFiltro` en una tarjeta o cuadro de texto vinculado, justo como el aviso de la web. Para colorear el gráfico igual que en la web (LUMBRA / tienda física / solo online), usa formato condicional en el color de las barras con una regla basada en `TieneTiendaEnCiudadSel`.

### Medidas DAX — resumen recalculado (ya existentes, sin cambio en su lógica)

```dax
PVP_Vigente = SELECTEDVALUE('Categorias FW27'[PVP VIGENTE])

DeltaEUR_vsLUMBRA = 
VAR PvpFila = SELECTEDVALUE('Benchmark_PVP'[PVP Corregido])
RETURN IF(SELECTEDVALUE('Benchmark_PVP'[Tipo]) = "LUMBRA", BLANK(), PvpFila - [PVP_Vigente])

DeltaPct_vsLUMBRA = DIVIDE([DeltaEUR_vsLUMBRA], [PVP_Vigente])

Factibilidad = 
VAR Coste = SELECTEDVALUE('Categorias FW27'[Coste unitario])
VAR Minimo = SELECTEDVALUE('Analisis Prospeccion'[Mínimo mercado])
VAR PvpVigente = [PVP_Vigente]
RETURN
    IF(Minimo - Coste > 0, "Factible (sólida)",
        IF(PvpVigente - Coste > 0, "Factible (ajustada)", "No factible"))
```

`PVP Corregido` es la columna que arma el merge que ya hiciste (PVP por fila, sustituyendo el planner por el PVP vigente solo en la fila de LUMBRA).

**Resumen en tabla** (no tarjetas KPI): Indicador | Valor, con PVP vigente, mínimo, media, máximo, Δ vs. media, Posición (marcada "según planner, referencia") y Factibilidad.

### Medidas DAX — mínimo/media/máximo recalculados con el filtro activo

```dax
CompetidorVisibleEnFiltro = 
VAR MarcaActual = SELECTEDVALUE('Benchmark_PVP'[Marca])
VAR PaisSel = SELECTEDVALUE(Sel_Pais[Pais])
VAR CiudadSel = SELECTEDVALUE(Sel_Ciudad[Ciudad])
VAR PaisDeCiudad = SELECTEDVALUE(Sel_Ciudad[Pais])
VAR OnlineEnPaisSel = CALCULATE(COUNTROWS(presencia_marcas), presencia_marcas[Marca]=MarcaActual, presencia_marcas[Canal]="Online", presencia_marcas[Pais]=PaisSel) > 0
VAR PresenteParaCiudad = [TieneTiendaEnCiudadSel] || CALCULATE(COUNTROWS(presencia_marcas), presencia_marcas[Marca]=MarcaActual, presencia_marcas[Canal]="Online", presencia_marcas[Pais]=PaisDeCiudad) > 0
RETURN
    (ISBLANK(PaisSel) || OnlineEnPaisSel) && (ISBLANK(CiudadSel) || PresenteParaCiudad)

MinimoFiltro = CALCULATE(MIN('Benchmark_PVP'[PVP Corregido]), FILTER('Benchmark_PVP', [CompetidorVisibleEnFiltro]), 'Benchmark_PVP'[Tipo] <> "LUMBRA")
MediaFiltro = CALCULATE(AVERAGE('Benchmark_PVP'[PVP Corregido]), FILTER('Benchmark_PVP', [CompetidorVisibleEnFiltro]), 'Benchmark_PVP'[Tipo] <> "LUMBRA")
MaximoFiltro = CALCULATE(MAX('Benchmark_PVP'[PVP Corregido]), FILTER('Benchmark_PVP', [CompetidorVisibleEnFiltro]), 'Benchmark_PVP'[Tipo] <> "LUMBRA")
DeltaMediaFiltro_EUR = [PVP_Vigente] - [MediaFiltro]
DeltaMediaFiltro_Pct = DIVIDE([DeltaMediaFiltro_EUR], [MediaFiltro])
```

Usa `MinimoFiltro`/`MediaFiltro`/`MaximoFiltro` en la tabla de Resumen en vez de las columnas fijas de `Analisis Prospeccion` cuando haya país o ciudad seleccionados — igual que en la web, `Posición` y `Lectura comercial` siguen siendo las del Excel (referencia sobre todo el mercado), nunca se recalculan.


**Cuota de mercado:** trae `Dim_Marca[CuotaMercadoTexto]` y `Dim_Marca[CuotaMercadoFuente]` directo a la tabla de detalle — sin necesitar ninguna medida, son columnas ya calculadas.

**Aviso visible sobre la lectura comercial:** el mismo de siempre — un cuadro de texto encima de `Lectura comercial` señalando que ese texto fue redactado sobre el precio del planner.


---

## Página 5 — Tendencias por país (nueva, categoría de calzado dominante)

**Tabla a importar:** `data/processed/categoria_dominante_pais.csv` (10 filas: los
10 países del estudio, con bandera en emoji, categoría de calzado más grande
según Statista Market Insights, cifra cuando estuvo disponible, y las marcas
notables de esa categoría con su propia fuente).

**Visual principal:** Tabla simple — Pais | Categoria | Cifra. Sin necesitar
ninguna medida DAX, son columnas ya calculadas. Resalta con negrita o color las
2 filas marcadas `excepcion = TRUE` (Holanda=Sneakers, Finlandia=Textile & Other
Footwear) usando formato condicional sobre la columna `excepcion`.

**Cuadro de texto 1 — la tendencia de temporada:** pega el texto de
`TENDENCIA_BOTAS_ALTAS` (está en `build_dashboard.py`, o cópialo del HTML ya
generado) — cita real de Who What Wear y Marie Claire sobre botas altas como
tendencia dominante de Otoño/Invierno 2026, mencionando que Vagabond y Camper
(ya en tu benchmark) están en ese espacio.

**Cuadro de texto 2 — la recomendación de negocio:** conecta este hallazgo con
la Factibilidad ya calculada del Botín plano en la página de Competencia — Boots
es la categoría más grande en 7 de 8 países con dato, y la tendencia apunta a
versiones altas, mientras el Botín plano actual de LUMBRA es bajo.

## Página 6 — España (NUEVA)

> **Guía vigente para esta página: `kit_paginas_6_y_7.md`** (consultas de Power Query, medidas DAX y especificación de cada visual, listas para copiar). Lo de abajo quedó como contexto; si hay diferencias, manda el kit.


**Tablas a importar** (las 4, generadas por `exportar_espana_csv()` en
`build_dashboard.py`):
- `espana_canal_b2b.csv` — 5 filas (Cadenas, Hiper, Outlets, Online, Multimarca), columnas `canal`, `pct`
- `espana_export_ranking.csv` — 3 filas (Italia/Francia/Alemania), con `devolucion_tipo` marcando cuál es real y cuál estimada
- `espana_margen_export.csv` — 18 filas (3 países × 6 categorías), margen neto y % por categoría
- `espana_margen_canal.csv` — 6 filas (una por categoría), margen Online vs. Presencial dentro de España

No hace falta crear relaciones entre estas 4 tablas — cada una alimenta un
visual distinto, igual que el patrón de la página de Resumen.

**Visual 1 — Reparto de canal:** gráfico de barras sobre `espana_canal_b2b`,
eje X = `canal`, valor = `pct`. Resalta la barra de Multimarca y Online con un
color distinto (son el argumento central: mayorista cae, online sube).

**Visual 2 — Ranking de exportación:** gráfico de barras sobre
`espana_export_ranking`, eje X = `pais` (ordenado por `posicion`), valor =
`crecimiento_2024_pct`. Agrega una tarjeta de texto con `precio_historico_2018`
de Alemania, destacando que paga más por par que Italia o Francia.

**Visual 3 — Margen por país, 3 categorías:** gráfico de barras agrupado sobre
`espana_margen_export`, eje X = `Categoria`, leyenda = `Pais`, valor =
`PctMargen`. Usa `DevolucionTipo` en el tooltip para que quede visible cuál
país tiene dato real (Alemania) y cuáles estimado (Francia/Italia).

**Visual 4 — Margen por canal en España:** gráfico de barras agrupado sobre
`espana_margen_canal`, eje X = `categoria`, dos series: `pct_margen_online` y
`pct_margen_presencial`. El canal presencial usa el multiplicador wholesale
(2,5x, definido en `build_dashboard.py` como `MULTIPLICADOR_WHOLESALE`
— si vuelve a cambiar, solo hay que editar esa constante, no hace falta tocar
ninguna fórmula DAX de esta página).

**Tarjetas KPI (4):** replican las de la web — tamaño del Fashion eCommerce,
relación de devolución online/física, % de calzado en lo devuelto, y tamaño
del canal multimarca — cada una respondiendo una pregunta de negocio concreta,
no solo mostrando el número.

## Página 7 — Ranking de modelos (NUEVA)

> **Guía vigente para esta página: `kit_paginas_6_y_7.md`** (consultas de Power Query, medidas DAX y especificación de cada visual, listas para copiar). Lo de abajo quedó como contexto; si hay diferencias, manda el kit.


Margen de los 24 modelos activos (ficticios) de la colección FW27 (agrupados en 10
perfiles económicos, por coste/PVP idéntico), en cada uno de los 11 países del
estudio. Fuente: `data/raw/BuyPlan_Maestro_Modelos_Lumbra.xlsx`, hoja
"Maestro Modelos" — ya incorporado al proyecto, no depende de que alguien lo
vuelva a subir.

**Tabla a importar:** `ranking_margen_modelos.csv` — 110 filas (10 perfiles ×
11 países), columnas `Familia`, `Modelos` (nombres de modelo de BuyPlan agrupados),
`PVP`, `Coste`, `Pais`, `PctMargen`, `PromedioModelo`.

**Visual 1 — Ranking general:** barras sobre `PromedioModelo` por `Familia`
(agrupar/quitar duplicados, ya que se repite 11 veces por fila-país — usar
`Modelos` o `Familia`+`PVP` como clave de agrupación). Resaltar en un color distinto las filas de familia "Barefoot" para comparar su margen con el resto.

**Visual 2 — Por país:** mismo gráfico, pero filtrado por un segmentador de
`Pais`, usando `PctMargen` en vez de `PromedioModelo`.

**Visual 3 — Tabla completa:** `Familia`, `Modelos`, `PVP`, `Pais`, `PctMargen`
— para auditar cualquier celda hasta el modelo exacto de BuyPlan que la compone.

**El hallazgo a destacar en un cuadro de texto** (se calcula en `build_dashboard.py`, función `texto_hallazgo_ranking`):

> Con datos ficticios del sistema de compras de LUMBRA (BuyPlan, 24 modelos de la colección FW27), la familia con mejor margen promedio en los 11 países es **Bailarina** (35,4%) y la de peor margen es **Barefoot** (24,0%). El perfil de mayor margen es «Slip Fácil» (38,6% promedio); el de menor margen es «Botín Barefoot Monte» (20,8% promedio, llegando a solo 9,6% en Alemania): cuesta 54€ de producción y se vende a 179€ — es un problema de coste de producción, no de precio de venta.
