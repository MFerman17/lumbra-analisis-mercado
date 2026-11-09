# Kit para construir las páginas 6 (España) y 7 (Ranking de modelos) en Power BI

**Esta es la guía vigente para esas dos páginas.** Si algo difiere de `estructura_powerbi.md`, manda este documento (se escribió después, con los CSV ya definitivos).

Tiempo estimado: 45-60 minutos la primera vez. Todo es copiar y pegar; lo único manual es arrastrar campos a los visuales.

> **Qué no se hizo y por qué:** el archivo `.pbix` no se incluye en el repositorio (binario propietario); este kit describe cómo reconstruir las páginas. Lo que sí está hecho: los datos (CSV), las consultas, las medidas y la especificación de cada visual.

---

## 0. Antes de empezar (una sola vez)

1. Corre `python src/build_dashboard.py` (o el notebook de Colab) para que `data/processed/` tenga los CSV al día. Los que usa este kit:
   `espana_kpis.csv`, `espana_canal_b2b.csv`, `espana_export_ranking.csv`, `espana_margen_export.csv`, `espana_margen_canal.csv`, `espana_competencia_barefoot.csv`, `ranking_margen_modelos.csv`.
2. En Power BI Desktop abre tu `.pbix` actual. **Archivo → Opciones → Configuración regional**: déjalo como esté; el kit ya fuerza el formato numérico correcto (ver nota abajo).
3. Aplica el tema `powerbi/tema_lumbra.json` si no lo tenías aplicado (Vista → Temas → Examinar temas).

> **Nota importante sobre decimales.** Los CSV usan punto decimal (`0.537`). Con Windows en español, Power BI puede leer eso como `537` sin avisar. Por eso la consulta `fnCsv` de abajo pasa `"en-US"` al convertir tipos: **no la quites**.

---

## 1. Consultas de Power Query

**Inicio → Transformar datos.** Todo se pega en **Editor avanzado** (Inicio → Editor avanzado, o clic derecho en la consulta).

### 1.1 Parámetro `RutaDatos`

Inicio → Administrar parámetros → Nuevo parámetro:

- Nombre: `RutaDatos`
- Tipo: Texto
- Valor actual (debe terminar en `\`):

```
C:\ruta\al\proyecto\data\processed\
```

Si algún día la carpeta cambia de lugar (por ejemplo a una carpeta sincronizada de Drive u OneDrive), **solo se cambia este parámetro** y las 7 tablas siguen funcionando.

### 1.2 Función `fnCsv`

Nueva consulta → Consulta en blanco → Editor avanzado → pegar → renombrar la consulta a `fnCsv`:

```powerquery
(archivo as text, tipos as list) as table =>
let
    Origen = Csv.Document(
        File.Contents(RutaDatos & archivo),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),
    SinBOM = Table.TransformColumnNames(
        Encabezados,
        each Text.Replace(_, Character.FromNumber(65279), "")
    ),
    Tipado = Table.TransformColumnTypes(SinBOM, tipos, "en-US")
in
    Tipado
```

### 1.3 Las 7 tablas

Para cada una: Nueva consulta → Consulta en blanco → Editor avanzado → pegar → renombrar con el nombre indicado (el nombre importa: las medidas lo usan).

**`espana_kpis`**
```powerquery
let
    Origen = fnCsv("espana_kpis.csv", {{"kpi", type text}, {"pregunta", type text}, {"valor_num", type number}, {"valor_texto", type text}, {"detalle", type text}})
in
    Origen
```

**`espana_canal_b2b`**
```powerquery
let
    Origen = fnCsv("espana_canal_b2b.csv", {{"canal", type text}, {"pct", type number}})
in
    Origen
```

**`espana_export_ranking`**
```powerquery
let
    Origen = fnCsv("espana_export_ranking.csv", {{"pais", type text}, {"posicion", Int64.Type}, {"pct_valor_exportado", type number}, {"crecimiento_2024_pct", type number}, {"precio_historico_2018", type number}, {"iva", type number}, {"devolucion", type number}, {"devolucion_tipo", type text}})
in
    Origen
```

**`espana_margen_export`**
```powerquery
let
    Origen = fnCsv("espana_margen_export.csv", {{"Pais", type text}, {"DevolucionTipo", type text}, {"Categoria", type text}, {"PVP", type number}, {"Coste", type number}, {"MargenNeto", type number}, {"PctMargen", type number}})
in
    Origen
```

**`espana_margen_canal`**
```powerquery
let
    Origen = fnCsv("espana_margen_canal.csv", {{"categoria", type text}, {"pvp", type number}, {"coste", type number}, {"precio_wholesale", type number}, {"margen_online", type number}, {"pct_margen_online", type number}, {"margen_presencial", type number}, {"pct_margen_presencial", type number}})
in
    Origen
```

**`espana_competencia_barefoot`**
```powerquery
let
    Origen = fnCsv("espana_competencia_barefoot.csv", {{"marca", type text}, {"modelo", type text}, {"pvp", type number}, {"orden", Int64.Type}, {"nota", type text}})
in
    Origen
```

**`ranking_margen_modelos`**
```powerquery
let
    Origen = fnCsv("ranking_margen_modelos.csv", {{"Familia", type text}, {"NombreModelo", type text}, {"EsBarefoot", Int64.Type}, {"Modelos", type text}, {"PVP", type number}, {"Coste", type number}, {"Pais", type text}, {"PctMargen", type number}, {"PromedioModelo", type number}})
in
    Origen
```

Cuando termines: **Cerrar y aplicar**.

**No crees relaciones entre estas 7 tablas.** Cada una alimenta sus propios visuales; relacionarlas solo generaría ambigüedad. Oculta la columna `ranking_margen_modelos[PromedioModelo]` (no se usa; las medidas calculan el promedio ellas mismas).

---

## 2. Columnas calculadas y medidas

**Cómo crear cada una:** selecciona la tabla indicada en el panel Datos → pestaña **Herramientas de tabla / Modelado** → **Nueva medida** (o **Nueva columna**) → pegar → Enter. Después, en la cinta, ajusta el **Formato** que se indica.

### 2.1 Tabla `espana_kpis` (las 4 tarjetas)

| Medida | Formato |
|---|---|
| (texto, sin formato) | — |

```dax
KPI Mercado Online =
CALCULATE ( SELECTEDVALUE ( espana_kpis[valor_texto] ), espana_kpis[kpi] = "fashion_ecommerce" )
```
```dax
KPI Devolución Online vs Física =
CALCULATE ( SELECTEDVALUE ( espana_kpis[valor_texto] ), espana_kpis[kpi] = "devolucion_online_vs_fisica" )
```
```dax
KPI Calzado en Devoluciones =
CALCULATE ( SELECTEDVALUE ( espana_kpis[valor_texto] ), espana_kpis[kpi] = "calzado_pct_devuelto" )
```
```dax
KPI Canal Multimarca =
CALCULATE ( SELECTEDVALUE ( espana_kpis[valor_texto] ), espana_kpis[kpi] = "canal_multimarca" )
```

### 2.2 Tabla `espana_canal_b2b`

```dax
Cuota del Mercado = SUM ( espana_canal_b2b[pct] ) / 100
```
Formato: Porcentaje, 1 decimal.

```dax
Color Canal =
SWITCH (
    SELECTEDVALUE ( espana_canal_b2b[canal] ),
    "Multimarca (B2B tradicional, cayendo)", "#B8703F",
    "Online (creciendo)", "#5C7A52",
    "#A8724A"
)
```

### 2.3 Tabla `espana_export_ranking`

Nueva **columna** (no medida):
```dax
Etiqueta País = espana_export_ranking[posicion] & "º " & espana_export_ranking[pais]
```
Luego: selecciona la columna `Etiqueta País` → Ordenar por columna → `posicion`.

```dax
Crecimiento Exportación 2024 = AVERAGE ( espana_export_ranking[crecimiento_2024_pct] ) / 100
```
Formato: Porcentaje, 1 decimal.

```dax
Color Crecimiento =
IF ( [Crecimiento Exportación 2024] >= 0, "#5C7A52", "#B8703F" )
```
```dax
Precio por Par Alemania 2018 =
CALCULATE ( MAX ( espana_export_ranking[precio_historico_2018] ), espana_export_ranking[pais] = "Alemania" )
```
Formato: Número decimal, 2 decimales, sufijo ` €`.

### 2.4 Tabla `espana_margen_export`

```dax
Margen Neto % = AVERAGE ( espana_margen_export[PctMargen] )
```
Formato: Porcentaje, 1 decimal.

### 2.5 Tabla `espana_margen_canal`

```dax
Margen % Online = AVERAGE ( espana_margen_canal[pct_margen_online] )
```
```dax
Margen % Presencial/B2B = AVERAGE ( espana_margen_canal[pct_margen_presencial] )
```
Formato de las dos: Porcentaje, 1 decimal.

### 2.6 Tabla `espana_competencia_barefoot`

```dax
PVP Barefoot = AVERAGE ( espana_competencia_barefoot[pvp] )
```
Formato: Número decimal, 2 decimales, sufijo ` €`.

```dax
Color Marca =
SWITCH (
    SELECTEDVALUE ( espana_competencia_barefoot[marca] ),
    "LUMBRA", "#A8724A",
    "Pikolinos", "#9C7A3C",
    "Hispanitas", "#5C7A52",
    "#999999"
)
```

### 2.7 Tabla `ranking_margen_modelos`

```dax
Margen Promedio 11 Países =
CALCULATE (
    AVERAGE ( ranking_margen_modelos[PctMargen] ),
    REMOVEFILTERS ( ranking_margen_modelos[Pais] )
)
```
```dax
Margen en País Seleccionado = AVERAGE ( ranking_margen_modelos[PctMargen] )
```
Formato de las dos: Porcentaje, 1 decimal.

> La primera ignora el selector de país (siempre promedia los 11); la segunda sí lo obedece. Por eso una sola tabla alimenta los dos gráficos de la página.

```dax
Color Barra Modelo =
IF ( MAX ( ranking_margen_modelos[EsBarefoot] ) = 1, "#B8703F", "#5C7A52" )
```

Las cuatro medidas de las tarjetas de "mejor" y "peor" modelo:

```dax
Peor Modelo (Nombre) =
VAR t = ADDCOLUMNS ( VALUES ( ranking_margen_modelos[NombreModelo] ), "@m", [Margen Promedio 11 Países] )
RETURN CONCATENATEX ( TOPN ( 1, t, [@m], ASC ), ranking_margen_modelos[NombreModelo] )
```
```dax
Peor Modelo (Margen) =
VAR t = ADDCOLUMNS ( VALUES ( ranking_margen_modelos[NombreModelo] ), "@m", [Margen Promedio 11 Países] )
RETURN MINX ( t, [@m] )
```
```dax
Mejor Modelo (Nombre) =
VAR t = ADDCOLUMNS ( VALUES ( ranking_margen_modelos[NombreModelo] ), "@m", [Margen Promedio 11 Países] )
RETURN CONCATENATEX ( TOPN ( 1, t, [@m], DESC ), ranking_margen_modelos[NombreModelo] )
```
```dax
Mejor Modelo (Margen) =
VAR t = ADDCOLUMNS ( VALUES ( ranking_margen_modelos[NombreModelo] ), "@m", [Margen Promedio 11 Países] )
RETURN MAXX ( t, [@m] )
```
Formato de las dos de "Margen": Porcentaje, 1 decimal.

---

## 3. Página 6 — España

Página nueva → renombrar a **España**. Lienzo 16:9. Orden sugerido (de arriba hacia abajo):

```
┌──────────┬──────────┬──────────┬──────────┐
│ Tarjeta 1│ Tarjeta 2│ Tarjeta 3│ Tarjeta 4│   fila de 4 tarjetas
├──────────┴─────┬────┴──────────┴──────────┤
│  Visual A      │       Visual B           │   2 gráficos
├────────────────┴──────────────────────────┤
│  Visual C (ancho completo)                │
├───────────────────────┬───────────────────┤
│  Visual D             │  Visual E         │
└───────────────────────┴───────────────────┘
```

### Tarjetas (4) — visual "Tarjeta"

| # | Campo | Título de la tarjeta (cuadro de texto encima, o Título del visual) |
|---|---|---|
| 1 | `KPI Mercado Online` | ¿Qué tan grande es el mercado de moda online en España? |
| 2 | `KPI Devolución Online vs Física` | ¿Cuánto más se devuelve comprando online que en tienda? |
| 3 | `KPI Calzado en Devoluciones` | ¿Qué parte de lo que se devuelve en España es calzado? |
| 4 | `KPI Canal Multimarca` | ¿Qué parte del mercado de moda sigue pasando por el canal mayorista tradicional? |

Etiqueta de categoría: desactívala (el título ya explica). Fuente de cada cifra: viene en la columna `espana_kpis[detalle]`; ponla como **tooltip** o en una nota al pie de la página.

### Visual A — Reparto del mercado por canal
- Tipo: **Gráfico de barras agrupadas** (horizontal)
- Eje Y: `espana_canal_b2b[canal]`
- Valores: `Cuota del Mercado`
- Colores → Barras → botón *fx* → Formato por **Valor de campo** → `Color Canal`
- Etiquetas de datos: activadas
- Título: ¿Cómo se reparte el mercado de moda en España?
- Subtítulo (cuadro de texto): *Fuente: Acotex, «El Comercio Textil en Cifras 2025» — mercado total 11.107M€. El canal multimarca cae a 7% mientras el online sube a 17%.*

### Visual B — Ranking de exportación
- Tipo: **Gráfico de columnas agrupadas**
- Eje X: `espana_export_ranking[Etiqueta País]`
- Valores: `Crecimiento Exportación 2024`
- Colores → *fx* → Valor de campo → `Color Crecimiento`
- Etiquetas de datos: activadas
- Título: ¿A quién le exporta España, y quién paga más por par?
- Encima o debajo, una **tarjeta** con `Precio por Par Alemania 2018` y el texto: *Alemania es el 3er comprador de calzado español y el que mejor paga por par (20,92€ en 2018, frente a 16,47€ de Italia y 14,74€ de Francia). Fuente: Revista del Calzado / Dirección General de Aduanas 2024; IVACE 2019 (29€/par UE vs 11€/par terceros países).*

### Visual C — Margen por categoría: Alemania, Francia, Italia
- Tipo: **Gráfico de columnas agrupadas**
- Eje X: `espana_margen_export[Categoria]`
- Leyenda: `espana_margen_export[Pais]`
- Valores: `Margen Neto %`
- Información sobre herramientas: añade `espana_margen_export[DevolucionTipo]`
- Cuadro de texto bajo el título (déjalo así, es la advertencia obligatoria): *IVA y devolución reales por país. **Alemania: devolución real (Statista 2022). Francia e Italia: devolución estimada (27,5%), sin dato propio confirmado** — se incluyen porque son el 1º y 2º comprador de calzado español, con la salvedad marcada.*
- La leyenda queda solo con el nombre del país (sin paréntesis).

### Visual D — España: margen por canal
- Tipo: **Gráfico de columnas agrupadas**
- Eje X: `espana_margen_canal[categoria]`
- Valores: `Margen % Online` y `Margen % Presencial/B2B`
- Renombra las series en el visual a «Online» y «Presencial/B2B» (clic derecho sobre el campo → Cambiar nombre para este visual)
- Cuadro de texto bajo el título: *Canal presencial usa el multiplicador wholesale vigente (2,5x) y descuenta la comisión del agente de ventas (12%) y el coste de transporte (2,40€/par), igual que el Simulador. Devolución real por canal: Online 23,7% / Presencial 5,6% (ZigZag + Retail Economics, 2025).*

### Visual E — Competencia en España (Barefoot)
- Tipo: **Gráfico de barras agrupadas**
- Eje Y: `espana_competencia_barefoot[marca]` → Ordenar eje por `orden` ascendente
- Valores: `PVP Barefoot`
- Colores → *fx* → Valor de campo → `Color Marca`
- Etiquetas de datos: activadas
- Información sobre herramientas: `modelo` y `nota`
- Título: Competencia en España — precio del Barefoot, LUMBRA vs. las 2 marcas españolas comparables
- Cuadro de texto: *Pikolinos (modelo Betera) e Hispanitas (Transitional Barefoot) son marcas españolas de piel/confort tradicional con línea «Barefoot». Ambas compiten más abajo en precio que LUMBRA (159€).*

---

## 4. Página 7 — Ranking de modelos

Página nueva → renombrar a **Ranking de modelos**.

```
┌──────────────┬──────────────┬──────────────┬──────────────┐
│ Mejor modelo │ Margen mejor │ Peor modelo  │ Margen peor  │   4 tarjetas
├──────────────┴──────────────┴──────────────┴──────────────┤
│ Cuadro de texto: el hallazgo (Barefoot)                   │
├────────────────────────────────┬───────────────────────────┤
│ Visual F  (ranking 11 países)  │ Selector País + Visual G  │
├────────────────────────────────┴───────────────────────────┤
│ Visual H  (matriz modelo × país)                          │
└────────────────────────────────────────────────────────────┘
```

### Tarjetas (4)
`Mejor Modelo (Nombre)`, `Mejor Modelo (Margen)`, `Peor Modelo (Nombre)`, `Peor Modelo (Margen)`.
Para que estas cuatro no cambien al usar el selector de país, desactiva la interacción del selector sobre ellas (ver «Editar interacciones» más abajo); de todos modos usan siempre el promedio de los 11 países.

### Cuadro de texto — el hallazgo (copiar tal cual)

> **Qué modelos dejan más margen — y cuáles menos.** Con datos ficticios del sistema de compras de LUMBRA (BuyPlan, 24 modelos de la colección FW27), la familia con mejor margen promedio en los 11 países es **Bailarina** (35,4%) y la de peor margen es **Barefoot** (24,0%). El perfil de mayor margen es «Slip Fácil» (38,6% promedio); el de menor margen es «Botín Barefoot Monte» (20,8% promedio, llegando a solo 9,6% en Alemania): cuesta 54€ de producción y se vende a 179€ — es un problema de coste de producción, no de precio de venta.

### Visual F — Ranking de margen promedio (11 países)
- Tipo: **Gráfico de barras agrupadas** (horizontal)
- Eje Y: `ranking_margen_modelos[NombreModelo]`
- Valores: `Margen Promedio 11 Países`
- Orden: ⋯ → Ordenar eje → `Margen Promedio 11 Países` → Descendente
- Colores → *fx* → Valor de campo → `Color Barra Modelo` (Barefoot sale en terracota, el resto en verde)
- Información sobre herramientas: añade `ranking_margen_modelos[Modelos]` (los nombres exactos de BuyPlan que componen cada barra)
- Título: Ranking de margen promedio por modelo (los 11 países)

> Se usa **NombreModelo**, no Familia, a propósito: la familia se repetía en varias barras y parecía duplicado.

### Visual G — ¿Cómo cambia el margen según el país?
- Un **Segmentador** (dropdown) con `ranking_margen_modelos[Pais]`: Formato → Configuración del segmentador → Estilo **Lista desplegable**, selección única.
- Un **Gráfico de barras agrupadas**: Eje Y `NombreModelo`, Valores `Margen en País Seleccionado`, ordenado descendente por esa medida, color `Color Barra Modelo`.
- Mientras no se elija país, el gráfico muestra el promedio de los 11.
- **Cuadro de texto obligatorio bajo este gráfico y la matriz** (es la advertencia de datos estimados; no lo omitas): *IVA real de cada país y su tasa de devolución. Estimación, no dato propio: Noruega, Finlandia, Holanda, Bélgica y Luxemburgo no tienen tasa de devolución confirmada, así que usan el promedio de las tasas reales (33,1%).*

### Visual H — Matriz completa
- Tipo: **Matriz**
- Filas: `NombreModelo` · Columnas: `Pais` · Valores: `Margen en País Seleccionado`
- Formato condicional: Valores → Color de fondo → Escala de colores (mínimo rojo claro, máximo verde)
- Desactiva las subtotales de fila y de columna.
- **Editar interacciones:** selecciona el segmentador de país → Formato → Editar interacciones → en esta matriz pon el icono «Ninguno» (para que muestre siempre los 11 países).

---

## 5. Valores de control (si coinciden, quedó bien)

Estos números salen de los mismos CSV que la web; deben verse **idénticos** en Power BI.

| Dónde | Qué mirar | Debe dar |
|---|---|---|
| Tarjeta 1 | KPI Mercado Online | €7,842M |
| Tarjeta 2 | Devolución online vs física | 4.2x |
| Tarjeta 3 | Calzado en devoluciones | 27% |
| Tarjeta 4 | Canal multimarca | 7% |
| Visual B | Crecimiento Italia / Francia / Alemania | 39,0% / −6,2% / ≈7,1% |
| Visual B | Precio por par Alemania 2018 | 20,92 € |
| Visual C | Categoría Zueco: Alemania / Francia / Italia | 16,6% / 34,8% / 34,1% |
| Visual D | Categoría Zueco: Online / Presencial | ≈38,7% / ≈22,9% |
| Visual E | PVP LUMBRA / Pikolinos / Hispanitas | 159 € / 119,95 € / 119,90 € |
| Página 7 | Peor modelo | Botín Barefoot Monte — 20,8% |
| Página 7 | Mejor modelo | Slip Fácil — 38,6% |
| Visual H | Botín Barefoot Monte × Alemania | 9,6% |
| Visual H | Clog Nube × Suecia | 41,6% |

Si algún número no coincide, lo más probable es el formato decimal: revisa que la consulta `fnCsv` conserve el `"en-US"`.

---

## 6. Problemas frecuentes

- **Los porcentajes salen ×100 o ×1000:** falta el `"en-US"` en `fnCsv`, o la medida tiene formato «Número» en vez de «Porcentaje».
- **`DataSource.NotFound` al actualizar:** el parámetro `RutaDatos` no apunta a la carpeta correcta, o no termina en `\`.
- **El segmentador de país cambia las tarjetas o la matriz:** revisa «Editar interacciones» (sección 4).
- **Una barra no toma color:** en *fx* tiene que estar elegido «Valor de campo» y la medida de color, no «Escala de colores».
- **Después de actualizar los datos (nuevo precio, nueva tasa):** primero corre el notebook de Colab o `build_dashboard.py` para regenerar los CSV, luego en Power BI pulsa **Actualizar**.
- **Cuando el proyecto pase a otra carpeta o a la nube:** solo cambia el valor del parámetro `RutaDatos`.
