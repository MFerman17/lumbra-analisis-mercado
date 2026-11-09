import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from pathlib import Path

wb = openpyxl.Workbook()

FONT = "Arial"
AZUL_INPUT = Font(name=FONT, size=10, color="0000FF")
NEGRO_FORMULA = Font(name=FONT, size=10, color="000000")
NEGRITA = Font(name=FONT, size=10, bold=True)
TITULO = Font(name=FONT, size=13, bold=True)
SUBTITULO = Font(name=FONT, size=10, italic=True, color="4A3B2E")
AMARILLO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GRIS_HEADER = PatternFill(start_color="2B2018", end_color="2B2018", fill_type="solid")
HEADER_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PENDIENTE_FILL = PatternFill(start_color="F0E6D6", end_color="F0E6D6", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="E8E2D5", end_color="E8E2D5", fill_type="solid")

# 11 paises (10 del estudio + Espana), cada uno con su segmento para poder
# seguir agrupando/filtrando por region si hace falta.
PAISES = [
    ("Suecia", "Nordicos"), ("Dinamarca", "Nordicos"), ("Noruega", "Nordicos"), ("Finlandia", "Nordicos"),
    ("Alemania", "DACH"), ("Austria", "DACH"), ("Suiza", "DACH"),
    ("Holanda", "Benelux"), ("Belgica", "Benelux"), ("Luxemburgo", "Benelux"),
    ("España", "España"),
]

# Unidades Año 1 por defecto: reparto EQUITATIVO del total que ya tenia cada
# segmento (800 Nordicos/4, 900 DACH/3, 500 Benelux/3), Espana se queda en 300.
# Es un supuesto de partida transparente, no un dato medido - cambialo en la
# hoja Parametros si tienes un reparto real distinto en mente.
UNIDADES_DEFECTO = {
    "Suecia": 200, "Dinamarca": 200, "Noruega": 200, "Finlandia": 200,
    "Alemania": 300, "Austria": 300, "Suiza": 300,
    "Holanda": 167, "Belgica": 167, "Luxemburgo": 166,
    "España": 300,
}
CRECIMIENTO_DEFECTO = {p: 0.20 for p, _ in PAISES}
CRECIMIENTO_DEFECTO["España"] = 0.25
ONLINE_DEFECTO = {
    "Suecia": 0.75, "Dinamarca": 0.75, "Noruega": 0.75, "Finlandia": 0.75,
    "Alemania": 0.70, "Austria": 0.70, "Suiza": 0.70,
    "Holanda": 0.75, "Belgica": 0.75, "Luxemburgo": 0.75,
    "España": 0.60,
}
NOTA_PAIS = {
    "Suecia": "Tasa de devolución real (Statista)", "Dinamarca": "Tasa de devolución real (Statista)",
    "Noruega": "Sin tasa de devolución real confirmada", "Finlandia": "Sin tasa de devolución real confirmada",
    "Alemania": "Tasa de devolución real (Statista)", "Austria": "Tasa de devolución real (Statista)",
    "Suiza": "Tasa de devolución real (Statista)",
    "Holanda": "Sin tasa de devolución real confirmada", "Belgica": "Sin tasa de devolución real confirmada",
    "Luxemburgo": "Sin tasa de devolución real confirmada",
    "España": "PENDIENTE — tasa de devolución real 23,7% (ZigZag+Retail Economics 2025) ya confirmada; resto del estudio B2B/B2C en curso",
}

# ============================================================
# HOJA 1: Parametros (celdas de entrada editables - AZUL, fondo amarillo)
# ============================================================
ws_p = wb.active
ws_p.title = "Parametros"
ws_p["A1"] = "Parámetros del presupuesto, por país (edita las celdas en azul con fondo amarillo)"
ws_p["A1"].font = TITULO
ws_p["A2"] = ("Unidades Año 1 por país = reparto EQUITATIVO del total que antes vivía a nivel de segmento — "
              "es un supuesto de partida, no un dato medido. Ajusta cada país a tu propio criterio.")
ws_p["A2"].font = SUBTITULO
ws_p["A3"] = "Leyenda: azul + fondo amarillo = celda editable (tu supuesto) · negro = fórmula, no tocar · relleno beige = España, pendiente del estudio propio"
ws_p["A3"].font = Font(name=FONT, size=9, italic=True, color="4A3B2E")

headers = ["Segmento", "País", "Unidades Año 1 (tu supuesto)", "Crecimiento Año 2 (%) (tu supuesto)",
           "% Canal Online (tu supuesto)", "Nota",
           "Fuente/Método (prospectos, tiendas contactadas, ferias, benchmark...)",
           "Fecha de la estimación", "Responsable", "Nivel de confianza", "Estado"]
HEADER_ROW_PARAM = 4
for c, h in enumerate(headers, start=1):
    cell = ws_p.cell(row=HEADER_ROW_PARAM, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = BORDER

PARAM_START_ROW = 5
for i, (pais, seg) in enumerate(PAISES):
    r = PARAM_START_ROW + i
    ws_p.cell(row=r, column=1, value=seg).font = NEGRITA
    ws_p.cell(row=r, column=2, value=pais).font = NEGRITA
    c_unid = ws_p.cell(row=r, column=3, value=UNIDADES_DEFECTO[pais]); c_unid.font = AZUL_INPUT; c_unid.fill = AMARILLO
    c_crec = ws_p.cell(row=r, column=4, value=CRECIMIENTO_DEFECTO[pais]); c_crec.font = AZUL_INPUT; c_crec.fill = AMARILLO
    c_crec.number_format = "0.0%"
    c_online = ws_p.cell(row=r, column=5, value=ONLINE_DEFECTO[pais]); c_online.font = AZUL_INPUT; c_online.fill = AMARILLO
    c_online.number_format = "0.0%"
    ws_p.cell(row=r, column=6, value=NOTA_PAIS[pais]).font = Font(name=FONT, size=9, italic=True, color="4A3B2E")
    # Columnas nuevas de trazabilidad (auditoria) - en blanco a proposito: solo
    # el equipo comercial de LUMBRA tiene esta informacion, no se puede inventar.
    for col_idx in [7, 8, 9]:
        c_blank = ws_p.cell(row=r, column=col_idx, value="Pendiente de completar")
        c_blank.font = AZUL_INPUT; c_blank.fill = AMARILLO
    c_conf = ws_p.cell(row=r, column=10, value="Sin definir")
    c_conf.font = AZUL_INPUT; c_conf.fill = AMARILLO
    c_estado = ws_p.cell(row=r, column=11, value="Estimado")
    c_estado.font = AZUL_INPUT; c_estado.fill = AMARILLO
    if pais == "España":
        for c in range(1, 12):
            ws_p.cell(row=r, column=c).fill = PENDIENTE_FILL
    for c in range(1, 12):
        ws_p.cell(row=r, column=c).border = BORDER
PARAM_END_ROW = PARAM_START_ROW + len(PAISES) - 1

for col, width in zip("GHIJK", [32, 16, 16, 16, 14]):
    ws_p.column_dimensions[col].width = width
ws_p.cell(row=PARAM_END_ROW+2, column=1,
          value=("Nivel de confianza: usa Alto / Medio / Bajo. Estado: usa Confirmado / Estimado / Pendiente. "
                 "Mientras el equipo comercial no las complete, estas 5 columnas dejan explícito que las unidades "
                 "y el crecimiento de arriba son un punto de partida, no una demanda confirmada.")).font = Font(name=FONT, size=9, italic=True, color="B8703F")
ws_p.cell(row=PARAM_END_ROW+2, column=1).alignment = Alignment(wrap_text=True)

ws_p.column_dimensions["A"].width = 12
ws_p.column_dimensions["B"].width = 14
ws_p.column_dimensions["C"].width = 20
ws_p.column_dimensions["D"].width = 20
ws_p.column_dimensions["E"].width = 18
ws_p.column_dimensions["F"].width = 60

COSTE_LOG_ROW = PARAM_END_ROW + 2
ws_p.cell(row=COSTE_LOG_ROW, column=1,
          value="Coste logístico por devolución (€/unidad) — tu supuesto, no confirmado como dato real de LUMBRA").font = Font(name=FONT, size=9, italic=True)
c_costelog = ws_p.cell(row=COSTE_LOG_ROW, column=3, value=13); c_costelog.font = AZUL_INPUT; c_costelog.fill = AMARILLO
CELDA_COSTE_LOGISTICO = f"Parametros!$C${COSTE_LOG_ROW}"

print("Hoja Parametros lista —", len(PAISES), "países")

# ============================================================
# HOJA 2: Datos_Base (datos reales ya verificados en el proyecto, con fuente)
# ============================================================
ws_d = wb.create_sheet("Datos_Base")
ws_d["A1"] = "Datos base reales (verificados en el estudio — no editar sin revisar la fuente)"
ws_d["A1"].font = TITULO

ws_d["A3"] = "Categorías FW27 — coste, PVP y plan de compra (Collection Planner, datos ficticios)"
ws_d["A3"].font = NEGRITA
cat_headers = ["Categoría", "Coste unitario (€)", "PVP retail c/IVA (€)", "Uds. de compra (plan)"]
HEADER_ROW_CAT = 4
for c, h in enumerate(cat_headers, start=1):
    cell = ws_d.cell(row=HEADER_ROW_CAT, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER; cell.border = BORDER
# Categorias del caso (ficticias): coste y PVP de cada fila = promedio ponderado por
# unidades de los modelos de esa familia en el maestro BuyPlan (24 modelos x 220 pares
# = 5.280 uds), no un promedio simple.
categorias_reales = [
    ("Zueco", 27.0, 129, 880), ("Bamba / Zapatilla", 29.0, 152.3333, 1320), ("Zapato", 52.0, 189, 660),
    ("Barefoot", 43.3333, 165.6667, 660), ("Botín plano", 46.1667, 192.3333, 1320), ("Bailarina", 22.0, 139, 440),
]
CAT_START = 5
for i, (cat, coste, pvp, uds_compra) in enumerate(categorias_reales):
    r = CAT_START + i
    ws_d.cell(row=r, column=1, value=cat).border = BORDER
    ws_d.cell(row=r, column=2, value=coste).border = BORDER
    ws_d.cell(row=r, column=3, value=pvp).border = BORDER
    ws_d.cell(row=r, column=4, value=uds_compra).number_format = "#,##0"; ws_d.cell(row=r, column=4).border = BORDER
r_prom = CAT_START + len(categorias_reales)
ws_d.cell(row=r_prom, column=1, value="Promedio PONDERADO por el plan de compra (no ponderación igual)").font = NEGRITA
f_coste_prom = ws_d.cell(row=r_prom, column=2,
                          value=f"=SUMPRODUCT(B{CAT_START}:B{r_prom-1},D{CAT_START}:D{r_prom-1})/SUM(D{CAT_START}:D{r_prom-1})")
f_coste_prom.font = NEGRO_FORMULA
f_pvp_prom = ws_d.cell(row=r_prom, column=3,
                        value=f"=SUMPRODUCT(C{CAT_START}:C{r_prom-1},D{CAT_START}:D{r_prom-1})/SUM(D{CAT_START}:D{r_prom-1})")
f_pvp_prom.font = NEGRO_FORMULA
f_total_uds_compra = ws_d.cell(row=r_prom, column=4, value=f"=SUM(D{CAT_START}:D{r_prom-1})")
f_total_uds_compra.font = NEGRITA; f_total_uds_compra.number_format = "#,##0"
for c in range(1, 5):
    ws_d.cell(row=r_prom, column=c).border = BORDER
CELDA_COSTE_PROM = f"Datos_Base!$B${r_prom}"
CELDA_PVP_PROM = f"Datos_Base!$C${r_prom}"
CELDA_TOTAL_UDS_COMPRA = f"Datos_Base!$D${r_prom}"

ws_d.cell(row=r_prom+2, column=1, value="Wholesale = Coste × multiplicador (BuyPlan Config: 'Objetivo MG LUMBRA')").font = NEGRITA
ws_d.cell(row=r_prom+3, column=1, value="Multiplicador wholesale (supuesto ficticio)")
c_mult = ws_d.cell(row=r_prom+3, column=2, value=2.5); c_mult.font = AZUL_INPUT
CELDA_MULTIPLICADOR = f"Datos_Base!$B${r_prom+3}"
ws_d.cell(row=r_prom+4, column=1, value="Coste adicional de aprovisionamiento (€/unidad) — supuesto del caso")
c_costeadic = ws_d.cell(row=r_prom+4, column=2, value=10); c_costeadic.font = AZUL_INPUT
CELDA_COSTE_ADICIONAL = f"Datos_Base!$B${r_prom+4}"

IVA_TITLE_ROW = r_prom + 7
ws_d.cell(row=IVA_TITLE_ROW, column=1, value="IVA por país (real, con fuente oficial nombrada)").font = NEGRITA
iva_headers = ["País", "Segmento", "IVA", "Fuente"]
HEADER_ROW_IVA = IVA_TITLE_ROW + 1
for c, h in enumerate(iva_headers, start=1):
    cell = ws_d.cell(row=HEADER_ROW_IVA, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER; cell.border = BORDER
IVA_POR_PAIS = {
    "Suecia": 0.25, "Dinamarca": 0.25, "Noruega": 0.25, "Finlandia": 0.255,
    "Alemania": 0.19, "Austria": 0.20, "Suiza": 0.081,
    "Holanda": 0.21, "Belgica": 0.21, "Luxemburgo": 0.17, "España": 0.21,
}
FUENTE_IVA_UE = "Comisión Europea, TEDB — Taxes in Europe Database (ec.europa.eu/taxation_customs/tedb)"
FUENTE_IVA_POR_PAIS = {
    "Suiza": "Administración Federal de Contribuciones (ESTV), tipo normal 8,1% — estv.admin.ch/estv/en/home/value-added-tax/vat-rates-switzerland.html",
    "Noruega": "Skatteetaten (Noruega no es UE, no está en TEDB) — tipo general 25%",
}
IVA_START = HEADER_ROW_IVA + 1
for i, (pais, seg) in enumerate(PAISES):
    r = IVA_START + i
    ws_d.cell(row=r, column=1, value=pais).border = BORDER
    ws_d.cell(row=r, column=2, value=seg).border = BORDER
    c_iva = ws_d.cell(row=r, column=3, value=IVA_POR_PAIS[pais]); c_iva.number_format = "0.0%"; c_iva.border = BORDER
    c_fuente = ws_d.cell(row=r, column=4, value=FUENTE_IVA_POR_PAIS.get(pais, FUENTE_IVA_UE))
    c_fuente.font = Font(name=FONT, size=8, italic=True, color="4A3B2E")
    c_fuente.border = BORDER
IVA_END = IVA_START + len(PAISES) - 1
ws_d.column_dimensions["D"].width = 70

DEVOL_TITLE_ROW = IVA_END + 3
ws_d.cell(row=DEVOL_TITLE_ROW, column=1,
          value="Tasa de devolución real, por país — solo donde se confirmó con fuente nombrada; el resto queda sin dato").font = NEGRITA
devol_headers = ["País", "Segmento", "Tasa devolución real"]
HEADER_ROW_DEVOL = DEVOL_TITLE_ROW + 1
for c, h in enumerate(devol_headers, start=1):
    cell = ws_d.cell(row=HEADER_ROW_DEVOL, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER; cell.border = BORDER
# Solo los paises con tasa REAL confirmada y nombrada - Noruega/Finlandia/Holanda/
# Belgica/Luxemburgo se quedan fuera de esta lista a proposito (sin dato real),
# la formula de Presupuesto cae al promedio general cuando el pais no aparece aqui.
devol_reales = [
    ("Alemania", "DACH", 0.44), ("Austria", "DACH", 0.36), ("Suiza", "DACH", 0.45),
    ("Suecia", "Nordicos", 0.20), ("Dinamarca", "Nordicos", 0.30),
    ("España", "España", 0.237),  # ZigZag + Retail Economics, Informe Benchmark Anual de Devoluciones Espana 2025
]
DEVOL_START = HEADER_ROW_DEVOL + 1
for i, (pais, seg, tasa) in enumerate(devol_reales):
    r = DEVOL_START + i
    ws_d.cell(row=r, column=1, value=pais).border = BORDER
    ws_d.cell(row=r, column=2, value=seg).border = BORDER
    c_t = ws_d.cell(row=r, column=3, value=tasa); c_t.number_format = "0.0%"; c_t.border = BORDER
DEVOL_END = DEVOL_START + len(devol_reales) - 1

ws_d.column_dimensions["A"].width = 32
ws_d.column_dimensions["B"].width = 16
ws_d.column_dimensions["C"].width = 20

print("Hoja Datos_Base lista")

# ============================================================
# HOJA 3: Presupuesto (resultado, TODO con formulas, por pais)
# ============================================================
ws_r = wb.create_sheet("Presupuesto")
ws_r["A1"] = "Presupuesto / forecast a 2 años — LUMBRA, por país"
ws_r["A1"].font = TITULO
ws_r["A2"] = "Todo lo de abajo son fórmulas que leen Parametros y Datos_Base — cambia un supuesto ahí y esto se recalcula solo."
ws_r["A2"].font = SUBTITULO
ws_r["A3"] = ("IMPORTANTE: 'Margen de contribución' descuenta coste de producto, IVA y devoluciones — todavía NO descuenta "
              "gastos comerciales, marketing, ni generales (ver hoja Capital_Inversion). No es el beneficio neto final.")
ws_r["A3"].font = Font(name=FONT, size=9, bold=True, color="B8703F")

headers_r = ["Segmento", "País", "Año", "Unidades", "% Online", "IVA", "Tasa devolución", "Tipo de dato (devolución)",
             "PVP prom. (€)", "Coste prom. (€)", "Precio wholesale (€)",
             "Ingreso bruto online (€)", "Ingreso bruto presencial (€)", "Ingreso bruto total (€)",
             "IVA a remitir (€)", "Ingreso neto (€)", "Coste producción total (€)",
             "Unidades devueltas", "Ingreso perdido devolución (€)", "Coste logístico devolución (€)",
             "Margen de contribución (€)", "Ticket medio (€)",
             "% Margen de contribución (s/ Ingreso neto)", "% Coste logístico devolución (s/ Ingreso neto)"]
HEADER_ROW = 5
for c, h in enumerate(headers_r, start=1):
    cell = ws_r.cell(row=HEADER_ROW, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = BORDER

PARAM_RANGE_PAIS = f"Parametros!$B${PARAM_START_ROW}:$B${PARAM_END_ROW}"
PARAM_RANGE_UNID = f"Parametros!$C${PARAM_START_ROW}:$C${PARAM_END_ROW}"
PARAM_RANGE_CREC = f"Parametros!$D${PARAM_START_ROW}:$D${PARAM_END_ROW}"
PARAM_RANGE_ONLINE = f"Parametros!$E${PARAM_START_ROW}:$E${PARAM_END_ROW}"

from openpyxl.utils import get_column_letter as GCL
# Columnas con nombre, para que insertar/mover una columna nunca desalinee una
# formula por una letra escrita a mano.
COL = {"segmento":1,"pais":2,"anio":3,"unidades":4,"online":5,"iva":6,"devol":7,
       "tipo_devol":8,"pvp":9,"coste":10,"wholesale":11,"ing_online":12,"ing_presencial":13,
       "ing_bruto_total":14,"iva_remitir":15,"ing_neto":16,"coste_prod":17,"uds_devueltas":18,
       "ing_perdido":19,"coste_log_devol":20,"margen":21,"ticket":22,
       "margen_pct":23,"coste_log_pct":24}
L = {k: GCL(v) for k, v in COL.items()}

DATA_START = HEADER_ROW + 1
row = DATA_START
PRIMERA_FILA_POR_PAIS = {}
for pais, seg in PAISES:
    PRIMERA_FILA_POR_PAIS[pais] = row
    for anio in [1, 2]:
        r = row
        ws_r.cell(row=r, column=COL["segmento"], value=seg).border = BORDER
        ws_r.cell(row=r, column=COL["pais"], value=pais).border = BORDER
        ws_r.cell(row=r, column=COL["anio"], value=anio).border = BORDER

        unidades_anio1 = f"INDEX({PARAM_RANGE_UNID},MATCH(${L['pais']}{r},{PARAM_RANGE_PAIS},0))"
        crecimiento = f"INDEX({PARAM_RANGE_CREC},MATCH(${L['pais']}{r},{PARAM_RANGE_PAIS},0))"
        f_unid = f"={unidades_anio1}" if anio == 1 else f"={unidades_anio1}*(1+{crecimiento})"
        ws_r.cell(row=r, column=COL["unidades"], value=f_unid).border = BORDER

        f_online = f"=INDEX({PARAM_RANGE_ONLINE},MATCH(${L['pais']}{r},{PARAM_RANGE_PAIS},0))"
        ws_r.cell(row=r, column=COL["online"], value=f_online).number_format = "0.0%"
        ws_r.cell(row=r, column=COL["online"]).border = BORDER

        f_iva = f"=VLOOKUP(${L['pais']}{r},Datos_Base!$A${IVA_START}:$C${IVA_END},3,FALSE)"
        ws_r.cell(row=r, column=COL["iva"], value=f_iva).number_format = "0.0%"
        ws_r.cell(row=r, column=COL["iva"]).border = BORDER

        f_devol = (f"=IFERROR(VLOOKUP(${L['pais']}{r},Datos_Base!$A${DEVOL_START}:$C${DEVOL_END},3,FALSE),"
                   f"AVERAGE(Datos_Base!$C${DEVOL_START}:$C${DEVOL_END}))")
        ws_r.cell(row=r, column=COL["devol"], value=f_devol).number_format = "0.0%"
        ws_r.cell(row=r, column=COL["devol"]).border = BORDER

        # Nueva columna: deja explicito si la tasa de devolucion es real (esta en la
        # lista nombrada de Datos_Base) o el promedio general de referencia (fallback).
        f_tipo_devol = (f'=IFERROR(IF(VLOOKUP(${L["pais"]}{r},Datos_Base!$A${DEVOL_START}:$C${DEVOL_END},3,FALSE)<>"","Real (fuente nombrada)"),'
                        f'"PROXY, no es dato de LUMBRA — rango real conocido en otros países: "&TEXT(MIN(Datos_Base!$C${DEVOL_START}:$C${DEVOL_END}),"0.0%")&"-"&TEXT(MAX(Datos_Base!$C${DEVOL_START}:$C${DEVOL_END}),"0.0%")&", promedio usado: "&TEXT(AVERAGE(Datos_Base!$C${DEVOL_START}:$C${DEVOL_END}),"0.0%"))')
        ws_r.cell(row=r, column=COL["tipo_devol"], value=f_tipo_devol).border = BORDER
        ws_r.cell(row=r, column=COL["tipo_devol"]).font = Font(name=FONT, size=9, italic=True)

        ws_r.cell(row=r, column=COL["pvp"], value=f"={CELDA_PVP_PROM}").border = BORDER
        ws_r.cell(row=r, column=COL["coste"], value=f"={CELDA_COSTE_PROM}").border = BORDER
        ws_r.cell(row=r, column=COL["wholesale"], value=f"={L['coste']}{r}*{CELDA_MULTIPLICADOR}").border = BORDER

        ws_r.cell(row=r, column=COL["ing_online"], value=f"={L['unidades']}{r}*{L['online']}{r}*{L['pvp']}{r}").border = BORDER
        ws_r.cell(row=r, column=COL["ing_presencial"], value=f"={L['unidades']}{r}*(1-{L['online']}{r})*{L['wholesale']}{r}").border = BORDER
        ws_r.cell(row=r, column=COL["ing_bruto_total"], value=f"={L['ing_online']}{r}+{L['ing_presencial']}{r}").border = BORDER

        ws_r.cell(row=r, column=COL["iva_remitir"], value=f"={L['ing_online']}{r}*{L['iva']}{r}/(1+{L['iva']}{r})").border = BORDER
        ws_r.cell(row=r, column=COL["ing_neto"], value=f"={L['ing_bruto_total']}{r}-{L['iva_remitir']}{r}").border = BORDER
        ws_r.cell(row=r, column=COL["coste_prod"], value=f"={L['unidades']}{r}*({L['coste']}{r}+{CELDA_COSTE_ADICIONAL})").border = BORDER

        ws_r.cell(row=r, column=COL["uds_devueltas"], value=f"={L['unidades']}{r}*{L['devol']}{r}").border = BORDER
        precio_neto_mix = f"({L['online']}{r}*({L['pvp']}{r}/(1+{L['iva']}{r}))+(1-{L['online']}{r})*{L['wholesale']}{r})"
        ws_r.cell(row=r, column=COL["ing_perdido"], value=f"={L['uds_devueltas']}{r}*{precio_neto_mix}").border = BORDER
        ws_r.cell(row=r, column=COL["coste_log_devol"], value=f"={L['uds_devueltas']}{r}*{CELDA_COSTE_LOGISTICO}").border = BORDER

        ws_r.cell(row=r, column=COL["margen"],
                  value=f"={L['ing_neto']}{r}-{L['ing_perdido']}{r}-{L['coste_prod']}{r}-{L['coste_log_devol']}{r}").border = BORDER
        ws_r.cell(row=r, column=COL["ticket"], value=f"={L['ing_bruto_total']}{r}/{L['unidades']}{r}").border = BORDER

        ws_r.cell(row=r, column=COL["margen_pct"], value=f"={L['margen']}{r}/{L['ing_neto']}{r}").border = BORDER
        ws_r.cell(row=r, column=COL["coste_log_pct"], value=f"={L['coste_log_devol']}{r}/{L['ing_neto']}{r}").border = BORDER

        for k in ["pvp","coste","wholesale","ing_online","ing_presencial","ing_bruto_total",
                  "iva_remitir","ing_neto","ing_perdido","coste_log_devol","margen","ticket"]:
            ws_r.cell(row=r, column=COL[k]).number_format = '#,##0.00" €"'
        for k in ["coste_prod","unidades","uds_devueltas"]:
            ws_r.cell(row=r, column=COL[k]).number_format = "#,##0"
        for k in ["margen_pct","coste_log_pct"]:
            ws_r.cell(row=r, column=COL[k]).number_format = "0.0%"

        if pais == "España":
            for c in range(1, len(headers_r) + 1):
                ws_r.cell(row=r, column=c).fill = PENDIENTE_FILL

        row += 1
ULTIMA_FILA_PRESUPUESTO = row - 1

for col_letter, width in zip("ABCDEFGHIJKLMNOPQRSTUVWX",
                              [11,11,6,10,10,8,12,22,11,11,13,15,16,14,12,13,15,12,15,14,13,12,15,17]):
    ws_r.column_dimensions[col_letter].width = width

print("Hoja Presupuesto lista —", row - DATA_START, "filas (11 países × 2 años)")

# ============================================================
# HOJA 4: Consolidado (suma de los 11 paises, + participacion por pais)
# ============================================================
ws_c = wb.create_sheet("Consolidado")
ws_c["A1"] = "Consolidado LUMBRA — los 11 países juntos"
ws_c["A1"].font = TITULO
ws_c["A2"] = "Suma directa de la hoja Presupuesto — ninguna cifra nueva, solo el total y el % que cada país representa."
ws_c["A2"].font = SUBTITULO

headers_c = ["Segmento", "País", "Ingreso neto Año 1 (€)", "Ingreso neto Año 2 (€)",
             "Margen de contribución Año 1 (€)", "Margen de contribución Año 2 (€)",
             "% Participación en ingreso (Año 1)", "% Participación en ingreso (Año 2)"]
HEADER_ROW_C = 4
for c, h in enumerate(headers_c, start=1):
    cell = ws_c.cell(row=HEADER_ROW_C, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = BORDER

C_START = HEADER_ROW_C + 1
for i, (pais, seg) in enumerate(PAISES):
    r = C_START + i
    fila_anio1 = PRIMERA_FILA_POR_PAIS[pais]
    fila_anio2 = fila_anio1 + 1
    ws_c.cell(row=r, column=1, value=seg).border = BORDER
    ws_c.cell(row=r, column=2, value=pais).border = BORDER
    ws_c.cell(row=r, column=3, value=f"=Presupuesto!{L['ing_neto']}{fila_anio1}").border = BORDER
    ws_c.cell(row=r, column=4, value=f"=Presupuesto!{L['ing_neto']}{fila_anio2}").border = BORDER
    ws_c.cell(row=r, column=5, value=f"=Presupuesto!{L['margen']}{fila_anio1}").border = BORDER
    ws_c.cell(row=r, column=6, value=f"=Presupuesto!{L['margen']}{fila_anio2}").border = BORDER
    for c in [3,4,5,6]:
        ws_c.cell(row=r, column=c).number_format = '#,##0.00" €"'
    if pais == "España":
        for c in range(1, 9):
            ws_c.cell(row=r, column=c).fill = PENDIENTE_FILL
    for c in range(1, 9):
        ws_c.cell(row=r, column=c).border = BORDER
C_END = C_START + len(PAISES) - 1

TOTAL_ROW = C_END + 1
ws_c.cell(row=TOTAL_ROW, column=2, value="TOTAL LUMBRA (11 países)").font = NEGRITA
for c in [3,4,5,6]:
    col = ws_c.cell(row=TOTAL_ROW, column=c, value=f"=SUM({ws_c.cell(row=C_START,column=c).coordinate}:{ws_c.cell(row=C_END,column=c).coordinate})")
    col.font = NEGRITA; col.number_format = '#,##0.00" €"'
for c in range(1, 9):
    ws_c.cell(row=TOTAL_ROW, column=c).fill = TOTAL_FILL
    ws_c.cell(row=TOTAL_ROW, column=c).border = BORDER

# % de participacion, ahora que ya existe la fila de Total para dividir
for i in range(len(PAISES)):
    r = C_START + i
    ws_c.cell(row=r, column=7, value=f"=C{r}/C{TOTAL_ROW}").number_format = "0.0%"
    ws_c.cell(row=r, column=8, value=f"=D{r}/D{TOTAL_ROW}").number_format = "0.0%"
    ws_c.cell(row=r, column=7).border = BORDER
    ws_c.cell(row=r, column=8).border = BORDER

for col_letter, width in zip("ABCDEFGH", [11,14,20,20,18,18,22,22]):
    ws_c.column_dimensions[col_letter].width = width

print("Hoja Consolidado lista —", len(PAISES), "países + 1 fila de total")

# ============================================================
# HOJA 5: Validacion (revisa que los datos de entrada sean completos y
# razonables, y que ningun calculo se haya roto por falta de dato)
# ============================================================
ws_v = wb.create_sheet("Validacion")
ws_v["A1"] = "Validación de datos — revisa esto antes de presentar el presupuesto"
ws_v["A1"].font = TITULO
ws_v["A2"] = "Cada fila revisa una condición sola. Si alguna dice 'Revisar', el número final puede estar mal o incompleto."
ws_v["A2"].font = SUBTITULO

headers_v = ["Revisión", "Resultado", "Detalle"]
HEADER_ROW_V = 4
for c, h in enumerate(headers_v, start=1):
    cell = ws_v.cell(row=HEADER_ROW_V, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER; cell.border = BORDER

checks = [
    ("Unidades Año 1 — sin huecos, todas mayores que 0",
     f'=IF(AND(COUNTBLANK(Parametros!C{PARAM_START_ROW}:C{PARAM_END_ROW})=0,COUNTIF(Parametros!C{PARAM_START_ROW}:C{PARAM_END_ROW},"<=0")=0),"✓ OK","⚠️ Revisar")',
     f'=COUNTBLANK(Parametros!C{PARAM_START_ROW}:C{PARAM_END_ROW})&" hueco(s), "&COUNTIF(Parametros!C{PARAM_START_ROW}:C{PARAM_END_ROW},"<=0")&" valor(es) en 0 o negativo"'),
    ("Crecimiento Año 2 — sin huecos, entre 0% y 300%",
     f'=IF(AND(COUNTBLANK(Parametros!D{PARAM_START_ROW}:D{PARAM_END_ROW})=0,COUNTIF(Parametros!D{PARAM_START_ROW}:D{PARAM_END_ROW},"<0")=0,COUNTIF(Parametros!D{PARAM_START_ROW}:D{PARAM_END_ROW},">3")=0),"✓ OK","⚠️ Revisar")',
     f'=COUNTBLANK(Parametros!D{PARAM_START_ROW}:D{PARAM_END_ROW})&" hueco(s), "&(COUNTIF(Parametros!D{PARAM_START_ROW}:D{PARAM_END_ROW},"<0")+COUNTIF(Parametros!D{PARAM_START_ROW}:D{PARAM_END_ROW},">3"))&" fuera de rango"'),
    ("% Canal Online — sin huecos, entre 0% y 100%",
     f'=IF(AND(COUNTBLANK(Parametros!E{PARAM_START_ROW}:E{PARAM_END_ROW})=0,COUNTIF(Parametros!E{PARAM_START_ROW}:E{PARAM_END_ROW},"<0")=0,COUNTIF(Parametros!E{PARAM_START_ROW}:E{PARAM_END_ROW},">1")=0),"✓ OK","⚠️ Revisar")',
     f'=COUNTBLANK(Parametros!E{PARAM_START_ROW}:E{PARAM_END_ROW})&" hueco(s), "&(COUNTIF(Parametros!E{PARAM_START_ROW}:E{PARAM_END_ROW},"<0")+COUNTIF(Parametros!E{PARAM_START_ROW}:E{PARAM_END_ROW},">1"))&" fuera de rango"'),
    ("Coste logístico de devolución — no está vacío ni en 0",
     f'=IF(AND(ISBLANK({CELDA_COSTE_LOGISTICO})=FALSE,{CELDA_COSTE_LOGISTICO}>0),"✓ OK","⚠️ Revisar")',
     '=""'),
    ("Hoja Presupuesto — ningún cálculo roto (error de fórmula)",
     f'=IF(SUMPRODUCT(--ISERROR(Presupuesto!D{DATA_START}:X{ULTIMA_FILA_PRESUPUESTO}))=0,"✓ OK","⚠️ Revisar")',
     f'=SUMPRODUCT(--ISERROR(Presupuesto!D{DATA_START}:X{ULTIMA_FILA_PRESUPUESTO}))&" celda(s) con error"'),
    ("Hoja Consolidado — ningún cálculo roto (error de fórmula)",
     f'=IF(SUMPRODUCT(--ISERROR(Consolidado!C{C_START}:H{TOTAL_ROW}))=0,"✓ OK","⚠️ Revisar")',
     f'=SUMPRODUCT(--ISERROR(Consolidado!C{C_START}:H{TOTAL_ROW}))&" celda(s) con error"'),
    ("% de participación en ingreso — suma 100% entre los 11 países",
     f'=IF(ROUND(SUM(Consolidado!G{C_START}:G{C_END}),3)=1,"✓ OK","⚠️ Revisar")',
     f'=TEXT(SUM(Consolidado!G{C_START}:G{C_END}),"0.0%")&" en vez de 100.0%"'),
]
VAL_START = HEADER_ROW_V + 1
for i, (nombre, f_resultado, f_detalle) in enumerate(checks):
    r = VAL_START + i
    ws_v.cell(row=r, column=1, value=nombre).border = BORDER
    c_res = ws_v.cell(row=r, column=2, value=f_resultado); c_res.border = BORDER
    c_res.alignment = Alignment(horizontal="center")
    ws_v.cell(row=r, column=3, value=f_detalle).border = BORDER
    ws_v.cell(row=r, column=3).font = Font(name=FONT, size=9, italic=True, color="4A3B2E")
VAL_END = VAL_START + len(checks) - 1

ESTADO_ROW = VAL_END + 2
ws_v.cell(row=ESTADO_ROW, column=1, value="ESTADO GENERAL").font = NEGRITA
f_estado = f'=IF(COUNTIF(B{VAL_START}:B{VAL_END},"⚠️ Revisar")=0,"✓ Validaciones básicas superadas — supuestos de volumen, coste y devolución todavía pendientes de confirmar con datos reales","⚠️ Hay "&COUNTIF(B{VAL_START}:B{VAL_END},"⚠️ Revisar")&" punto(s) que revisar antes de presentar")'
c_estado = ws_v.cell(row=ESTADO_ROW, column=2, value=f_estado)
c_estado.font = Font(name=FONT, size=11, bold=True)
for c in range(1, 4):
    ws_v.cell(row=ESTADO_ROW, column=c).fill = TOTAL_FILL
    ws_v.cell(row=ESTADO_ROW, column=c).border = BORDER

ws_v.column_dimensions["A"].width = 48
ws_v.column_dimensions["B"].width = 16
ws_v.column_dimensions["C"].width = 45

print("Hoja Validacion lista —", len(checks), "revisiones + estado general")

# ============================================================
# HOJA 6: Capital_Inversion (produccion + 15 categorias de gasto comercial/
# operativo, organizadas en 5 bloques - ninguna tiene dato real de LUMBRA salvo
# donde se indica, todas editables)
# ============================================================
ws_i = wb.create_sheet("Capital_Inversion")
ws_i["A1"] = "Capital de inversión necesario — 2 años"
ws_i["A1"].font = TITULO
ws_i["A2"] = ("Producción viene de la hoja Presupuesto (dato real del modelo). Todo lo demás son supuestos "
              "nuevos, marcados en azul/amarillo — ninguno es un dato confirmado de LUMBRA todavía. Algunas líneas "
              "sí usan un dato real como base (unidades, ingreso online/presencial) multiplicado por un % o "
              "coste unitario que tú defines.")
ws_i["A2"].font = SUBTITULO

# Totales reales del modelo, para usarlos como base de varias lineas de gasto
UNIDADES_TOTAL_ANIO = {a: f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={a})*Presupuesto!${L['unidades']}${DATA_START}:${L['unidades']}${ULTIMA_FILA_PRESUPUESTO})" for a in [1,2]}
ING_ONLINE_TOTAL_ANIO = {a: f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={a})*Presupuesto!${L['ing_online']}${DATA_START}:${L['ing_online']}${ULTIMA_FILA_PRESUPUESTO})" for a in [1,2]}
ING_PRESENCIAL_TOTAL_ANIO = {a: f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={a})*Presupuesto!${L['ing_presencial']}${DATA_START}:${L['ing_presencial']}${ULTIMA_FILA_PRESUPUESTO})" for a in [1,2]}

ws_i.cell(row=4, column=1, value="1. Producción (de la hoja Presupuesto — Coste producción total, 11 países)").font = NEGRITA
ws_i.cell(row=5, column=1, value="Año 1")
ws_i.cell(row=5, column=2, value=f"=SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}=1)*Presupuesto!${L['coste_prod']}${DATA_START}:${L['coste_prod']}${ULTIMA_FILA_PRESUPUESTO})").number_format = '#,##0.00" €"'
ws_i.cell(row=6, column=1, value="Año 2")
ws_i.cell(row=6, column=2, value=f"=SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}=2)*Presupuesto!${L['coste_prod']}${DATA_START}:${L['coste_prod']}${ULTIMA_FILA_PRESUPUESTO})").number_format = '#,##0.00" €"'
CELDA_PRODUCCION_A1, CELDA_PRODUCCION_A2 = "Capital_Inversion!$B$5", "Capital_Inversion!$B$6"

r = 8
ws_i.cell(row=r, column=1, value="2. Influencers").font = NEGRITA; r += 1
ws_i.cell(row=r, column=1, value="Nº influencers gifting (2 años) — perfil 5.000-25.000 seguidores, Alemania/Austria")
c = ws_i.cell(row=r, column=2, value=30); c.font=AZUL_INPUT; c.fill=AMARILLO; FILA_N_GIFTING=r; r+=1
ws_i.cell(row=r, column=1, value="Pares regalados por influencer")
c = ws_i.cell(row=r, column=2, value=1); c.font=AZUL_INPUT; c.fill=AMARILLO; FILA_PARES_GIFTING=r; r+=1
ws_i.cell(row=r, column=1, value="Coste del par regalado — promedio real Zueco/Barefoot (categorías de la campaña de octubre)")
ws_i.cell(row=r, column=2, value="=AVERAGE(Datos_Base!B5,Datos_Base!B8)").number_format = '#,##0.00" €"'; FILA_COSTE_PAR=r; r+=1
ws_i.cell(row=r, column=1, value="Influencers: productos de regalo — Total")
ws_i.cell(row=r, column=2, value=f"=B{FILA_N_GIFTING}*B{FILA_PARES_GIFTING}*B{FILA_COSTE_PAR}").font=NEGRITA
ws_i.cell(row=r, column=2).number_format = '#,##0.00" €"'; FILA_INFL_PRODUCTO=r; r+=1
ws_i.cell(row=r, column=1, value="Influencers: envío — coste de envío unitario (tu supuesto)")
c = ws_i.cell(row=r, column=2, value=8); c.font=AZUL_INPUT; c.fill=AMARILLO; FILA_ENVIO_UNIT=r; r+=1
ws_i.cell(row=r, column=1, value="Influencers: envío — Total")
ws_i.cell(row=r, column=2, value=f"=B{FILA_N_GIFTING}*B{FILA_PARES_GIFTING}*B{FILA_ENVIO_UNIT}").font=NEGRITA
ws_i.cell(row=r, column=2).number_format = '#,##0.00" €"'; FILA_INFL_ENVIO=r; r+=1
ws_i.cell(row=r, column=1, value="Nº influencers con colaboración REMUNERADA (2 años, tu supuesto — todavía no hay ninguno confirmado)")
c = ws_i.cell(row=r, column=2, value=5); c.font=AZUL_INPUT; c.fill=AMARILLO; FILA_N_PAGADOS=r; r+=1
ws_i.cell(row=r, column=1, value="Honorario medio por colaboración remunerada (tu supuesto)")
c = ws_i.cell(row=r, column=2, value=300); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_HONORARIO=r; r+=1
ws_i.cell(row=r, column=1, value="Influencers: honorarios o colaboración remunerada — Total")
ws_i.cell(row=r, column=2, value=f"=B{FILA_N_PAGADOS}*B{FILA_HONORARIO}").font=NEGRITA
ws_i.cell(row=r, column=2).number_format = '#,##0.00" €"'; FILA_INFL_HONORARIOS=r; r+=2

ws_i.cell(row=r, column=1, value="3. Marketing").font = NEGRITA; r += 1
ws_i.cell(row=r, column=1, value="Medios pagados / campañas digitales (2 años, tu supuesto)")
c = ws_i.cell(row=r, column=2, value=4000); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_MEDIOS_PAGADOS=r; r+=2

ws_i.cell(row=r, column=1, value="4. Viajes comerciales").font = NEGRITA; r += 1
ws_i.cell(row=r, column=1, value="Nº de viajes comerciales (2 años, tu supuesto)")
c = ws_i.cell(row=r, column=2, value=6); c.font=AZUL_INPUT; c.fill=AMARILLO; FILA_N_VIAJES=r; r+=1
ws_i.cell(row=r, column=1, value="Transporte — coste medio por viaje (tu supuesto)")
c = ws_i.cell(row=r, column=2, value=350); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_TRANSPORTE_UNIT=r; r+=1
ws_i.cell(row=r, column=1, value="Viajes comerciales: transporte — Total")
ws_i.cell(row=r, column=2, value=f"=B{FILA_N_VIAJES}*B{FILA_TRANSPORTE_UNIT}").font=NEGRITA
ws_i.cell(row=r, column=2).number_format='#,##0.00" €"'; FILA_VIAJE_TRANSPORTE=r; r+=1
ws_i.cell(row=r, column=1, value="Alojamiento — coste medio por viaje (tu supuesto)")
c = ws_i.cell(row=r, column=2, value=180); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_ALOJ_UNIT=r; r+=1
ws_i.cell(row=r, column=1, value="Viajes comerciales: alojamiento — Total")
ws_i.cell(row=r, column=2, value=f"=B{FILA_N_VIAJES}*B{FILA_ALOJ_UNIT}").font=NEGRITA
ws_i.cell(row=r, column=2).number_format='#,##0.00" €"'; FILA_VIAJE_ALOJ=r; r+=1
ws_i.cell(row=r, column=1, value="Comidas / transporte local — coste medio por viaje (tu supuesto)")
c = ws_i.cell(row=r, column=2, value=70); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_COMIDAS_UNIT=r; r+=1
ws_i.cell(row=r, column=1, value="Viajes comerciales: comidas / transporte local — Total")
ws_i.cell(row=r, column=2, value=f"=B{FILA_N_VIAJES}*B{FILA_COMIDAS_UNIT}").font=NEGRITA
ws_i.cell(row=r, column=2).number_format='#,##0.00" €"'; FILA_VIAJE_COMIDAS=r; r+=2

ws_i.cell(row=r, column=1, value="5. Comercial / Cliente").font = NEGRITA; r += 1
ws_i.cell(row=r, column=1, value="Reuniones con clientes / atención en showroom (2 años, tu supuesto)")
c = ws_i.cell(row=r, column=2, value=1500); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_REUNIONES=r; r+=1
ws_i.cell(row=r, column=1, value="Ferias comerciales / costes de showroom (tu supuesto — ferias reales del sector: MICAM, Who's Next, CIFF Copenhague)")
c = ws_i.cell(row=r, column=2, value=3000); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_FERIAS=r; r+=1
ws_i.cell(row=r, column=1, value="Comisiones de venta / agentes — % sobre el ingreso bruto PRESENCIAL (canal donde actúan los agentes)")
c = ws_i.cell(row=r, column=2, value=0.08); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format="0.0%"; FILA_PCT_COMISION_AGENTE=r; r+=1
ws_i.cell(row=r, column=1, value="Comisiones de venta / agentes — Total (2 años, sobre ingreso presencial real del modelo)")
f_comision_agente = f"=B{FILA_PCT_COMISION_AGENTE}*(({ING_PRESENCIAL_TOTAL_ANIO[1]})+({ING_PRESENCIAL_TOTAL_ANIO[2]}))"
ws_i.cell(row=r, column=2, value=f_comision_agente).font=NEGRITA; ws_i.cell(row=r, column=2).number_format='#,##0.00" €"'
FILA_COMISION_AGENTE=r; r+=2

ws_i.cell(row=r, column=1, value="6. Operaciones").font = NEGRITA; r += 1
ws_i.cell(row=r, column=1, value="Gestión de pedidos (fulfillment) / preparación / embalaje — € por unidad vendida (tu supuesto)")
c = ws_i.cell(row=r, column=2, value=1.5); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_PCT_FULFILLMENT=r; r+=1
ws_i.cell(row=r, column=1, value="Gestión de pedidos (fulfillment) — Total (2 años, sobre unidades reales del modelo)")
f_fulfillment = f"=B{FILA_PCT_FULFILLMENT}*(({UNIDADES_TOTAL_ANIO[1]})+({UNIDADES_TOTAL_ANIO[2]}))"
ws_i.cell(row=r, column=2, value=f_fulfillment).font=NEGRITA; ws_i.cell(row=r, column=2).number_format='#,##0.00" €"'
FILA_FULFILLMENT=r; r+=1
ws_i.cell(row=r, column=1, value="Comisiones de plataformas de pago / marketplaces — % sobre el ingreso bruto ONLINE")
c = ws_i.cell(row=r, column=2, value=0.03); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format="0.0%"; FILA_PCT_MARKETPLACE=r; r+=1
ws_i.cell(row=r, column=1, value="Comisiones de plataformas de pago / marketplaces — Total (2 años, sobre ingreso online real del modelo)")
f_marketplace = f"=B{FILA_PCT_MARKETPLACE}*(({ING_ONLINE_TOTAL_ANIO[1]})+({ING_ONLINE_TOTAL_ANIO[2]}))"
ws_i.cell(row=r, column=2, value=f_marketplace).font=NEGRITA; ws_i.cell(row=r, column=2).number_format='#,##0.00" €"'
FILA_MARKETPLACE=r; r+=1
ws_i.cell(row=r, column=1, value="Almacenamiento / depósito (2 años, tu supuesto)")
c = ws_i.cell(row=r, column=2, value=2400); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_ALMACEN=r; r+=1
ws_i.cell(row=r, column=1, value="Muestras / logística comercial (2 años, tu supuesto)")
c = ws_i.cell(row=r, column=2, value=800); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_MUESTRAS=r; r+=1
ws_i.cell(row=r, column=1, value="Otros gastos operativos comerciales (2 años, tu supuesto)")
c = ws_i.cell(row=r, column=2, value=1000); c.font=AZUL_INPUT; c.fill=AMARILLO; c.number_format='#,##0.00" €"'; FILA_OTROS=r; r+=2

# IMPORTANTE (corregido tras auditoria): se separan en DOS subtotales distintos,
# nunca se vuelven a sumar juntos fuera de esta hoja - sumarlos ambos en otro
# lado junto con Margen de Contribucion (que YA resta produccion) duplicaria el
# coste de produccion. Consolidado usa SOLO la fila "sin produccion" de abajo.
filas_a_sumar = [FILA_INFL_PRODUCTO, FILA_INFL_ENVIO, FILA_INFL_HONORARIOS, FILA_MEDIOS_PAGADOS,
                 FILA_VIAJE_TRANSPORTE, FILA_VIAJE_ALOJ, FILA_VIAJE_COMIDAS, FILA_REUNIONES, FILA_FERIAS,
                 FILA_COMISION_AGENTE, FILA_FULFILLMENT, FILA_MARKETPLACE, FILA_ALMACEN, FILA_MUESTRAS, FILA_OTROS]

GASTOS_SIN_PRODUCCION_ROW = r + 1
ws_i.cell(row=GASTOS_SIN_PRODUCCION_ROW, column=1,
          value=("OPEX — GASTOS COMERCIALES / MARKETING / OPERATIVOS — TOTAL, SIN PRODUCCIÓN (2 años) — "
                 "esta es la fila a usar para calcular el resultado final, la producción ya está en el Margen de Contribución")).font = Font(name=FONT, size=11, bold=True)
ws_i.cell(row=GASTOS_SIN_PRODUCCION_ROW, column=1).alignment = Alignment(wrap_text=True)
f_gastos_sin_prod = "=" + "+".join(f"B{fr}" for fr in filas_a_sumar)
c_gsp = ws_i.cell(row=GASTOS_SIN_PRODUCCION_ROW, column=2, value=f_gastos_sin_prod)
c_gsp.font = Font(name=FONT, size=12, bold=True); c_gsp.number_format = '#,##0.00" €"'
for c in range(1, 3):
    ws_i.cell(row=GASTOS_SIN_PRODUCCION_ROW, column=c).fill = TOTAL_FILL
    ws_i.cell(row=GASTOS_SIN_PRODUCCION_ROW, column=c).border = BORDER
CELDA_GASTOS_SIN_PRODUCCION = f"Capital_Inversion!$B${GASTOS_SIN_PRODUCCION_ROW}"

TOTAL_CAPITAL_ROW = GASTOS_SIN_PRODUCCION_ROW + 2
ws_i.cell(row=TOTAL_CAPITAL_ROW, column=1,
          value=("COBERTURA PRESUPUESTARIA TOTAL — PRODUCCIÓN + los gastos de arriba — 2 AÑOS "
                 "(uso distinto: esto es 'cuánto costaría todo si nada se cobrara todavía', NO el capital de caja "
                 "real — ver hoja Flujo_Mensual — y NO se debe restar de nuevo del Margen de Contribución)")).font = Font(name=FONT, size=11, bold=True)
ws_i.cell(row=TOTAL_CAPITAL_ROW, column=1).alignment = Alignment(wrap_text=True)
f_total_capital = f"={CELDA_PRODUCCION_A1}+{CELDA_PRODUCCION_A2}+{CELDA_GASTOS_SIN_PRODUCCION}"
c_totalcap = ws_i.cell(row=TOTAL_CAPITAL_ROW, column=2, value=f_total_capital)
c_totalcap.font = Font(name=FONT, size=12, bold=True); c_totalcap.number_format = '#,##0.00" €"'
for c in range(1, 3):
    ws_i.cell(row=TOTAL_CAPITAL_ROW, column=c).fill = TOTAL_FILL
    ws_i.cell(row=TOTAL_CAPITAL_ROW, column=c).border = BORDER

for rr in range(4, TOTAL_CAPITAL_ROW + 1):
    for c in [1, 2]:
        ws_i.cell(row=rr, column=c).border = BORDER

ws_i.column_dimensions["A"].width = 78
ws_i.column_dimensions["B"].width = 18

print("Hoja Capital_Inversion lista — 15 categorías en 5 bloques")

# Resultado Final, en Consolidado: Margen de Contribucion total (variable, por
# pais) MENOS Capital_Inversion total (fijo, a nivel de marca) - el "despues de
# todo" que faltaba. Se agrega aqui porque recien ahora existe TOTAL_CAPITAL_ROW.
RESULTADO_FINAL_ROW = TOTAL_ROW + 2
ws_c.cell(row=RESULTADO_FINAL_ROW, column=2,
          value="Margen de Contribución TOTAL (los 11 países, 2 años)").font = NEGRITA
f_margen_total = f"=E{TOTAL_ROW}+F{TOTAL_ROW}"
ws_c.cell(row=RESULTADO_FINAL_ROW, column=3, value=f_margen_total).number_format = '#,##0.00" €"'
ws_c.cell(row=RESULTADO_FINAL_ROW+1, column=2,
          value=("OPEX — Gastos comerciales/marketing/operativos TOTAL (SIN producción, ya está en el Margen de arriba)")).font = NEGRITA
ws_c.cell(row=RESULTADO_FINAL_ROW+1, column=3, value=f"={CELDA_GASTOS_SIN_PRODUCCION}").number_format = '#,##0.00" €"'
ws_c.cell(row=RESULTADO_FINAL_ROW+2, column=2, value="EBITDA (Margen de Contribución − OPEX)").font = Font(name=FONT, size=12, bold=True)
f_resultado_final = f"=C{RESULTADO_FINAL_ROW}-C{RESULTADO_FINAL_ROW+1}"
c_rf = ws_c.cell(row=RESULTADO_FINAL_ROW+2, column=3, value=f_resultado_final)
c_rf.font = Font(name=FONT, size=12, bold=True); c_rf.number_format = '#,##0.00" €"'
for rr in range(RESULTADO_FINAL_ROW, RESULTADO_FINAL_ROW+3):
    for c in [2, 3]:
        ws_c.cell(row=rr, column=c).border = BORDER
    if rr == RESULTADO_FINAL_ROW+2:
        ws_c.cell(row=rr, column=2).fill = TOTAL_FILL
        ws_c.cell(row=rr, column=3).fill = TOTAL_FILL

print("Consolidado: Resultado Final agregado (Margen total - Capital_Inversion total)")

# Reconciliacion en Validacion: el forecast de unidades vendidas + las
# unidades de gifting NO deberian superar el plan de compra (5.280,
# BuyPlan, Maestro Modelos) - si lo hacen, el
# forecast esta vendiendo mas de lo que se va a producir.
RECONCIL_ROW = ESTADO_ROW + 3
ws_v.cell(row=RECONCIL_ROW, column=1, value="Reconciliación con el plan de compra (BuyPlan)").font = NEGRITA
ws_v.cell(row=RECONCIL_ROW+1, column=1, value="Unidades del plan de compra (24 modelos activos, BuyPlan)")
ws_v.cell(row=RECONCIL_ROW+1, column=2, value=f"={CELDA_TOTAL_UDS_COMPRA}").number_format = "#,##0"
ws_v.cell(row=RECONCIL_ROW+2, column=1, value="Unidades que el forecast proyecta VENDER (Año 1 + Año 2, los 11 países)")
f_uds_forecast = f"=SUM(Presupuesto!${L['unidades']}${DATA_START}:${L['unidades']}${ULTIMA_FILA_PRESUPUESTO})"
ws_v.cell(row=RECONCIL_ROW+2, column=2, value=f_uds_forecast).number_format = "#,##0"
ws_v.cell(row=RECONCIL_ROW+3, column=1, value="Unidades de gifting a influencers (se regalan, no se venden — reducen inventario disponible igual)")
ws_v.cell(row=RECONCIL_ROW+3, column=2, value=f"=Capital_Inversion!B{FILA_N_GIFTING}*Capital_Inversion!B{FILA_PARES_GIFTING}").number_format = "#,##0"
ws_v.cell(row=RECONCIL_ROW+4, column=1, value="Total unidades que hacen falta (venta + gifting)")
ws_v.cell(row=RECONCIL_ROW+4, column=2, value=f"=B{RECONCIL_ROW+2}+B{RECONCIL_ROW+3}").font = NEGRITA
ws_v.cell(row=RECONCIL_ROW+4, column=2).number_format = "#,##0"
ws_v.cell(row=RECONCIL_ROW+5, column=1, value="Diferencia vs. plan de compra (negativo = el forecast pide MÁS de lo que hay comprado)")
f_diff = f"=B{RECONCIL_ROW+1}-B{RECONCIL_ROW+4}"
c_diff = ws_v.cell(row=RECONCIL_ROW+5, column=2, value=f_diff)
c_diff.font = Font(name=FONT, size=11, bold=True); c_diff.number_format = "#,##0"
for rr in range(RECONCIL_ROW, RECONCIL_ROW+6):
    for c in [1, 2]:
        ws_v.cell(row=rr, column=c).border = BORDER
    if rr == RECONCIL_ROW+5:
        ws_v.cell(row=rr, column=1).fill = TOTAL_FILL
        ws_v.cell(row=rr, column=2).fill = TOTAL_FILL

# Desglose Espana vs. resto (BuyPlan tampoco dice explicitamente si sus 5.280
# unidades contemplan Espana o no - mismo tipo de pregunta abierta que antes,
# ahora sobre el dato nuevo).
DESGLOSE_ESPANA_ROW = RECONCIL_ROW + 7
ws_v.cell(row=DESGLOSE_ESPANA_ROW, column=1,
          value="¿De dónde sale el déficit? Desglose España vs. el resto").font = NEGRITA
ws_v.cell(row=DESGLOSE_ESPANA_ROW+1, column=1, value="Unidades del forecast SOLO España (Año 1 + Año 2)")
f_uds_espana = f"=SUMPRODUCT((Presupuesto!$B${DATA_START}:$B${ULTIMA_FILA_PRESUPUESTO}=\"España\")*Presupuesto!${L['unidades']}${DATA_START}:${L['unidades']}${ULTIMA_FILA_PRESUPUESTO})"
ws_v.cell(row=DESGLOSE_ESPANA_ROW+1, column=2, value=f_uds_espana).number_format = "#,##0"
ws_v.cell(row=DESGLOSE_ESPANA_ROW+2, column=1, value="Unidades del forecast SIN España (los 10 países del estudio original)")
ws_v.cell(row=DESGLOSE_ESPANA_ROW+2, column=2, value=f"=B{RECONCIL_ROW+2}-B{DESGLOSE_ESPANA_ROW+1}").number_format = "#,##0"
ws_v.cell(row=DESGLOSE_ESPANA_ROW+3, column=1, value="Diferencia vs. plan de compra, SIN España (positivo = sobra margen)")
c_diff2 = ws_v.cell(row=DESGLOSE_ESPANA_ROW+3, column=2, value=f"=B{RECONCIL_ROW+1}-B{DESGLOSE_ESPANA_ROW+2}")
c_diff2.font = NEGRITA; c_diff2.number_format = "#,##0"
ws_v.cell(row=DESGLOSE_ESPANA_ROW+4, column=1,
          value=("AVISO: el maestro BuyPlan fija 5.280 uds (24 modelos), pero no especifica si eso contempla venta en España — "
                 "confirmar si son exclusivas de los 10 países originales o la producción total de la marca. Además, el plan "
                 "actual cuesta 195.580€ de coste de producción y el presupuesto de compra aprobado (OTB) es de 180.000€ — "
                 "un desfase de 15.580€ propio del plan de compra, no de España, que conviene revisar.")).font = Font(name=FONT, size=9, italic=True, color="B8703F")
ws_v.cell(row=DESGLOSE_ESPANA_ROW+4, column=1).alignment = Alignment(wrap_text=True)
for rr in range(DESGLOSE_ESPANA_ROW, DESGLOSE_ESPANA_ROW+4):
    for c in [1, 2]:
        ws_v.cell(row=rr, column=c).border = BORDER
    if rr == DESGLOSE_ESPANA_ROW+3:
        ws_v.cell(row=rr, column=1).fill = TOTAL_FILL
        ws_v.cell(row=rr, column=2).fill = TOTAL_FILL

print("Validacion: reconciliación de unidades vs. plan de compra agregada, con desglose España")

# ============================================================
# HOJA 7: Flujo_Mensual (TRANSPUESTA: cada columna es un mes, Ene-27 a Dic-28
# -- estacionalidad real + cuando se paga produccion vs cuando entra el dinero
# de ventas -> el verdadero capital necesario)
# ============================================================
ws_f = wb.create_sheet("Flujo_Mensual")
ws_f["A1"] = "Flujo de caja mensual — LUMBRA consolidado (los 11 países juntos)"
ws_f["A1"].font = TITULO
ws_f["A2"] = ("Reparte el total anual (de Consolidado) en 12 meses usando dos curvas separadas y editables: "
              "cuándo se vende (estacionalidad real de moda) y cuándo se paga a producción (antes del lanzamiento). "
              "Son supuestos, no datos históricos de LUMBRA — no ha vendido todavía.")
ws_f["A2"].font = SUBTITULO
ws_f["A3"] = ("El capital real que hace falta no es el coste total del año — es el punto más profundo del "
              "flujo acumulado (fila de abajo), porque la producción se paga antes de que entre el dinero de las ventas.")
ws_f["A3"].font = Font(name=FONT, size=9, bold=True, color="B8703F")

MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
MESES_ABREV = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
ANIOS_CALENDARIO = {1: 27, 2: 28}  # Año 1 = 2027, Año 2 = 2028
# ACTUALIZADO (29/sept/2026): ya NO es una curva inventada - BuyPlan Config
# confirma la temporada real: Inicio 01/07/2027, Fin 31/12/2027 (6 meses,
# jul-dic). Ventas en 0% de ene-jun, repartidas solo jul-dic (sigue siendo UN
# supuesto el REPARTO dentro de esos 6 meses - eso si no viene de BuyPlan).
ESTACIONALIDAD_VENTAS_DEFECTO = [0.00,0.00,0.00,0.00,0.00,0.00,0.08,0.10,0.15,0.20,0.27,0.20]
# Calendario de PAGO A PRODUCCION: tiene que estar listo antes del inicio de
# temporada (01/07/2027) - se concentra ene-jun, con mas peso cerca de la
# fecha de entrada. El reparto exacto dentro de esos 6 meses sigue siendo
# supuesto mio, pero ahora la VENTANA (ene-jun) es real, no inventada.
TIMING_PRODUCCION_DEFECTO = [0.05,0.10,0.20,0.25,0.25,0.15,0.00,0.00,0.00,0.00,0.00,0.00]

HEADER_ROW_F = 5
ws_f.cell(row=HEADER_ROW_F, column=1, value="Concepto").font = HEADER_FONT
ws_f.cell(row=HEADER_ROW_F, column=1).fill = GRIS_HEADER
ws_f.cell(row=HEADER_ROW_F, column=1).border = BORDER

# 24 columnas de mes, B en adelante
COL_MES_START = 2
meses_cols = []  # lista de (columna, anio, mes_idx)
col = COL_MES_START
for anio in [1, 2]:
    for mes_idx in range(12):
        meses_cols.append((col, anio, mes_idx))
        etiqueta = f"{MESES_ABREV[mes_idx]}-{ANIOS_CALENDARIO[anio]}"
        cell = ws_f.cell(row=HEADER_ROW_F, column=col, value=etiqueta)
        cell.font = HEADER_FONT; cell.fill = GRIS_HEADER; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")
        col += 1
COL_MES_END = col - 1
L_MES = {c: GCL(c) for c, _, _ in meses_cols}

# Totales anuales de referencia (de Presupuesto) - CORREGIDO tras auditoria:
# el ingreso usado aqui ya resta el impacto de las devoluciones (antes usaba
# Ingreso Neto SIN descontar devolucion, lo que inflaba la caja disponible).
CELDA_COSTE_PROD_TOTAL_ANIO = {}
CELDA_ING_NETO_POST_DEVOL_ANIO = {}
for anio in [1, 2]:
    CELDA_COSTE_PROD_TOTAL_ANIO[anio] = (f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})"
                                          f"*Presupuesto!${L['coste_prod']}${DATA_START}:${L['coste_prod']}${ULTIMA_FILA_PRESUPUESTO})")
    ing_neto_a = (f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})"
                  f"*Presupuesto!${L['ing_neto']}${DATA_START}:${L['ing_neto']}${ULTIMA_FILA_PRESUPUESTO})")
    ing_perdido_a = (f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})"
                      f"*Presupuesto!${L['ing_perdido']}${DATA_START}:${L['ing_perdido']}${ULTIMA_FILA_PRESUPUESTO})")
    CELDA_ING_NETO_POST_DEVOL_ANIO[anio] = f"(({ing_neto_a})-({ing_perdido_a}))"
CELDA_INGRESO_NETO_TOTAL = CELDA_ING_NETO_POST_DEVOL_ANIO
# CORREGIDO tras auditoria: los gastos LIGADOS A VENTA (comision de agentes,
# fulfillment, comision de marketplace) no deberian repartirse parejo - suben
# y bajan con las ventas del mes. El resto (ferias, viajes, medios, almacen,
# muestras, otros, reuniones) sigue en partes iguales por falta de un
# calendario real de cuando se pagan.
CELDA_GASTOS_LIGADOS_VENTA_TOTAL = f"(Capital_Inversion!B{FILA_COMISION_AGENTE}+Capital_Inversion!B{FILA_FULFILLMENT}+Capital_Inversion!B{FILA_MARKETPLACE})"
CELDA_GASTOS_FIJOS_TOTAL = f"({CELDA_GASTOS_SIN_PRODUCCION}-{CELDA_GASTOS_LIGADOS_VENTA_TOTAL})"
CELDA_GASTOS_FIJOS_MES = f"({CELDA_GASTOS_FIJOS_TOTAL})/24"

FILAS_F = {"estac": HEADER_ROW_F+1, "timing": HEADER_ROW_F+2, "unidades": HEADER_ROW_F+3,
           "ing_neto": HEADER_ROW_F+4, "pct_ing_neto": HEADER_ROW_F+5,
           "coste_prod": HEADER_ROW_F+6, "pct_coste_prod": HEADER_ROW_F+7,
           "gastos_fijos": HEADER_ROW_F+8, "gastos_venta": HEADER_ROW_F+9,
           "flujo_mes": HEADER_ROW_F+10, "flujo_acum": HEADER_ROW_F+11}

etiquetas_fila = {
    "estac": "% Estacionalidad ventas (tu supuesto)",
    "timing": "% Timing pago producción (tu supuesto)",
    "unidades": "Unidades vendidas del mes",
    "ing_neto": "Ingreso neto del mes, YA descontada la devolución (€)",
    "pct_ing_neto": "% del Ingreso neto (post-devolución) anual",
    "coste_prod": "Coste producción del mes (€)",
    "pct_coste_prod": "% del Coste de producción anual",
    "gastos_fijos": "Gastos fijos/puntuales del mes (ferias, viajes, medios...) — repartidos en partes iguales, tu supuesto",
    "gastos_venta": "Gastos ligados a venta del mes (comisión agentes, fulfillment, marketplace) — siguen el ritmo de ventas",
    "flujo_mes": "Flujo de caja del mes (€)",
    "flujo_acum": "Flujo acumulado (€)",
}
for key, r in FILAS_F.items():
    cell = ws_f.cell(row=r, column=1, value=etiquetas_fila[key])
    cell.font = NEGRITA if key in ("flujo_mes","flujo_acum") else Font(name=FONT, size=10)
    cell.border = BORDER

primera_col_anio = {1: COL_MES_START, 2: COL_MES_START + 12}
for idx, (colnum, anio, mes_idx) in enumerate(meses_cols):
    cl = L_MES[colnum]
    c_estac = ws_f.cell(row=FILAS_F["estac"], column=colnum, value=ESTACIONALIDAD_VENTAS_DEFECTO[mes_idx])
    c_estac.font = AZUL_INPUT; c_estac.fill = AMARILLO; c_estac.number_format = "0.0%"; c_estac.border = BORDER
    c_timing = ws_f.cell(row=FILAS_F["timing"], column=colnum, value=TIMING_PRODUCCION_DEFECTO[mes_idx])
    c_timing.font = AZUL_INPUT; c_timing.fill = AMARILLO; c_timing.number_format = "0.0%"; c_timing.border = BORDER

    total_unidades_anio = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['unidades']}${DATA_START}:${L['unidades']}${ULTIMA_FILA_PRESUPUESTO})"
    c = ws_f.cell(row=FILAS_F["unidades"], column=colnum, value=f"={total_unidades_anio}*{cl}{FILAS_F['estac']}")
    c.number_format = "#,##0"; c.border = BORDER

    c = ws_f.cell(row=FILAS_F["ing_neto"], column=colnum, value=f"={CELDA_INGRESO_NETO_TOTAL[anio]}*{cl}{FILAS_F['estac']}")
    c.number_format = '#,##0.00" €"'; c.border = BORDER
    c = ws_f.cell(row=FILAS_F["pct_ing_neto"], column=colnum, value=f"={cl}{FILAS_F['ing_neto']}/{CELDA_INGRESO_NETO_TOTAL[anio]}")
    c.number_format = "0.0%"; c.border = BORDER; c.font = Font(name=FONT, size=9, italic=True)

    c = ws_f.cell(row=FILAS_F["coste_prod"], column=colnum, value=f"={CELDA_COSTE_PROD_TOTAL_ANIO[anio]}*{cl}{FILAS_F['timing']}")
    c.number_format = '#,##0.00" €"'; c.border = BORDER
    c = ws_f.cell(row=FILAS_F["pct_coste_prod"], column=colnum, value=f"={cl}{FILAS_F['coste_prod']}/({CELDA_COSTE_PROD_TOTAL_ANIO[anio]})")
    c.number_format = "0.0%"; c.border = BORDER; c.font = Font(name=FONT, size=9, italic=True)

    c = ws_f.cell(row=FILAS_F["gastos_fijos"], column=colnum, value=f"={CELDA_GASTOS_FIJOS_MES}")
    c.number_format = '#,##0.00" €"'; c.border = BORDER
    # Gastos ligados a venta: usan la MISMA curva de estacionalidad que el ingreso,
    # sobre el total de los 2 años (no distingue Año 1/Año 2 porque estos gastos
    # ya son un total conjunto en Capital_Inversion).
    c = ws_f.cell(row=FILAS_F["gastos_venta"], column=colnum, value=f"={CELDA_GASTOS_LIGADOS_VENTA_TOTAL}*{cl}{FILAS_F['estac']}/2")
    c.number_format = '#,##0.00" €"'; c.border = BORDER

    c = ws_f.cell(row=FILAS_F["flujo_mes"], column=colnum,
                  value=f"={cl}{FILAS_F['ing_neto']}-{cl}{FILAS_F['coste_prod']}-{cl}{FILAS_F['gastos_fijos']}-{cl}{FILAS_F['gastos_venta']}")
    c.number_format = '#,##0.00" €"'; c.border = BORDER; c.font = NEGRITA

    if colnum == COL_MES_START:
        f_acum = f"={cl}{FILAS_F['flujo_mes']}"
    else:
        cl_prev = L_MES[colnum - 1]
        f_acum = f"={cl_prev}{FILAS_F['flujo_acum']}+{cl}{FILAS_F['flujo_mes']}"
    c = ws_f.cell(row=FILAS_F["flujo_acum"], column=colnum, value=f_acum)
    c.number_format = '#,##0.00" €"'; c.border = BORDER; c.font = NEGRITA

# Aviso si alguna de las dos curvas de estacionalidad no suma 100%, por año
CHECK_ROW_1 = FILAS_F["flujo_acum"] + 2
CHECK_ROW_2 = CHECK_ROW_1 + 1
col_a1_ini, col_a1_fin = L_MES[primera_col_anio[1]], L_MES[primera_col_anio[1]+11]
col_a2_ini, col_a2_fin = L_MES[primera_col_anio[2]], L_MES[primera_col_anio[2]+11]
ws_f.cell(row=CHECK_ROW_1, column=1, value="Suma % Estacionalidad ventas — Año 1 / Año 2 (debe dar 100% cada una)")
ws_f.cell(row=CHECK_ROW_1, column=2, value=f"=SUM({col_a1_ini}{FILAS_F['estac']}:{col_a1_fin}{FILAS_F['estac']})").number_format = "0.0%"
ws_f.cell(row=CHECK_ROW_1, column=3, value=f"=SUM({col_a2_ini}{FILAS_F['estac']}:{col_a2_fin}{FILAS_F['estac']})").number_format = "0.0%"
ws_f.cell(row=CHECK_ROW_2, column=1, value="Suma % Timing producción — Año 1 / Año 2 (debe dar 100% cada una)")
ws_f.cell(row=CHECK_ROW_2, column=2, value=f"=SUM({col_a1_ini}{FILAS_F['timing']}:{col_a1_fin}{FILAS_F['timing']})").number_format = "0.0%"
ws_f.cell(row=CHECK_ROW_2, column=3, value=f"=SUM({col_a2_ini}{FILAS_F['timing']}:{col_a2_fin}{FILAS_F['timing']})").number_format = "0.0%"

CAPITAL_REAL_ROW = CHECK_ROW_2 + 2
ws_f.cell(row=CAPITAL_REAL_ROW, column=1, value="CAPITAL REAL NECESARIO (el punto más profundo del flujo acumulado)").font = Font(name=FONT, size=12, bold=True)
col_flujo_ini, col_flujo_fin = L_MES[COL_MES_START], L_MES[COL_MES_END]
c_capreal = ws_f.cell(row=CAPITAL_REAL_ROW, column=2,
                       value=f'=IF(MIN({col_flujo_ini}{FILAS_F["flujo_acum"]}:{col_flujo_fin}{FILAS_F["flujo_acum"]})<0,-MIN({col_flujo_ini}{FILAS_F["flujo_acum"]}:{col_flujo_fin}{FILAS_F["flujo_acum"]}),0)')
c_capreal.font = Font(name=FONT, size=12, bold=True); c_capreal.number_format = '#,##0.00" €"'
for c in range(1, 3):
    ws_f.cell(row=CAPITAL_REAL_ROW, column=c).fill = TOTAL_FILL
    ws_f.cell(row=CAPITAL_REAL_ROW, column=c).border = BORDER
ws_f.cell(row=CAPITAL_REAL_ROW+1, column=1,
          value=("AVISO (actualizado 29/sept/2026): la VENTANA de venta (jul-dic) y de pago a producción (ene-jun) "
                 "ya viene de BuyPlan Config (Inicio/Fin de temporada real) — ya no es inventada. El REPARTO exacto "
                 "dentro de esos 6 meses, y el reparto igual de gastos comerciales entre los 24 meses, sigue siendo "
                 "supuesto mío. Este número ya descuenta devoluciones. Sigue asumiendo que ya hay inventario listo "
                 "para vender en el primer mes de venta sin haber pagado su producción dentro de este mismo "
                 "calendario — en la práctica probablemente haga falta producir antes, lo que subiría el capital "
                 "real todavía más.")).font = Font(name=FONT, size=9, italic=True, color="B8703F")
ws_f.cell(row=CAPITAL_REAL_ROW+1, column=1).alignment = Alignment(wrap_text=True)

ws_f.column_dimensions["A"].width = 34
for c, _, _ in meses_cols:
    ws_f.column_dimensions[GCL(c)].width = 11
ws_f.freeze_panes = ws_f.cell(row=HEADER_ROW_F+1, column=COL_MES_START).coordinate

print("Hoja Flujo_Mensual lista —", len(meses_cols), "meses, en columnas")

# ============================================================
# HOJA 8: Proyeccion_Conservadora (segunda vision, derivada del caso base
# con formulas -- no son numeros nuevos inventados, son los mismos ratios
# reales del caso base aplicados a un escenario mas duro)
# ============================================================
ws_pc = wb.create_sheet("Proyeccion_Conservadora")
ws_pc["A1"] = "Proyección conservadora — segunda visión, junto al Forecast base"
ws_pc["A1"].font = TITULO
ws_pc["A2"] = ("Todo aquí se DERIVA del caso base con fórmulas (mismo coste por unidad, mismo ingreso neto por "
               "unidad que ya salió del modelo) — solo se ajustan 3 supuestos: menos unidades, más devolución, "
               "más gasto comercial. No son cifras nuevas inventadas.")
ws_pc["A2"].font = SUBTITULO

ws_pc.cell(row=4, column=1, value="Ajustes conservadores (edita estos 3 supuestos)").font = NEGRITA
ws_pc.cell(row=5, column=1, value="Reducción de unidades vendidas (vs. caso base)")
c_reduccion = ws_pc.cell(row=5, column=2, value=0.20); c_reduccion.font = AZUL_INPUT; c_reduccion.fill = AMARILLO
c_reduccion.number_format = "0.0%"
ws_pc.cell(row=6, column=1, value="Puntos porcentuales adicionales de devolución (vs. caso base)")
c_pp_devol = ws_pc.cell(row=6, column=2, value=0.05); c_pp_devol.font = AZUL_INPUT; c_pp_devol.fill = AMARILLO
c_pp_devol.number_format = "0.0%"
ws_pc.cell(row=7, column=1, value="Incremento en gasto comercial/capital (vs. caso base)")
c_inc_comercial = ws_pc.cell(row=7, column=2, value=0.20); c_inc_comercial.font = AZUL_INPUT; c_inc_comercial.fill = AMARILLO
c_inc_comercial.number_format = "0.0%"
for r in range(5, 8):
    for c in [1, 2]:
        ws_pc.cell(row=r, column=c).border = BORDER
CELDA_REDUCCION, CELDA_PP_DEVOL, CELDA_INC_COMERCIAL = "Proyeccion_Conservadora!$B$5", "Proyeccion_Conservadora!$B$6", "Proyeccion_Conservadora!$B$7"

# Tabla comparativa Base vs Conservador, por año - todo formula, nada a mano
headers_pc = ["Indicador", "Año", "Caso Base", "Proyección Conservadora", "Diferencia"]
HEADER_ROW_PC = 10
for c, h in enumerate(headers_pc, start=1):
    cell = ws_pc.cell(row=HEADER_ROW_PC, column=c, value=h)
    cell.font = HEADER_FONT; cell.fill = GRIS_HEADER; cell.border = BORDER

filas_pc = []
r = HEADER_ROW_PC + 1
for anio in [1, 2]:
    unidades_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['unidades']}${DATA_START}:${L['unidades']}${ULTIMA_FILA_PRESUPUESTO})"
    ing_neto_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['ing_neto']}${DATA_START}:${L['ing_neto']}${ULTIMA_FILA_PRESUPUESTO})"
    coste_prod_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['coste_prod']}${DATA_START}:${L['coste_prod']}${ULTIMA_FILA_PRESUPUESTO})"
    uds_devueltas_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['uds_devueltas']}${DATA_START}:${L['uds_devueltas']}${ULTIMA_FILA_PRESUPUESTO})"
    ing_perdido_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['ing_perdido']}${DATA_START}:${L['ing_perdido']}${ULTIMA_FILA_PRESUPUESTO})"
    coste_log_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['coste_log_devol']}${DATA_START}:${L['coste_log_devol']}${ULTIMA_FILA_PRESUPUESTO})"
    margen_base = f"SUMPRODUCT((Presupuesto!$C${DATA_START}:$C${ULTIMA_FILA_PRESUPUESTO}={anio})*Presupuesto!${L['margen']}${DATA_START}:${L['margen']}${ULTIMA_FILA_PRESUPUESTO})"

    # Ratios implicitos del caso base (derivados, no supuestos nuevos)
    tasadevol_implicita = f"({uds_devueltas_base})/({unidades_base})"
    precioneto_implicito = f"({ing_perdido_base})/({uds_devueltas_base})"
    costeunit_implicito = f"({coste_prod_base})/({unidades_base})"
    ingresounit_implicito = f"({ing_neto_base})/({unidades_base})"

    uds_cons = f"({unidades_base})*(1-{CELDA_REDUCCION})"
    tasadevol_cons = f"({tasadevol_implicita})+{CELDA_PP_DEVOL}"
    ing_neto_cons = f"({uds_cons})*({ingresounit_implicito})"
    coste_prod_cons = f"({uds_cons})*({costeunit_implicito})"
    uds_devueltas_cons = f"({uds_cons})*({tasadevol_cons})"
    ing_perdido_cons = f"({uds_devueltas_cons})*({precioneto_implicito})"
    coste_log_cons = f"({uds_devueltas_cons})*{CELDA_COSTE_LOGISTICO}"
    margen_cons = f"({ing_neto_cons})-({coste_prod_cons})-({ing_perdido_cons})-({coste_log_cons})"

    filas_pc.append(("Unidades", anio, unidades_base, uds_cons, "#,##0"))
    filas_pc.append(("Ingreso neto (€)", anio, ing_neto_base, ing_neto_cons, '#,##0.00" €"'))
    filas_pc.append(("Margen de contribución (€)", anio, margen_base, margen_cons, '#,##0.00" €"'))

MARGEN_PC_ROWS = {}
for nombre, anio, f_base, f_cons, fmt in filas_pc:
    ws_pc.cell(row=r, column=1, value=nombre).border = BORDER
    ws_pc.cell(row=r, column=2, value=anio).border = BORDER
    c_b = ws_pc.cell(row=r, column=3, value=f"={f_base}"); c_b.number_format = fmt; c_b.border = BORDER
    c_c = ws_pc.cell(row=r, column=4, value=f"={f_cons}"); c_c.number_format = fmt; c_c.border = BORDER
    c_d = ws_pc.cell(row=r, column=5, value=f"=D{r}-C{r}"); c_d.number_format = fmt; c_d.border = BORDER
    if nombre.startswith("Margen"):
        MARGEN_PC_ROWS[anio] = r
    r += 1
ULTIMA_FILA_PC = r - 1

ws_pc.column_dimensions["A"].width = 28
ws_pc.column_dimensions["B"].width = 8
for col in "CDE":
    ws_pc.column_dimensions[col].width = 20

# Capital conservador (aplica el incremento de gasto comercial sobre el capital real del Flujo_Mensual)
CAP_PC_ROW = ULTIMA_FILA_PC + 2
ws_pc.cell(row=CAP_PC_ROW, column=1, value="Capital real necesario — Base vs. Conservador (aplica el % de gasto comercial extra)").font = NEGRITA
ws_pc.cell(row=CAP_PC_ROW+1, column=1, value="Caso Base (de Flujo_Mensual)")
ws_pc.cell(row=CAP_PC_ROW+1, column=2, value=f"=Flujo_Mensual!B{CAPITAL_REAL_ROW}").number_format = '#,##0.00" €"'
ws_pc.cell(row=CAP_PC_ROW+2, column=1, value="Proyección Conservadora")
ws_pc.cell(row=CAP_PC_ROW+2, column=2, value=f"=Flujo_Mensual!B{CAPITAL_REAL_ROW}*(1+{CELDA_INC_COMERCIAL})").number_format = '#,##0.00" €"'
for rr in [CAP_PC_ROW+1, CAP_PC_ROW+2]:
    for c in [1, 2]:
        ws_pc.cell(row=rr, column=c).border = BORDER

# NUEVO (punto 5 de la auditoria): el resultado final tambien debe reflejar el
# gasto comercial extra del escenario conservador, no solo el capital.
RESULTADO_PC_ROW = CAP_PC_ROW + 4
ws_pc.cell(row=RESULTADO_PC_ROW, column=1, value="EBITDA (Margen de Contribución − OPEX) — Base vs. Conservador").font = NEGRITA
ws_pc.cell(row=RESULTADO_PC_ROW+1, column=1, value="Gastos comerciales/operativos — Caso Base (Capital_Inversion, sin producción)")
ws_pc.cell(row=RESULTADO_PC_ROW+1, column=2, value=f"={CELDA_GASTOS_SIN_PRODUCCION}").number_format = '#,##0.00" €"'
ws_pc.cell(row=RESULTADO_PC_ROW+2, column=1, value="Gastos comerciales/operativos — Proyección Conservadora (aplica el % de incremento)")
ws_pc.cell(row=RESULTADO_PC_ROW+2, column=2, value=f"={CELDA_GASTOS_SIN_PRODUCCION}*(1+{CELDA_INC_COMERCIAL})").number_format = '#,##0.00" €"'
_fila_a1, _fila_a2 = MARGEN_PC_ROWS[1], MARGEN_PC_ROWS[2]
ws_pc.cell(row=RESULTADO_PC_ROW+3, column=1, value="Margen de Contribución TOTAL (2 años) — Caso Base")
f_margen_total_base_pc = f"=C{_fila_a1}+C{_fila_a2}"
ws_pc.cell(row=RESULTADO_PC_ROW+3, column=2, value=f_margen_total_base_pc).number_format = '#,##0.00" €"'
ws_pc.cell(row=RESULTADO_PC_ROW+4, column=1, value="Margen de Contribución TOTAL (2 años) — Proyección Conservadora")
f_margen_total_cons_pc = f"=D{_fila_a1}+D{_fila_a2}"
ws_pc.cell(row=RESULTADO_PC_ROW+4, column=2, value=f_margen_total_cons_pc).number_format = '#,##0.00" €"'
ws_pc.cell(row=RESULTADO_PC_ROW+5, column=1, value="EBITDA — Caso Base").font = Font(name=FONT, size=11, bold=True)
ws_pc.cell(row=RESULTADO_PC_ROW+5, column=2, value=f"=B{RESULTADO_PC_ROW+3}-B{RESULTADO_PC_ROW+1}").font = Font(name=FONT, size=11, bold=True)
ws_pc.cell(row=RESULTADO_PC_ROW+5, column=2).number_format = '#,##0.00" €"'
ws_pc.cell(row=RESULTADO_PC_ROW+6, column=1, value="EBITDA — Proyección Conservadora").font = Font(name=FONT, size=11, bold=True)
ws_pc.cell(row=RESULTADO_PC_ROW+6, column=2, value=f"=B{RESULTADO_PC_ROW+4}-B{RESULTADO_PC_ROW+2}").font = Font(name=FONT, size=11, bold=True)
ws_pc.cell(row=RESULTADO_PC_ROW+6, column=2).number_format = '#,##0.00" €"'
for rr in range(RESULTADO_PC_ROW+1, RESULTADO_PC_ROW+7):
    for c in [1, 2]:
        ws_pc.cell(row=rr, column=c).border = BORDER
    if rr >= RESULTADO_PC_ROW+5:
        ws_pc.cell(row=rr, column=1).fill = TOTAL_FILL
        ws_pc.cell(row=rr, column=2).fill = TOTAL_FILL

print("Hoja Proyeccion_Conservadora lista — con resultado final por escenario y referencia de capital corregida")

# Grafico de barras nativo en Excel, comparando Base vs Conservador (Margen de
# contribucion, los dos anios) - para que se entienda de un vistazo, sin
# depender de que alguien abra Power BI.
chart = BarChart()
chart.type = "col"
chart.title = "Margen de contribución — Base vs. Proyección Conservadora"
chart.y_axis.title = "€"
chart.x_axis.title = "Año"
# Solo las 2 filas de Margen de contribucion (Año 1 y Año 2) - no mezclar con
# Unidades ni Ingreso, que tienen escalas muy distintas.
fila_margen_min = min(MARGEN_PC_ROWS.values())
fila_margen_max = max(MARGEN_PC_ROWS.values())
data_vals = Reference(ws_pc, min_col=3, max_col=4, min_row=fila_margen_min, max_row=fila_margen_max)
chart.add_data(data_vals, titles_from_data=False)
chart.set_categories(Reference(ws_pc, min_col=2, min_row=fila_margen_min, max_row=fila_margen_max))
chart.series[0].tx = openpyxl.chart.series.SeriesLabel(v="Caso Base")
chart.series[1].tx = openpyxl.chart.series.SeriesLabel(v="Proyección Conservadora")
chart.height = 8
chart.width = 16
ws_pc.add_chart(chart, f"A{CAP_PC_ROW+5}")

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output" / "presupuesto"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
wb.save(str(OUTPUT_DIR / "LUMBRA_Presupuesto_2anios.xlsx"))
print("Guardado.")
