"""
build_dashboard.py (v2)
------------------------
Genera docs/index.html: dashboard interactivo autocontenido con tres pestañas
(Resumen, Simulador de producto, Explorar datos), gráficos curados que responden
preguntas concretas del estudio de mercado, un simulador simple de margen, y
exportación de gráficos (PNG) y tablas (CSV) para armar una presentación.

Uso:
    python src/build_dashboard.py
"""
import json
from pathlib import Path
import pandas as pd
import openpyxl

PROCESSED_PATH = Path(__file__).resolve().parent.parent / "data" / "processed" / "lumbra_datos_limpios.csv"
OUTPUT_HTML = Path(__file__).resolve().parent.parent / "docs" / "index.html"
BENCHMARK_XLSX = Path(__file__).resolve().parent.parent / "data" / "raw" / "LUMBRA_Benchmark_PVP_FW27_RANKING.xlsx"
BUYPLAN_XLSM = Path(__file__).resolve().parent.parent / "data" / "raw" / "BuyPlan_Maestro_Modelos_Lumbra.xlsx"

POBLACION_MILLONES = {
    "Suecia": 10.588, "Dinamarca": 5.993, "Noruega": 5.627, "Finlandia": 5.636,
    "Alemania": 83.577, "Austria": 9.197, "Suiza": 9.100, "Holanda": 18.044,
    "Belgica": 11.900, "Luxemburgo": 0.682,
}

SCATTER_DATA = [
    {"pais": "Suecia", "segmento": "Nordicos", "tamano_usd": 15590, "tamano_label": "Mercado total eCommerce (2025, Statista)",
     "crecimiento_pct": 5.90, "crecimiento_label": "CAGR eCommerce total 2025-2030 (Statista)"},
    {"pais": "Dinamarca", "segmento": "Nordicos", "tamano_usd": 1680, "tamano_label": "Mercado Fashion eCommerce (2024, Statista)",
     "crecimiento_pct": 7.66, "crecimiento_label": "CAGR Fashion eCommerce 2024-2029 (Statista)"},
    {"pais": "Finlandia", "segmento": "Nordicos", "tamano_usd": 1610, "tamano_label": "Mercado Fashion eCommerce (2023, Statista)",
     "crecimiento_pct": 8.61, "crecimiento_label": "CAGR Fashion eCommerce 2023-2027 (Statista)"},
    {"pais": "Alemania", "segmento": "DACH", "tamano_usd": 17060, "tamano_label": "Mercado Luxury Goods (2024, Statista)",
     "crecimiento_pct": 1.93, "crecimiento_label": "CAGR Luxury Goods 2024-2029 (Statista)"},
    {"pais": "Austria", "segmento": "DACH", "tamano_usd": 11462.22, "tamano_label": "Mercado Apparel y Footwear (2025, Euromonitor)",
     "crecimiento_pct": 1.57, "crecimiento_label": "CAGR derivado de proyecciones Euromonitor 2025-2030"},
    {"pais": "Suiza", "segmento": "DACH", "tamano_usd": 13130, "tamano_label": "Mercado Apparel (2025, Statista)",
     "crecimiento_pct": 3.94, "crecimiento_label": "CAGR Apparel 2025-2029 (Statista)"},
    {"pais": "Belgica", "segmento": "Benelux", "tamano_usd": 2910, "tamano_label": "Mercado Luxury Goods (2024, Statista)",
     "crecimiento_pct": 2.23, "crecimiento_label": "CAGR Luxury Goods 2024-2029 (Statista)"},
    {"pais": "Luxemburgo", "segmento": "Benelux", "tamano_usd": 1170, "tamano_label": "Mercado Apparel proyectado (2029, Statista)",
     "crecimiento_pct": 1.24, "crecimiento_label": "CAGR Apparel 2024-2029 (Statista)"},
]


def cargar_datos_json() -> str:
    df = pd.read_csv(PROCESSED_PATH)
    df = df.where(pd.notnull(df), None)
    return json.dumps(df.to_dict(orient="records"), ensure_ascii=False)


# Multiplicador wholesale: precio de venta al retailer = coste x multiplicador.
# Valor ficticio de demostración. Cambiar SOLO este número: se inyecta solo en el
# dashboard, en el Simulador y en la pestaña España (no hace falta tocar el HTML/JS).
MULTIPLICADOR_WHOLESALE = 2.5

# Canal Presencial/B2B: comision del agente de ventas y coste medio de transporte por par
# (supuestos ficticios de demostracion). UNICA fuente: el
# Simulador de la web y la pestaña España leen estos dos valores, asi no pueden
# contradecirse entre si. Cambiar SOLO aqui si se actualizan.
COMISION_AGENTE_B2B_PCT = 12
COSTE_TRANSPORTE_B2B = 2.40

BRAND_COUNTRY = {
    "LUMBRA": "España", "Birkenstock": "Alemania", "UGG": "Estados Unidos",
    "Satorisan": "España", "Saye": "España", "Veja": "Francia",
    "Alohas": "España", "Vivobarefoot": "Reino Unido", "Be Lenka": "Eslovaquia",
    "Angulus": "Dinamarca", "INUIKII": "Suiza", "ATP Atelier": "Suecia",
    "Clarks": "Reino Unido", "Camper": "España", "Blundstone": "Australia",
    "Flattered": "Suecia", "Groundies": "Alemania", "Vagabond": "Suecia",
}

# Ciudades del estudio cualitativo original (Nordicos/DACH/Benelux), usadas como
# lista de referencia para verificar presencia fisica real de cada marca.
CIUDADES_ESTUDIO = [
    "Estocolmo", "Copenhague", "Oslo", "Helsinki", "Aarhus", "Malmo", "Gotemburgo",
    "Munich", "Berlin", "Dusseldorf", "Colonia", "Viena", "Zurich",
    "Amsterdam", "Bruselas", "Luxemburgo", "La Haya", "Amberes",
]

PAISES_ESTUDIO = [
    "Suecia", "Dinamarca", "Noruega", "Finlandia",
    "Alemania", "Austria", "Suiza",
    "Holanda", "Belgica", "Luxemburgo",
]

# Presencia fisica CONFIRMADA (tienda propia o revendedor con direccion real),
# verificada via el sitio oficial de cada marca o su localizador de tiendas.
# Lista vacia = no se encontro evidencia, NO significa que la marca no venda ahi.
BRAND_CIUDADES = {
    "Birkenstock": ["Bruselas"],
    "UGG": ["Berlin", "Munich"],
    "Veja": ["Copenhague", "Berlin"],
    "Clarks": ["Viena", "Copenhague"],
    "Camper": ["Copenhague"],
    "ATP Atelier": ["Estocolmo"],
    "INUIKII": ["Zurich"],
    "Flattered": ["Estocolmo"],
    "Blundstone": ["Copenhague", "Estocolmo"],
    "Be Lenka": ["Bruselas", "Amberes"],
    "Angulus": ["Copenhague"],
}

# Cobertura de venta online, en texto (la calidad del dato varia mucho por marca
# - algunas tienen selector de pais explicito, otras solo "envio a la UE" generico).
BRAND_COBERTURA_ONLINE = {
    "Birkenstock": "10 de 10 países (tienda oficial por país)",
    "UGG": "10 de 10 países (tienda oficial por país)",
    "Satorisan": "8 de 10 (sin Noruega ni Suiza)",
    "Saye": "Amplia en la UE (sin verificar Noruega/Suiza, no son UE)",
    "Veja": "Amplia (distribuida en 100+ países)",
    "Alohas": "Principalmente online, cobertura amplia sin detalle verificado",
    "Vivobarefoot": "Amplia en Europa, sin detalle país por país verificado",
    "Be Lenka": "Unión Europea y Estados Unidos",
    "Angulus": "Principalmente Dinamarca (marca doméstica)",
    "INUIKII": "Sin detalle verificado",
    "ATP Atelier": "Sin detalle verificado",
    "Clarks": "Amplia (tiendas propias multi-país)",
    "Camper": "Amplia (marca global)",
    "Blundstone": "Vía distribuidores locales",
    "Flattered": "Sin detalle verificado",
    "Groundies": "Amplia en la UE (envío, sin detalle país por país)",
    "Vagabond": "Amplia (marca líder en Suecia)",
}

# A que pais pertenece cada ciudad del estudio - para poder inferir cobertura de
# pais a partir de una presencia de tienda ya confirmada por ciudad.
CIUDAD_A_PAIS = {
    "Estocolmo": "Suecia", "Malmo": "Suecia", "Gotemburgo": "Suecia",
    "Copenhague": "Dinamarca", "Aarhus": "Dinamarca",
    "Oslo": "Noruega", "Helsinki": "Finlandia",
    "Munich": "Alemania", "Berlin": "Alemania", "Dusseldorf": "Alemania", "Colonia": "Alemania",
    "Viena": "Austria", "Zurich": "Suiza",
    "Amsterdam": "Holanda", "La Haya": "Holanda",
    "Bruselas": "Belgica", "Amberes": "Belgica", "Luxemburgo": "Luxemburgo",
}

# Cobertura de pais EXPLICITA por venta ONLINE, con su fuente. Zalando opera en
# los 10 paises del estudio (verificado: zalando.de/.at/.be/.dk/.fi/.lu/.nl/.no/
# .se/.ch) - si una marca esta confirmada ahi, se cuenta cobertura online en los
# 10, con la fuente citada. Las que tienen su propio selector de pais exacto
# (Birkenstock, UGG) usan esa lista en vez de la de Zalando.
TODOS_LOS_10 = ["Suecia","Dinamarca","Noruega","Finlandia","Alemania","Austria","Suiza","Holanda","Belgica","Luxemburgo"]

BRAND_ONLINE = {
    "Birkenstock": {"paises": TODOS_LOS_10, "fuente": "Tienda oficial propia por país"},
    "UGG": {"paises": TODOS_LOS_10, "fuente": "Tienda oficial propia por país"},
    "Veja": {"paises": TODOS_LOS_10, "fuente": "Distribución global propia (100+ países)"},
    "Clarks": {"paises": TODOS_LOS_10, "fuente": "Tienda oficial propia multi-país"},
    "Camper": {"paises": TODOS_LOS_10, "fuente": "Tienda oficial propia (marca global)"},
    "Satorisan": {"paises": TODOS_LOS_10, "fuente": "Tienda propia (8/10 países) + Zalando.es/.de"},
    "Saye": {"paises": TODOS_LOS_10, "fuente": "Tienda propia (UE) + Zalando.de/.es"},
    "Alohas": {"paises": TODOS_LOS_10, "fuente": "Zalando.de"},
    "Vagabond": {"paises": TODOS_LOS_10, "fuente": "Zalando.de"},
    "Vivobarefoot": {"paises": TODOS_LOS_10, "fuente": "Tienda propia + Zalando.de"},
    "Angulus": {"paises": TODOS_LOS_10, "fuente": "Zalando.de"},
    "Groundies": {"paises": TODOS_LOS_10, "fuente": "Zalando.de"},
    "INUIKII": {"paises": TODOS_LOS_10, "fuente": "Zalando.de"},
    "ATP Atelier": {"paises": ["Suecia","Dinamarca"], "fuente": "Retailer nórdico nobodesign.com"},
    "Flattered": {"paises": ["Suecia","Dinamarca"], "fuente": "Retailer nórdico nobodesign.com"},
    "Blundstone": {"paises": ["Suecia","Dinamarca"], "fuente": "Distribuidor oficial (Vernon Sports)"},
    "Be Lenka": {"paises": ["Belgica"], "fuente": "Tienda propia (envío UE/EEUU) + revendedor Barefoot Alliance"},
}

# Marcas donde la investigacion encontro evidencia EXPLICITA de que su presencia
# fisica es solo domestica (no es falta de busqueda, es un hallazgo real) - estas
# van en el nivel intermedio de la escala, no en "pendiente de verificar".
BRAND_FISICO_SIN_INDICIOS_INTL = {
    "Satorisan": "Confirmado: solo 2 tiendas propias, ambas en España (Valencia y Madrid)",
    "Alohas": "Confirmado: nació 100% online, recién abrió su primera tienda física",
}

def calcular_estado_canal(paises_o_ciudades_confirmadas, marca=None, sin_indicios_dict=None):
    """3 niveles: Confirmada (hay evidencia real) / Sin indicios de presencia
    internacional (se investigo y se encontro evidencia de que NO la tienen,
    un hallazgo real) / Pendiente de verificar (no se encontro evidencia en
    ningun sentido)."""
    if paises_o_ciudades_confirmadas:
        return "Confirmada"
    if sin_indicios_dict and marca in sin_indicios_dict:
        return "Sin indicios de presencia internacional"
    return "Pendiente de verificar"

def calcular_paises_online(marca):
    return BRAND_ONLINE.get(marca, {}).get("paises", [])

def calcular_paises_marca(marca):
    """Cobertura combinada (online O fisica) - usada por el filtro de pais, para
    que un competidor no desaparezca del filtro solo por canal."""
    online = set(calcular_paises_online(marca))
    ciudades = BRAND_CIUDADES.get(marca, [])
    fisico = set(CIUDAD_A_PAIS[c] for c in ciudades if c in CIUDAD_A_PAIS)
    return sorted(online | fisico)

# Cuota de mercado real encontrada - solo Birkenstock tuvo un dato especifico de
# uno de nuestros 10 paises. El resto de las marcas grandes solo tenian cifras
# globales o de mercados fuera de nuestro estudio (EEUU, Reino Unido) - no se
# fuerzan aqui.
BRAND_CUOTA_MERCADO = {
    "Birkenstock": {
        "texto": "11% de las ventas de sandalias en Alemania (2023)",
        "fuente": "Federación Alemana del Calzado (via Business Research Insights)",
    },
}

# Categoria de calzado dominante por pais (investigacion nueva) - fuente: Statista
# Market Insights, pagina especifica de cada pais (statista.com/outlook/cmo/footwear/[pais]).
# Varias cifras exactas estan bloqueadas por Statista salvo cuenta de pago - se deja
# el campo "cifra" vacio en esos casos. Luxemburgo no arrojo ningun dato pese a
# varios intentos de busqueda - se deja honestamente vacio.
BANDERA_PAIS = {
    "Alemania": "🇩🇪", "Suiza": "🇨🇭", "Dinamarca": "🇩🇰", "Austria": "🇦🇹",
    "Suecia": "🇸🇪", "Noruega": "🇳🇴", "Belgica": "🇧🇪", "Holanda": "🇳🇱",
    "Finlandia": "🇫🇮", "Luxemburgo": "🇱🇺",
}

# Codigo ISO 3166-1 alpha-2 por pais, para construir la URL de imagen de bandera
# via flagcdn.com (servicio publico gratuito) - usado solo en Power BI, donde no
# se puede incrustar el SVG hecho a mano como en la web.
ISO_PAIS = {
    "Alemania": "de", "Suiza": "ch", "Dinamarca": "dk", "Austria": "at",
    "Suecia": "se", "Noruega": "no", "Belgica": "be", "Holanda": "nl",
    "Finlandia": "fi", "Luxemburgo": "lu",
}
def url_bandera(pais):
    codigo = ISO_PAIS.get(pais)
    return f"https://flagcdn.com/w40/{codigo}.png" if codigo else ""


# El emoji de bandera falla en algunos Windows (cae al codigo de 2 letras en vez
# de la imagen) - se dibuja la bandera en SVG en su lugar, para que se vea igual
# en cualquier sistema sin depender de que el sistema tenga la fuente de emojis.
BANDERA_SVG = {
    "Alemania": '<svg width="20" height="13" viewBox="0 0 3 2"><rect width="3" height="2" fill="#000"/><rect y=".667" width="3" height=".667" fill="#DD0000"/><rect y="1.333" width="3" height=".667" fill="#FFCE00"/></svg>',
    "Suiza": '<svg width="14" height="14" viewBox="0 0 32 32"><rect width="32" height="32" fill="#D52B1E"/><rect x="13" y="6" width="6" height="20" fill="#fff"/><rect x="6" y="13" width="20" height="6" fill="#fff"/></svg>',
    "Dinamarca": '<svg width="19" height="14" viewBox="0 0 28 21"><rect width="28" height="21" fill="#C60C30"/><rect x="9" width="4" height="21" fill="#fff"/><rect y="8.5" width="28" height="4" fill="#fff"/></svg>',
    "Austria": '<svg width="20" height="13" viewBox="0 0 3 2"><rect width="3" height="2" fill="#ED2939"/><rect y=".667" width="3" height=".667" fill="#fff"/></svg>',
    "Suecia": '<svg width="20" height="13" viewBox="0 0 16 10"><rect width="16" height="10" fill="#006AA7"/><rect x="5" width="2" height="10" fill="#FECC00"/><rect y="4" width="16" height="2" fill="#FECC00"/></svg>',
    "Noruega": '<svg width="19" height="14" viewBox="0 0 22 16"><rect width="22" height="16" fill="#EF2B2D"/><rect x="6" width="6" height="16" fill="#fff"/><rect y="5" width="22" height="6" fill="#fff"/><rect x="7.5" width="3" height="16" fill="#002868"/><rect y="6.5" width="22" height="3" fill="#002868"/></svg>',
    "Belgica": '<svg width="20" height="13" viewBox="0 0 3 2"><rect width="1" height="2" fill="#000"/><rect x="1" width="1" height="2" fill="#FAE042"/><rect x="2" width="1" height="2" fill="#ED2939"/></svg>',
    "Holanda": '<svg width="20" height="13" viewBox="0 0 3 2"><rect width="3" height="2" fill="#21468B"/><rect width="3" height="1.333" fill="#fff"/><rect width="3" height=".667" fill="#AE1C28"/></svg>',
    "Finlandia": '<svg width="20" height="12" viewBox="0 0 18 11"><rect width="18" height="11" fill="#fff"/><rect x="5" width="3" height="11" fill="#003580"/><rect y="4" width="18" height="3" fill="#003580"/></svg>',
    "Luxemburgo": '<svg width="20" height="13" viewBox="0 0 3 2"><rect width="3" height=".667" fill="#ED2939"/><rect y=".667" width="3" height=".667" fill="#fff"/><rect y="1.333" width="3" height=".667" fill="#00A1DE"/></svg>',
}


# Marcas notables por categoria - Boots y Sneakers citan textual "key players" /
# "leading companies" de la propia Statista para ese segmento. Textile & Other
# Footwear (zuecos) no tuvo una cita textual equivalente de Statista - Crocs es
# la marca mas asociada a este segmento en varias fuentes independientes, pero
# con menos certeza que las otras dos, marcado como tal.
MARCAS_NOTABLES_POR_CATEGORIA = {
    "Boots": {"marcas": "Timberland, Dr. Martens, Caterpillar", "certeza": "alta",
              "fuente": "Statista Market Insights (\"key players\" del segmento Boots)"},
    "Sneakers": {"marcas": "Nike, Adidas, Skechers, New Balance", "certeza": "alta",
                 "fuente": "Statista Market Insights (\"leading companies\" del segmento Sneakers)"},
    "Textile & Other Footwear (incluye zuecos de madera)": {
        "marcas": "Crocs", "certeza": "media",
        "fuente": "Varias fuentes de la industria (no es una cita textual de Statista para este segmento)"},
}

CATEGORIA_DOMINANTE_POR_PAIS = [
    {"pais": "Alemania", "categoria": "Boots", "cifra": "US$5,90bn (2025)", "excepcion": False},
    {"pais": "Suiza", "categoria": "Boots", "cifra": "US$593,8M (2025)", "excepcion": False},
    {"pais": "Dinamarca", "categoria": "Boots", "cifra": "US$468,4M (2024)", "excepcion": False},
    {"pais": "Austria", "categoria": "Boots", "cifra": "US$0,8bn (2024)", "excepcion": False},
    {"pais": "Suecia", "categoria": "Boots", "cifra": "Cifra bloqueada por la fuente", "excepcion": False},
    {"pais": "Noruega", "categoria": "Boots", "cifra": "Cifra bloqueada por la fuente", "excepcion": False},
    {"pais": "Belgica", "categoria": "Boots", "cifra": "Cifra bloqueada por la fuente", "excepcion": False},
    {"pais": "Holanda", "categoria": "Sneakers", "cifra": "Cifra bloqueada por la fuente", "excepcion": True},
    {"pais": "Finlandia", "categoria": "Textile & Other Footwear (incluye zuecos de madera)", "cifra": "US$0,28bn (2024)", "excepcion": True},
    {"pais": "Luxemburgo", "categoria": "Sin dato encontrado", "cifra": "", "excepcion": False},
]
for _fila in CATEGORIA_DOMINANTE_POR_PAIS:
    _fila["bandera"] = BANDERA_PAIS.get(_fila["pais"], "")
    _fila["bandera_svg"] = BANDERA_SVG.get(_fila["pais"], "")
    _fila["bandera_url"] = url_bandera(_fila["pais"])
    _marcas = MARCAS_NOTABLES_POR_CATEGORIA.get(_fila["categoria"])
    _fila["marcas_notables"] = _marcas["marcas"] if _marcas else ""
    _fila["marcas_fuente"] = _marcas["fuente"] if _marcas else ""

FUENTE_CATEGORIA_DOMINANTE = "Statista Market Insights (statista.com/outlook/cmo/footwear/[país])"

TENDENCIA_BOTAS_ALTAS = {
    "titulo": "Botas altas: la tendencia dominante de la temporada",
    "texto": "Múltiples publicaciones de moda (Who What Wear, Marie Claire, finales de 2025 / 2026) "
             "confirman que las botas altas (knee-high, hasta el muslo) son la tendencia dominante de "
             "Otoño/Invierno 2026 en Europa, desplazando a las botas cortas — una de esas notas incluso "
             "menciona a Vagabond y Camper (dos competidores ya presentes en nuestro benchmark) entre las "
             "marcas donde el público ya está comprando ese estilo.",
    "recomendacion": "El Botín plano actual de LUMBRA es bajo. Boots es la categoría de calzado más grande "
                      "en 7 de los 8 países donde se consiguió el dato, y la tendencia de temporada apunta "
                      "específicamente a botas altas — hay una oportunidad de producto real y respaldada en "
                      "datos para evaluar una versión más alta del Botín plano.",
}


def calcular_factibilidad(coste, pvp_reunion, minimo_mercado):
    """Regla confirmada: 'sólida' si el margen aguanta incluso al precio minimo
    de mercado (el escenario mas competido); 'ajustada' si solo aguanta al PVP
    de la reunion; 'No factible' si ni siquiera eso deja margen positivo."""
    if None in (coste, pvp_reunion, minimo_mercado):
        return None
    if (minimo_mercado - coste) > 0:
        return "Factible (sólida)"
    if (pvp_reunion - coste) > 0:
        return "Factible (ajustada)"
    return "No factible"


# ============================================================
# ESPAÑA — bloque nuevo (B2C, B2B, Alemania/Francia/Italia, competencia)
# ============================================================
ESPANA_KPIS = [
    {"indicador": "PIB per cápita", "valor": "$38.290", "anio": "2025", "fuente": "Statista"},
    {"indicador": "Apparel per cápita", "valor": "$531,90", "anio": "2024", "fuente": "Statista"},
    {"indicador": "Luxury Fashion per cápita", "valor": "$66,24", "anio": "2024", "fuente": "Statista"},
    {"indicador": "Fashion eCommerce", "valor": "$9.080M (CAGR 8,04%)", "anio": "2025", "fuente": "Statista"},
    {"indicador": "Luxury Footwear", "valor": "$0,87bn ($18,28 per cápita)", "anio": "2024", "fuente": "Statista"},
    {"indicador": "Athletic Footwear", "valor": "$0,59bn", "anio": "2023", "fuente": "Statista"},
    {"indicador": "Textile & Other Footwear", "valor": "$3,21bn", "anio": "2022", "fuente": "Statista"},
    {"indicador": "Tasa de devolución ONLINE", "valor": "23,7% (calzado: 27% del valor devuelto)", "anio": "2025", "fuente": "ZigZag + Retail Economics"},
    {"indicador": "Tasa de devolución TIENDA FÍSICA", "valor": "5,6%", "anio": "2025", "fuente": "ZigZag + Retail Economics"},
]
ESPANA_LUXURY_GOODS_NOTA = ("Mercado total de Luxury Goods (todas las categorías): ~$9.750M, proyectado desde 2020 — "
                             "fuente elegida: Statista, por coherencia con el resto del estudio (otras fuentes conflictivas: "
                             "IMARC $4.800M, Mordor Intelligence $7.070M, descartadas por consistencia metodológica).")

ESPANA_CANAL_RETAIL = [
    {"canal": "Cadenas especializadas", "pct": 0.41, "valor_m": 4554},
    {"canal": "Hipermercados/supermercados", "pct": 0.264, "valor_m": 2932},
    {"canal": "Outlets/factory store", "pct": 0.17, "valor_m": 1888},
    {"canal": "Online (en crecimiento)", "pct": 0.17, "valor_m": 1888},
    {"canal": "Multimarca (canal B2B tradicional, en caída)", "pct": 0.07, "valor_m": 777},
]
ESPANA_CANAL_FUENTE = "Acotex — El Comercio Textil en Cifras 2025 (mercado total 11.107M€); valores en M€ derivados (% × total)"

# Margen por categoria en Alemania (dato solido: IVA y devolucion reales) vs.
# Francia/Italia (devolucion ESTIMADA, sin confirmar por pais - ver aviso).
MARGEN_PAIS_EXPORTACION = [
    {"categoria": "Zueco", "alemania": 21.1, "francia": 39.2, "italia": 38.6},
    {"categoria": "Bamba / Zapatilla", "alemania": 18.0, "francia": 35.8, "italia": 35.2},
    {"categoria": "Zapato", "alemania": 13.0, "francia": 30.6, "italia": 29.9},
    {"categoria": "Barefoot", "alemania": 16.8, "francia": 34.5, "italia": 33.8},
    {"categoria": "Botín plano", "alemania": 10.7, "francia": 28.1, "italia": 27.4},
    {"categoria": "Bailarina", "alemania": 23.5, "francia": 41.4, "italia": 40.9},
]
EXPORTACION_RANKING = [
    {"pais": "Italia", "posicion": 1, "pct_valor": "20%", "crecimiento_2024": "+39%", "precio_medio_historico": "16,47€/par (2018)", "margen_certeza": "Real (IVA/devolución Alemania confirmados; Francia/Italia con devolución estimada)"},
    {"pais": "Francia", "posicion": 2, "pct_valor": None, "crecimiento_2024": "-6,2%", "precio_medio_historico": "14,74€/par (2018)", "margen_certeza": "Estimado — devolución 27,5% sin confirmar por país"},
    {"pais": "Alemania", "posicion": 3, "pct_valor": None, "crecimiento_2024": "+5,9% a +8,2%", "precio_medio_historico": "20,92€/par (2018) — el más alto de los 3", "margen_certeza": "Real — IVA 19% y devolución 44% confirmados"},
]
EXPORTACION_FUENTE = "Dirección General de Aduanas, vía Revista del Calzado (ranking y crecimiento 2024); precio medio histórico: Modaes.es (2018)"
EXPORTACION_NOTA_DESTACADA = ("Alemania es el 3er país que más compra calzado español, con el precio medio histórico más "
                               "alto de los tres (20,92€/par) — paga más por par aunque crezca menos en volumen que Italia.")

# Competencia confirmada en España - 19 marcas (reusa el pais de origen y la
# confirmacion ya investigada; Pikolinos/Hispanitas son de origen espanol).
BRAND_ESPANA_CONFIRMADO = {
    "Satorisan": "Origen España (obviamente presente)", "Saye": "Origen España (obviamente presente)",
    "Alohas": "Zalando España (zalando.es)", "Camper": "Origen España (obviamente presente)",
    "Birkenstock": "Zalando España", "UGG": "Zalando España", "Clarks": "Zalando España",
    "Groundies": "Zalando España", "Vagabond": "Zalando España", "Be Lenka": "Zalando España",
    "INUIKII": "Zalando España", "Veja": "Zalando España",
    "Vivobarefoot": "Tienda propia en español + Decathlon España",
    "Angulus": "Miinto España (es.miinto.com)",
    "Blundstone": "Lyst.es + Miinto España + tienda propia blundstone.es",
    "Flattered": "Miinto España", "ATP Atelier": "Lyst.es + Miinto España",
    "Pikolinos": "Marca española, 29+ tiendas físicas confirmadas",
    "Hispanitas": "Marca española (fabrica en Alicante/Portugal), confirmada en Zalando España",
}

# Competidores agregados el 29/sept/2026, a peticion del usuario, para el
# analisis de competencia especifico de Espana (Pikolinos, Hispanitas) - NO
# vienen del Excel maestro LUMBRA_Benchmark_PVP_FW27_RANKING.xlsx, se marcan
# explicitamente para que se puedan revisar/incorporar al Excel oficial
# cuando el usuario confirme los datos.
COMPETIDORES_ESPANA_NUEVOS = [
    {"marca": "Pikolinos", "modelo": "Betera (sandalia barefoot)", "categoria": "Barefoot",
     "pvp": 119.95, "pais_origen": "España",
     "nota": "Piel premium, certificado por Centro de Biomecánica UMANA. Presencia física: 29+ tiendas en España.",
     "fuente_url": "https://www.pikolinos.com/es-es/mujer/colecciones/barefoot/sandalias.html"},
    {"marca": "Hispanitas", "modelo": "Transitional Barefoot (mocasín)", "categoria": "Barefoot",
     "pvp": 119.90, "pais_origen": "España",
     "nota": "Piel + 20% goma reciclada, forro con polímero de maíz reciclado. Confirmada en Zalando España.",
     "fuente_url": "https://www.elcorteingles.es/moda-mujer/MP_0021725_HV264818-mocasines-planos-de-mujer-de-transicion-barefoot/"},
]


def cargar_categorias_info():
    """Coste y PVP vigente por categoria - una sola fuente de verdad
    (hoja 'Categorias FW27' del Excel de benchmark). Reutilizada por Competencia
    y por el bloque de Espana (margen por pais)."""
    wb = openpyxl.load_workbook(BENCHMARK_XLSX, data_only=True)
    ws_cat = wb["Categorias FW27"]
    categorias_info = {}
    for r in range(6, 12):
        cat = ws_cat.cell(row=r, column=4).value
        categorias_info[cat] = {
            "coste": ws_cat.cell(row=r, column=8).value,
            "pvp_reunion": ws_cat.cell(row=r, column=13).value,
        }
    return categorias_info


def cargar_benchmark_competencia():
    """Lee directo del Excel de benchmark de LUMBRA (pestañas 'Categorias FW27',
    'Benchmark PVP' y 'Analisis Prospeccion') - una sola fuente de verdad, no se
    copia a mano. Usa el PVP vigente, no el del planner."""
    wb = openpyxl.load_workbook(BENCHMARK_XLSX, data_only=True)

    categorias_info = cargar_categorias_info()

    ws_p = wb["Analisis Prospeccion"]
    resumen = []
    factibilidad_por_categoria = {}
    for r in range(6, 12):
        categoria = ws_p.cell(row=r, column=2).value
        info = categorias_info.get(categoria, {})
        pvp_reunion = info.get("pvp_reunion")
        coste = info.get("coste")
        minimo = ws_p.cell(row=r, column=5).value
        media = ws_p.cell(row=r, column=6).value
        maximo = ws_p.cell(row=r, column=7).value
        delta_media_eur = (pvp_reunion - media) if (pvp_reunion is not None and media) else None
        delta_media_pct = (delta_media_eur / media) if (delta_media_eur is not None and media) else None
        factibilidad = calcular_factibilidad(coste, pvp_reunion, minimo)
        factibilidad_por_categoria[categoria] = factibilidad
        resumen.append({
            "categoria": categoria,
            "pvp_lumbra": pvp_reunion,
            "n_comparables": ws_p.cell(row=r, column=4).value,
            "minimo": round(minimo, 2) if minimo is not None else None,
            "media": round(media, 2) if media is not None else None,
            "maximo": round(maximo, 2) if maximo is not None else None,
            "marcas_por_encima": ws_p.cell(row=r, column=8).value,
            "delta_media_eur": round(delta_media_eur, 2) if delta_media_eur is not None else None,
            "delta_media_pct": round(delta_media_pct, 4) if delta_media_pct is not None else None,
            "posicion": ws_p.cell(row=r, column=11).value,
            "lectura": ws_p.cell(row=r, column=12).value,
            "lectura_calculada_con_planner": False,
            "factibilidad": factibilidad,
        })

    ws_b = wb["Benchmark PVP"]
    filas = []
    for r in range(6, ws_b.max_row + 1):
        ref = ws_b.cell(row=r, column=2).value
        if not ref:
            continue
        categoria = ws_b.cell(row=r, column=3).value
        tipo = ws_b.cell(row=r, column=4).value
        marca = ws_b.cell(row=r, column=6).value
        info = categorias_info.get(categoria, {})
        pvp_reunion = info.get("pvp_reunion")

        # Para la fila propia de LUMBRA, usar el precio vigente (reunion), no el
        # planner que trae guardado el Excel en la columna de PVP zona euro.
        if tipo == "LUMBRA":
            pvp_zona_euro = pvp_reunion
        else:
            pvp_zona_euro = ws_b.cell(row=r, column=10).value

        delta_eur = (pvp_zona_euro - pvp_reunion) if (pvp_zona_euro is not None and pvp_reunion is not None) else None
        delta_pct = (delta_eur / pvp_reunion) if (delta_eur is not None and pvp_reunion) else None

        filas.append({
            "ref": ref,
            "categoria": categoria,
            "tipo": tipo,
            "bloque": ws_b.cell(row=r, column=5).value,
            "marca": marca,
            "pais_marca": BRAND_COUNTRY.get(marca, "Sin confirmar"),
            "ciudades": BRAND_CIUDADES.get(marca, []),
            "n_ciudades_confirmadas": len(BRAND_CIUDADES.get(marca, [])),
            "estado_fisico": calcular_estado_canal(BRAND_CIUDADES.get(marca, []), marca, BRAND_FISICO_SIN_INDICIOS_INTL),
            "estado_fisico_nota": BRAND_FISICO_SIN_INDICIOS_INTL.get(marca),
            "paises_online": calcular_paises_online(marca),
            "estado_online": calcular_estado_canal(calcular_paises_online(marca)),
            "online_fuente": BRAND_ONLINE.get(marca, {}).get("fuente"),
            "paises_venta": calcular_paises_marca(marca),
            "cobertura_online": BRAND_COBERTURA_ONLINE.get(marca, "Sin verificar"),
            "cuota_mercado": BRAND_CUOTA_MERCADO.get(marca, {}).get("texto"),
            "cuota_mercado_fuente": BRAND_CUOTA_MERCADO.get(marca, {}).get("fuente"),
            "modelo": ws_b.cell(row=r, column=7).value,
            "pvp_zona_euro": round(pvp_zona_euro, 2) if pvp_zona_euro is not None else None,
            "delta_eur": round(delta_eur, 2) if delta_eur is not None else None,
            "delta_pct": round(delta_pct, 4) if delta_pct is not None else None,
            "factibilidad": factibilidad_por_categoria.get(categoria),
            "fuente_url": ws_b.cell(row=r, column=22).value,
            "agregado_fuera_excel": False,
        })

    for c in COMPETIDORES_ESPANA_NUEVOS:
        info = categorias_info.get(c["categoria"], {})
        pvp_reunion = info.get("pvp_reunion")
        delta_eur = c["pvp"] - pvp_reunion if pvp_reunion is not None else None
        delta_pct = delta_eur / pvp_reunion if (delta_eur is not None and pvp_reunion) else None
        filas.append({
            "ref": f"ESP-{c['marca'][:3].upper()}", "categoria": c["categoria"], "tipo": "Competencia",
            "bloque": "España (agregado)", "marca": c["marca"], "pais_marca": c["pais_origen"],
            "ciudades": [], "n_ciudades_confirmadas": 0,
            "estado_fisico": "Confirmada" if c["marca"] == "Pikolinos" else "Pendiente de verificar",
            "estado_fisico_nota": c["nota"],
            "paises_online": ["España"], "estado_online": "Confirmada",
            "online_fuente": "Sitio oficial de la marca (verificado 29/sept/2026)",
            "paises_venta": ["España"], "cobertura_online": "Confirmada en España",
            "cuota_mercado": None, "cuota_mercado_fuente": None,
            "modelo": c["modelo"], "pvp_zona_euro": c["pvp"],
            "delta_eur": round(delta_eur, 2) if delta_eur is not None else None,
            "delta_pct": round(delta_pct, 4) if delta_pct is not None else None,
            "factibilidad": factibilidad_por_categoria.get(c["categoria"]),
            "fuente_url": c["fuente_url"],
            "agregado_fuera_excel": True,
        })

    return filas, resumen


# ============================================================
# ESPAÑA — B2C, B2B, y margen por país de exportación (Alemania/Francia/Italia)
# ============================================================

# B2C real, con fuente por cada indicador - Statista salvo donde se indica otra.
ESPANA_B2C = {
    "pib_per_capita": {"valor": 38290, "unidad": "USD", "anio": 2025, "fuente": "Statista"},
    "apparel_per_capita": {"valor": 531.90, "unidad": "USD", "anio": 2024, "fuente": "Statista"},
    "luxury_fashion_per_capita": {"valor": 66.24, "unidad": "USD", "anio": 2024, "fuente": "Statista"},
    "fashion_ecommerce": {"valor": 9080, "valor_eur": round(9080 / 1.1578, 0), "unidad": "USD millones", "anio": 2025, "cagr": 8.04, "fuente": "Statista (convertido a EUR, tipo de cambio BCE 02/09/2026)"},
    "luxury_goods_total": {"valor": 9750, "unidad": "USD millones (proyectado desde 2020)", "anio": 2025, "fuente": "Statista — decisión del equipo: se usa esta en vez de IMARC (4.800M) o Mordor (7.070M), por consistencia de fuente con el resto del estudio"},
    "devolucion_online": {"valor": 23.7, "unidad": "%", "anio": 2025, "fuente": "ZigZag + Retail Economics, Informe Benchmark Anual de Devoluciones España 2025"},
    "devolucion_fisica": {"valor": 5.6, "unidad": "%", "anio": 2025, "fuente": "ZigZag + Retail Economics, Informe Benchmark Anual de Devoluciones España 2025"},
    "calzado_pct_valor_devuelto": {"valor": 27, "unidad": "%", "anio": 2025, "fuente": "ZigZag + Retail Economics, Informe Benchmark Anual de Devoluciones España 2025"},
    "luxury_footwear": {"valor": 0.87, "unidad": "USD mil millones", "per_capita": 18.28, "cagr": 5.06, "anio": 2024, "fuente": "Statista"},
    "athletic_footwear": {"valor": 0.59, "unidad": "USD mil millones", "cagr": 4.06, "anio": 2023, "fuente": "Statista"},
    "textile_other_footwear": {"valor": 3.21, "unidad": "USD mil millones", "cagr": 4.10, "anio": 2022, "fuente": "Statista"},
}

# B2B real - estructura de canal (Acotex) y exportacion (Revista del Calzado /
# Direccion General de Aduanas, datos 2024).
ESPANA_CANAL_B2B = [
    {"canal": "Cadenas especializadas", "pct": 41.0},
    {"canal": "Hipermercados/supermercados", "pct": 26.4},
    {"canal": "Outlets/factory store", "pct": 17.0},
    {"canal": "Online (creciendo)", "pct": 17.0},
    {"canal": "Multimarca (B2B tradicional, cayendo)", "pct": 7.0},
]
ACOTEX_MERCADO_TOTAL_MEUR = 11107  # "El Comercio Textil en Cifras 2025"

# Ranking de exportacion de calzado espanol - Italia/Francia/Alemania, con
# precio historico (2018, unica comparacion directa encontrada) y devolucion
# (real para Alemania, ESTIMADA para Francia/Italia - sin dato propio
# confirmado, se marca explicitamente como tal, no se oculta).
EXPORT_RANKING_PAISES = [
    {"pais": "Italia", "posicion": 1, "pct_valor_exportado": 20.0, "crecimiento_2024_pct": 39.0,
     "precio_historico_2018": 16.47, "iva": 0.22, "devolucion": 0.275, "devolucion_tipo": "ESTIMADO — sin dato propio confirmado para Italia"},
    {"pais": "Francia", "posicion": 2, "pct_valor_exportado": None, "crecimiento_2024_pct": -6.2,
     "precio_historico_2018": 14.74, "iva": 0.20, "devolucion": 0.275, "devolucion_tipo": "ESTIMADO — sin dato propio confirmado para Francia"},
    {"pais": "Alemania", "posicion": 3, "pct_valor_exportado": None, "crecimiento_2024_pct": 7.05,
     "precio_historico_2018": 20.92, "iva": 0.19, "devolucion": 0.44, "devolucion_tipo": "Real (Statista 2022)"},
]
IVACE_PRECIO_EU_PAR = 29.0      # EU-origin footwear en Alemania, 2019, informe IVACE
IVACE_PRECIO_NOEU_PAR = 11.0    # terceros paises, mismo informe

CATEGORIAS_6 = ["Zueco", "Bamba / zapatilla", "Zapato", "Barefoot", "Botín plano", "Bailarina"]

def calcular_margen_pais(pais_info, coste_adicional=10, coste_logistico_devol=13):
    """Margen neto real por las 6 categorias de LUMBRA en un pais de exportacion
    (Alemania/Francia/Italia) - canal 100% online (zona euro), con el
    multiplicador wholesale solo sirve para el canal presencial, aqui no
    aplica. Misma formula que el Simulador: IVA, coste, devolucion."""
    categorias_info = cargar_categorias_info()
    iva = pais_info["iva"]
    devol = pais_info["devolucion"]
    filas = []
    for cat in CATEGORIAS_6:
        info = categorias_info.get(cat, {})
        coste = info.get("coste")
        pvp = info.get("pvp_reunion")
        if coste is None or pvp is None:
            continue
        ingreso_neto = pvp / (1 + iva)
        coste_total_unit = coste + coste_adicional
        ingreso_perdido = devol * ingreso_neto
        coste_devol = devol * coste_logistico_devol
        margen_neto = ingreso_neto - coste_total_unit - ingreso_perdido - coste_devol
        pct_margen = margen_neto / ingreso_neto if ingreso_neto else None
        filas.append({
            "categoria": cat, "pvp": pvp, "coste": coste,
            "margen_neto": round(margen_neto, 2), "pct_margen": round(pct_margen, 4) if pct_margen is not None else None,
        })
    return filas


def calcular_margen_espana_canales(coste_adicional=10, coste_logistico_devol=13):
    """Margen por categoria DENTRO de Espana, separando Online (devolucion real
    23.7%, ZigZag 2025) y Presencial/B2B multimarca (devolucion real 5.6%,
    misma fuente; usa MULTIPLICADOR_WHOLESALE, definido arriba)."""
    categorias_info = cargar_categorias_info()
    iva = 0.21  # Espana
    filas = []
    for cat in CATEGORIAS_6:
        info = categorias_info.get(cat, {})
        coste = info.get("coste")
        pvp = info.get("pvp_reunion")
        if coste is None or pvp is None:
            continue
        coste_total_unit = coste + coste_adicional
        # Online
        ingreso_neto_on = pvp / (1 + iva)
        devol_on = 0.237
        margen_on = ingreso_neto_on - coste_total_unit - devol_on * ingreso_neto_on - devol_on * coste_logistico_devol
        # Presencial/B2B (wholesale x MULTIPLICADOR_WHOLESALE, menos comision de agente y transporte)
        precio_wholesale = coste * MULTIPLICADOR_WHOLESALE
        devol_fis = 0.056
        comision = precio_wholesale * COMISION_AGENTE_B2B_PCT / 100
        margen_fis = (precio_wholesale - coste_total_unit - devol_fis * precio_wholesale
                      - devol_fis * coste_logistico_devol - comision - COSTE_TRANSPORTE_B2B)
        filas.append({
            "categoria": cat, "pvp": pvp, "coste": coste, "precio_wholesale": round(precio_wholesale, 2),
            "margen_online": round(margen_on, 2), "pct_margen_online": round(margen_on / ingreso_neto_on, 4),
            "margen_presencial": round(margen_fis, 2), "pct_margen_presencial": round(margen_fis / precio_wholesale, 4),
        })
    return filas


# ============================================================
# RANKING DE MODELOS — margen real por modelo de la coleccion FW27 y por los
# 11 paises del estudio. Fuente: BuyPlan, hoja
# "Maestro Modelos" (24 modelos activos reales, coste y PVP por modelo).
# Agregado 2/oct/2026, a peticion del usuario: saber que modelo deja mejor
# margen, a pesar de que Barefoot es el modelo insignia del lanzamiento.
# ============================================================
PAISES_11 = {
    "Suecia": {"iva": 0.25, "devolucion": 0.20, "tipo": "Real"},
    "Dinamarca": {"iva": 0.25, "devolucion": 0.30, "tipo": "Real"},
    "Noruega": {"iva": 0.25, "devolucion": None, "tipo": "Estimado"},
    "Finlandia": {"iva": 0.255, "devolucion": None, "tipo": "Estimado"},
    "Alemania": {"iva": 0.19, "devolucion": 0.44, "tipo": "Real"},
    "Austria": {"iva": 0.20, "devolucion": 0.36, "tipo": "Real"},
    "Suiza": {"iva": 0.081, "devolucion": 0.45, "tipo": "Real"},
    "Holanda": {"iva": 0.21, "devolucion": None, "tipo": "Estimado"},
    "Bélgica": {"iva": 0.21, "devolucion": None, "tipo": "Estimado"},
    "Luxemburgo": {"iva": 0.17, "devolucion": None, "tipo": "Estimado"},
    "España": {"iva": 0.21, "devolucion": 0.237, "tipo": "Real"},
}
_DEVOL_PROMEDIO_11 = sum(p["devolucion"] for p in PAISES_11.values() if p["devolucion"] is not None) / \
                     sum(1 for p in PAISES_11.values() if p["devolucion"] is not None)


def cargar_modelos_buyplan():
    """Lee los modelos activos reales de la coleccion FW27 desde el sistema de
    compras de LUMBRA (BuyPlan, hoja 'Maestro Modelos') - una sola fuente de
    verdad, no se copia a mano. Columnas reales: 1=Activo, 3=Familia,
    5=Modelo, 6=Coste, 10=PVP retail c/IVA, 11=Pares."""
    wb = openpyxl.load_workbook(BUYPLAN_XLSM, data_only=True, keep_vba=False)
    ws = wb["Maestro Modelos"]
    modelos = []
    for r in range(6, 65):
        if ws.cell(row=r, column=1).value == "Sí":
            modelos.append({
                "familia": ws.cell(row=r, column=3).value.replace("_", " ").title(),
                "modelo": ws.cell(row=r, column=5).value,
                "coste": ws.cell(row=r, column=6).value,
                "pvp": ws.cell(row=r, column=10).value,
                "pares": ws.cell(row=r, column=11).value,
            })
    return modelos


def agrupar_perfiles_economicos(modelos):
    """Varios modelos de BuyPlan comparten exactamente el mismo coste y PVP
    (ej. los 4 'Clog Nube A/B/C/D') - se agrupan en un solo perfil porque dan
    el mismo margen, en vez de repetir filas identicas."""
    perfiles = {}
    for m in modelos:
        key = (m["familia"], m["coste"], m["pvp"])
        if key not in perfiles:
            perfiles[key] = {"familia": m["familia"], "coste": m["coste"], "pvp": m["pvp"], "modelos": []}
        perfiles[key]["modelos"].append(m["modelo"])
    return list(perfiles.values())


def nombre_representativo_perfil(lista_modelos):
    """A partir de los nombres reales de BuyPlan de un perfil (ej. 'Clog Nube A',
    'Clog Nube B', 'Clog Nube C', 'Clog Nube D'), devuelve un nombre corto sin
    la letra de variante final, para no repetir la familia en el grafico (el
    problema que senalo el usuario: 'Bamba' o 'Botin Plano' se repetian varias
    veces como si fueran lo mismo). Si el perfil mezcla dos nombres base
    distintos (ej. Marea + Brisa), los une con '/'."""
    import re
    bases = []
    for m in lista_modelos:
        base = re.sub(r"\s+[A-Za-z]$", "", m).strip()  # quita " A", " B", " b", etc. al final
        bases.append(base)
    vistos = []
    for b in bases:
        if b not in vistos:
            vistos.append(b)
    return "/".join(vistos)


def calcular_matriz_margen_modelos(coste_adicional=10, coste_logistico_devol=13):
    """Margen neto (%) de cada perfil economico, en cada uno de los 11 paises
    del estudio - misma formula del Simulador (IVA, devolucion real o el
    promedio general marcado como estimado, coste adicional de
    aprovisionamiento). Devuelve tambien el promedio y el pais min/max de
    cada perfil, para el ranking general."""
    modelos = cargar_modelos_buyplan()
    perfiles = agrupar_perfiles_economicos(modelos)
    filas = []
    for p in perfiles:
        modelos_unicos = sorted(set(p["modelos"]))
        fila = {"familia": p["familia"], "coste": p["coste"], "pvp": p["pvp"],
                "modelos": ", ".join(modelos_unicos),
                "nombre_display": nombre_representativo_perfil(modelos_unicos)}
        pcts = []
        for pais, info in PAISES_11.items():
            devol = info["devolucion"] if info["devolucion"] is not None else _DEVOL_PROMEDIO_11
            ingreso_neto = p["pvp"] / (1 + info["iva"])
            coste_total = p["coste"] + coste_adicional
            margen = ingreso_neto - coste_total - devol * ingreso_neto - devol * coste_logistico_devol
            pct = margen / ingreso_neto
            fila[pais] = round(pct * 100, 1)
            pcts.append(pct * 100)
        fila["promedio_11_paises"] = round(sum(pcts) / len(pcts), 1)
        fila["min_pais"] = round(min(pcts), 1)
        fila["max_pais"] = round(max(pcts), 1)
        filas.append(fila)
    filas.sort(key=lambda f: -f["promedio_11_paises"])
    return filas


def construir_espana_margenes():
    """Margen por categoria para los 3 paises de exportacion (Alemania real,
    Francia/Italia con devolucion estimada) - usado en la pestana Espana."""
    resultado = {}
    for p in EXPORT_RANKING_PAISES:
        resultado[p["pais"]] = {
            "filas": calcular_margen_pais(p),
            "devolucion_tipo": p["devolucion_tipo"],
            "iva": p["iva"],
            "devolucion": p["devolucion"],
        }
    return resultado


def exportar_scatter_csv():
    """Exporta la misma curaduria del grafico de dispersion como CSV, para Power BI
    (o cualquier otra herramienta) - una sola fuente de verdad, no se copia a mano.
    Incluye USD y EUR ya calculados: Power BI no puede convertir sobre la marcha
    como hace el HTML, asi que ambas monedas viven en el archivo."""
    df_pib = pd.read_csv(PROCESSED_PATH)
    df_pib = df_pib[df_pib["Indicador"] == "PIB per capita nominal"][["Pais", "Valor USD"]]
    df_pib = df_pib.rename(columns={"Valor USD": "PibPerCapitaUSD"})

    df_scatter = pd.DataFrame(SCATTER_DATA).rename(columns={
        "pais": "Pais", "segmento": "Segmento", "tamano_usd": "TamanoUSD",
        "tamano_label": "TamanoLabel", "crecimiento_pct": "CrecimientoPct",
        "crecimiento_label": "CrecimientoLabel",
    })
    df_scatter = df_scatter.merge(df_pib, on="Pais", how="left")

    # Misma tasa EUR_USD que usa etl.py y el dashboard HTML (BCE, 02/09/2026)
    EUR_USD = 1.1578
    df_scatter["TamanoEUR"] = (df_scatter["TamanoUSD"] / EUR_USD).round(2)
    df_scatter["PibPerCapitaEUR"] = (df_scatter["PibPerCapitaUSD"] / EUR_USD).round(2)

    columnas_orden = ["Pais", "Segmento", "TamanoUSD", "TamanoEUR", "TamanoLabel",
                       "CrecimientoPct", "CrecimientoLabel", "PibPerCapitaUSD", "PibPerCapitaEUR"]
    df_scatter = df_scatter[columnas_orden]

    out_path = PROCESSED_PATH.parent / "lumbra_scatter_oportunidad.csv"
    df_scatter.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Tabla auxiliar del mapa de oportunidad exportada: {out_path}")


def exportar_espana_csv():
    """Exporta los datos de la pestana Espana como CSV, para que Power BI arme
    la misma pagina con los mismos numeros - igual patron que el resto del
    proyecto (una sola fuente de verdad en Python, nunca copiado a mano)."""
    out_dir = PROCESSED_PATH.parent

    pd.DataFrame(ESPANA_CANAL_B2B).to_csv(out_dir / "espana_canal_b2b.csv", index=False, encoding="utf-8-sig")

    pd.DataFrame(EXPORT_RANKING_PAISES).to_csv(out_dir / "espana_export_ranking.csv", index=False, encoding="utf-8-sig")

    filas_margen = []
    for pais, data in construir_espana_margenes().items():
        for f in data["filas"]:
            filas_margen.append({"Pais": pais, "DevolucionTipo": data["devolucion_tipo"],
                                  "Categoria": f["categoria"], "PVP": f["pvp"], "Coste": f["coste"],
                                  "MargenNeto": f["margen_neto"], "PctMargen": f["pct_margen"]})
    pd.DataFrame(filas_margen).to_csv(out_dir / "espana_margen_export.csv", index=False, encoding="utf-8-sig")

    pd.DataFrame(calcular_margen_espana_canales()).to_csv(out_dir / "espana_margen_canal.csv", index=False, encoding="utf-8-sig")

    # Tarjetas KPI de la pestaña (mismos valores y fuentes que la web)
    EUR_USD_CSV = 1.1578  # BCE, 02/09/2026 - igual que la web
    b2c = ESPANA_B2C
    kpis = [
        {"kpi": "fashion_ecommerce", "pregunta": "¿Qué tan grande es el mercado de moda online en España?",
         "valor_num": round(b2c["fashion_ecommerce"]["valor"] / EUR_USD_CSV, 0),
         "valor_texto": f"€{b2c['fashion_ecommerce']['valor'] / EUR_USD_CSV:,.0f}M",
         "detalle": f"Fashion eCommerce {b2c['fashion_ecommerce']['anio']}, CAGR {b2c['fashion_ecommerce']['cagr']}% — {b2c['fashion_ecommerce']['fuente']} (original: ${b2c['fashion_ecommerce']['valor']}M)"},
        {"kpi": "devolucion_online_vs_fisica", "pregunta": "¿Cuánto más se devuelve comprando online que en tienda?",
         "valor_num": round(b2c["devolucion_online"]["valor"] / b2c["devolucion_fisica"]["valor"], 1),
         "valor_texto": f"{b2c['devolucion_online']['valor'] / b2c['devolucion_fisica']['valor']:.1f}x",
         "detalle": f"{b2c['devolucion_online']['valor']}% online vs {b2c['devolucion_fisica']['valor']}% física — {b2c['devolucion_online']['fuente']}"},
        {"kpi": "calzado_pct_devuelto", "pregunta": "¿Qué parte de lo que se devuelve en España es calzado?",
         "valor_num": b2c["calzado_pct_valor_devuelto"]["valor"],
         "valor_texto": f"{b2c['calzado_pct_valor_devuelto']['valor']}%",
         "detalle": f"del valor total devuelto — {b2c['calzado_pct_valor_devuelto']['fuente']}"},
        {"kpi": "canal_multimarca", "pregunta": "¿Qué parte del mercado de moda sigue pasando por el canal mayorista tradicional?",
         "valor_num": 7.0, "valor_texto": "7%",
         "detalle": "≈777M€ del mercado total (11.107M€), y cayendo — Acotex 2025"},
    ]
    pd.DataFrame(kpis).to_csv(out_dir / "espana_kpis.csv", index=False, encoding="utf-8-sig")

    # Competencia en España: Barefoot de LUMBRA vs las 2 marcas españolas comparables
    pvp_lumbra = cargar_categorias_info().get("Barefoot", {}).get("pvp_reunion")
    comp = [{"marca": "LUMBRA", "modelo": "Barefoot", "pvp": pvp_lumbra, "orden": 1, "nota": "PVP vigente (planner)"}]
    for c in COMPETIDORES_ESPANA_NUEVOS:
        comp.append({"marca": c["marca"], "modelo": c["modelo"], "pvp": c["pvp"],
                     "orden": len(comp) + 1, "nota": c["nota"]})
    pd.DataFrame(comp).to_csv(out_dir / "espana_competencia_barefoot.csv", index=False, encoding="utf-8-sig")

    print(f"CSV de España exportados: {out_dir}/espana_*.csv (6 archivos)")


def exportar_ranking_modelos_csv():
    """Exporta la matriz de margen por modelo y pais como CSV, para Power BI -
    misma fuente que alimenta la pestana Ranking de modelos de la web."""
    out_dir = PROCESSED_PATH.parent
    filas = calcular_matriz_margen_modelos()
    filas_planas = []
    for f in filas:
        for pais in PAISES_11.keys():
            filas_planas.append({
                "Familia": f["familia"].split(" ", 1)[-1] if f["familia"][0].isdigit() else f["familia"],
                "NombreModelo": f["nombre_display"],
                "EsBarefoot": 1 if "Barefoot" in f["familia"] else 0,
                "Modelos": f["modelos"], "PVP": f["pvp"], "Coste": f["coste"],
                "Pais": pais, "PctMargen": round(f[pais] / 100, 4),  # fraccion (0,537), no puntos: asi el formato % de Power BI funciona directo
                "PromedioModelo": round(f["promedio_11_paises"] / 100, 4),
            })
    pd.DataFrame(filas_planas).to_csv(out_dir / "ranking_margen_modelos.csv", index=False, encoding="utf-8-sig")
    print(f"CSV de Ranking de modelos exportado: {out_dir}/ranking_margen_modelos.csv")


def cargar_chartjs() -> str:
    """Lee Chart.js desde src/vendor/ (descargado una vez via npm) para embeberlo
    directamente en el HTML. Asi el dashboard funciona sin conexion a internet,
    sin importar si la red bloquea CDNs externas (lo que causo el problema original)."""
    vendor_path = Path(__file__).resolve().parent / "vendor" / "chart.umd.js"
    if not vendor_path.exists():
        raise FileNotFoundError(
            f"No se encontro {vendor_path}. Chart.js debe vivir en src/vendor/chart.umd.js"
        )
    return vendor_path.read_text(encoding="utf-8")


def construir_product_categories():
    """Coste y PVP por categoria para el Simulador: salen del Excel de benchmark (hoja
    'Categorias FW27'), no se escriben a mano en el JS. Claves como las muestra el selector."""
    info = cargar_categorias_info()
    por_min = {k.lower(): v for k, v in info.items()}
    claves = {"Zueco": "zueco", "Bamba / Zapatilla": "bamba / zapatilla", "Zapato": "zapato",
              "Barefoot": "barefoot", "Botín plano": "botín plano", "Bailarina": "bailarina"}
    return {k: {"coste": por_min[v]["coste"], "pvp": por_min[v]["pvp_reunion"]} for k, v in claves.items()}


def texto_hallazgo_ranking():
    """Titulo y parrafo de la pestana Ranking, calculados desde la matriz de margenes
    (antes eran texto escrito a mano con cifras que no se actualizaban)."""
    filas = calcular_matriz_margen_modelos()
    def fmt(x): return f"{x:.1f}".replace(".", ",")
    def pais_min(f): return min(PAISES_11.keys(), key=lambda p: f[p])
    mejor, peor = filas[0], filas[-1]
    fam = {}
    for f in filas:
        fam.setdefault(f["familia"].split(" ", 1)[-1] if f["familia"][0].isdigit() else f["familia"], []).append(f["promedio_11_paises"])
    fam_prom = sorted(((k, sum(v) / len(v)) for k, v in fam.items()), key=lambda kv: -kv[1])
    n_mod = sum(len(f["modelos"].split(", ")) for f in filas)
    pm = pais_min(peor)
    causa = ("es un problema de coste de producción, no de precio de venta" if peor["coste"] / peor["pvp"] > 0.25
             else "el coste es contenido: el margen bajo viene de la estructura de impuestos y devoluciones")
    titulo = "Qué modelos dejan más margen — y cuáles menos"
    texto = (f"Con datos ficticios del sistema de compras de LUMBRA (BuyPlan, {n_mod} modelos de la colección FW27), "
             f"la familia con mejor margen promedio en los 11 países es <strong>{fam_prom[0][0]}</strong> ({fmt(fam_prom[0][1])}%) "
             f"y la de peor margen es <strong>{fam_prom[-1][0]}</strong> ({fmt(fam_prom[-1][1])}%). "
             f"El perfil de mayor margen es «{mejor['nombre_display']}» ({fmt(mejor['promedio_11_paises'])}% promedio); "
             f"el de menor margen es «{peor['nombre_display']}» ({fmt(peor['promedio_11_paises'])}% promedio, llegando a solo "
             f"{fmt(peor[pm])}% en {pm}): cuesta {peor['coste']:.0f}€ de producción y se vende a {peor['pvp']:.0f}€ — {causa}.")
    return titulo, texto


def main():
    plantilla_path = Path(__file__).resolve().parent / "dashboard_template.html"
    html = plantilla_path.read_text(encoding="utf-8")
    html = html.replace("__CHARTJS_SOURCE__", cargar_chartjs())
    html = html.replace("__DATA_JSON__", cargar_datos_json())
    benchmark_rows, prospeccion_resumen = cargar_benchmark_competencia()
    html = html.replace("__BENCHMARK_JSON__", json.dumps(benchmark_rows, ensure_ascii=False))
    html = html.replace("__PROSPECCION_JSON__", json.dumps(prospeccion_resumen, ensure_ascii=False))
    html = html.replace("__CIUDADES_JSON__", json.dumps(CIUDADES_ESTUDIO, ensure_ascii=False))
    html = html.replace("__PAISES_JSON__", json.dumps(PAISES_ESTUDIO, ensure_ascii=False))
    html = html.replace("__CIUDAD_A_PAIS_JSON__", json.dumps(CIUDAD_A_PAIS, ensure_ascii=False))
    html = html.replace("__CATEGORIA_DOMINANTE_JSON__", json.dumps(CATEGORIA_DOMINANTE_POR_PAIS, ensure_ascii=False))
    html = html.replace("__MULTIPLICADOR_WHOLESALE__", str(MULTIPLICADOR_WHOLESALE))
    html = html.replace("__PRODUCT_CATEGORIES_JSON__", json.dumps(construir_product_categories(), ensure_ascii=False))
    _tit, _txt = texto_hallazgo_ranking()
    html = html.replace("__RANKING_HALLAZGO_TITULO__", _tit).replace("__RANKING_HALLAZGO_TEXTO__", _txt)
    html = html.replace("__PVP_BAREFOOT__", f"{construir_product_categories()['Barefoot']['pvp']:.0f}")
    est = [p for p, info in PAISES_11.items() if info["devolucion"] is None]
    prom = f"{_DEVOL_PROMEDIO_11 * 100:.1f}".replace(".", ",")
    nota = (f"IVA real de cada país y su tasa de devolución. Estimación, no dato propio: {', '.join(est[:-1])} y {est[-1]} "
            f"no tienen tasa de devolución confirmada, así que usan el promedio de las tasas reales ({prom}%).")
    html = html.replace("__RANKING_NOTA_ESTIMADOS__", nota)
    html = html.replace("__MULT_TXT__", str(MULTIPLICADOR_WHOLESALE).replace(".", ","))
    html = html.replace("__COMISION_TXT__", str(COMISION_AGENTE_B2B_PCT).replace(".", ","))
    html = html.replace("__TRANSPORTE_TXT__", f"{COSTE_TRANSPORTE_B2B:.2f}".replace(".", ","))
    html = html.replace("__COMISION_AGENTE_PCT__", str(COMISION_AGENTE_B2B_PCT))
    html = html.replace("__COSTE_TRANSPORTE__", str(COSTE_TRANSPORTE_B2B))
    html = html.replace("__ESPANA_B2C_JSON__", json.dumps(ESPANA_B2C, ensure_ascii=False))
    html = html.replace("__ESPANA_CANAL_B2B_JSON__", json.dumps(ESPANA_CANAL_B2B, ensure_ascii=False))
    html = html.replace("__EXPORT_RANKING_JSON__", json.dumps(EXPORT_RANKING_PAISES, ensure_ascii=False))
    html = html.replace("__ESPANA_MARGENES_EXPORT_JSON__", json.dumps(construir_espana_margenes(), ensure_ascii=False))
    html = html.replace("__ESPANA_MARGEN_CANAL_JSON__", json.dumps(calcular_margen_espana_canales(), ensure_ascii=False))
    html = html.replace("__MATRIZ_MARGEN_MODELOS_JSON__", json.dumps(calcular_matriz_margen_modelos(), ensure_ascii=False))
    html = html.replace("__PAISES_11_LISTA_JSON__", json.dumps(list(PAISES_11.keys()), ensure_ascii=False))
    html = html.replace("__TENDENCIA_BOTAS_JSON__", json.dumps(TENDENCIA_BOTAS_ALTAS, ensure_ascii=False))
    html = html.replace("__SCATTER_JSON__", json.dumps(SCATTER_DATA, ensure_ascii=False))
    html = html.replace("__POBLACION_JSON__", json.dumps(POBLACION_MILLONES))
    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"Dashboard generado: {OUTPUT_HTML}")
    exportar_scatter_csv()
    exportar_espana_csv()
    exportar_ranking_modelos_csv()



if __name__ == "__main__":
    main()
