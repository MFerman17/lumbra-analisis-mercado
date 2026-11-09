"""
analysis.py
-----------
Calcula un modelo TAM-SAM-SOM simplificado por pais para la propuesta de valor de LUMBRA,
y un ranking de oportunidad relativa basado en indicadores normalizados.

IMPORTANTE - esto es un PUNTO DE PARTIDA, no una cifra definitiva:
- Los supuestos (SAM_ONLINE_SHARE, SOM_CUOTA_OBJETIVO) son editables abajo y deben
  discutirse con el equipo antes de usarse en la entrega final.
- La poblacion es la unica variable que no viene del Excel: se agrego aqui con fuentes
  oficiales (Eurostat 1-ene-2025 para los 8 paises UE; oficinas nacionales de Suiza y
  Noruega para esos dos). Ver el diccionario POBLACION_MILLONES.

Uso:
    python src/analysis.py
"""
from pathlib import Path
import pandas as pd

PROCESSED_PATH = Path(__file__).resolve().parent.parent / "data" / "processed" / "lumbra_datos_limpios.csv"
OUTPUT_PATH = Path(__file__).resolve().parent.parent / "output" / "tam_sam_som_por_pais.csv"

# Poblacion 2025, en millones. Fuentes:
# - 8 paises UE: Eurostat, poblacion a 1 de enero de 2025 (publicado julio 2025)
# - Suiza: Oficina Federal de Estadistica (BFS), poblacion residente permanente, dic. 2025
# - Noruega: Statistics Norway (SSB) / countryeconomy.com, poblacion a fin de 2025
POBLACION_MILLONES = {
    "Suecia": 10.588,
    "Dinamarca": 5.993,
    "Noruega": 5.627,
    "Finlandia": 5.636,
    "Alemania": 83.577,
    "Austria": 9.197,
    "Suiza": 9.100,
    "Holanda": 18.044,
    "Belgica": 11.900,
    "Luxemburgo": 0.682,
}

# Supuestos del modelo (EDITABLES - documentar cualquier cambio para la entrega)
SAM_ONLINE_SHARE_DEFAULT = 0.30  # % del TAM que se transacciona online, si el pais no tiene un dato propio ya en el dataset
SOM_CUOTA_OBJETIVO = 0.0003      # 0.03% de cuota del SAM online: punto de partida conservador para una marca DTC nueva

# Cuota online real por pais, tomada del propio dataset cuando esta disponible
# (columna "Cuota online del retail de moda" o "Cuota de moda/apparel en eCommerce total")
CUOTA_ONLINE_CONOCIDA = {
    "Dinamarca": 0.207,   # dato propio del dataset (FashionUnited)
    "Holanda": 0.2567,    # dato propio del dataset (Mordor Intelligence)
}


def cargar_datos_limpios() -> pd.DataFrame:
    if not PROCESSED_PATH.exists():
        raise FileNotFoundError(
            f"No se encontro {PROCESSED_PATH}. Corre primero: python src/etl.py"
        )
    return pd.read_csv(PROCESSED_PATH)


def obtener_valor(df: pd.DataFrame, pais: str, indicador_contiene: str, columna: str):
    """Busca la primera fila de un pais cuyo Indicador contenga cierto texto y devuelve una columna."""
    filtro = (df["Pais"] == pais) & (df["Indicador"].str.contains(indicador_contiene, case=False, na=False))
    resultado = df.loc[filtro, columna]
    return resultado.iloc[0] if not resultado.empty else None


def calcular_tam_sam_som(df: pd.DataFrame) -> pd.DataFrame:
    filas = []
    for pais, poblacion_m in POBLACION_MILLONES.items():
        apparel_per_capita_usd = obtener_valor(df, pais, "Ingreso per capita mercado Apparel", "Valor USD")
        cuota_lujo_pct = obtener_valor(df, pais, "Cuota de lujo en Apparel", "Valor Porcentaje")

        if apparel_per_capita_usd is None or cuota_lujo_pct is None:
            # Paises sin ambos indicadores directos (p.ej. Alemania, Suiza usan Footwear/Apparel distintos)
            filas.append({
                "Pais": pais, "Poblacion (millones)": poblacion_m,
                "TAM lujo/premium (M USD)": None, "SAM online (M USD)": None,
                "SOM estimado (M USD)": None, "Nota": "Faltan indicadores directos de Apparel per capita o cuota de lujo; completar manualmente",
            })
            continue

        tam_total_m_usd = apparel_per_capita_usd * poblacion_m  # (USD/persona) x (millones de personas) = millones de USD
        tam_lujo_m_usd = tam_total_m_usd * (cuota_lujo_pct / 100)

        cuota_online = CUOTA_ONLINE_CONOCIDA.get(pais, SAM_ONLINE_SHARE_DEFAULT)
        sam_online_m_usd = tam_lujo_m_usd * cuota_online

        som_m_usd = sam_online_m_usd * SOM_CUOTA_OBJETIVO

        filas.append({
            "Pais": pais,
            "Poblacion (millones)": poblacion_m,
            "TAM lujo/premium (M USD)": round(tam_lujo_m_usd, 2),
            "Cuota online usada (%)": round(cuota_online * 100, 2),
            "SAM online (M USD)": round(sam_online_m_usd, 3),
            "SOM estimado (M USD)": round(som_m_usd, 4),
            "Nota": "Cuota online real del dataset" if pais in CUOTA_ONLINE_CONOCIDA else f"Cuota online supuesta ({SAM_ONLINE_SHARE_DEFAULT:.0%}), sin dato propio en el dataset",
        })

    return pd.DataFrame(filas)


def main():
    df = cargar_datos_limpios()
    resultado = calcular_tam_sam_som(df)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    resultado.to_csv(OUTPUT_PATH, index=False, encoding="utf-8-sig")

    print("Modelo TAM-SAM-SOM (supuestos: SOM = 0.03% del SAM online, ver comentarios en el codigo)\n")
    print(resultado.to_string(index=False))
    print(f"\nGuardado en: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
