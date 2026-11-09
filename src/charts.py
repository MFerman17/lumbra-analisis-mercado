"""
charts.py
---------
Genera un par de graficos de barras rapidos (PNG) a partir del dataset limpio,
para revisar visualmente antes de pasar a Power BI o al dashboard web.

Uso:
    python src/charts.py
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

PROCESSED_PATH = Path(__file__).resolve().parent.parent / "data" / "processed" / "lumbra_datos_limpios.csv"
CHARTS_DIR = Path(__file__).resolve().parent.parent / "output" / "charts"

COLOR_SEGMENTO = {
    "Nordicos": "#5B7FA6",
    "DACH": "#B8703F",
    "Benelux": "#5C7A52",
}


def cargar_datos() -> pd.DataFrame:
    return pd.read_csv(PROCESSED_PATH)


def grafico_pib_per_capita(df: pd.DataFrame):
    filtro = df["Indicador"] == "PIB per capita nominal"
    datos = df.loc[filtro, ["Pais", "Segmento", "Valor USD"]].dropna().sort_values("Valor USD")

    colores = [COLOR_SEGMENTO.get(s, "#888888") for s in datos["Segmento"]]
    fig, ax = plt.subplots(figsize=(9, 5.5))
    ax.barh(datos["Pais"], datos["Valor USD"], color=colores)
    ax.set_xlabel("USD")
    ax.set_title("PIB per cápita nominal 2025 — los 10 mercados LUMBRA")
    ax.xaxis.set_major_formatter(lambda x, _: f"${x:,.0f}")
    fig.tight_layout()
    fig.savefig(CHARTS_DIR / "pib_per_capita.png", dpi=150)
    plt.close(fig)


def grafico_cuota_lujo(df: pd.DataFrame):
    filtro = df["Indicador"].str.contains("Cuota de lujo", case=False, na=False)
    datos = df.loc[filtro, ["Pais", "Segmento", "Valor Porcentaje"]].dropna().sort_values("Valor Porcentaje")

    colores = [COLOR_SEGMENTO.get(s, "#888888") for s in datos["Segmento"]]
    fig, ax = plt.subplots(figsize=(9, 5.5))
    ax.barh(datos["Pais"], datos["Valor Porcentaje"], color=colores)
    ax.set_xlabel("% de ventas de lujo dentro del mercado de Apparel/Footwear")
    ax.set_title("Cuota de lujo por país (donde el dato está disponible)")
    fig.tight_layout()
    fig.savefig(CHARTS_DIR / "cuota_lujo.png", dpi=150)
    plt.close(fig)


def main():
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    df = cargar_datos()
    grafico_pib_per_capita(df)
    grafico_cuota_lujo(df)
    print(f"Graficos guardados en: {CHARTS_DIR}")


if __name__ == "__main__":
    main()
