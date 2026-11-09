"""
etl.py
------
Lee el Excel crudo (data/raw/LUMBRA_datos_cuantitativos_FW27.xlsx), replica en Python
la misma logica de conversion de moneda que ya vive en las formulas del Excel
(hoja "Tipos de cambio"), y exporta un CSV limpio y listo para Power BI en
data/processed/lumbra_datos_limpios.csv.

Por que replicar la conversion aqui y no solo leer los valores ya calculados del Excel:
Power BI lee el VALOR EN CACHE de una formula de Excel, no la formula en si. Si alguien
cambia el tipo de cambio en el Excel y no lo vuelve a abrir con Excel/LibreOffice antes
de refrescar Power BI, el dato queda desactualizado sin que nadie lo note. Haciendo el
calculo aqui, basta con correr este script de nuevo.

Uso:
    python src/etl.py
"""
from pathlib import Path
import pandas as pd

RAW_PATH = Path(__file__).resolve().parent.parent / "data" / "raw" / "LUMBRA_datos_cuantitativos_FW27.xlsx"
PROCESSED_PATH = Path(__file__).resolve().parent.parent / "data" / "processed" / "lumbra_datos_limpios.csv"

# Mismos tipos de cambio que la hoja "Tipos de cambio" del Excel.
# Fuente: BCE (Banco Central Europeo) - Euro foreign exchange reference rates.
# Actualiza estos tres valores si tienes una fecha de referencia mas reciente.
TIPOS_DE_CAMBIO = {
    "EUR_USD": 1.1578,   # 1 EUR = 1.1578 USD (BCE, 02/09/2026)
    "EUR_CHF": 0.9424,   # 1 EUR = 0.9424 CHF (BCE, 02/09/2026)
    "EUR_NOK": 10.9325,  # 1 EUR = 10.9325 NOK (BCE, 14/08/2026)
}


def leer_datos_crudos() -> pd.DataFrame:
    """Lee la hoja 'Datos por pais' del Excel (valores en cache, sin formulas)."""
    df = pd.read_excel(RAW_PATH, sheet_name="Datos por pais")
    df.columns = [c.strip() for c in df.columns]
    return df


def convertir_a_usd_y_eur(row: pd.Series, fx: dict) -> tuple[float, float]:
    """Devuelve (valor_usd, valor_eur) para una fila monetaria; (None, None) si es porcentaje."""
    if row["Tipo"] != "Monetario":
        return (None, None)

    moneda = row["Moneda original"]
    valor = row["Valor original"]

    if moneda == "USD":
        valor_usd = valor
        valor_eur = valor / fx["EUR_USD"]
    elif moneda == "EUR":
        valor_eur = valor
        valor_usd = valor * fx["EUR_USD"]
    elif moneda == "CHF":
        valor_eur = valor / fx["EUR_CHF"]
        valor_usd = valor_eur * fx["EUR_USD"]
    elif moneda == "NOK":
        valor_eur = valor / fx["EUR_NOK"]
        valor_usd = valor_eur * fx["EUR_USD"]
    else:
        raise ValueError(f"Moneda no reconocida: {moneda}")

    return (round(valor_usd, 2), round(valor_eur, 2))


def construir_dataset_limpio() -> pd.DataFrame:
    df = leer_datos_crudos()

    conversiones = df.apply(lambda r: convertir_a_usd_y_eur(r, TIPOS_DE_CAMBIO), axis=1)
    df["Valor USD"] = [c[0] for c in conversiones]
    df["Valor EUR"] = [c[1] for c in conversiones]

    # Para filas de tipo Porcentaje, el valor ya vive en "Valor original"
    df["Valor Porcentaje"] = df.apply(
        lambda r: r["Valor original"] if r["Tipo"] == "Porcentaje" else None, axis=1
    )

    columnas_finales = [
        "Segmento", "Pais", "Indicador", "Tipo", "Escala", "Moneda original",
        "Valor original", "Valor USD", "Valor EUR", "Valor Porcentaje",
        "Anio", "Fuente", "URL",
    ]
    return df[columnas_finales]


def main():
    df_limpio = construir_dataset_limpio()
    PROCESSED_PATH.parent.mkdir(parents=True, exist_ok=True)
    df_limpio.to_csv(PROCESSED_PATH, index=False, encoding="utf-8-sig")
    print(f"Dataset limpio exportado: {PROCESSED_PATH}")
    print(f"Filas: {len(df_limpio)}  |  Columnas: {list(df_limpio.columns)}")


if __name__ == "__main__":
    main()
