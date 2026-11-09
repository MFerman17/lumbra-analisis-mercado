"""
montecarlo.py
--------------
Simulacion Monte Carlo del margen neto unitario de un producto LUMBRA, muestreando
la tasa de devolucion dentro de un rango en vez de usar un unico valor fijo.

Correccion deliberada frente a un modelo de referencia que revisamos (que asignaba
un rango de devolucion "plausible" a cada region sin citar fuente): aqui SOLO
Alemania, Austria y Suiza reciben un rango por defecto, derivado de su dato real
de Statista (ver RETURN_RATE_REAL) con una banda de +/- BANDA_INCERTIDUMBRE_PP
puntos porcentuales alrededor de ese valor. Para el resto de paises no se ofrece
ningun rango por defecto: el usuario debe indicarlo explicitamente (mismo criterio
"tu dato / fuente" que ya usa el dashboard).

Uso:
    python src/montecarlo.py --pais Suecia --precio 155 --coste 52 \
        --tasa-min 18 --tasa-max 25 --coste-devolucion 13
"""
import argparse
from pathlib import Path
import numpy as np
import pandas as pd

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data" / "processed"

# Mismo dato real que ya usa el dashboard (Statista, encuesta 2022, recogida por Prime AI)
RETURN_RATE_REAL = {"Alemania": 44.0, "Austria": 36.0, "Suiza": 45.0}
BANDA_INCERTIDUMBRE_PP = 4.0  # banda +/- alrededor del dato real, EDITABLE, no es un dato de fuente
N_SIMULACIONES_DEFAULT = 10000
RANDOM_SEED = 42


def rango_por_defecto(pais: str):
    """Devuelve (min, max, es_dato_real) para un pais. Solo DACH tiene default."""
    if pais in RETURN_RATE_REAL:
        base = RETURN_RATE_REAL[pais]
        return (max(0, base - BANDA_INCERTIDUMBRE_PP), min(100, base + BANDA_INCERTIDUMBRE_PP), True)
    return (None, None, False)


def simular(pais: str, precio: float, coste_produccion: float, tasa_min_pct: float,
            tasa_max_pct: float, coste_logistico_devolucion: float,
            n_simulaciones: int = N_SIMULACIONES_DEFAULT, seed: int = RANDOM_SEED) -> pd.DataFrame:
    """Corre la simulacion Monte Carlo para un pais y devuelve un DataFrame, una fila por iteracion."""
    rng = np.random.default_rng(seed)
    tasas = rng.uniform(tasa_min_pct / 100, tasa_max_pct / 100, n_simulaciones)

    ingreso_efectivo = precio * (1 - tasas)
    coste_devolucion_promedio = tasas * coste_logistico_devolucion
    margen_neto_unidad = ingreso_efectivo - coste_produccion - coste_devolucion_promedio
    margen_pct = (margen_neto_unidad / precio) * 100

    return pd.DataFrame({
        "id_simulacion": [f"{pais[:3].upper()}_{i+1:05d}" for i in range(n_simulaciones)],
        "pais": pais,
        "tasa_devolucion_pct": tasas * 100,
        "margen_neto_unidad": margen_neto_unidad,
        "margen_pct": margen_pct,
    })


def resumen(df: pd.DataFrame) -> pd.DataFrame:
    """Estadisticas agregadas por pais: media, min, max, desviacion estandar, percentiles 10/90."""
    return df.groupby("pais").agg(
        margen_medio=("margen_neto_unidad", "mean"),
        margen_min=("margen_neto_unidad", "min"),
        margen_max=("margen_neto_unidad", "max"),
        margen_std=("margen_neto_unidad", "std"),
        margen_p10=("margen_neto_unidad", lambda x: np.percentile(x, 10)),
        margen_p90=("margen_neto_unidad", lambda x: np.percentile(x, 90)),
        margen_pct_medio=("margen_pct", "mean"),
    ).reset_index()


def main():
    parser = argparse.ArgumentParser(description="Simulacion Monte Carlo de margen neto por pais")
    parser.add_argument("--pais", required=True)
    parser.add_argument("--precio", type=float, required=True, help="Precio de venta por unidad")
    parser.add_argument("--coste", type=float, required=True, help="Coste de produccion por unidad")
    parser.add_argument("--coste-devolucion", type=float, required=True, help="Coste logistico por devolucion")
    parser.add_argument("--tasa-min", type=float, help="Tasa de devolucion minima (%%). Si se omite, se usa el default real (solo Alemania/Austria/Suiza)")
    parser.add_argument("--tasa-max", type=float, help="Tasa de devolucion maxima (%%). Si se omite, se usa el default real (solo Alemania/Austria/Suiza)")
    parser.add_argument("--n", type=int, default=N_SIMULACIONES_DEFAULT)
    args = parser.parse_args()

    tasa_min, tasa_max = args.tasa_min, args.tasa_max
    if tasa_min is None or tasa_max is None:
        tasa_min_def, tasa_max_def, es_real = rango_por_defecto(args.pais)
        if tasa_min_def is None:
            raise SystemExit(
                f"'{args.pais}' no tiene un rango de devolucion por defecto (solo Alemania, Austria "
                f"y Suiza lo tienen, porque son los unicos con dato real de Statista). "
                f"Debes indicar --tasa-min y --tasa-max explicitamente."
            )
        tasa_min = tasa_min if tasa_min is not None else tasa_min_def
        tasa_max = tasa_max if tasa_max is not None else tasa_max_def
        print(f"[i] Usando rango por defecto para {args.pais}: {tasa_min:.1f}%-{tasa_max:.1f}% "
              f"(banda de +/-{BANDA_INCERTIDUMBRE_PP}pp alrededor del dato real de Statista)")

    df = simular(args.pais, args.precio, args.coste, tasa_min, tasa_max, args.coste_devolucion, args.n)
    df_resumen = resumen(df)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT_DIR / "simulacion_montecarlo.csv", index=False, encoding="utf-8-sig")
    df_resumen.to_csv(OUTPUT_DIR / "resumen_montecarlo.csv", index=False, encoding="utf-8-sig")

    print(f"\n=== Resumen Monte Carlo — {args.pais} ({args.n} simulaciones) ===")
    print(df_resumen.to_string(index=False))
    print(f"\nGuardado en: {OUTPUT_DIR / 'simulacion_montecarlo.csv'}")
    print(f"Guardado en: {OUTPUT_DIR / 'resumen_montecarlo.csv'}")


if __name__ == "__main__":
    main()
