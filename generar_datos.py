"""
Script auxiliar para generar el archivo de datos sintéticos.
Base_datos_franquicias_estadistica.xlsx
120 filas: 60 días × 2 franquicias (Norte / Sur)
"""

import numpy as np
import pandas as pd
from datetime import date, timedelta

np.random.seed(42)

n = 60
fechas = [date(2024, 1, 1) + timedelta(days=i) for i in range(n)]

canales = ["App", "Web", "Teléfono", "Presencial"]
turnos  = ["Mañana", "Tarde", "Noche"]

# ---------- Franquicia Norte ----------
rng_n = np.random.default_rng(10)
pedidos_n   = rng_n.integers(80, 160, n)
ventas_n    = rng_n.normal(4_500_000, 600_000, n).clip(2_500_000, 7_000_000)
entrega_n   = rng_n.normal(35, 8, n).clip(18, 60)
reclamos_n  = rng_n.integers(1, 12, n)
satisf_n    = rng_n.choice([1, 2, 3, 4, 5], n, p=[0.04, 0.08, 0.20, 0.40, 0.28])
canal_n     = rng_n.choice(canales, n, p=[0.40, 0.30, 0.15, 0.15])
turno_n     = rng_n.choice(turnos, n, p=[0.35, 0.45, 0.20])

df_norte = pd.DataFrame({
    "Fecha": fechas,
    "Franquicia": "Norte",
    "Pedidos diarios": pedidos_n,
    "Ventas diarias CLP": ventas_n.round(0).astype(int),
    "Tiempo promedio de entrega (min)": entrega_n.round(2),
    "Reclamos diarios": reclamos_n,
    "Satisfacción cliente (1-5)": satisf_n,
    "Canal principal de venta": canal_n,
    "Turno mayor demanda": turno_n,
})

# ---------- Franquicia Sur ----------
rng_s = np.random.default_rng(20)
pedidos_s   = rng_s.integers(60, 130, n)
ventas_s    = rng_s.normal(3_800_000, 750_000, n).clip(1_800_000, 6_200_000)
entrega_s   = rng_s.normal(42, 11, n).clip(20, 75)
reclamos_s  = rng_s.integers(3, 18, n)
satisf_s    = rng_s.choice([1, 2, 3, 4, 5], n, p=[0.07, 0.15, 0.28, 0.33, 0.17])
canal_s     = rng_s.choice(canales, n, p=[0.25, 0.35, 0.20, 0.20])
turno_s     = rng_s.choice(turnos, n, p=[0.40, 0.40, 0.20])

df_sur = pd.DataFrame({
    "Fecha": fechas,
    "Franquicia": "Sur",
    "Pedidos diarios": pedidos_s,
    "Ventas diarias CLP": ventas_s.round(0).astype(int),
    "Tiempo promedio de entrega (min)": entrega_s.round(2),
    "Reclamos diarios": reclamos_s,
    "Satisfacción cliente (1-5)": satisf_s,
    "Canal principal de venta": canal_s,
    "Turno mayor demanda": turno_s,
})

df = pd.concat([df_norte, df_sur], ignore_index=True)

with pd.ExcelWriter("Base_datos_franquicias_estadistica.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Base de datos", index=False)

print("Archivo generado: Base_datos_franquicias_estadistica.xlsx")
print(f"Filas totales: {len(df)}")
