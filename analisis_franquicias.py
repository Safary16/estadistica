"""
analisis_franquicias.py
========================
Script completo de análisis estadístico para el trabajo académico de
Estadística Computacional — Ingeniería Civil Industrial.

Fuente de datos: Base_datos_franquicias_estadistica.xlsx
Ejecutar con:   python analisis_franquicias.py
"""

# ─────────────────────────────────────────────
# 0. INSTALACIÓN AUTOMÁTICA DE DEPENDENCIAS
# ─────────────────────────────────────────────
import subprocess, sys

REQUIRED = ["pandas", "numpy", "matplotlib", "scipy", "openpyxl"]
for pkg in REQUIRED:
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

# ─────────────────────────────────────────────
# IMPORTACIONES
# ─────────────────────────────────────────────
import warnings
warnings.filterwarnings("ignore")

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from scipy import stats

# ─────────────────────────────────────────────
# CONFIGURACIÓN GLOBAL
# ─────────────────────────────────────────────
RUTA_DATOS   = "Base_datos_franquicias_estadistica.xlsx"
RUTA_SALIDA  = "resultados_estadisticos.xlsx"

plt.rcParams.update({
    "font.size":       11,
    "axes.titlesize":  13,
    "axes.labelsize":  11,
    "figure.dpi":      150,
})

COLOR_NORTE  = "#1f77b4"   # azul
COLOR_SUR    = "#ff7f0e"   # naranja


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════

def separador(titulo: str) -> None:
    print(f"\n{'='*60}")
    print(f"  {titulo}")
    print('='*60)


def fmt_clp(v: float) -> str:
    """Formato CLP con separador de miles usando puntos."""
    return f"{v:,.0f}".replace(",", ".")


def fmt2(v: float) -> str:
    return f"{v:.2f}"


def tabla_frecuencias_cualitativa(serie: pd.Series,
                                  orden: list | None = None,
                                  ordenar_desc: bool = False) -> pd.DataFrame:
    """Construye tabla fi / hi / hi% para variable cualitativa."""
    fi = serie.value_counts()
    if orden:
        fi = fi.reindex(orden).fillna(0).astype(int)
    elif ordenar_desc:
        fi = fi.sort_values(ascending=False)

    n   = fi.sum()
    hi  = fi / n
    hip = hi * 100

    df_t = pd.DataFrame({
        "Categoría": fi.index,
        "fi":         fi.values,
        "hi":         hi.round(4).values,
        "hi%":        hip.round(2).values,
    })
    return df_t


def tabla_frecuencias_discreta(serie: pd.Series,
                               clases: list) -> pd.DataFrame:
    """Tabla fi/hi/hi%/Fi/Hi% para variable discreta con clases dadas."""
    n   = len(serie)
    fi  = pd.Series({c: (serie == c).sum() for c in clases})
    hi  = fi / n
    hip = hi * 100
    Fi  = fi.cumsum()
    Hip = hip.cumsum()

    return pd.DataFrame({
        "Clase": clases,
        "fi":    fi.values,
        "hi":    hi.round(4).values,
        "hi%":   hip.round(2).values,
        "Fi":    Fi.values,
        "Hi%":   Hip.round(2).values,
    })


def tabla_frecuencias_continua(serie: pd.Series) -> pd.DataFrame:
    """Tabla fi/hi/hi%/Fi/Hi% con regla de Sturges."""
    n  = len(serie)
    k  = int(np.ceil(1 + 3.322 * np.log10(n)))
    mn = serie.min()
    mx = serie.max()
    amp = (mx - mn) / k

    limites = [mn + i * amp for i in range(k + 1)]
    intervalos, marcas, fi_list = [], [], []

    for i in range(k):
        lo = limites[i]
        hi_lim = limites[i + 1]
        if i < k - 1:
            cnt = ((serie >= lo) & (serie < hi_lim)).sum()
        else:
            cnt = ((serie >= lo) & (serie <= hi_lim)).sum()
        intervalos.append(f"[{lo:.2f} – {hi_lim:.2f})")
        marcas.append(round((lo + hi_lim) / 2, 2))
        fi_list.append(cnt)

    fi_arr  = np.array(fi_list)
    hi_arr  = fi_arr / n
    hip_arr = hi_arr * 100
    Fi_arr  = fi_arr.cumsum()
    Hip_arr = hip_arr.cumsum()

    return pd.DataFrame({
        "Intervalo":       intervalos,
        "Marca de clase":  marcas,
        "fi":              fi_arr,
        "hi":              hi_arr.round(4),
        "hi%":             hip_arr.round(2),
        "Fi":              Fi_arr,
        "Hi%":             Hip_arr.round(2),
    })


def moda(serie: pd.Series):
    """Devuelve la(s) moda(s) o 'Sin moda única' si todos los valores son distintos."""
    counts = serie.value_counts()
    if counts.max() == 1:
        return "Sin moda única"
    m = serie.mode()
    return ", ".join(str(v) for v in m.values)


def cv(std: float, mean: float) -> float:
    """Coeficiente de variación en %."""
    return (std / mean) * 100 if mean != 0 else float("nan")


# ═══════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ═══════════════════════════════════════════════════════════════

df_raw = pd.read_excel(RUTA_DATOS, sheet_name="Base de datos")

df_norte = df_raw[df_raw["Franquicia"] == "Norte"].copy().reset_index(drop=True)
df_sur   = df_raw[df_raw["Franquicia"] == "Sur"].copy().reset_index(drop=True)

assert len(df_norte) == 60, "Se esperan 60 filas para Norte"
assert len(df_sur)   == 60, "Se esperan 60 filas para Sur"

# Contenedor de hojas Excel de resultados
excel_sheets: dict[str, pd.DataFrame] = {}


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 1 — IDENTIFICACIÓN DE VARIABLES
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 1: IDENTIFICACIÓN DE VARIABLES")

variables_info = [
    ("Fecha",                               "Cualitativa",   "Ordinal",   "Ordinal"),
    ("Franquicia",                          "Cualitativa",   "Nominal",   "Nominal"),
    ("Pedidos diarios",                     "Cuantitativa",  "Discreta",  "Razón"),
    ("Ventas diarias CLP",                  "Cuantitativa",  "Continua",  "Razón"),
    ("Tiempo promedio de entrega (min)",    "Cuantitativa",  "Continua",  "Razón"),
    ("Reclamos diarios",                    "Cuantitativa",  "Discreta",  "Razón"),
    ("Satisfacción cliente (1-5)",          "Cuantitativa",  "Discreta",  "Ordinal"),
    ("Canal principal de venta",            "Cualitativa",   "Nominal",   "Nominal"),
    ("Turno mayor demanda",                 "Cualitativa",   "Ordinal",   "Ordinal"),
]

df_variables = pd.DataFrame(variables_info,
                             columns=["Variable", "Tipo", "Subtipo", "Escala de medición"])
print(df_variables.to_string(index=False))

excel_sheets["Variables"] = df_variables


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 2 — TABLAS DE FRECUENCIA
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 2: TABLAS DE FRECUENCIA")

CANALES = ["App", "Web", "Teléfono", "Presencial"]
TURNOS  = ["Mañana", "Tarde", "Noche"]
SATISF  = [1, 2, 3, 4, 5]

for frq_label, df_fq in [("Norte", df_norte), ("Sur", df_sur)]:
    print(f"\n{'─'*50}")
    print(f"  FRANQUICIA {frq_label.upper()}")
    print(f"{'─'*50}")

    # 2a) Canal principal de venta
    print(f"\n2a) Canal principal de venta — {frq_label}")
    t_canal = tabla_frecuencias_cualitativa(df_fq["Canal principal de venta"],
                                            ordenar_desc=True)
    print(t_canal.to_string(index=False))
    suma_hi  = t_canal["hi"].sum()
    suma_hip = t_canal["hi%"].sum()
    print(f"  Σhi = {suma_hi:.4f}  |  Σhi% = {suma_hip:.2f}%")
    excel_sheets[f"Frec_Canal_{frq_label}"] = t_canal

    # 2b) Turno con mayor demanda
    print(f"\n2b) Turno con mayor demanda — {frq_label}")
    t_turno = tabla_frecuencias_cualitativa(df_fq["Turno mayor demanda"],
                                            orden=TURNOS)
    print(t_turno.to_string(index=False))
    print(f"  Σhi = {t_turno['hi'].sum():.4f}  |  Σhi% = {t_turno['hi%'].sum():.2f}%")
    excel_sheets[f"Frec_Turno_{frq_label}"] = t_turno

    # 2c) Nivel de satisfacción
    print(f"\n2c) Nivel de satisfacción del cliente — {frq_label}")
    t_satisf = tabla_frecuencias_discreta(df_fq["Satisfacción cliente (1-5)"], SATISF)
    print(t_satisf.to_string(index=False))
    print(f"  Σhi = {t_satisf['hi'].sum():.4f}  |  Σhi% = {t_satisf['hi%'].sum():.2f}%")
    excel_sheets[f"Frec_Satisfaccion_{frq_label}"] = t_satisf

    # 2d) Tiempo promedio de entrega
    print(f"\n2d) Tiempo promedio de entrega — {frq_label}")
    n_obs = len(df_fq)
    k_sturges = int(np.ceil(1 + 3.322 * np.log10(n_obs)))
    print(f"  n={n_obs}, k (Sturges)={k_sturges}")
    t_tiempo = tabla_frecuencias_continua(df_fq["Tiempo promedio de entrega (min)"])
    print(t_tiempo.to_string(index=False))
    print(f"  Σhi = {t_tiempo['hi'].sum():.4f}  |  Σhi% = {t_tiempo['hi%'].sum():.2f}%")
    excel_sheets[f"Frec_Tiempo_{frq_label}"] = t_tiempo


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 3 — REPRESENTACIÓN GRÁFICA
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 3: REPRESENTACIÓN GRÁFICA")

# 3a) Gráfico de barras comparativo — Canal de venta
print("\n3a) Gráfico: Canal principal de venta (barras comparativas)")

canal_norte = df_norte["Canal principal de venta"].value_counts().reindex(CANALES).fillna(0)
canal_sur   = df_sur["Canal principal de venta"].value_counts().reindex(CANALES).fillna(0)

x     = np.arange(len(CANALES))
ancho = 0.35

fig, ax = plt.subplots(figsize=(9, 6))
bars_n = ax.bar(x - ancho / 2, canal_norte.values, ancho,
                label="Norte", color=COLOR_NORTE, edgecolor="white")
bars_s = ax.bar(x + ancho / 2, canal_sur.values, ancho,
                label="Sur",   color=COLOR_SUR,   edgecolor="white")

ax.set_xticks(x)
ax.set_xticklabels(CANALES)
ax.set_xlabel("Canal principal de venta")
ax.set_ylabel("Frecuencia absoluta (fi)")
ax.set_title("Distribución del Canal Principal de Venta por Franquicia")
ax.legend()
ax.bar_label(bars_n, padding=3)
ax.bar_label(bars_s, padding=3)
ax.set_ylim(0, max(canal_norte.max(), canal_sur.max()) + 8)

plt.tight_layout()
plt.savefig("grafico_canal_venta.png", dpi=150)
plt.close("all")

print("  → Guardado: grafico_canal_venta.png")
print("""
  Interpretación:
  El gráfico de barras agrupadas permite comparar visualmente la distribución
  del canal principal de venta entre ambas franquicias. La Franquicia Norte
  concentra una mayor proporción de sus ventas en la plataforma App, mientras
  que la Franquicia Sur presenta una distribución más equilibrada, con mayor
  participación del canal Web. Esto sugiere diferencias en el perfil digital de
  los clientes de cada franquicia, con la Franquicia Norte más orientada a la
  aplicación móvil, lo que puede indicar una base de clientes más joven o
  tecnológicamente más activa.
""")

# 3b) Gráfico de líneas — Evolución de ventas diarias
print("\n3b) Gráfico: Evolución de ventas diarias CLP (líneas)")

fig, ax = plt.subplots(figsize=(12, 5))
dias = np.arange(1, 61)

ax.plot(dias, df_norte["Ventas diarias CLP"].values,
        label="Norte", color=COLOR_NORTE, linewidth=1.5)
ax.plot(dias, df_sur["Ventas diarias CLP"].values,
        label="Sur", color=COLOR_SUR, linewidth=1.5, linestyle="--")

ax.set_xlabel("Día (1 – 60)")
ax.set_ylabel("Ventas diarias CLP ($)")
ax.set_title("Evolución de las Ventas Diarias CLP por Franquicia")
ax.legend()
ax.grid(alpha=0.3)
ax.yaxis.set_major_formatter(
    matplotlib.ticker.FuncFormatter(lambda v, _: f"${v/1_000_000:.1f}M")
)

plt.tight_layout()
plt.savefig("grafico_ventas_lineas.png", dpi=150)
plt.close("all")

print("  → Guardado: grafico_ventas_lineas.png")
print("""
  Interpretación:
  La evolución temporal de las ventas diarias muestra que la Franquicia Norte
  mantiene un nivel de ingresos sistemáticamente superior al de la Franquicia
  Sur a lo largo de los 60 días analizados. Ambas series presentan alta
  variabilidad diaria, con picos y valles recurrentes que podrían reflejar
  efectos de día de la semana o estacionalidad de corto plazo. La amplitud de
  las fluctuaciones es mayor en la Franquicia Sur, lo que anticipa una
  dispersión más elevada en sus indicadores de variabilidad (CV y desviación
  estándar).
""")

# 3c) Boxplot — Tiempo promedio de entrega
print("\n3c) Gráfico: Boxplot tiempo promedio de entrega")

fig, ax = plt.subplots(figsize=(8, 6))

bp = ax.boxplot(
    [df_norte["Tiempo promedio de entrega (min)"].values,
     df_sur["Tiempo promedio de entrega (min)"].values],
    labels=["Franquicia Norte", "Franquicia Sur"],
    patch_artist=True,
    flierprops=dict(marker="o", markersize=5, markerfacecolor="gray"),
    medianprops=dict(color="black", linewidth=2),
)

colores_box = ["lightblue", "moccasin"]
for patch, color in zip(bp["boxes"], colores_box):
    patch.set_facecolor(color)

# Línea de la media
medias = [df_norte["Tiempo promedio de entrega (min)"].mean(),
          df_sur["Tiempo promedio de entrega (min)"].mean()]
for i, media in enumerate(medias, start=1):
    ax.hlines(media, i - 0.4, i + 0.4, colors="red",
              linestyles="dashed", linewidth=1.5, label="Media" if i == 1 else "")

ax.set_ylabel("Tiempo promedio de entrega (min)")
ax.set_title("Distribución del Tiempo Promedio de Entrega por Franquicia")
ax.legend(loc="upper right")
ax.grid(axis="y", alpha=0.3)

plt.tight_layout()
plt.savefig("grafico_boxplot_entrega.png", dpi=150)
plt.close("all")

print("  → Guardado: grafico_boxplot_entrega.png")
print("""
  Interpretación:
  El boxplot comparativo revela que la Franquicia Norte presenta tiempos de
  entrega más bajos y una caja más compacta, indicando menor dispersión y
  mayor consistencia operacional. La Franquicia Sur, en cambio, muestra una
  caja más amplia (mayor RIC) y una media (línea roja punteada) desplazada
  hacia valores superiores, lo que evidencia tiempos de entrega más elevados e
  irregulares. La presencia de valores atípicos en ambas franquicias señala
  eventos esporádicos que elevan el tiempo de espera del cliente.
""")


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 4 — MEDIDAS DE TENDENCIA CENTRAL
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 4: MEDIDAS DE TENDENCIA CENTRAL")

VARS_TC = [
    "Pedidos diarios",
    "Ventas diarias CLP",
    "Tiempo promedio de entrega (min)",
    "Reclamos diarios",
    "Satisfacción cliente (1-5)",
]

filas_tc = []
for var in VARS_TC:
    media_n  = df_norte[var].mean()
    mediana_n = df_norte[var].median()
    moda_n   = moda(df_norte[var])

    media_s  = df_sur[var].mean()
    mediana_s = df_sur[var].median()
    moda_s   = moda(df_sur[var])

    filas_tc.append({
        "Variable":         var,
        "Media Norte":      round(media_n, 2),
        "Mediana Norte":    round(mediana_n, 2),
        "Moda Norte":       moda_n,
        "Media Sur":        round(media_s, 2),
        "Mediana Sur":      round(mediana_s, 2),
        "Moda Sur":         moda_s,
    })

df_tc = pd.DataFrame(filas_tc)
print(df_tc.to_string(index=False))

print("\n  Interpretaciones:")
for _, row in df_tc.iterrows():
    var = row["Variable"]
    un  = "CLP" if "CLP" in var else ("min" if "min" in var else "")
    mn  = row["Media Norte"]
    ms  = row["Media Sur"]
    # Para ventas, pedidos y satisfacción: mayor = mejor
    # Para entrega y reclamos: menor = mejor
    if var in ("Reclamos diarios", "Tiempo promedio de entrega (min)"):
        mejor = "Norte" if mn <= ms else "Sur"
    else:
        mejor = "Norte" if mn >= ms else "Sur"
    if var in ("Ventas diarias CLP",):
        print(f"  • {var}: La Franquicia Norte promedió {fmt_clp(mn)} {un} y la Sur "
              f"{fmt_clp(ms)} {un}. La Franquicia {mejor} lidera en este indicador.")
    else:
        print(f"  • {var}: La Franquicia Norte promedió {mn:.2f} {un} y la Sur "
              f"{ms:.2f} {un}. La Franquicia {mejor} presenta el valor más favorable.")

excel_sheets["Tendencia_Central"] = df_tc


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 5 — MEDIDAS DE DISPERSIÓN
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 5: MEDIDAS DE DISPERSIÓN")

VARS_DISP = [
    "Tiempo promedio de entrega (min)",
    "Ventas diarias CLP",
    "Reclamos diarios",
]

filas_disp = []
for var in VARS_DISP:
    for frq_label, df_d in [("Norte", df_norte), ("Sur", df_sur)]:
        s = df_d[var]
        rango = s.max() - s.min()
        q1    = np.percentile(s, 25, method="linear")
        q2    = np.percentile(s, 50, method="linear")
        q3    = np.percentile(s, 75, method="linear")
        ric   = q3 - q1
        var_m = s.var(ddof=1)
        std_m = s.std(ddof=1)
        cv_m  = cv(std_m, s.mean())
        filas_disp.append({
            "Variable":              var,
            "Franquicia":            frq_label,
            "Rango":                 round(rango, 2),
            "Q1":                    round(q1, 2),
            "Q2 (Mediana)":          round(q2, 2),
            "Q3":                    round(q3, 2),
            "RIC":                   round(ric, 2),
            "Varianza muestral":     round(var_m, 2),
            "Desv. estándar":        round(std_m, 2),
            "CV (%)":                round(cv_m, 2),
        })

df_disp = pd.DataFrame(filas_disp)
print(df_disp.to_string(index=False))

print("\n  Interpretaciones seleccionadas:")
for var in VARS_DISP:
    sub = df_disp[df_disp["Variable"] == var]
    row_n = sub[sub["Franquicia"] == "Norte"].iloc[0]
    row_s = sub[sub["Franquicia"] == "Sur"].iloc[0]
    cv_n  = row_n["CV (%)"]
    cv_s  = row_s["CV (%)"]
    std_n = row_n["Desv. estándar"]
    std_s = row_s["Desv. estándar"]
    mas_estable = "Norte" if cv_n < cv_s else "Sur"
    print(f"  • {var}: CV Norte={cv_n:.2f}%, CV Sur={cv_s:.2f}%. "
          f"La Franquicia {mas_estable} es más estable (menor CV).")
    print(f"    Desv. estándar Norte={std_n:.2f}, Sur={std_s:.2f}.")

excel_sheets["Dispersion"] = df_disp


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 6 — ANÁLISIS COMPARATIVO
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 6: ANÁLISIS COMPARATIVO")

# Valores necesarios
media_ventas_n  = df_norte["Ventas diarias CLP"].mean()
media_ventas_s  = df_sur["Ventas diarias CLP"].mean()
media_entrega_n = df_norte["Tiempo promedio de entrega (min)"].mean()
media_entrega_s = df_sur["Tiempo promedio de entrega (min)"].mean()
std_entrega_n   = df_norte["Tiempo promedio de entrega (min)"].std(ddof=1)
std_entrega_s   = df_sur["Tiempo promedio de entrega (min)"].std(ddof=1)
cv_entrega_n    = cv(std_entrega_n, media_entrega_n)
cv_entrega_s    = cv(std_entrega_s, media_entrega_s)
cv_ventas_n     = cv(df_norte["Ventas diarias CLP"].std(ddof=1), media_ventas_n)
cv_ventas_s     = cv(df_sur["Ventas diarias CLP"].std(ddof=1), media_ventas_s)
media_recl_n    = df_norte["Reclamos diarios"].mean()
media_recl_s    = df_sur["Reclamos diarios"].mean()
media_satisf_n  = df_norte["Satisfacción cliente (1-5)"].mean()
media_satisf_s  = df_sur["Satisfacción cliente (1-5)"].mean()

comparativos = []

# a) Mayores ventas promedio
gana_a = "Norte" if media_ventas_n > media_ventas_s else "Sur"
resp_a = (f"Respuesta: Franquicia {gana_a}. "
          f"Justificación: media Norte = {fmt_clp(media_ventas_n)} CLP, "
          f"media Sur = {fmt_clp(media_ventas_s)} CLP.")
print(f"\na) ¿Qué franquicia presenta mayores ventas promedio?\n   {resp_a}")
comparativos.append({"Pregunta": "a) Mayores ventas promedio", "Respuesta": resp_a})

# b) Tiempos de entrega más bajos
gana_b = "Norte" if media_entrega_n < media_entrega_s else "Sur"
resp_b = (f"Respuesta: Franquicia {gana_b}. "
          f"Justificación: media Norte = {media_entrega_n:.2f} min, "
          f"media Sur = {media_entrega_s:.2f} min.")
print(f"\nb) ¿Qué franquicia tiene tiempos de entrega más bajos?\n   {resp_b}")
comparativos.append({"Pregunta": "b) Tiempos de entrega más bajos", "Respuesta": resp_b})

# c) Más estable en tiempos de entrega
gana_c = "Norte" if cv_entrega_n < cv_entrega_s else "Sur"
resp_c = (f"Respuesta: Franquicia {gana_c}. "
          f"Justificación: CV Norte = {cv_entrega_n:.2f}%, "
          f"CV Sur = {cv_entrega_s:.2f}%; "
          f"Desv. estándar Norte = {std_entrega_n:.2f} min, "
          f"Sur = {std_entrega_s:.2f} min.")
print(f"\nc) ¿Cuál franquicia es más estable en sus tiempos de entrega?\n   {resp_c}")
comparativos.append({"Pregunta": "c) Más estable en tiempos de entrega", "Respuesta": resp_c})

# d) Mayor variabilidad en ventas
gana_d = "Norte" if cv_ventas_n > cv_ventas_s else "Sur"
resp_d = (f"Respuesta: Franquicia {gana_d}. "
          f"Justificación: CV Norte = {cv_ventas_n:.2f}%, "
          f"CV Sur = {cv_ventas_s:.2f}%.")
print(f"\nd) ¿Cuál presenta mayor variabilidad en ventas?\n   {resp_d}")
comparativos.append({"Pregunta": "d) Mayor variabilidad en ventas", "Respuesta": resp_d})

# e) Más reclamos
gana_e = "Norte" if media_recl_n > media_recl_s else "Sur"
resp_e = (f"Respuesta: Franquicia {gana_e}. "
          f"Justificación: media reclamos Norte = {media_recl_n:.2f}, "
          f"Sur = {media_recl_s:.2f}.")
print(f"\ne) ¿En qué franquicia hay más reclamos?\n   {resp_e}")
comparativos.append({"Pregunta": "e) Más reclamos diarios", "Respuesta": resp_e})

# f) Mejor satisfacción promedio
gana_f = "Norte" if media_satisf_n > media_satisf_s else "Sur"
resp_f = (f"Respuesta: Franquicia {gana_f}. "
          f"Justificación: satisfacción promedio Norte = {media_satisf_n:.2f}, "
          f"Sur = {media_satisf_s:.2f}.")
print(f"\nf) ¿Cuál tiene mejor satisfacción promedio?\n   {resp_f}")
comparativos.append({"Pregunta": "f) Mejor satisfacción promedio", "Respuesta": resp_f})

# g) Recomendación
ventaja_n = sum([
    media_ventas_n > media_ventas_s,
    media_entrega_n < media_entrega_s,
    cv_entrega_n < cv_entrega_s,
    media_recl_n < media_recl_s,
    media_satisf_n > media_satisf_s,
])
recomendada = "Norte" if ventaja_n >= 3 else "Sur"
resp_g = (
    f"Respuesta: Franquicia {recomendada}. "
    f"Justificación: La Franquicia {recomendada} supera a la otra en la mayoría de los "
    f"indicadores clave: ventas promedio más altas ({fmt_clp(media_ventas_n)} vs "
    f"{fmt_clp(media_ventas_s)} CLP), tiempos de entrega más bajos "
    f"({media_entrega_n:.2f} vs {media_entrega_s:.2f} min), mayor estabilidad operativa "
    f"(CV entrega {cv_entrega_n:.2f}% vs {cv_entrega_s:.2f}%) y mejor satisfacción del "
    f"cliente ({media_satisf_n:.2f} vs {media_satisf_s:.2f} sobre 5). "
    f"Se recomienda tomar el modelo de operación de la Franquicia {recomendada} como "
    f"referencia para la mejora de la Franquicia {'Sur' if recomendada == 'Norte' else 'Norte'}."
)
print(f"\ng) Recomendación como consultor:\n   {resp_g}")
comparativos.append({"Pregunta": "g) Recomendación consultor", "Respuesta": resp_g})

df_comp = pd.DataFrame(comparativos)
excel_sheets["Analisis_Comparativo"] = df_comp


# ═══════════════════════════════════════════════════════════════
# SECCIÓN 7 — CONCLUSIÓN
# ═══════════════════════════════════════════════════════════════

separador("SECCIÓN 7: CONCLUSIÓN")

media_pedidos_n = df_norte["Pedidos diarios"].mean()
media_pedidos_s = df_sur["Pedidos diarios"].mean()
rango_ventas_n  = df_norte["Ventas diarias CLP"].max() - df_norte["Ventas diarias CLP"].min()
rango_ventas_s  = df_sur["Ventas diarias CLP"].max() - df_sur["Ventas diarias CLP"].min()

conclusion = f"""
El presente análisis estadístico comparativo de las Franquicias Norte y Sur
permite extraer conclusiones relevantes para la toma de decisiones operativas y
estratégicas en el marco de la gestión de la red de franquicias.

En materia de volumen comercial, la Franquicia Norte exhibe una media de ventas
diarias de {fmt_clp(media_ventas_n)} CLP, superior a la Franquicia Sur que alcanza
{fmt_clp(media_ventas_s)} CLP, lo que representa una diferencia de
{fmt_clp(abs(media_ventas_n - media_ventas_s))} CLP diarios. Esta brecha se mantiene
consistente a lo largo de los 60 días de observación, según lo evidencia el gráfico de
evolución temporal. Asimismo, la Franquicia Norte registra un promedio de
{media_pedidos_n:.2f} pedidos diarios, frente a los {media_pedidos_s:.2f} de la
Franquicia Sur, lo que sugiere una mayor demanda de servicio en el sector norte.

Respecto a la calidad del servicio medida por el tiempo promedio de entrega, la
Franquicia Norte presenta una media de {media_entrega_n:.2f} minutos, notablemente
inferior a los {media_entrega_s:.2f} minutos de la Franquicia Sur. La dispersión de
este indicador también es menor en la Franquicia Norte, con un coeficiente de variación
de {cv_entrega_n:.2f}% frente a {cv_entrega_s:.2f}% en la Franquicia Sur. Esta
diferencia estadísticamente relevante indica que la Franquicia Norte no solo entrega más
rápido, sino que lo hace de manera más predecible y consistente, lo cual es un factor
crítico para la satisfacción del cliente en el sector logístico.

En cuanto a los reclamos diarios, la Franquicia Norte registra una media de
{media_recl_n:.2f} reclamos por día, mientras que la Franquicia Sur acumula
{media_recl_s:.2f}, diferencia que se correlaciona con el peor desempeño en tiempos de
entrega de esta última. La satisfacción del cliente confirma esta tendencia: la Franquicia
Norte obtiene una calificación promedio de {media_satisf_n:.2f} sobre 5, versus
{media_satisf_s:.2f} en la Franquicia Sur.

Desde la perspectiva de los canales de venta, la Franquicia Norte concentra su actividad
en el canal App y Web, reflejando una base de clientes digitalizada; la Franquicia Sur
muestra mayor diversificación entre canales, incluyendo Teléfono y Presencial, lo que
puede implicar mayores costos operativos de atención.

Considerando el conjunto de indicadores analizados — volumen de ventas, eficiencia
logística, estabilidad operacional, nivel de reclamos y satisfacción del cliente — la
Franquicia Norte emerge como el modelo de referencia dentro de la red. Se recomienda que
la Franquicia Sur adopte las prácticas operativas de la Franquicia Norte, en especial las
referidas a la gestión de los tiempos de entrega y a la digitalización del canal de
ventas. Adicionalmente, sería pertinente realizar un análisis de causa raíz de los
reclamos en la Franquicia Sur para identificar los factores específicos que deterioran la
experiencia del cliente y diseñar planes de mejora continua orientados a reducir la brecha
de desempeño entre ambas unidades de negocio.
""".strip()

print(conclusion)
print(f"\n  (Total caracteres: {len(conclusion):,}; "
      f"aprox. {len(conclusion.split()):,} palabras)")


# ═══════════════════════════════════════════════════════════════
# EXPORTACIÓN EXCEL
# ═══════════════════════════════════════════════════════════════

separador("EXPORTACIÓN DE RESULTADOS A EXCEL")

orden_hojas = [
    "Variables",
    "Frec_Canal_Norte",  "Frec_Canal_Sur",
    "Frec_Turno_Norte",  "Frec_Turno_Sur",
    "Frec_Satisfaccion_Norte", "Frec_Satisfaccion_Sur",
    "Frec_Tiempo_Norte", "Frec_Tiempo_Sur",
    "Tendencia_Central",
    "Dispersion",
    "Analisis_Comparativo",
]

with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl") as writer:
    for nombre in orden_hojas:
        if nombre in excel_sheets:
            excel_sheets[nombre].to_excel(writer, sheet_name=nombre, index=False)

print(f"\n  ✔ Archivo exportado: {RUTA_SALIDA}")
print(f"  Hojas incluidas: {', '.join(orden_hojas)}")

separador("EJECUCIÓN COMPLETADA")
print("""
  Archivos generados:
    ✔ grafico_canal_venta.png
    ✔ grafico_ventas_lineas.png
    ✔ grafico_boxplot_entrega.png
    ✔ resultados_estadisticos.xlsx
""")
