import streamlit as st
import numpy as np
import pandas as pd
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # Importar para gráficos 3D

st.header("Cálculo Exacto de la Intersección: PMV = 0 y tdb = tr para diferentes Humedades y Velocidades del Aire")

# Parámetros fijos
clo = 0.5
met = 1.2

# Definir valores de humedad relativa (%) y velocidad del aire (m/s)
rh_values = [20, 30, 40, 50, 60, 70, 80]  # Ejemplo: 20, 30, ... 80%
v_values = [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]  # Ejemplo: de 0 a 0.9 m/s

results = []


# Función que calcula PMV cuando tdb = tr = t, dada una humedad y velocidad
def f_tdb(t, rh, v):
    result = pmv_ppd_iso(
        tdb=t,
        tr=t,
        vr=v,
        rh=rh,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    )
    return result.pmv


# Para cada combinación de humedad y velocidad, se busca la raíz de f_tdb(t) = 0
for rh in rh_values:
    for v in v_values:
        try:
            t_intersection = brentq(f_tdb, 18, 35, args=(rh, v))
            results.append({
                "Humedad (%)": rh,
                "Velocidad (m/s)": v,
                "T_intersección (°C)": t_intersection
            })
        except ValueError:
            results.append({
                "Humedad (%)": rh,
                "Velocidad (m/s)": v,
                "T_intersección (°C)": np.nan
            })

# Crear un DataFrame para mostrar los resultados
df = pd.DataFrame(results)
st.write("Resultados del cálculo de intersección (donde PMV = 0 y tdb = tr):")
st.dataframe(df)

###############################################
# Gráfico Adicional: Superficie de T_intersección vs. Humedad y Velocidad
###############################################
st.header("Superficie de T_intersección (PMV = 0, tdb = tr) vs. Humedad y Velocidad del Aire")

# Definir mallas de humedad y velocidad
rh_vals_plot = np.linspace(min(rh_values), max(rh_values), 50)
v_vals_plot = np.linspace(min(v_values), max(v_values), 50)
RH, V = np.meshgrid(rh_vals_plot, v_vals_plot)

# Crear una matriz para almacenar T_intersección
T_intersection = np.zeros_like(RH)

# Calcular T_intersección para cada combinación de (RH, V)
for i in range(RH.shape[0]):
    for j in range(RH.shape[1]):
        try:
            T_intersection[i, j] = brentq(f_tdb, 18, 35, args=(RH[i, j], V[i, j]))
        except ValueError:
            T_intersection[i, j] = np.nan

# Crear gráfico 3D de la superficie
fig = plt.figure(figsize=(8, 6))
ax = fig.add_subplot(111, projection='3d')
surf = ax.plot_surface(RH, V, T_intersection, cmap='viridis', edgecolor='none', alpha=0.8)
ax.set_xlabel("Humedad (%)")
ax.set_ylabel("Velocidad (m/s)")
ax.set_zlabel("T_intersección (°C)")
ax.set_title("Superficie de T_intersección (PMV=0) vs. Humedad y Velocidad")
fig.colorbar(surf, shrink=0.5, aspect=5)
st.pyplot(fig)

###############################################
# Análisis y decisión para intervenir
###############################################
st.header("Interpretación y Decisión de Intervención")

st.write("""
La superficie anterior muestra, para cada combinación de humedad y velocidad del aire, la temperatura neutral (T_intersección) en la que se alcanza PMV = 0 asumiendo que tdb = tr. Esta temperatura es un indicador del "objetivo" para el ambiente.

**Cómo interpretar para intervenir:**
- **Si en tu ambiente real la temperatura del aire (tdb) o la temperatura radiante (tr) se aleja significativamente de T_intersección**, entonces el componente correspondiente puede estar desbalanceado.
  - Por ejemplo, si mides un tdb mucho mayor que T_intersección, es probable que debas intervenir en el sistema HVAC para reducir el aire caliente.
  - Si, en cambio, tr es mayor (o menor) que T_intersección, podrías considerar soluciones sobre la envolvente del edificio (aislamiento, materiales, control de radiación solar).
- **La velocidad del aire (vr) también modula T_intersección.**  
  En la superficie se observa que a mayor vr, T_intersección tiende a aumentar. Esto significa que, en ambientes con alta velocidad del aire, la influencia de tdb es menor (por el enfriamiento convectivo), por lo que en ocasiones aumentar vr puede ser una intervención relativamente sencilla para acercar el ambiente a la condición de confort.

Con estos análisis, se puede identificar cuál de los parámetros (tdb, tr o vr) está más desviado del valor neutro y, por tanto, es el más indicado para intervenir.
""")
