import streamlit as st
import numpy as np
import pandas as pd
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq

st.header("Cálculo Exacto de la Intersección: PMV = 0 y tdb = tr para diferentes Humedades y Velocidades del Aire")

# Parámetros fijos
clo = 0.5
met = 1.2

# Definir valores de humedad relativa (%) y velocidad del aire (m/s)
rh_values = [20, 30, 40, 50,60, 70, 80]        # Ejemplo: 25%, 40% y 60%
v_values = [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]   # Ejemplo: 0, 0.2, 0.5 y 0.8 m/s

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
