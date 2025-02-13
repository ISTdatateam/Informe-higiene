import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq

# Constantes y parámetros fijos
clo = 0.5
met = 1.2
rh = 25  # Humedad relativa (%)

st.title("Gráficas de Límites de Confort Térmico (PMV entre -1 y 1)")
st.write("Constantes: clo = 0.5, met = 1.2 y humedad relativa = 25%")
st.write("Nota: La velocidad del aire debe ser menor a 1.0 m/s, excepto en el gráfico 4 (v = 0 m/s).")

num_points = 100  # Resolución de la malla

###############################
# 1. Gráfico: tdb vs tr (v fija por slider)
###############################
st.header("1. Temperatura de Bulbo Seco vs Temperatura Radiante")
v_fixed = st.slider("Selecciona la velocidad del aire fija (m/s):", min_value=0.1, max_value=0.99, value=0.8, step=0.01)

tdb_min = st.number_input("Rango inferior de Temperatura de Bulbo Seco (°C):", value=20.0, key='tdb_min')
tdb_max = st.number_input("Rango superior de Temperatura de Bulbo Seco (°C):", value=35.0, key='tdb_max')
tr_min = st.number_input("Rango inferior de Temperatura Radiante (°C):", value=20.0, key='tr_min')
tr_max = st.number_input("Rango superior de Temperatura Radiante (°C):", value=35.0, key='tr_max')

tdb_vals = np.linspace(tdb_min, tdb_max, num_points)
tr_vals = np.linspace(tr_min, tr_max, num_points)
TDB, TR = np.meshgrid(tdb_vals, tr_vals)

PMV = np.zeros_like(TDB)
for i in range(TDB.shape[0]):
    for j in range(TDB.shape[1]):
        result = pmv_ppd_iso(
            tdb=TDB[i, j],
            tr=TR[i, j],
            vr=v_fixed,
            rh=rh,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        )
        PMV[i, j] = result.pmv

fig1, ax1 = plt.subplots(figsize=(6, 4))
CS = ax1.contour(TDB, TR, PMV, levels=[-1, 0, 1], colors=['red', 'blue', 'red'])
ax1.clabel(CS, inline=True, fontsize=10)
ax1.set_xlabel("Temperatura de Bulbo Seco (°C)")
ax1.set_ylabel("Temperatura Radiante (°C)")
ax1.set_title(f"Contornos de PMV = -1, 0 y 1\n(v = {v_fixed:.2f} m/s)")
st.pyplot(fig1)

###############################
# 2. Gráfico: tdb vs v (tr fija)
###############################
st.header("2. Temperatura de Bulbo Seco vs Velocidad del Aire")
tr_fixed = st.number_input("Selecciona la Temperatura Radiante fija (°C):", value=30.0, key='tr_fixed')

tdb_min2 = st.number_input("Rango inferior de Temperatura de Bulbo Seco (°C):", value=20.0, key='tdb2_min')
tdb_max2 = st.number_input("Rango superior de Temperatura de Bulbo Seco (°C):", value=35.0, key='tdb2_max')
v_min = st.number_input("Rango inferior de Velocidad del Aire (m/s):", value=0.1, key='v_min')
v_max = st.number_input("Rango superior de Velocidad del Aire (m/s):", value=0.99, key='v_max')

tdb_vals2 = np.linspace(tdb_min2, tdb_max2, num_points)
v_vals = np.linspace(v_min, v_max, num_points)
TDB2, V = np.meshgrid(tdb_vals2, v_vals)

PMV2 = np.zeros_like(TDB2)
for i in range(TDB2.shape[0]):
    for j in range(TDB2.shape[1]):
        result = pmv_ppd_iso(
            tdb=TDB2[i, j],
            tr=tr_fixed,
            vr=V[i, j],
            rh=rh,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        )
        PMV2[i, j] = result.pmv

fig2, ax2 = plt.subplots(figsize=(6, 4))
CS2 = ax2.contour(TDB2, V, PMV2, levels=[-1, 0, 1], colors=['red', 'blue', 'red'])
ax2.clabel(CS2, inline=True, fontsize=10)
ax2.set_xlabel("Temperatura de Bulbo Seco (°C)")
ax2.set_ylabel("Velocidad del Aire (m/s)")
ax2.set_title(f"Contornos de PMV = -1, 0 y 1\n(tr = {tr_fixed:.1f} °C)")
st.pyplot(fig2)

###############################
# 3. Gráfico: tr vs v (tdb fija)
###############################
st.header("3. Temperatura Radiante vs Velocidad del Aire")
tdb_fixed = st.number_input("Selecciona la Temperatura de Bulbo Seco fija (°C):", value=30.0, key='tdb_fixed')

tr_min3 = st.number_input("Rango inferior de Temperatura Radiante (°C):", value=20.0, key='tr3_min')
tr_max3 = st.number_input("Rango superior de Temperatura Radiante (°C):", value=35.0, key='tr3_max')
v_min2 = st.number_input("Rango inferior de Velocidad del Aire (m/s):", value=0.1, key='v2_min')
v_max2 = st.number_input("Rango superior de Velocidad del Aire (m/s):", value=0.99, key='v2_max')

tr_vals2 = np.linspace(tr_min3, tr_max3, num_points)
v_vals2 = np.linspace(v_min2, v_max2, num_points)
TR2, V2 = np.meshgrid(tr_vals2, v_vals2)

PMV3 = np.zeros_like(TR2)
for i in range(TR2.shape[0]):
    for j in range(TR2.shape[1]):
        result = pmv_ppd_iso(
            tdb=tdb_fixed,
            tr=TR2[i, j],
            vr=V2[i, j],
            rh=rh,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        )
        PMV3[i, j] = result.pmv

fig3, ax3 = plt.subplots(figsize=(6, 4))
CS3 = ax3.contour(TR2, V2, PMV3, levels=[-1, 0, 1], colors=['red', 'blue', 'red'])
ax3.clabel(CS3, inline=True, fontsize=10)
ax3.set_xlabel("Temperatura Radiante (°C)")
ax3.set_ylabel("Velocidad del Aire (m/s)")
ax3.set_title(f"Contornos de PMV = -1, 0 y 1\n(tdb = {tdb_fixed:.1f} °C)")
st.pyplot(fig3)

###############################
# 4. Gráfico: tdb vs tr (v = 0 m/s)
###############################
st.header("4. Temperatura de Bulbo Seco vs Temperatura Radiante (v = 0 m/s)")
tdb_min4 = st.number_input("Rango inferior de Temperatura de Bulbo Seco (°C) [Gráfico 4]:", value=20.0, key='tdb4_min')
tdb_max4 = st.number_input("Rango superior de Temperatura de Bulbo Seco (°C) [Gráfico 4]:", value=35.0, key='tdb4_max')
tr_min4 = st.number_input("Rango inferior de Temperatura Radiante (°C) [Gráfico 4]:", value=20.0, key='tr4_min')
tr_max4 = st.number_input("Rango superior de Temperatura Radiante (°C) [Gráfico 4]:", value=35.0, key='tr4_max')

tdb_vals4 = np.linspace(tdb_min4, tdb_max4, num_points)
tr_vals4 = np.linspace(tr_min4, tr_max4, num_points)
TDB4, TR4 = np.meshgrid(tdb_vals4, tr_vals4)

PMV4 = np.zeros_like(TDB4)
for i in range(TDB4.shape[0]):
    for j in range(TDB4.shape[1]):
        result = pmv_ppd_iso(
            tdb=TDB4[i, j],
            tr=TR4[i, j],
            vr=0.0,  # velocidad del aire fija en 0 m/s
            rh=rh,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        )
        PMV4[i, j] = result.pmv

fig4, ax4 = plt.subplots(figsize=(6, 4))
CS4 = ax4.contour(TDB4, TR4, PMV4, levels=[-1, 0, 1], colors=['red', 'blue', 'red'])
ax4.clabel(CS4, inline=True, fontsize=10)
ax4.set_xlabel("Temperatura de Bulbo Seco (°C)")
ax4.set_ylabel("Temperatura Radiante (°C)")
ax4.set_title("Contornos de PMV = -1, 0 y 1 (v = 0 m/s)")
st.pyplot(fig4)

###############################
# 5. Gráfico: Curva de PMV = 0 (Extracción numérica y Regresión Lineal)
###############################
st.header("5. Curva de PMV = 0 (Extracción numérica y Regresión Lineal)")

# Función que retorna PMV para un tdb y tr dados, con v = 0 m/s
def f(tr, tdb):
    result = pmv_ppd_iso(
        tdb=tdb,
        tr=tr,
        vr=0.0,  # velocidad del aire en 0 m/s
        rh=rh,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    )
    return result.pmv

# Rango de tdb entre 18 y 35°C
tdb_range = np.linspace(18, 35, 100)
tr_neutral = []

# Para cada valor de tdb, se busca el valor de tr que hace que PMV = 0
for tdb in tdb_range:
    try:
        tr_root = brentq(f, 18, 35, args=(tdb,))
        tr_neutral.append(tr_root)
    except ValueError:
        tr_neutral.append(np.nan)

# Convertir a arrays y eliminar NaN para la regresión
tdb_array = np.array(tdb_range)
tr_array = np.array(tr_neutral)
mask = ~np.isnan(tr_array)
tdb_valid = tdb_array[mask]
tr_valid = tr_array[mask]

# Regresión lineal simple: tr = slope * tdb + intercept
slope, intercept = np.polyfit(tdb_valid, tr_valid, 1)
regression_line = slope * tdb_valid + intercept

fig5, ax5 = plt.subplots(figsize=(6, 4))
ax5.plot(tdb_range, tr_neutral, 'b-', label="PMV = 0 (numérico)")
ax5.plot(tdb_valid, regression_line, 'r--',
         label=f"Regresión: tr = {slope:.2f} * tdb + {intercept:.2f}")
ax5.set_xlabel("Temperatura de Bulbo Seco (°C)")
ax5.set_ylabel("Temperatura Radiante (°C)")
ax5.set_title("Curva de PMV = 0 (v = 0 m/s)")
ax5.legend()
ax5.grid(True)
st.pyplot(fig5)

st.write(f"La fórmula de la curva de PMV = 0 es: tr = {slope:.2f} * tdb + {intercept:.2f}")

###############################
# 6. Gráfico: Intersección de la curva PMV = 0 con la línea tr = tdb
###############################
st.header("6. Intersección: PMV = 0 vs tr = tdb")

# Convertir a arrays (ya definidos en el Gráfico 5)
tdb_array = np.array(tdb_range)
tr_array = np.array(tr_neutral)

# Calcular la diferencia absoluta entre la curva PMV = 0 y la línea tr = tdb
diferencia = np.abs(tr_array - tdb_array)
indice_min = np.argmin(diferencia)
tdb_igual = tdb_array[indice_min]
tr_igual = tr_array[indice_min]

st.write(f"El punto de intersección se encuentra aproximadamente en: tdb = {tdb_igual:.2f} °C y tr = {tr_igual:.2f} °C")

fig6, ax6 = plt.subplots(figsize=(6, 4))
ax6.plot(tdb_range, tr_neutral, 'b-', label="Curva PMV = 0")
ax6.plot(tdb_range, tdb_range, 'k--', label="tr = tdb")
ax6.plot(tdb_igual, tr_igual, 'ro', markersize=8, label=f"Intersección ({tdb_igual:.2f}°C)")
ax6.set_xlabel("Temperatura de Bulbo Seco (°C)")
ax6.set_ylabel("Temperatura Radiante (°C)")
ax6.set_title("Intersección: PMV = 0 y tr = tdb")
ax6.legend()
ax6.grid(True)
st.pyplot(fig6)



###############################
# 7. Cálculo Exacto de la Intersección: PMV = 0 y tdb = tr
###############################
st.header("7. Cálculo Exacto de la Intersección: PMV = 0 y tdb = tr")

# Definir función para calcular PMV cuando tdb = tr
def f_tdb(t):
    result = pmv_ppd_iso(
        tdb=t,
        tr=t,
        vr=0.0,  # velocidad del aire en 0 m/s
        rh=rh,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    )
    return result.pmv

# Usar brentq para encontrar la raíz en el intervalo [18, 35] °C
tdb_intersection = brentq(f_tdb, 18, 35)
st.write(f"El punto exacto donde tdb = tr y PMV = 0 es: {tdb_intersection:.2f} °C")


