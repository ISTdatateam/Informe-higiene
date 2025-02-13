import streamlit as st
import numpy as np
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq

st.header("Recomendación para Ingresar al Área de Confort (PMV entre -1 y +1)")

st.write("""
Este módulo permite ingresar las condiciones medidas de un ambiente (tdb, tr, humedad, velocidad del aire y met) y propone, de forma individual, 
los valores a los que cada parámetro debería ajustarse para alcanzar PMV = 0 (aproximadamente el centro de la zona de confort). 
Se recomienda intervenir la variable que se encuentre más alejada de su valor objetivo, respetando que la velocidad del aire no supere 1 m/s.
""")

# Parámetros fijos o ingresables
clo = 0.5
met = st.number_input("Tasa Metabólica (met):", value=1.2, step=0.1)

st.subheader("Ingreso de Variables Medidas")
tdb_measured = st.number_input("Temperatura de Bulbo Seco (tdb) (°C):", value=26.0, step=0.1)
tr_measured = st.number_input("Temperatura Radiante (tr) (°C):", value=26.0, step=0.1)
rh_measured = st.number_input("Humedad Relativa (%) :", value=40.0, step=1.0)
v_measured = st.number_input("Velocidad del Aire (vr) (m/s):", value=0.3, min_value=0.0, max_value=1.0, step=0.1)

# Calcular PMV con las condiciones medidas
pmv_current = pmv_ppd_iso(
    tdb=tdb_measured,
    tr=tr_measured,
    vr=v_measured,
    rh=rh_measured,
    met=met,
    clo=clo,
    model="7730-2005",
    limit_inputs=False,
    round_output=True
).pmv
st.write(f"**PMV actual:** {pmv_current:.2f}")

# Funciones para ajustar cada variable individualmente (las demás se mantienen fijas)
def f_adj_tdb(x):
    """Ajuste de tdb: se varía tdb manteniendo tr, vr, rh fijos."""
    return pmv_ppd_iso(
        tdb=x,
        tr=tr_measured,
        vr=v_measured,
        rh=rh_measured,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    ).pmv

def f_adj_tr(x):
    """Ajuste de tr: se varía tr manteniendo tdb, vr, rh fijos."""
    return pmv_ppd_iso(
        tdb=tdb_measured,
        tr=x,
        vr=v_measured,
        rh=rh_measured,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    ).pmv

def f_adj_rh(x):
    """Ajuste de humedad: se varía rh (en %) manteniendo tdb, tr, vr fijos."""
    return pmv_ppd_iso(
        tdb=tdb_measured,
        tr=tr_measured,
        vr=v_measured,
        rh=x,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    ).pmv

def f_adj_v(x):
    """Ajuste de velocidad: se varía vr manteniendo tdb, tr, rh fijos.
       x se busca en [0, 1]."""
    return pmv_ppd_iso(
        tdb=tdb_measured,
        tr=tr_measured,
        vr=x,
        rh=rh_measured,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=False
    ).pmv

# Resolver individualmente para PMV = 0
# Se supone que al ajustar cada variable individualmente se busca PMV=0
try:
    tdb_target = brentq(f_adj_tdb, 18, 35)
except ValueError:
    tdb_target = np.nan

try:
    tr_target = brentq(f_adj_tr, 18, 35)
except ValueError:
    tr_target = np.nan

# Para humedad se busca en un rango razonable (por ejemplo, 20% a 80%)
try:
    rh_target = brentq(f_adj_rh, 20, 80)
except ValueError:
    rh_target = np.nan

# Para velocidad, se busca en [0,1] (recordando la limitante)
try:
    v_target = brentq(f_adj_v, 0.0, 1.0)
except ValueError:
    v_target = np.nan

# Calcular las desviaciones (valor medido - valor objetivo)
diff_tdb = tdb_measured - tdb_target if not np.isnan(tdb_target) else None
diff_tr = tr_measured - tr_target if not np.isnan(tr_target) else None
diff_rh = rh_measured - rh_target if not np.isnan(rh_target) else None
diff_v  = v_measured - v_target   if not np.isnan(v_target) else None

st.write("### Valores Objetivo (para PMV=0) y Diferencias")
if tdb_target is not np.nan:
    st.write(f"**Tdb:** Medido = {tdb_measured:.2f} °C, Objetivo = {tdb_target:.2f} °C, Diferencia = {diff_tdb:+.2f} °C")
if tr_target is not np.nan:
    st.write(f"**Tr:** Medido = {tr_measured:.2f} °C, Objetivo = {tr_target:.2f} °C, Diferencia = {diff_tr:+.2f} °C")
if rh_target is not np.nan:
    st.write(f"**Humedad (%):** Medido = {rh_measured:.2f} %, Objetivo = {rh_target:.2f} %, Diferencia = {diff_rh:+.2f} %")
if v_target is not np.nan:
    st.write(f"**Velocidad (m/s):** Medido = {v_measured:.2f} m/s, Objetivo = {v_target:.2f} m/s, Diferencia = {diff_v:+.2f} m/s")

# Determinar cuál variable presenta la mayor desviación (valor absoluto)
deviations = {
    "Tdb": abs(diff_tdb) if diff_tdb is not None else 0,
    "Tr": abs(diff_tr) if diff_tr is not None else 0,
    "Humedad": abs(diff_rh) if diff_rh is not None else 0,
    "Velocidad": abs(diff_v) if diff_v is not None else 0
}
var_critica = max(deviations, key=deviations.get)

st.write("### Recomendación de Intervención")
st.write(f"El parámetro que presenta la mayor desviación es **{var_critica}**.")
st.write("Se recomienda intervenir, preferentemente, este parámetro para acercar las condiciones al área de confort (PMV entre -1 y +1).")
st.write("""
**Sugerencias:**
- Si **Tdb** está muy alejada del valor objetivo, se debe ajustar el sistema HVAC para modificar la temperatura del aire.
- Si **Tr** difiere significativamente, se debe considerar modificar la envolvente del edificio (aislamiento, radiación, etc.).
- Si la **Humedad** está fuera de rango, se puede intervenir con humidificadores o deshumidificadores.
- Si la **Velocidad del Aire** difiere, se deben ajustar los sistemas de ventilación, siempre respetando que no supere 1 m/s.
""")
