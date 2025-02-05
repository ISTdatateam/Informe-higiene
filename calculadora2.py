import streamlit as st
from pythermalcomfort.models import pmv_ppd_iso

# Título de la aplicación
st.title("Calculadora de PMV y PPD")
st.write("Ajusta los parámetros para ver los resultados recalculados automáticamente.")

# Entradas de usuario con valores por defecto (variables de valor fijo)
tdb = st.number_input("Temperatura de bulbo seco (°C):", value=30.0)
tr = st.number_input("Temperatura radiante (°C):", value=30.0)
rh = st.number_input("Humedad relativa (%):", value=32.0)
v = st.number_input("Velocidad del aire (m/s):", value=0.8)
met = st.number_input("Tasa metabólica (met):", value=1.1)
clo_dynamic = st.number_input("Aislamiento de la ropa (clo):", value=0.5)

# Aquí calculamos PMV y PPD de forma automática (SIN botón)

# 2. Calcular PMV y PPD
results = pmv_ppd_iso(
    tdb=tdb,
    tr=tr,
    vr=v,
    rh=rh,
    met=met,
    clo=clo_dynamic,
    model="7730-2005",
    limit_inputs=False,
    round_output=True
)

# Mostrar los resultados en pantalla
st.subheader("Resultados")
st.write(f"**PMV:** {results.pmv}")
st.write(f"**PPD:** {results.ppd}")
