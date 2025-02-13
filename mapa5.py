import streamlit as st
import numpy as np
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq

st.header("Recomendación para Ingresar al Área de Confort (PMV entre -1 y +1)")

st.write("""
Este módulo permite ingresar las condiciones medidas de un ambiente (tdb, tr, humedad, velocidad del aire y met) y propone, de forma individual, 
los valores a los que cada parámetro debería ajustarse para alcanzar un PMV cercano a 0.99 (dentro del área de confort). Se recomienda intervenir la variable 
que se encuentre más alejada del valor objetivo, siempre respetando que la velocidad del aire no supere 1 m/s.
""")

# Usaremos un formulario para que la recalculación se ejecute al presionar el botón.
with st.form("form_parametros"):
    clo = 0.5  # Fijo
    met = st.number_input("Tasa Metabólica (met):", value=1.2, step=0.1)

    st.subheader("Ingreso de Variables Medidas")
    tdb_measured = st.number_input("Temperatura de Bulbo Seco (tdb) (°C):", value=29.60, step=0.1)
    tr_measured = st.number_input("Temperatura Radiante (tr) (°C):", value=30.30, step=0.1)
    rh_measured = st.number_input("Humedad Relativa (%):", value=38.0, step=1.0)
    v_measured = st.number_input("Velocidad del Aire (vr) (m/s):", value=0.20, min_value=0.0, max_value=1.0, step=0.1)

    submit = st.form_submit_button("Recalcular Indicaciones")

if submit:
    # 1. Cálculo del PMV actual
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


    # 2. Calcular la temperatura neutral (PMV = 0 con tdb = tr)
    def f_tdb(t, rh, v, met):
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


    try:
        t_neutral = brentq(f_tdb, 18, 35, args=(rh_measured, v_measured, met))
        st.write(f"**Temperatura Neutral (PMV=0, tdb=tr):** {t_neutral:.2f} °C")
    except ValueError:
        st.write("No se pudo calcular la temperatura neutral para los valores ingresados.")
        t_neutral = None

    # 3. Cálculo de valores objetivo individuales (para PMV = 0) dejando fijas las demás variables
    try:
        tdb_target = brentq(lambda x: pmv_ppd_iso(
            tdb=x,
            tr=tr_measured,
            vr=v_measured,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv, 18, 35)
    except ValueError:
        tdb_target = np.nan

    try:
        tr_target = brentq(lambda x: pmv_ppd_iso(
            tdb=tdb_measured,
            tr=x,
            vr=v_measured,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv, 18, 35)
    except ValueError:
        tr_target = np.nan

    try:
        rh_target = brentq(lambda x: pmv_ppd_iso(
            tdb=tdb_measured,
            tr=tr_measured,
            vr=v_measured,
            rh=x,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv, 20, 80)
    except ValueError:
        rh_target = np.nan

    try:
        v_target = brentq(lambda x: pmv_ppd_iso(
            tdb=tdb_measured,
            tr=tr_measured,
            vr=x,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv, 0.0, 1.0)
    except ValueError:
        v_target = np.nan

    diff_tdb = tdb_measured - tdb_target if not np.isnan(tdb_target) else None
    diff_tr = tr_measured - tr_target if not np.isnan(tr_target) else None
    diff_rh = rh_measured - rh_target if not np.isnan(rh_target) else None
    diff_v = v_measured - v_target if not np.isnan(v_target) else None

    st.write("### Valores Objetivo (para PMV=0) y Diferencias")
    if tdb_target is not np.nan:
        st.write(
            f"**Tdb:** Medido = {tdb_measured:.2f} °C, Objetivo = {tdb_target:.2f} °C, Diferencia = {diff_tdb:+.2f} °C")
    if tr_target is not np.nan:
        st.write(
            f"**Tr:** Medido = {tr_measured:.2f} °C, Objetivo = {tr_target:.2f} °C, Diferencia = {diff_tr:+.2f} °C")
    if rh_target is not np.nan:
        st.write(
            f"**Humedad (%):** Medido = {rh_measured:.2f} %, Objetivo = {rh_target:.2f} %, Diferencia = {diff_rh:+.2f} %")
    if v_target is not np.nan:
        st.write(
            f"**Velocidad (m/s):** Medido = {v_measured:.2f} m/s, Objetivo = {v_target:.2f} m/s, Diferencia = {diff_v:+.2f} m/s")

    deviations = {
        "Tdb": abs(diff_tdb) if diff_tdb is not None else 0,
        "Tr": abs(diff_tr) if diff_tr is not None else 0,
        "Humedad": abs(diff_rh) if diff_rh is not None else 0,
        "Velocidad": abs(diff_v) if diff_v is not None else 0
    }
    var_critica = max(deviations, key=deviations.get)

    st.write("### Recomendación de Intervención para PMV = 0")
    st.write(f"El parámetro que presenta la mayor desviación es **{var_critica}**.")
    st.write(
        "Se recomienda intervenir preferentemente este parámetro para acercar las condiciones al área de confort (PMV entre -1 y +1).")

    st.write("---")
    st.header("Propuesta para Alcanzar PMV = 0.99")
    st.write("""
    Dado que el PMV actual es elevado, se propone como objetivo práctico alcanzar un PMV de 0.99. En esta propuesta se calcularán 
    individualmente los valores a los que deberían ajustarse Tdb, Tr y Vr, manteniendo las demás variables fijas.
    """)
    target_pmv = 0.99

    try:
        tdb_target_new = brentq(lambda x: pmv_ppd_iso(
            tdb=x,
            tr=tr_measured,
            vr=v_measured,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv, 18, 35)
    except ValueError:
        tdb_target_new = np.nan

    try:
        tr_target_new = brentq(lambda x: pmv_ppd_iso(
            tdb=tdb_measured,
            tr=x,
            vr=v_measured,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv, 18, 35)
    except ValueError:
        tr_target_new = np.nan

    try:
        v_target_new = brentq(lambda x: pmv_ppd_iso(
            tdb=tdb_measured,
            tr=tr_measured,
            vr=x,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv, 0.0, 1.0)
    except ValueError:
        v_target_new = np.nan

    st.write("### Valores Objetivo para PMV = 0.99")
    if not np.isnan(tdb_target_new):
        st.write(f"**Tdb:** Medido = {tdb_measured:.2f} °C → Requerido = {tdb_target_new:.2f} °C")
    if not np.isnan(tr_target_new):
        st.write(f"**Tr:** Medido = {tr_measured:.2f} °C → Requerido = {tr_target_new:.2f} °C")
    if not np.isnan(v_target_new):
        st.write(f"**Vr:** Medido = {v_measured:.2f} m/s → Requerido = {v_target_new:.2f} m/s")

    diff_tdb_new = tdb_measured - tdb_target_new if not np.isnan(tdb_target_new) else None
    diff_tr_new = tr_measured - tr_target_new if not np.isnan(tr_target_new) else None
    diff_v_new = v_measured - v_target_new if not np.isnan(v_target_new) else None

    st.write("### Diferencias para alcanzar PMV = 0.99")
    if diff_tdb_new is not None:
        st.write(f"Diferencia en Tdb: {diff_tdb_new:+.2f} °C")
    if diff_tr_new is not None:
        st.write(f"Diferencia en Tr: {diff_tr_new:+.2f} °C")
    if diff_v_new is not None:
        st.write(f"Diferencia en Vr: {diff_v_new:+.2f} m/s")

    # --- Nueva Propuesta: Distribución de Ajustes ---
    st.write("---")
    st.header("Propuesta Distribuida de Ajustes")
    st.write("""
    Dado que en algunos casos la diferencia es muy grande, se propone repartir el ajuste entre las variables. Como norma, 
    se recomienda elevar la velocidad del aire al menos a 0.50 m/s (si no lo está) y luego ajustar Tdb y Tr.
    """)
    # Si la velocidad medida es menor a 0.50, se propone fijarla en 0.50 m/s
    if v_measured < 0.5:
        v_target_dist = 0.5
        st.write(f"**Velocidad del Aire:** Se propone aumentar la velocidad de {v_measured:.2f} m/s a 0.50 m/s.")
    else:
        v_target_dist = v_measured
        st.write(f"**Velocidad del Aire:** La velocidad actual de {v_measured:.2f} m/s se mantiene.")

    # Con Vr ajustado a v_target_dist, recalcular objetivos para Tdb y Tr para alcanzar PMV = 0.99
    try:
        tdb_target_dist = brentq(lambda x: pmv_ppd_iso(
            tdb=x,
            tr=tr_measured,
            vr=v_target_dist,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv, 18, 35)
    except ValueError:
        tdb_target_dist = np.nan

    try:
        tr_target_dist = brentq(lambda x: pmv_ppd_iso(
            tdb=tdb_measured,
            tr=x,
            vr=v_target_dist,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv, 18, 35)
    except ValueError:
        tr_target_dist = np.nan

    diff_tdb_dist = tdb_measured - tdb_target_dist if not np.isnan(tdb_target_dist) else None
    diff_tr_dist = tr_measured - tr_target_dist if not np.isnan(tr_target_dist) else None

    if not np.isnan(tdb_target_dist):
        st.write(
            f"**Tdb:** Se propone reducir de {tdb_measured:.2f} °C a {tdb_target_dist:.2f} °C (diferencia: {diff_tdb_dist:+.2f} °C)")
    if not np.isnan(tr_target_dist):
        st.write(
            f"**Tr:** Se propone reducir de {tr_measured:.2f} °C a {tr_target_dist:.2f} °C (diferencia: {diff_tr_dist:+.2f} °C)")

    st.write("### Recomendación Final Distribuida")
    st.write("""
    Se recomienda en primer lugar elevar la velocidad del aire a 0.50 m/s (si es menor), y luego ajustar tanto la Temperatura de Bulbo Seco (Tdb) 
    como la Temperatura Radiante (Tr) para acercar el PMV al objetivo de 0.99. Este enfoque distribuido permite no cargar todo el ajuste en una sola variable.
    """)
