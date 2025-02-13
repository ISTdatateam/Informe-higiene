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

# Se agrupan los inputs en un formulario
with st.form("form_parametros"):
    clo = 0.5  # Valor fijo
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


    # 2. (Opcional) Se muestran otros cálculos previos...
    # Por ejemplo, la Temperatura Neutral (PMV=0 con tdb = tr)
    def f_tdb(t, rh, v, met):
        return pmv_ppd_iso(
            tdb=t,
            tr=t,
            vr=v,
            rh=rh,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv


    try:
        t_neutral = brentq(f_tdb, 18, 35, args=(rh_measured, v_measured, met))
        st.write(f"**Temperatura Neutral (PMV=0, tdb=tr):** {t_neutral:.2f} °C")
    except ValueError:
        st.write("No se pudo calcular la temperatura neutral para los valores ingresados.")

    st.write("---")
    st.header("Propuesta para Alcanzar PMV = 0.99 (Ajuste Recursivo)")
    st.write("""
    Se aplicará un método iterativo de hasta 3 iteraciones: en cada iteración se identifica la variable (Tdb, Tr o Vr) cuya diferencia
    respecto al valor objetivo para alcanzar PMV = 0.99 es mayor, y se reduce esa diferencia a la mitad. Como norma, si la velocidad del aire es menor a 0.5 m/s,
    se propone fijarla en 0.50 m/s.
    """)

    # Establecer el objetivo PMV
    target_pmv = 0.99
    max_iter = 3

    # Inicializar variables de ajuste con los valores medidos
    tdb_adj = tdb_measured
    tr_adj = tr_measured
    v_adj = v_measured
    # Asegurarse de que v no sea menor a 0.5 m/s (según la norma)
    if v_adj < 0.5:
        v_adj = 0.5


    # Funciones para calcular PMV al ajustar individualmente
    def f_target_tdb(x, tr_val, v_val):
        return pmv_ppd_iso(
            tdb=x,
            tr=tr_val,
            vr=v_val,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv


    def f_target_tr(x, tdb_val, v_val):
        return pmv_ppd_iso(
            tdb=tdb_val,
            tr=x,
            vr=v_val,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv


    def f_target_v(x, tdb_val, tr_val):
        return pmv_ppd_iso(
            tdb=tdb_val,
            tr=tr_val,
            vr=x,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv - target_pmv


    # Iteración recursiva
    for i in range(1, max_iter + 1):
        # Calcular el PMV actual con los valores ajustados
        pmv_current_adj = pmv_ppd_iso(
            tdb=tdb_adj,
            tr=tr_adj,
            vr=v_adj,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv

        st.write(f"**Iteración {i}:**")
        st.write(f"   Tdb = {tdb_adj:.2f} °C, Tr = {tr_adj:.2f} °C, Vr = {v_adj:.2f} m/s")
        st.write(f"   PMV = {pmv_current_adj:.2f}")

        # Si el PMV ya está dentro del rango de confort (-1 a +1), se finaliza el proceso.
        if -1 < pmv_current_adj < 1:
            st.write("El PMV ya se encuentra dentro del rango de confort.")
            break

        # Calcular el valor objetivo para cada variable (dejando las otras fijas) para alcanzar target_pmv
        try:
            tdb_target_iter = brentq(f_target_tdb, 18, 35, args=(tr_adj, v_adj))
        except ValueError:
            tdb_target_iter = tdb_adj
        try:
            tr_target_iter = brentq(f_target_tr, 18, 35, args=(tdb_adj, v_adj))
        except ValueError:
            tr_target_iter = tr_adj
        try:
            v_target_iter = brentq(f_target_v, 0.0, 1.0, args=(tdb_adj, tr_adj))
        except ValueError:
            v_target_iter = v_adj

        # Calcular las diferencias (medido - objetivo)
        diff_tdb_iter = tdb_adj - tdb_target_iter
        diff_tr_iter = tr_adj - tr_target_iter
        diff_v_iter = v_adj - v_target_iter

        # Seleccionar la variable con mayor diferencia absoluta
        differences = {
            "Tdb": abs(diff_tdb_iter),
            "Tr": abs(diff_tr_iter),
            "Vr": abs(diff_v_iter)
        }
        var_to_adjust = max(differences, key=differences.get)

        # Aplicar ajuste: se reduce la diferencia a la mitad
        if var_to_adjust == "Tdb":
            tdb_adj = tdb_adj - diff_tdb_iter / 2
            st.write(f"   Se ajusta Tdb: se reduce la diferencia de {diff_tdb_iter:+.2f} °C a la mitad.")
        elif var_to_adjust == "Tr":
            tr_adj = tr_adj - diff_tr_iter / 2
            st.write(f"   Se ajusta Tr: se reduce la diferencia de {diff_tr_iter:+.2f} °C a la mitad.")
        elif var_to_adjust == "Vr":
            v_adj = v_adj - diff_v_iter / 2
            # Si el ajuste hace que Vr caiga por debajo de 0.5 m/s, forzar a 0.5 m/s
            if v_adj < 0.5:
                v_adj = 0.5
            st.write(f"   Se ajusta Vr: se reduce la diferencia de {diff_v_iter:+.2f} m/s a la mitad.")

    # Mostrar el PMV final y los valores finales después de la iteración
    pmv_final = pmv_ppd_iso(
        tdb=tdb_adj,
        tr=tr_adj,
        vr=v_adj,
        rh=rh_measured,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=True
    ).pmv
    st.write("---")
    st.header("Resultado Final del Ajuste Recursivo")
    st.write(f"**Valores finales propuestos:**")
    st.write(f"   Tdb = {tdb_adj:.2f} °C")
    st.write(f"   Tr = {tr_adj:.2f} °C")
    st.write(f"   Vr = {v_adj:.2f} m/s")
    st.write(f"**PMV final calculado:** {pmv_final:.2f}")

    st.write("""
    Estas modificaciones se proponen de forma iterativa (hasta 3 ajustes) para acercar el PMV al objetivo de 0.99. 
    Se recomienda evaluar estas propuestas en conjunto y considerar intervenciones combinadas en el ambiente.
    """)

