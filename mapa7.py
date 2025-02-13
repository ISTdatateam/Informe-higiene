import streamlit as st
import numpy as np
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq

st.header("Propuesta de Ajuste Simplificado: Algoritmo de Reducción de Temperatura (50% de Corrección)")

st.write("""
Este algoritmo tiene como objetivo alcanzar un PMV de 0.99 mediante ajustes en las temperaturas.
Para ello se calcula:

- **Candidate 2:** Se resuelve la ecuación para obtener el valor que produce PMV = 0.99 asumiendo Tdb = Tr.
- **Candidate 1:** Se toma el parámetro con el mayor valor actual (Tdb o Tr) y se resuelve la ecuación, dejando fija la otra, para obtener el valor que produce PMV = 0.99.

Luego se define:

      target_avg = (Candidate 1 + Candidate 2) / 2

y se actualiza la variable (Tdb o Tr) que esté más alejada del valor actual respecto a target_avg aplicando una corrección del 50%:

      Nuevo Valor = Valor Actual - 0.5*(Valor Actual - target_avg)

**Importante:**
- El PMV inicial se calcula con la velocidad del aire ingresada.
- Para el proceso iterativo, si la velocidad del aire es menor a 0.5 m/s, se fuerza a 0.5 m/s.

**Recomendaciones:**
- Si se debe ajustar Tdb, revise el sistema HVAC para modificar la temperatura del aire.
- Si se debe ajustar Tr, se recomienda mejorar el aislamiento o controlar la radiación solar.
- Mantener una velocidad del aire por encima de 0.5 m/s favorece el confort.
""")

with st.form("form_ajuste_simplificado"):
    clo = 0.5  # Valor fijo
    met = st.number_input("Tasa Metabólica (met):", value=1.2, step=0.1)

    st.subheader("Ingreso de Variables Medidas")
    tdb_measured = st.number_input("Temperatura de Bulbo Seco (Tdb) (°C):", value=29.60, step=0.1)
    tr_measured = st.number_input("Temperatura Radiante (Tr) (°C):", value=30.30, step=0.1)
    rh_measured = st.number_input("Humedad Relativa (%):", value=38.0, step=1.0)
    v_measured = st.number_input("Velocidad del Aire (Vr) (m/s):", value=0.20, min_value=0.0, max_value=1.0, step=0.1)

    submit = st.form_submit_button("Calcular Propuesta de Ajuste")

if submit:
    # Calcular PMV actual utilizando la velocidad del aire entregada
    pmv_actual = pmv_ppd_iso(
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
    st.write(f"**PMV actual (con Vr = {v_measured:.2f} m/s):** {pmv_actual:.2f}")

    # Para el proceso iterativo, forzamos Vr a 0.5 m/s si el valor entregado es menor
    v_adj = v_measured if v_measured >= 0.5 else 0.5

    st.write("---")
    st.header("Ajuste Iterativo para Alcanzar PMV = 0.99")
    st.write("""
    Se realizarán hasta 10 iteraciones. En cada iteración se:

      1. Calcula **Candidate 2:** el valor x que produce PMV = 0.99 asumiendo Tdb = Tr = x.
      2. Se determina cuál de las dos temperaturas actuales es mayor.
         - Si Tdb es mayor, se calcula **Candidate 1** resolviendo para Tdb (dejando Tr fija).
         - Si Tr es mayor, se calcula Candidate 1 resolviendo para Tr (dejando Tdb fija).
      3. Se define:

             target_avg = (Candidate 1 + Candidate 2) / 2

      4. Se actualiza la variable (Tdb o Tr) que esté más alejada del valor actual respecto a target_avg, aplicando:

             Nuevo Valor = Valor Actual - 0.5*(Valor Actual - target_avg)

    El proceso se repite hasta 10 iteraciones o hasta que |PMV - 0.98| < 0.01.
    """)

    target_pmv = 0.98
    max_iter = 10
    tol = 0.01  # Tolerancia en PMV

    # Inicializamos Tdb y Tr con los valores medidos
    tdb_adj = tdb_measured
    tr_adj = tr_measured

    for i in range(1, max_iter + 1):
        pmv_iter = pmv_ppd_iso(
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
        st.write(f"   PMV = {pmv_iter:.2f}")
        if abs(pmv_iter - target_pmv) < tol:
            st.write("Objetivo alcanzado (dentro de la tolerancia).")
            break

        # Candidate 2: Ajustar ambas temperaturas (Tdb = Tr)
        try:
            candidate2 = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=x,
                    tr=x,
                    vr=v_adj,
                    rh=rh_measured,
                    met=met,
                    clo=clo,
                    model="7730-2005",
                    limit_inputs=False,
                    round_output=False
                ).pmv - target_pmv, 18, 35)
        except ValueError:
            candidate2 = (tdb_adj + tr_adj) / 2.0

        # Determinar cuál de las dos temperaturas actuales es mayor
        if tdb_adj >= tr_adj:
            # Candidate 1: Ajustar Tdb dejando Tr fija
            try:
                candidate1 = brentq(
                    lambda x: pmv_ppd_iso(
                        tdb=x,
                        tr=tr_adj,
                        vr=v_adj,
                        rh=rh_measured,
                        met=met,
                        clo=clo,
                        model="7730-2005",
                        limit_inputs=False,
                        round_output=False
                    ).pmv - target_pmv, 18, 35)
            except ValueError:
                candidate1 = tdb_adj
            param_to_adjust = "Tdb"
        else:
            # Candidate 1: Ajustar Tr dejando Tdb fija
            try:
                candidate1 = brentq(
                    lambda x: pmv_ppd_iso(
                        tdb=tdb_adj,
                        tr=x,
                        vr=v_adj,
                        rh=rh_measured,
                        met=met,
                        clo=clo,
                        model="7730-2005",
                        limit_inputs=False,
                        round_output=False
                    ).pmv - target_pmv, 18, 35)
            except ValueError:
                candidate1 = tr_adj
            param_to_adjust = "Tr"

        target_avg = (candidate1 + candidate2) / 2.0

        st.write(f"   Candidate 1 (ajustando solo {param_to_adjust}): {candidate1:.2f} °C")
        st.write(f"   Candidate 2 (ajustando ambas): {candidate2:.2f} °C")
        st.write(f"   **Target promedio = {target_avg:.2f} °C**")

        # Actualizar la variable que esté más alejada del target_avg
        if param_to_adjust == "Tdb":
            diff = tdb_adj - target_avg
            new_val = tdb_adj - 0.5 * diff  # Aplicar 50% del ajuste
            st.write(f"   Se actualiza Tdb: diferencia = {diff:+.2f} °C, nuevo Tdb = {new_val:.2f} °C")
            tdb_adj = new_val
        else:
            diff = tr_adj - target_avg
            new_val = tr_adj - 0.5 * diff  # Aplicar 50% del ajuste
            st.write(f"   Se actualiza Tr: diferencia = {diff:+.2f} °C, nuevo Tr = {new_val:.2f} °C")
            tr_adj = new_val

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
    st.header("Resultado Final del Ajuste")
    st.write(f"**Valores finales propuestos:**")
    st.write(f"   Tdb = {tdb_adj:.2f} °C")
    st.write(f"   Tr = {tr_adj:.2f} °C")
    st.write(f"   Vr = {v_adj:.2f} m/s")
    st.write(f"**PMV final calculado:** {pmv_final:.2f}")

    st.write("""
    **Recomendaciones:**
      - Si se debe ajustar Tdb, revise el sistema HVAC para modificar la temperatura del aire.
      - Si se debe ajustar Tr, se recomienda mejorar el aislamiento o controlar la radiación solar.
      - Se recomienda mantener la velocidad del aire por encima de 0.5 m/s para favorecer el confort.
    """)
