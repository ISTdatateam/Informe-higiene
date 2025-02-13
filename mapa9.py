import streamlit as st
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq
import numpy as np

st.header("Sistema de Control de Confort Térmico Inteligente")


def calcular_ajuste_optimo(tdb_initial, tr_initial, vr, rh, met, clo, target_pmv, max_iter=20):
    tdb_adj = tdb_initial
    tr_adj = tr_initial
    historial = []

    for i in range(max_iter):
        try:
            pmv = pmv_ppd_iso(
                tdb=tdb_adj, tr=tr_adj, vr=vr,
                rh=rh, met=met, clo=clo,
                model="7730-2005", limit_inputs=False, round_output=False
            ).pmv
        except:
            pmv = np.nan

        if -1 <= pmv <= 1 or np.isnan(pmv):
            break

        # Determinar dirección del ajuste y límites dinámicos
        necesita_calentar = pmv < target_pmv  # PMV demasiado bajo (ambiente frío)

        # Rangos seguros para evitar valores extremos
        if necesita_calentar:
            tdb_bounds = (tdb_adj, min(30, tdb_adj + 10))  # Máximo 30°C
            tr_bounds = (tr_adj, min(30, tr_adj + 10))
        else:
            tdb_bounds = (max(16, tdb_adj - 10), tdb_adj)  # Mínimo 16°C
            tr_bounds = (max(16, tr_adj - 10), tr_adj)

        # Cálculo de candidatos con protección contra errores
        try:
            candidate2 = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=x, tr=x, vr=vr, rh=rh,
                    met=met, clo=clo).pmv - target_pmv,
                *tdb_bounds, xtol=0.5
            )
        except:
            candidate2 = (tdb_adj + tr_adj) / 2

        try:
            candidate1_Tdb = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=x, tr=tr_adj, vr=vr, rh=rh,
                    met=met, clo=clo).pmv - target_pmv,
                *tdb_bounds, xtol=0.5
            )
        except:
            candidate1_Tdb = tdb_adj

        try:
            candidate1_Tr = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=tdb_adj, tr=x, vr=vr, rh=rh,
                    met=met, clo=clo).pmv - target_pmv,
                *tr_bounds, xtol=0.5
            )
        except:
            candidate1_Tr = tr_adj

        # Ajuste simultáneo de ambos parámetros con diferente intensidad
        factor_tdb = 0.7 if necesita_calentar else 0.5
        factor_tr = 0.7 if necesita_calentar else 0.5

        tdb_adj += factor_tdb * (candidate1_Tdb - tdb_adj)
        tr_adj += factor_tr * (candidate1_Tr - tr_adj)

        # Forzar diferencias mínimas
        tdb_adj = round(tdb_adj, 1)
        tr_adj = round(tr_adj, 1)

        historial.append({
            'Iteración': i + 1,
            'Tdb': tdb_adj,
            'Tr': tr_adj,
            'PMV': round(pmv, 2) if not np.isnan(pmv) else 'Error'
        })

    return tdb_adj, tr_adj, historial


with st.form("main_form"):
    st.subheader("Parámetros Ambientales")
    col1, col2 = st.columns(2)

    with col1:
        tdb = st.number_input("Temperatura del Aire (°C)", value=19.4, step=0.1)
        tr = st.number_input("Temperatura Radiante (°C)", value=19.6, step=0.1)
        vr = st.number_input("Velocidad del Aire (m/s)", value=0.26, step=0.01)

    with col2:
        rh = st.number_input("Humedad Relativa (%)", value=52, step=1)
        met = st.number_input("Tasa Metabólica (met)", value=1.2, step=0.1)
        clo = st.number_input("Aislamiento (CLO)", value=0.5, step=0.1)

    submit = st.form_submit_button("Optimizar Confort")

if submit:
    # Cálculo inicial y validación
    try:
        pmv_initial = pmv_ppd_iso(
            tdb=tdb, tr=tr, vr=vr, rh=rh,
            met=met, clo=clo, limit_inputs=False
        ).pmv
    except:
        st.error("Error en cálculo inicial de PMV. Verifique los valores ingresados.")
        st.stop()

    # Ajuste de ventilación para frío extremo
    if pmv_initial < -1.0 and vr < 0.1:
        vr_ajustado = 0.05  # Reducir corrientes de aire
    else:
        vr_ajustado = vr

    target_pmv = -1.0 if pmv_initial < -1 else 1.0

    # Ajuste térmico
    tdb_final, tr_final, historial = calcular_ajuste_optimo(
        tdb_initial=tdb,
        tr_initial=tr,
        vr=vr_ajustado,
        rh=rh,
        met=met,
        clo=clo,
        target_pmv=target_pmv
    )

    # Cálculo final
    try:
        pmv_final = pmv_ppd_iso(
            tdb=tdb_final, tr=tr_final, vr=vr_ajustado,
            rh=rh, met=met, clo=clo
        ).pmv
    except:
        pmv_final = np.nan

    # Visualización
    st.header("📈 Proceso de Optimización")

    if vr != vr_ajustado:
        st.subheader("Ajuste Inicial de Ventilación")
        cols = st.columns(2)
        cols[0].metric("Velocidad Aire Original", f"{vr} m/s")
        cols[1].metric("Nueva Velocidad Aire", f"{vr_ajustado} m/s")

    st.subheader("Evolución de Ajustes Térmicos")

    if historial:
        chart_data = {
            'Iteración': [x['Iteración'] for x in historial],
            'PMV': [x['PMV'] if isinstance(x['PMV'], float) else np.nan for x in historial],
            'Tdb': [x['Tdb'] for x in historial],
            'Tr': [x['Tr'] for x in historial]
        }

        st.line_chart(chart_data, x='Iteración', y=['PMV'], height=300)
        st.line_chart(chart_data, x='Iteración', y=['Tdb', 'Tr'], height=300)
    else:
        st.warning("No se realizaron ajustes térmicos")

    st.header("📋 Resultado Final")
    cols = st.columns(3)

    try:
        eficiencia = (abs(pmv_initial - pmv_final) / abs(pmv_initial)) * 100
    except:
        eficiencia = 0

    cols[0].metric("PMV Inicial", f"{pmv_initial:.2f}")
    cols[1].metric("PMV Final",
                   f"{pmv_final:.2f}" if not np.isnan(pmv_final) else "Error",
                   delta="✅ Confort" if -1 <= pmv_final <= 1 else "⚠️ Fuera de rango")
    cols[2].metric("Eficiencia",
                   f"{eficiencia:.0f}%" if not np.isnan(eficiencia) else "N/A")

    st.subheader("Recomendaciones de Ajuste")
    col_rec = st.columns(2)

    with col_rec[0]:
        st.write("**Temperatura del Aire**")
        st.metric("Ajuste requerido", f"{tdb} → {tdb_final}°C",
                  delta=f"{tdb_final - tdb:+.1f}°C")
        st.write("**Acciones:**")
        if tdb_final > tdb:
            st.write("- Activar sistema de calefacción")
            st.write("- Mejorar aislamiento térmico")
        else:
            st.write("- Optimizar sistema de refrigeración")

    with col_rec[1]:
        st.write("**Temperatura Radiante**")
        st.metric("Ajuste requerido", f"{tr} → {tr_final}°C",
                  delta=f"{tr_final - tr:+.1f}°C")
        st.write("**Acciones:**")
        if tr_final > tr:
            st.write("- Instalar paneles radiantes")
            st.write("- Reducir pérdidas térmicas")
        else:
            st.write("- Usar materiales reflectantes")
