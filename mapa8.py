import streamlit as st
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq

st.header("Propuesta de Ajuste para Ambientes Fríos/Calurosos: Alcanzar |PMV| = 0.98")

st.write("""
Este algoritmo tiene como objetivo alcanzar un PMV cuyo valor absoluto sea 0.98 (±0.01) mediante ajustes en las temperaturas (Tdb y Tr), sin modificar la velocidad del aire (Vr).

Se calculan dos candidatos:

- **Candidate 2:** Se resuelve la ecuación para obtener el valor que produce PMV = target, asumiendo Tdb = Tr = x.
- **Candidate1_Tdb:** Se resuelve la ecuación para obtener el valor que produce PMV = target al variar Tdb (dejando Tr fija).
- **Candidate1_Tr:** Se resuelve la ecuación para obtener el valor que produce PMV = target al variar Tr (dejando Tdb fija).

Se calculan las diferencias:

    diff_Tdb = |Tdb_actual - Candidate1_Tdb|
    diff_Tr  = |Tr_actual - Candidate1_Tr|

Y se selecciona el parámetro (Tdb o Tr) cuya diferencia sea mayor.
Luego se define:

      target_avg = (Candidate1 + Candidate2) / 2

y se actualiza el parámetro seleccionado aplicando:

      Nuevo Valor = Valor Actual + 0.66 * (target_avg - Valor Actual)

Esto incrementa la temperatura si target_avg es mayor (para ambientes fríos) o la reduce si es menor (para ambientes calurosos).

El target de PMV se define según el PMV actual:
  - Si PMV_actual < 0, target = –0.98
  - Si PMV_actual ≥ 0, target = 0.98

**Importante:**
- El PMV inicial se calcula con la velocidad del aire entregada.
- En este ajuste no se modifica la velocidad del aire.

**Recomendaciones:**
  - Si se debe ajustar Tdb, revise el sistema HVAC para modificar la temperatura del aire.
  - Si se debe ajustar Tr, se recomienda mejorar el aislamiento o controlar la radiación solar.
""")


def generar_recomendaciones(pmv, tdb, tr, vr, rh, met, clo):
    recomendaciones = []

    # 1. Análisis de Velocidad del Aire (VR)
    if pmv > 1.0 and vr < 0.5:
        recomendaciones.append({
            'tipo': 'ventilacion',
            'mensaje': f"Aumentar velocidad del aire a 0.5 m/s usando ventiladores (VR actual: {vr} m/s)",
            'accion': {'vr': 0.5}
        })
    elif pmv < -1.0 and vr < 0.1:
        recomendaciones.append({
            'tipo': 'ventilacion',
            'mensaje': "Reducir corrientes de aire frío (VR muy baja puede aumentar sensación de frío)",
            'accion': {'vr': 0.05}
        })

    # 2. Relación entre Temperatura Radiante (Tr) y del Aire (Tdb)
    dif_temp = tr - tdb
    if dif_temp > 3.0:
        recomendaciones.append({
            'tipo': 'temperatura_radiante',
            'mensaje': f"Temperatura radiante elevada (Tr-Tdb = {dif_temp:.1f}°C). Medidas sugeridas:",
            'acciones': [
                "Instalar superficies reflectantes",
                "Mejorar aislamiento térmico",
                "Controlar fuentes de radiación (ventanas, equipos)"
            ]
        })
    elif dif_temp < -2.0:
        recomendaciones.append({
            'tipo': 'temperatura_radiante',
            'mensaje': f"Temperatura radiante baja (Tr-Tdb = {dif_temp:.1f}°C). Medidas sugeridas:",
            'acciones': [
                "Mejorar aislamiento de paredes/ventanas",
                "Considerar calefacción radiante"
            ]
        })

    # 3. Estrategias según dirección del ajuste
    if pmv > 1.0:  # Ambiente caluroso
        recomendaciones.append({
            'tipo': 'enfriamiento',
            'mensaje': "Estrategias de enfriamiento:",
            'acciones': [
                f"Reducir Tdb actual ({tdb}°C) mediante HVAC",
                f"Reducir Tr ({tr}°C) con sombreado/aislamiento",
                "Ventilación cruzada para aumentar disipación térmica"
            ]
        })
    elif pmv < -1.0:  # Ambiente frío
        recomendaciones.append({
            'tipo': 'calefaccion',
            'mensaje': "Estrategias de calentamiento:",
            'acciones': [
                f"Aumentar Tdb ({tdb}°C) mediante calefacción",
                f"Aumentar Tr ({tr}°C) con superficies radiantes",
                "Reducir infiltraciones de aire frío"
            ]
        })

    # 4. Factores adicionales
    if met < 1.2:
        recomendaciones.append({
            'tipo': 'actividad',
            'mensaje': f"Actividad metabólica baja ({met} met). Considerar:",
            'acciones': [
                "Pausas activas para aumentar movimiento",
                "Adecuar vestimenta (actual CLO = {clo})"
            ]
        })

    if rh > 70:
        recomendaciones.append({
            'tipo': 'humedad',
            'mensaje': f"Humedad relativa elevada ({rh}%). Acciones:",
            'acciones': [
                "Usar deshumidificadores",
                "Mejorar ventilación natural/mecánica"
            ]
        })

    return recomendaciones




with st.form("form_ajuste_target"):
    clo = 0.5
    met = st.number_input("Tasa Metabólica (met):", value=1.10, step=0.1)

    st.subheader("Ingreso de Variables Medidas")
    tdb_measured = st.number_input("Temperatura de Bulbo Seco (Tdb) (°C):", value=19.70, step=0.1)
    tr_measured = st.number_input("Temperatura Radiante (Tr) (°C):", value=20.40, step=0.1)
    rh_measured = st.number_input("Humedad Relativa (%):", value=53.0, step=1.0)
    v_measured = st.number_input("Velocidad del Aire (Vr) (m/s):", value=0.00, min_value=0.0, max_value=1.0, step=0.1)

    submit = st.form_submit_button("Calcular Propuesta de Ajuste")

if submit:
    # Calcular PMV actual con Vr ingresado
    v_used = v_measured
    pmv_actual = pmv_ppd_iso(
        tdb=tdb_measured,
        tr=tr_measured,
        vr=v_used,
        rh=rh_measured,
        met=met,
        clo=clo,
        model="7730-2005",
        limit_inputs=False,
        round_output=True
    ).pmv

    # Aplicar ajuste automático de Vr si es necesario
    if pmv_actual > 1.0 and v_used < 0.5:
        v_used = 0.5
        st.write("## Ajuste Automático de Velocidad del Aire")
        st.write(f"**PMV inicial = {pmv_actual:.2f} > 1.0** y **Vr = {v_measured:.2f} < 0.5 m/s**")
        st.write("Se aplica **Vr = 0.50 m/s** para mejorar el enfriamiento.")

        # Recalcular PMV con nuevo Vr
        pmv_actual = pmv_ppd_iso(
            tdb=tdb_measured,
            tr=tr_measured,
            vr=v_used,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=True
        ).pmv
        st.write(f"**Nuevo PMV con Vr ajustado:** {pmv_actual:.2f}")
        st.write("---")

    st.write(f"**PMV actual (con Vr = {v_used:.2f} m/s):** {pmv_actual:.2f}")

    # Verificar si está en zona de confort
    if -1 <= pmv_actual <= 1:
        st.success("**El PMV está dentro de la zona de confort (-1 ≤ PMV ≤ 1). No se requieren ajustes adicionales.**")
        st.stop()  # Detener la ejecución aquí

    # Definir target_pmv según posición fuera de la zona de confort
    if pmv_actual < -1:
        target_pmv = -0.98  # Ajustar hacia el límite inferior del confort
        st.write("**PMV < -1:** Ajustando hacia zona de confort (Target PMV = -0.98 +/-0.01)")
    else:
        target_pmv = 0.98  # Ajustar hacia el límite superior del confort
        st.write("**PMV > 1:** Ajustando hacia zona de confort (Target PMV = 0.98 +/-0.01)")


    st.write("---")
    st.header("Ajuste Iterativo para Alcanzar PMV = Target")
    st.write("""
    Se realizarán hasta 20 iteraciones. En cada iteración se:

      1. Calcula **Candidate 2:** el valor x que produce PMV = target, asumiendo Tdb = Tr = x.
      2. Calcula **Candidate1_Tdb:** resolviendo para Tdb (dejando Tr fija).
      3. Calcula **Candidate1_Tr:** resolviendo para Tr (dejando Tdb fija).
      4. Se calculan:
             diff_Tdb = |Tdb_actual - Candidate1_Tdb|
             diff_Tr  = |Tr_actual  - Candidate1_Tr|
         y se selecciona el candidato individual (Candidate1) correspondiente al parámetro con mayor diferencia.
      5. Se define:

             target_avg = (Candidate1 + Candidate2) / 2

      6. Se actualiza el parámetro seleccionado aplicando:

             Nuevo Valor = Valor Actual + 0.66*(target_avg - Valor Actual)

    El proceso se repite hasta 20 iteraciones o hasta que |PMV – target| < 0.01.
    """)

    max_iter = 20
    tol = 0.01

    # Inicializamos con los valores medidos
    tdb_adj = tdb_measured
    tr_adj = tr_measured

    for i in range(1, max_iter + 1):
        pmv_iter = pmv_ppd_iso(
            tdb=tdb_adj,
            tr=tr_adj,
            vr=v_used,
            rh=rh_measured,
            met=met,
            clo=clo,
            model="7730-2005",
            limit_inputs=False,
            round_output=False
        ).pmv
        st.write(f"**Iteración {i}:**")
        st.write(f"   Tdb = {tdb_adj:.2f} °C, Tr = {tr_adj:.2f} °C, Vr = {v_measured:.2f} m/s")
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
                    vr=v_used,
                    rh=rh_measured,
                    met=met,
                    clo=clo,
                    model="7730-2005",
                    limit_inputs=False,
                    round_output=False
                ).pmv - target_pmv, 10, 50)
        except ValueError as e:
            st.error(f"Error en Candidate2: {e}. Usando promedio actual.")
            candidate2 = (tdb_adj + tr_adj) / 2.0

        # Candidate 1 para Tdb
        try:
            candidate1_Tdb = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=x,
                    tr=tr_adj,
                    vr=v_used,
                    rh=rh_measured,
                    met=met,
                    clo=clo,
                    model="7730-2005",
                    limit_inputs=False,
                    round_output=False
                ).pmv - target_pmv, 10, 50)
        except ValueError as e:
            st.error(f"Error en Candidate1_Tdb: {e}. Usando Tdb actual.")
            candidate1_Tdb = tdb_adj

        # Candidate 1 para Tr
        try:
            candidate1_Tr = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=tdb_adj,
                    tr=x,
                    vr=v_used,
                    rh=rh_measured,
                    met=met,
                    clo=clo,
                    model="7730-2005",
                    limit_inputs=False,
                    round_output=False
                ).pmv - target_pmv, 10, 50)
        except ValueError as e:
            st.error(f"Error en Candidate1_Tr: {e}. Usando Tr actual.")
            candidate1_Tr = tr_adj

        st.write(f"   Candidate1_Tdb: {candidate1_Tdb:.2f} °C, Candidate1_Tr: {candidate1_Tr:.2f} °C")

        # Determinar parámetro a ajustar
        if i == 1:
            # Primera iteración: seleccionar el parámetro más lejano
            diff_Tdb = abs(tdb_adj - candidate1_Tdb)
            diff_Tr = abs(tr_adj - candidate1_Tr)
            param_to_adjust = "Tdb" if diff_Tdb >= diff_Tr else "Tr"
        else:
            # Iteraciones posteriores: alternar entre parámetros
            param_to_adjust = "Tr" if last_param == "Tdb" else "Tdb"

        last_param = param_to_adjust  # Actualizar para próxima iteración

        # Seleccionar candidato1 correspondiente
        candidate1 = candidate1_Tdb if param_to_adjust == "Tdb" else candidate1_Tr
        target_avg = (candidate1 + candidate2) / 2.0

        st.write(f"   Candidate 1 (ajustando {param_to_adjust}): {candidate1:.2f} °C")
        st.write(f"   Candidate 2 (ajustando ambas): {candidate2:.2f} °C")
        st.write(f"   **Target promedio = {target_avg:.2f} °C**")

        # Aplicar ajuste al parámetro seleccionado
        if param_to_adjust == "Tdb":
            diff = target_avg - tdb_adj
            new_val = tdb_adj + 0.5 * diff
            st.write(f"   Se actualiza Tdb: diferencia = {diff:+.2f} °C, nuevo Tdb = {new_val:.2f} °C")
            tdb_adj = new_val
        else:
            diff = target_avg - tr_adj
            new_val = tr_adj + 0.5 * diff
            st.write(f"   Se actualiza Tr: diferencia = {diff:+.2f} °C, nuevo Tr = {new_val:.2f} °C")
            tr_adj = new_val

    pmv_final = pmv_ppd_iso(
        tdb=tdb_adj,
        tr=tr_adj,
        vr=v_used,
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
    st.write(f"   Vr = {v_measured:.2f} m/s")
    st.write(f"**PMV final calculado:** {pmv_final:.2f}")

    recomendaciones = generar_recomendaciones(
        pmv=pmv_actual,
        tdb=tdb_measured,
        tr=tr_measured,
        vr=v_used,
        rh=rh_measured,
        met=met,
        clo=0.5  # Valor fijo según tu código
    )

    # Mostrar recomendaciones
    st.header("Recomendaciones de Ajuste")
    cols = st.columns(2)
    priority_order = ['ventilacion', 'temperatura_radiante', 'enfriamiento', 'calefaccion', 'humedad', 'actividad']

    for category in priority_order:
        for rec in recomendaciones:
            if rec['tipo'] == category:
                with cols[0] if category in ['ventilacion', 'temperatura_radiante'] else cols[1]:
                    st.subheader(rec['mensaje'])
                    if 'acciones' in rec:
                        for accion in rec['acciones']:
                            st.write(f"- {accion}")
                    if 'accion' in rec:
                        for k, v in rec['accion'].items():
                            st.write(f"- Ajustar {k.upper()} a {v}")
