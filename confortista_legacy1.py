import streamlit as st
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq
import numpy as np
import pandas as pd

st.header("CONFORTISTA 1.0")
st.subheader("Asistente de Control de Confort Térmico")

def color_pmv(val):
    try:
        pmv = float(val)
        if -1 <= pmv <= 1:
            return 'background-color: #90EE90'  # Verde claro
        else:
            return 'background-color: #FFCCCB'  # Rojo claro
    except:
        return ''

def generar_recomendaciones(pmv, tdb_initial, tr_initial, vr, rh, met, clo, tdb_final, tr_final):
    recomendaciones = []

    # 1. Análisis de Velocidad del Aire (VR) - Se mantiene igual
    if pmv > 1.0 and vr < 0.2:
        recomendaciones.append({
            'tipo': 'ventilacion',
            'mensaje': f"Aumentar velocidad del aire al menos a 0.2 m/s usando ventiladores (VR actual: {vr} m/s)",
            'accion': {'velocidad del aire': 0.2}
        })
    elif pmv < -1.0 and vr > 1:
        recomendaciones.append({
            'tipo': 'ventilacion',
            'mensaje': "Reducir corrientes de aire frío bajo 1.0 m/s (aumentan sensación de frío)",
            'accion': {'velocidad del aire': 1.0}
        })

        # Modificar lógica de estrategias
        if pmv > 1.0:
            accion_tdb = f"Reducir Tdb desde {tdb_initial}°C a {tdb_final}°C" if tdb_final < tdb_initial else f"Mantener Tdb en {tdb_final}°C"
            accion_tr = f"Reducir Tr desde {tr_initial}°C a {tr_final}°C" if tr_final < tr_initial else f"Mantener Tr en {tr_final}°C"
        else:
            accion_tdb = f"Aumentar Tdb desde {tdb_initial}°C a {tdb_final}°C" if tdb_final < tdb_initial else f"Mantener Tdb en {tdb_final}°C"
            accion_tr = f"Aumentar Tr desde {tr_initial}°C a {tr_final}°C" if tr_final < tr_initial else f"Mantener Tr en {tr_final}°C"



    # 2. Relación entre Temperatura Radiante (Tr) y del Aire (Tdb) - Modificado
    dif_temp = tr - tdb
    if dif_temp > 3.0:
        recomendaciones.append({
            'tipo': 'temperatura_radiante',
            'mensaje': f"Temperatura radiante elevada (Tr-Tdb = {dif_temp:.1f}°C). Medidas sugeridas:",
            'acciones': [
                f"Reducir Tr desde {tr}°C hasta aproximadamente {tr_final}°C con aislamiento",
                "Instalar superficies reflectantes",
                "Controlar fuentes de radiación (ventanas, equipos)"
            ]
        })
    elif dif_temp < -2.0:
        recomendaciones.append({
            'tipo': 'temperatura_radiante',
            'mensaje': f"Temperatura radiante baja (Tr-Tdb = {dif_temp:.1f}°C). Medidas sugeridas:",
            'acciones': [
                f"Aumentar Tr desde {tr}°C hasta aproximadamente {tr_final}°C",
                "Considerar calefacción radiante"
            ]
        })

    # 3. Estrategias según dirección del ajuste - Modificado
    if pmv > 1.0:  # Ambiente caluroso
        recomendaciones.append({
            'tipo': 'enfriamiento',
            'mensaje': "Estrategias de enfriamiento:",
            'acciones': [
                f"Reducir Tdb desde {tdb}°C hasta aproximadamente {tdb_final}°C mediante HVAC" if tdb_final < tdb_initial else f"Mantener Tdb en {tdb_final}°C",
                f"Reducir Tr desde {tr}°C hasta aproximadamente {tr_final}°C con sombreado" if tr_final < tr_initial else f"Mantener Tr en {tr_final}°C",
                "Ventilación cruzada para aumentar disipación térmica"
            ]
        })
    elif pmv < -1.0:  # Ambiente frío
        recomendaciones.append({
            'tipo': 'calefaccion',
            'mensaje': "Estrategias de calentamiento:",
            'acciones': [
                f"Aumentar Tdb desde {tdb}°C hasta aproximadamente {tdb_final}°C mediante calefacción" if tdb_final < tdb_initial else f"Mantener Tdb en {tdb_final}°C",
                f"Aumentar Tr desde {tr}°C hasta aproximadamente {tr_final}°C con superficies radiantes" if tr_final < tr_initial else f"Mantener Tr en {tr_final}°C",
                "Reducir infiltraciones de aire frío"
            ]
        })

    # 4. Factores adicionales - Se mantiene igual
    if met < 1.2:
        recomendaciones.append({
            'tipo': 'actividad',
            'mensaje': f"Actividad metabólica baja ({met} met). Considerar:",
            'acciones': [
                "Pausas activas para aumentar movimiento",
                f"Adecuar vestimenta (actual CLO = {clo})"
            ]
        })

    if rh > 70:
        recomendaciones.append({
            'tipo': 'humedad',
            'mensaje': f"Humedad relativa elevada ({rh}%. Acciones:",
            'acciones': [
                "Usar deshumidificadores",
                "Mejorar ventilación natural/mecánica"
            ]
        })

    return recomendaciones

def calcular_ajuste_optimo(pmv_initial,tdb_initial, tr_initial, vr_initial, rh, met, clo, target_pmv, max_iter=20):

    # Parámetros ajuste
    tol_temp = 0.01
    tol_pmv = 0.001
    min_temp = 18.0
    max_temp = 28.0


    ajuste_ventilacion = []  # Variable para almacenar el mensaje de ajuste
    tdb_adj = tdb_initial
    tr_adj = tr_initial
    vr_adj = vr_initial  # Nueva variable para VR ajustado
    historial = []


    for i in range(max_iter):
        # --- Paso 1: Ajuste dinámico de VR dentro de la iteración ---
        current_pmv = pmv_ppd_iso(tdb_adj, tr_adj, vr_adj, rh, met, clo, limit_inputs=False).pmv

        # Regla de ajuste de VR solo si estamos fuera de confort
        if not (-1 < current_pmv < 1):
            if current_pmv > 1.0 and vr < 0.2:
                vr_adj = 0.2

            elif current_pmv < -1.0 and vr > 1:
                vr_adj = 1.0

        # --- Paso 2: Optimización de temperaturas considerando VR actual ---
        if target_pmv < current_pmv:  # Enfriar
            lower_bound = min_temp
            upper_bound = min(tdb_adj, tr_adj, max_temp)  # No superar max_temp
            upper_bound = max(lower_bound, upper_bound)  # Asegurar upper >= lower
        else:  # Calentar
            lower_bound = max(min(tdb_adj, tr_adj), min_temp)  # No bajar de min_temp
            upper_bound = max_temp

        # Definimos x = diff_pmv
        diff_pmv = abs(current_pmv - target_pmv)

        # Aplicamos la función cuadrática para diff_pmv ≤ 2
        if diff_pmv <= 2:
            factor = 0.09 * (diff_pmv ** 2) - 0.43 * diff_pmv + 1
        else:
            # Para diff_pmv > 2, mantenemos el valor mínimo alcanzado (0.5)
            factor = 0.5

        # Verificar que el intervalo es válido
        if lower_bound >= upper_bound:
            st.error(f"Intervalo inválido en la iteración {i+1}: [{lower_bound}, {upper_bound}]")
            break

        try:
            candidate2 = brentq(
                lambda x: pmv_ppd_iso(
                    tdb=x, tr=x, vr=vr, rh=rh, met=met, clo=clo, limit_inputs=False
                ).pmv - target_pmv,
                lower_bound, upper_bound, xtol=0.1
            )
        except ValueError:
            # Fallback a método de bisección si hay error de signos
            candidate2 = (lower_bound + upper_bound) / 2
        except Exception as e:
            st.error(f"Error Candidate2 en la iteración {i+1}: {str(e)}")
            candidate2 = (tdb_adj + tr_adj) / 2

        # Actualizar Tdb y Tr utilizando el factor de ajuste
        prev_tdb = tdb_adj  # <--- AÑADIR ESTA LÍNEA
        prev_tr = tr_adj  # Opcional: Para consistencia


        # Después - Limitar cambio máximo por iteración
        max_change = 2.0  # Máximo 2°C por iteración
        change_tdb = factor * (candidate2 - tdb_adj)
        change_tr = factor * (candidate2 - tr_adj)

        # Actualizar Tdb/Tr con límites físicos
        tdb_adj = np.clip(tdb_adj + change_tdb, min_temp, max_temp)
        tr_adj = np.clip(tr_adj + change_tr, min_temp, max_temp)
        tdb_adj = round(tdb_adj, 2)
        tr_adj = round(tr_adj, 2)

        #tdb_adj += max(min(change_tdb, max_change), -max_change)
        #tr_adj += max(min(change_tr, max_change), -max_change)

        # Recalcular PMV con los nuevos valores
        try:
            new_pmv = pmv_ppd_iso(tdb_adj, tr_adj, vr_adj, rh, met, clo, limit_inputs=False).pmv
        except Exception as e:
            st.error(f"Error al recalcular PMV en la iteración {i+1}: {str(e)}")
            new_pmv = np.nan

        # Registrar el historial de ajustes
        historial.append({
            'Iteracion': i + 1,
            'Tdb': round(tdb_adj, 2),
            'Tr': round(tr_adj, 2),
            'VR': vr_adj,
            'PMV': round(new_pmv, 3) if not np.isnan(new_pmv) else "Error"
        })

        # Condición de corte (usar new_pmv)
        if -1 < new_pmv < 1:
            break
        if abs(tdb_adj - prev_tdb) < tol_temp and abs(tr_adj - prev_tr) < tol_temp:
            break

        # Condición de corte: si el cambio en temperatura es mínimo, detener la iteración
        if (
                (abs(tdb_adj - prev_tdb) < tol_temp
                 and abs(tr_adj - prev_tr) < tol_temp)
                or abs(new_pmv - current_pmv) < tol_pmv
        ):
            break

    return tdb_adj, tr_adj, vr_adj, historial


def df_style(df):
    def format_value(val, fmt):
        if isinstance(val, (int, float)):
            return fmt.format(val)
        return val

    # Crear copia del DataFrame y asegurar columnas requeridas
    df = df.copy()

    # Verificar y agregar columnas faltantes con np.nan
    for col in ['Tdb', 'Tr', 'PMV']:
        if col not in df.columns:
            df[col] = np.nan

    # Aplicar formato a las columnas
    df['Tdb'] = df['Tdb'].apply(lambda x: format_value(x, '{:.1f}°C'))
    df['Tr'] = df['Tr'].apply(lambda x: format_value(x, '{:.1f}°C'))
    df['PMV'] = df['PMV'].apply(lambda x: format_value(x, '{:.2f}'))

    return df.style.applymap(color_pmv, subset=['PMV'])


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
    try:
        pmv_initial = pmv_ppd_iso(
            tdb=tdb, tr=tr, vr=vr,
            rh=rh, met=met, clo=clo,
            limit_inputs=False
        ).pmv
    except:
        st.error("Error en cálculo inicial. Verifique los valores ingresados.")
        st.stop()


    # Determinar target PMV
    if pmv_initial < -1:
        target_pmv = -0.99
    elif pmv_initial > 1:
        target_pmv = 0.99
    else:
        target_pmv = pmv_initial

    # Proceso de optimización
    tdb_final, tr_final, vr_final, historial = calcular_ajuste_optimo(pmv_initial,
        tdb, tr, vr, rh, met, clo, target_pmv
    )

    # Cálculo final
    try:
        pmv_final = pmv_ppd_iso(
            tdb=tdb_final, tr=tr_final, vr=vr_final,
            rh=rh, met=met, clo=clo
        ).pmv
    except:
        pmv_final = np.nan

    # Mostrar resultados
    st.header("📊 Resultados de la Optimización")

    cols = st.columns(3)
    cols[0].metric("PMV Inicial", f"{pmv_initial:.2f}")
    cols[1].metric("PMV Final",
                   f"{pmv_final:.2f}" if not np.isnan(pmv_final) else "Error",
                   "✅ En confort" if -1 < pmv_final < 1 else "⚠️ Fuera de rango")  # Exclusivo para (-1, 1)
    cols[2].metric("Iteraciones", len(historial))

    # Nuevo bloque añadido: Detalles de parámetros finales
    with st.expander("🔍 Parámetros utilizados para el cálculo final", expanded=True):
        param_cols = st.columns(3)
        param_cols[0].markdown(f"""
        **Variables térmicas:**  
        - T. Aire: `{tdb_final:.1f}°C`  
        - T. Radiante: `{tr_final:.1f}°C`  
        - Vel. Aire: `{vr_final:.2f} m/s`
        """)

        param_cols[1].markdown(f"""
        **Factores humanos:**  
        - Metabolismo: `{met:.1f} met`  
        - Vestimenta: `{clo:.1f} CLO`  
        - Humedad: `{rh:.0f}%`
        """)

        param_cols[2].markdown(f"""
        **Objetivos:**  
        - PMV Target: `{target_pmv:.2f}`  
        - Tolerancia: `±0.05°C`  
        - Máx iteraciones: `20`
        """)

    # Recomendaciones
    st.header("🔧 Recomendaciones de Ajuste")

    # Añadir recomendación de ajuste de ventilación si existe
    recomendaciones = []

    # Agregar demás recomendaciones
    recomendaciones += generar_recomendaciones(pmv_initial, tdb, tr, vr, rh, met, clo, tdb_final, tr_final)
    if recomendaciones:
        cols_rec = st.columns(2)
        for i, rec in enumerate(recomendaciones):
            with cols_rec[i % 2]:
                with st.expander(f"{rec['tipo'].capitalize()}: {rec['mensaje']}", expanded=True):
                    # Manejar ambos casos: 'acciones' (lista) y 'accion' (diccionario)
                    if 'acciones' in rec:
                        st.write("**Acciones recomendadas:**")
                        for accion in rec['acciones']:
                            st.write(f"- {accion}")
                    if 'accion' in rec:
                        st.write("**Ajuste específico:**")
                        for k, v in rec['accion'].items():  # Ahora funciona porque es diccionario
                            st.write(f"- Ajustar {k} a {v}")
    else:
        if not (-1 <= pmv_final <= 1):
            st.warning("No se encontraron ajustes adecuados. Considere modificar otros parámetros.")
        else:
            st.success(
                "¡No se requieren ajustes adicionales! Las condiciones actuales están dentro del rango de confort.")

    # Mostrar progreso detallado
    with st.expander("Ver detalles del proceso de ajuste"):
        st.write("**Historial de ajustes:**")
        df_historial = pd.DataFrame(historial)
        st.dataframe(df_style(df_historial), height=300)