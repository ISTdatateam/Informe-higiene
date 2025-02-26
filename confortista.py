import streamlit as st
from pythermalcomfort.models import pmv_ppd_iso
from scipy.optimize import brentq
import numpy as np
import pandas as pd

st.header("CONFORTISTA 1.1.20250226")
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


    # Capa 1: Ajustes específicos según parámetros
    ajustes = []

    # 1. Velocidad del Aire
    if pmv > 1.0 and vr < 0.2:
        ajustes.append({
            'tipo': 'ventilacion',
            'parametro': 'Velocidad del aire',
            'actual': vr,
            'objetivo': 0.2,
        })

    # 2. Temperatura Radiante
    dif_temp = tr_initial - tdb_initial
    if dif_temp > 2.0:
        ajustes.append({
            'tipo': 'temperatura',
            'parametro': 'Exceso de radiación (Tr-Tdb)',
            'actual': dif_temp,
            'objetivo': "≤2.0°C",
        })

    # 3. Temperatura del Aire
    dif_tdb = tdb_initial - tdb_final
    if dif_tdb > 1:
        ajustes.append({
            'tipo': 'temperatura',
            'parametro': 'Exceso de temperatura del aire',
            'actual': tdb_initial,
            'objetivo': tdb_final,
        })

    # 3. Temperatura del Aire
    dif_tr = tr_initial - tr_final
    if dif_tdb > 1:
        ajustes.append({
            'tipo': 'temperatura',
            'parametro': 'Exceso de calor radiante',
            'actual': tr_initial,
            'objetivo': tr_final,
        })

    if ajustes:
        recomendaciones.append({
            'tipo': 'ajustes_tecnicos',
            'categoria': 'Optimización',
            'mensaje': "Ajustes específicos por parámetros evaluados",
            'detalles': ajustes
        })

    # Capa 2: Estrategias según PMV
    if pmv > 1.0:  # Ambiente caluroso
        estrategias = {
            'enfriamiento': {
                'mensaje': "Refrigeración activa requerida",
                'acciones': [
                    "Adquirir enfriadores portátiles (1-2 por área)",
                    "Coordinar protocolos de recarga con proveedores",
                    "Instalar sistemas HVAC en áreas críticas"
                ],
                'plazo': '3-6 meses'
            }}
        if dif_temp > 2.0:
            estrategias = estrategias | {
            'aislamiento': {
                'mensaje': "Reducción de carga térmica",
                'acciones': [
                    "Instalar materiales aislantes en techos/paredes",
                    "Implementar protecciones solares reflectivas",
                    "Aislar fuentes de calor radiante"
                ],
                'plazo': '3 meses'
            }}
        if vr < 0.2:
            estrategias = estrategias | {
            'ventilacion': {
                'mensaje': f"VR actual: {vr} m/s - Aumentar ventilación",
                'acciones': [
                    "Instalar ventiladores (mínimo 2 por área)",
                    "Implementar sistemas de extracción forzada",
                    "Optimizar ventilación cruzada"
                ],
                'plazo': '3 meses'
                }}
    elif pmv < -1.0:  # Ambiente frío
        estrategias = {
            'calefaccion': {
                'mensaje': "Protección contra el frío",
                'acciones': [
                    "Implementar sistemas de calefacción radiante",
                    "Mejorar aislamiento térmico en envolvente",
                    "Optimizar sellado de infiltraciones"
                ],
                'plazo': '3 meses'
            }
        }
    else:
        estrategias = {}

    # Agregar estrategias principales
    for key, value in estrategias.items():
        recomendaciones.append({
            'tipo': key,
            'categoria': 'Estrategia Principal',
            'nivel': 'Prioridad 2',
            'mensaje': value['mensaje'],
            'acciones': value['acciones'],
            'plazo': value['plazo']
        })

    # Capa 3: Mantenimiento y Control
    mantenimiento = {
        'preventivo': [
            "Programar mantención de equipos con proveedores",
            "Calendario de limpieza de filtros/paneles",
            "Verificación mensual de sistemas"
        ],
        'correctivo': [
            "Reparación de sistemas de ventilación",
            "Ajuste de equipos de climatización",
            "Registro de intervenciones técnicas"
        ],
        'control': [
            f"Regulación térmica (23-26°C) con registro",
            "Monitoreo continuo de parámetros ambientales"
        ]
    }

    recomendaciones.append({
        'tipo': 'mantenimiento',
        'categoria': 'Gestión Técnica',
        'nivel': 'Prioridad 3',
        'mensaje': "Programa de mantenimiento integral",
        'acciones': {
            'preventivo': mantenimiento['preventivo'],
            'correctivo': mantenimiento['correctivo'],
            'control': mantenimiento['control']
        },
        'plazo': 'Continuo'
    })

    # Capa 4: Medidas Administrativas (siempre aplican)
    recomendaciones.append({
        'tipo': 'administrativa',
        'categoria': 'Comunicación',
        'nivel': 'Prioridad 3',
        'acciones': [
            "Informar formalmente a todo el personal sobre riesgos térmicos (DS N°44)",
            "Establecer registros firmados de capacitación"
        ],
        'plazo': '30 días'
    })

    return recomendaciones

def calcular_ajuste_optimo(pmv_initial,tdb_initial, tr_initial, vr_initial, rh, met, clo, target_pmv, max_iter=30):

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

        current_pmv = pmv_ppd_iso(tdb_adj, tr_adj, vr_adj, rh, met, clo, limit_inputs=False).pmv

        # --- Ajuste dinámico de VR dentro de la iteración ---
        # Regla de ajuste de VR solo si estamos fuera de confort
        if not (-1 < current_pmv < 1):
            if current_pmv > 1.0 and vr < 0.2:
                vr_adj = 0.2

            elif current_pmv < -1.0 and vr > 1:
                vr_adj = 1.0

        current_pmv = pmv_ppd_iso(tdb_adj, tr_adj, vr_adj, rh, met, clo, limit_inputs=False).pmv

        # --- Optimización de temperaturas considerando VR actual ---
        if target_pmv < current_pmv:  # Enfriar
            lower_bound = min_temp
            upper_bound = max(tdb_adj, tr_adj)
        else:  # Calentar
            lower_bound = min(tdb_adj, tr_adj)
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
                lower_bound, upper_bound, xtol=0.01
            )
        except Exception as e:
            st.error(f"Error Candidate2 en la iteración {i+1}: {str(e)}")

        # Actualizar Tdb y Tr utilizando el factor de ajuste
        prev_tdb = tdb_adj  # <--- AÑADIR ESTA LÍNEA
        prev_tr = tr_adj  # Opcional: Para consistencia


        # Después - Limitar cambio máximo por iteración
        max_change = 0.25 # Máximo 2°C por iteración
        change_tdb = factor * (candidate2 - tdb_adj)
        change_tr = factor * (candidate2 - tr_adj)

        # Actualizar Tdb/Tr con límites físicos
        tdb_adj = np.clip(tdb_adj + change_tdb, (tdb_adj-max_change), (tdb_adj+max_change))
        tdb_adj = round(tdb_adj, 2)
        tr_adj = np.clip(tr_adj + change_tr, (tr_adj-max_change), (tr_adj+max_change))
        tr_adj = round(tr_adj, 2)

        # Recalcular PMV con los nuevos valores
        try:
            new_pmv = pmv_ppd_iso(tdb_adj, tr_adj, vr_adj, rh, met, clo, limit_inputs=False).pmv
            new_ppd = pmv_ppd_iso(tdb_adj, tr_adj, vr_adj, rh, met, clo, limit_inputs=False).ppd
        except Exception as e:
            st.error(f"Error al recalcular PMV en la iteración {i+1}: {str(e)}")
            new_pmv = np.nan

        # Registrar el historial de ajustes
        historial.append({
            'Iteracion': i + 1,
            'Tdb': round(tdb_adj, 2),
            'Tr': round(tr_adj, 2),
            'VR': vr_adj,
            'PPD': round(new_ppd, 3),
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
        ppd_initial = pmv_ppd_iso(
            tdb=tdb, tr=tr, vr=vr,
            rh=rh, met=met, clo=clo,
            limit_inputs=False
        ).ppd
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
        ppd_final = pmv_ppd_iso(
            tdb=tdb_final, tr=tr_final, vr=vr_final,
            rh=rh, met=met, clo=clo
        ).ppd
    except:
        pmv_final = np.nan

    # Mostrar resultados

    st.header("📋 Condiciones iniciales evaluadas")

    cols = st.columns(2)
    cols[0].metric("PMV Inicial",
                   f"{pmv_initial:.2f}" if not np.isnan(pmv_initial) else "Error",
                   "✅ En confort" if -1 < pmv_initial < 1 else "⚠️ Fuera de rango",
                   delta_color="off")  # Exclusivo para (-1, 1)
    cols[1].metric("PPD Inicial", f"{ppd_initial:.2f}")

    # Nuevo bloque añadido: Detalles de parámetros finales
    with st.expander("🔍 Parámetros utilizados para el cálculo inicial", expanded=True):
        param_cols = st.columns(3)
        param_cols[0].markdown(f"""
            **Variables térmicas:**  
            - T. Aire: `{tdb:.1f}°C`  
            - T. Radiante: `{tr:.1f}°C`  
            - Vel. Aire: `{vr:.2f} m/s`
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


    st.header("📊 Resultados de la Optimización")

    cols = st.columns(2)
    cols[0].metric("PMV Final",
                   f"{pmv_final:.2f}" if not np.isnan(pmv_final) else "Error",
                   "✅ En confort" if -1 < pmv_final < 1 else "⚠️ Fuera de rango",
                   delta_color="off")  # Exclusivo para (-1, 1)
    cols[1].metric("PPD Final", f"{ppd_final:.2f}")

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

    # Agregar demás recomendaciones
    recomendaciones = generar_recomendaciones(pmv_initial, tdb, tr, vr, rh, met, clo, tdb_final, tr_final)
    if recomendaciones:
        cols_rec = st.columns(1)
        for i, rec in enumerate(recomendaciones):
            with cols_rec[i % 1]:
                # Construir título con emojis según tipo
                iconos = {
                    'administrativa': '📋',
                    'ventilacion': '🌀',
                    'enfriamiento': '❄️',
                    'aislamiento': '🛡️',
                    'mantenimiento': '🔧',
                    'ajustes_tecnicos': '⚙️'
                }
                icono = iconos.get(rec['tipo'], '📌')

                # Construir título de manera segura
                titulo = f"{iconos.get(rec['tipo'], '📌')} {rec['categoria']}"
                if 'mensaje' in rec:
                    titulo += f": {rec['mensaje']}"

                with st.expander(titulo, expanded=True):
                    # Sección de información básica
                    labels = []
                    if 'nivel' in rec:
                        labels.append(f"**Prioridad:** `{rec['nivel']}`")
                    if 'plazo' in rec:
                        labels.append(f"**Plazo:** `{rec['plazo']}`")
                    if labels:
                        st.markdown(" | ".join(labels))
                    # Manejo de diferentes estructuras de acciones

                    # Detalles técnicos de ajustes
                    if 'detalles' in rec:
                        st.write("**Parametros a modificar con las medidas prescritas:**")
                        for ajuste in rec['detalles']:
                            st.markdown(f"""
                            - **Parámetro:** {ajuste['parametro']}  
                              **Actual:** {ajuste['actual']}  
                              **Objetivo:** {ajuste['objetivo']}  
                            """)

                    if 'acciones' in rec:
                        if isinstance(rec['acciones'], dict):
                            st.write("**Plan de Acción:**")
                            for subtipo, acciones in rec['acciones'].items():
                                st.markdown(f"**{subtipo.capitalize()}:**")
                                for accion in acciones:
                                    st.write(f"- {accion}")
                        else:
                            st.write("**Acciones recomendadas:**")
                            for accion in rec['acciones']:
                                st.write(f"- {accion}")



                    # Acción inmediata si existe
                    if 'accion' in rec:
                        st.write("---")
                        st.write("**Acción prioritaria:**")
                        st.write(f"- {rec['accion']}")

    else:
        if not (-1 <= pmv_final <= 1):
            st.warning(
                "⚠️ No se encontraron ajustes adecuados. Considere modificar otros parámetros o implementar soluciones estructurales.")
        else:
            st.success("✅ ¡Condiciones óptimas! No se requieren ajustes adicionales.")

    # Mostrar progreso detallado
    with st.expander("Ver detalles del proceso de ajuste: " + str(len(historial)) + " iteraciones"):
        st.write("**Historial de ajustes:**")
        df_historial = pd.DataFrame(historial)
        st.dataframe(df_style(df_historial), height=300)