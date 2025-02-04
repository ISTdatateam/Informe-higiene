import streamlit as st
from datetime import datetime, time, date

st.set_page_config(page_title="Formulario Informe Confort Térmico", layout="wide")

st.title("Formulario para Informe de Evaluación de Confort Térmico")

# Usamos un formulario para agrupar la entrada de datos y un botón de envío
with st.form("informe_form"):
    # Utilizamos tabs para separar las secciones
    tabs = st.tabs(["Empresa y Centro de Trabajo", "Datos Generales", "Mediciones de Áreas"])

    # ================================
    # Pestaña 1: Empresa y Centro de Trabajo
    # ================================
    with tabs[0]:
        st.header("Información de la Empresa y Centro de Trabajo")
        col1, col2 = st.columns(2)
        with col1:
            razon_social = st.text_input("Razón Social")
            rut = st.text_input("RUT")
            nombre_local = st.text_input("Nombre de Local")
            direccion = st.text_input("Dirección")
        with col2:
            cuv = st.text_input("CUV")
            comuna = st.text_input("Comuna")
            region = st.text_input("Región")

    # ================================
    # Pestaña 2: Datos Generales de la Evaluación
    # ================================
    with tabs[1]:
        st.header("Datos Generales de la Evaluación")
        col1, col2 = st.columns(2)
        with col1:
            fecha_visita = st.date_input("Fecha de visita", value=date.today())
            hora_medicion = st.time_input("Hora de medición", value=time(hour=9, minute=0))
            temp_max = st.number_input("Temperatura máxima del día (°C)", min_value=-50.0, max_value=60.0, value=25.0,
                                       step=0.1)
            nombre_personal = st.text_input("Nombre del personal SMU")
            cargo = st.text_input("Cargo")
            correo_ist = st.text_input("Profesional IST (Correo)")
        with col2:
            cod_equipo_temp = st.text_input("Código equipo temperatura")
            cod_equipo_2 = st.text_input("Código equipo 2")
            verif_tbs_inicial = st.text_input("Verificación TBS inicial")
            verif_tbh_inicial = st.text_input("Verificación TBH inicial")
            verif_tg_inicial = st.text_input("Verificación TG inicial")
            verif_tbs_final = st.text_input("Verificación TBS final")
            verif_tbh_final = st.text_input("Verificación TBH final")
            verif_tg_final = st.text_input("Verificación TG final")

        st.markdown("### Calibración y otros datos")
        col3, col4 = st.columns(2)
        with col3:
            patron_calibracion = st.text_input("Patrón utilizado para calibrar")
            patron_tbs = st.text_input("Patrón TBS")
            patron_tbh = st.text_input("Patrón TBH")
        with col4:
            patron_tg = st.text_input("Patrón TG")
            vestimenta = st.text_input("Tipo de vestimenta utilizada")
            motivo_evaluacion = st.text_input("Motivo de evaluación")
            comentarios_finales = st.text_area("Comentarios finales de evaluación")

    # ================================
    # Pestaña 3: Mediciones de Áreas (Cardinalidad n)
    # ================================
    with tabs[2]:
        st.header("Mediciones de Áreas")
        st.info("Registre la información para cada área evaluada.")

        # Primero preguntamos cuántas áreas se desean registrar
        num_areas = st.number_input("Cantidad de áreas a evaluar", min_value=1, max_value=20, value=1, step=1)

        # Creamos una lista para almacenar la información de cada área.
        areas_data = []
        for i in range(1, int(num_areas) + 1):
            st.markdown(f"#### Área {i}")
            with st.expander(f"Completar datos del Área {i}", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    area_sector = st.text_input(f"Área o sector (Área {i})")
                    espec_sector = st.text_input(f"Especificación sector (Área {i})")
                    tbs = st.number_input(f"Temperatura bulbo seco (°C) - Área {i}", value=0.0, step=0.1,
                                          key=f"tbs_{i}")
                    tg = st.number_input(f"Temperatura globo (°C) - Área {i}", value=0.0, step=0.1, key=f"tg_{i}")
                    hr = st.number_input(f"Humedad relativa (%) - Área {i}", min_value=0.0, max_value=100.0, value=50.0,
                                         step=0.1, key=f"hr_{i}")
                    vel_aire = st.number_input(f"Velocidad del aire (m/s) - Área {i}", min_value=0.0, max_value=20.0,
                                               value=0.0, step=0.1, key=f"vel_{i}")
                with col2:
                    puesto_trabajo = st.text_input(f"Puesto de trabajo - Área {i}")
                    posicion_trabajador = st.selectbox(f"Trabajador (de pie o sentado) - Área {i}",
                                                       options=["", "De pie", "Sentado"], key=f"pos_{i}")
                    techumbre = st.text_input(f"Techumbre - Área {i}")
                    obs_techumbre = st.text_input(f"Observación techumbre - Área {i}")
                    paredes = st.text_input(f"Paredes - Área {i}")
                    obs_paredes = st.text_input(f"Observación paredes - Área {i}")
                    ventanales = st.text_input(f"Ventanales - Área {i}")
                    obs_ventanales = st.text_input(f"Observación ventanales - Área {i}")
                st.markdown("##### Otros aspectos")
                col3, col4 = st.columns(2)
                with col3:
                    aire_acond = st.text_input(f"Aire acondicionado - Área {i}")
                    obs_aire_acond = st.text_input(f"Observaciones aire acondicionado - Área {i}")
                    ventiladores = st.text_input(f"Ventiladores - Área {i}")
                    obs_ventiladores = st.text_input(f"Observaciones ventiladores - Área {i}")
                    inyeccion_extrac = st.text_input(f"Inyección y/o extracción de aire - Área {i}")
                    obs_inyeccion = st.text_input(f"Observaciones inyección/extracción de aire - Área {i}")
                with col4:
                    ventanas = st.text_input(f"Ventanas (ventilación natural) - Área {i}")
                    obs_ventanas = st.text_input(f"Observaciones ventanas - Área {i}")
                    puertas = st.text_input(f"Puertas (ventilación natural) - Área {i}")
                    obs_puertas = st.text_input(f"Observaciones puertas - Área {i}")
                    condiciones_disconfort = st.text_input(f"Otras condiciones de disconfort térmico - Área {i}")
                    obs_condiciones = st.text_input(f"Observaciones sobre disconfort térmico - Área {i}")
                evidencia = st.text_input(f"Evidencia fotográfica (URL(s) separadas por coma) - Área {i}")

                # Se almacena la información de este área en un diccionario
                areas_data.append({
                    "Area o sector": area_sector,
                    "Especificación sector": espec_sector,
                    "Temperatura bulbo seco": tbs,
                    "Temperatura globo": tg,
                    "Humedad relativa": hr,
                    "Velocidad del aire": vel_aire,
                    "Puesto de trabajo": puesto_trabajo,
                    "Trabajador de pie o sentado": posicion_trabajador,
                    "Techumbre": techumbre,
                    "Observación techumbre": obs_techumbre,
                    "Paredes": paredes,
                    "Observación paredes": obs_paredes,
                    "Ventanales": ventanales,
                    "Observación ventanales": obs_ventanales,
                    "Aire acondicionado": aire_acond,
                    "Observaciones aire acondicionado": obs_aire_acond,
                    "Ventiladores": ventiladores,
                    "Observaciones ventiladores": obs_ventiladores,
                    "Inyección y/o extracción de aire": inyeccion_extrac,
                    "Observaciones inyección/extracción de aire": obs_inyeccion,
                    "Ventanas (ventilación natural)": ventanas,
                    "Observaciones ventanas": obs_ventanas,
                    "Puertas (ventilación natural)": puertas,
                    "Observaciones puertas": obs_puertas,
                    "Otras condiciones de disconfort térmico": condiciones_disconfort,
                    "Observaciones sobre disconfort térmico": obs_condiciones,
                    "Evidencia fotográfica": evidencia,
                })

    # Botón de envío del formulario
    submitted = st.form_submit_button("Generar Informe")

    if submitted:
        # Aquí se pueden consolidar todos los datos en un diccionario o convertirlos a DataFrame
        datos_informe = {
            "empresa": {
                "Razón Social": razon_social,
                "RUT": rut,
                "CUV": cuv,
                "Nombre de Local": nombre_local,
                "Dirección": direccion,
                "Comuna": comuna,
                "Región": region,
            },
            "datos_generales": {
                "Fecha visita": fecha_visita.strftime("%d/%m/%Y"),
                "Hora medición": hora_medicion.strftime("%H:%M"),
                "Temperatura máxima del día": temp_max,
                "Nombre del personal SMU": nombre_personal,
                "Cargo": cargo,
                "Profesional IST (Correo)": correo_ist,
                "Código equipo temperatura": cod_equipo_temp,
                "Código equipo 2": cod_equipo_2,
                "Verificación TBS inicial": verif_tbs_inicial,
                "Verificación TBH inicial": verif_tbh_inicial,
                "Verificación TG inicial": verif_tg_inicial,
                "Verificación TBS final": verif_tbs_final,
                "Verificación TBH final": verif_tbh_final,
                "Verificación TG final": verif_tg_final,
                "Patrón utilizado para calibrar": patron_calibracion,
                "Patrón TBS": patron_tbs,
                "Patrón TBH": patron_tbh,
                "Patrón TG": patron_tg,
                "Tipo de vestimenta utilizada": vestimenta,
                "Motivo de evaluación": motivo_evaluacion,
                "Comentarios finales de evaluación": comentarios_finales,
            },
            "mediciones_areas": areas_data
        }

        st.success("¡Formulario enviado correctamente!")
        st.json(datos_informe)
        # Aquí se podría llamar a la función que genera el documento Word
        # Por ejemplo: buffer_doc = generar_informe_en_word(df_filtrado, df_info_cuv)
        # Y luego ofrecer la descarga del documento.
