import streamlit as st
import pandas as pd
from datetime import datetime, date, time

from data_access import get_data  # Función que obtiene el CSV principal
from doc_utils import generar_informe_en_word  # Función para generar el Word


st.set_page_config(page_title="Informes Confort Térmico", layout="wide")

def main():
    st.title("Informes Confort Térmico")
    st.write("")
    st.write("Versión 0.4.20250203")
    st.write("")

    # --- Carga de CSVs ---
    # CSV 1: Datos generales / mediciones (CSV principal)
    csv_url_main = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vTPdZTyxM6BDLmnlqe246tfBm7H06vXBdQKruh2mPg-rhQSD8olCS30ej4BdtJ1R__3W6K-Va3hm5Ax/"
        "pub?output=csv"
    )
    df_main = get_data(csv_url_main)  # DataFrame principal

    # CSV 2: Información de CUV (RUT, Razón Social, etc.)
    csv_url_cuv_info = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vSn2sEH86jBQNbjQEhtehoFIL54cFtdH3HST5zM257XbzkFx5V3VNDCO_CyIYiIWECrl1xoohSnC-lC/"
        "pub?output=csv"
    )
    df_cuv_info = pd.read_csv(csv_url_cuv_info)

    # --- Inicialización en session_state ---
    if "df_filtrado" not in st.session_state:
        st.session_state["df_filtrado"] = pd.DataFrame()
    if "df_info_cuv" not in st.session_state:
        st.session_state["df_info_cuv"] = pd.DataFrame()
    if "input_cuv_str" not in st.session_state:
        st.session_state["input_cuv_str"] = ""

    # --- Búsqueda por CUV ---
    input_cuv = st.text_input("Ingresa el CUV:")

    if st.button("Buscar"):
        st.session_state["input_cuv_str"] = input_cuv.strip()
        # Asegurarse de trabajar con strings sin espacios en ambos DataFrames
        df_main["CUV"] = df_main["CUV"].astype(str).str.strip()
        df_cuv_info["CUV"] = df_cuv_info["CUV"].astype(str).str.strip()

        st.session_state["df_filtrado"] = df_main[df_main["CUV"] == st.session_state["input_cuv_str"]]
        st.session_state["df_info_cuv"] = df_cuv_info[df_cuv_info["CUV"] == st.session_state["input_cuv_str"]]

    df_filtrado = st.session_state["df_filtrado"]
    df_info_cuv = st.session_state["df_info_cuv"]

    if not df_filtrado.empty:
        st.markdown("---")
        st.info("Completa o actualiza la información en el siguiente formulario para generar el informe.")

        # --- Formulario (st.form) ---
        with st.form("informe_form"):
            tabs = st.tabs([
                "Centro de Trabajo",
                "Inicio visita",
                "Calibración",
                "Mediciones de Áreas",
                "Fin de visita"
            ])

            # ==============================================
            # Pestaña 1: Datos generales
            # ==============================================
            with tabs[0]:
                #st.header("Datos generales")
                col1, col2 = st.columns(2)
                # Precarga de datos de la información CUV (si existe)
                if not df_info_cuv.empty:
                    cuv_info_row = df_info_cuv.iloc[0]
                    razon_social = col1.text_input("Razón Social", value=str(cuv_info_row.get("RAZÓN SOCIAL", "")))
                    rut = col1.text_input("RUT", value=str(cuv_info_row.get("RUT", "")))
                    nombre_local = col1.text_input("Nombre de Local", value=str(cuv_info_row.get("Nombre de Local", "")))
                    direccion = col1.text_input("Dirección", value=str(cuv_info_row.get("Dirección", "")))
                    cuv_val = col2.text_input("CUV", value=str(cuv_info_row.get("CUV", "")))
                    comuna = col2.text_input("Comuna", value=str(cuv_info_row.get("Comuna", "")))
                    region = col2.text_input("Región", value=str(cuv_info_row.get("Región", "")))
                else:
                    razon_social = col1.text_input("Razón Social")
                    rut = col1.text_input("RUT")
                    nombre_local = col1.text_input("Nombre de Local")
                    direccion = col1.text_input("Dirección")
                    cuv_val = col2.text_input("CUV")
                    comuna = col2.text_input("Comuna")
                    region = col2.text_input("Región")

            # ==============================================
            # Pestaña 2: Datos iniciales
            # ==============================================
            with tabs[1]:
                #st.header("Datos iniciales")
                col1 = st.columns(1)
                with col1[0]:
                    fecha_visita = st.date_input("Fecha de visita", value=date.today())
                    hora_medicion = st.time_input("Hora de medición", value=time(hour=9, minute=0))
                    temp_max = st.number_input("Temperatura máxima del día (°C)", min_value=-50.0, max_value=60.0,
                                               value=25.0, step=0.1)

                    motivo_evaluacion = st.selectbox("Motivo de evaluación",
                                                     options=["Programa anual", "Solicitud empresa", "Fiscalización"])
                    nombre_personal = st.text_input("Nombre del personal SMU")
                    cargo = st.text_input("Cargo")
                    correo_ist = st.text_input("Profesional IST (Correo)")
                    vestimenta = st.text_input("Tipo de vestimenta utilizada")
                    motivo_evaluacion = st.radio("Motivo de evaluación",
                                                 options=["Programa anual", "Solicitud empresa", "Fiscalización"])

            # ==============================================
            # Pestaña 3: Calibración y otros datos
            # ==============================================
            with tabs[2]:
                #st.header("Calibración y otros datos")
                cod_equipo_temp = st.text_input("Código equipo temperatura")
                cod_equipo_2 = st.text_input("Código equipo 2")
                patron_calibracion = st.text_input("Patrón utilizado para calibrar")

                col1, col2 = st.columns(2)
                with col1:
                    patron_tbs = st.number_input("Patrón TBS", value=46.4, step=0.1)
                    patron_tbh = st.number_input("Patrón TBH", value=12.7, step=0.1)
                    patron_tg = st.number_input("Patrón TG", value=69.8, step=0.1)
                with col2:
                    verif_tbs_inicial = st.number_input("Verificación TBS inicial", step=0.1)
                    verif_tbh_inicial = st.number_input("Verificación TBH inicial", step=0.1)
                    verif_tg_inicial = st.number_input("Verificación TG inicial", step=0.1)

            # ==============================================
            # Pestaña 4: Mediciones de Áreas (10 áreas fijas)
            # ==============================================
            with tabs[3]:
                #st.header("Mediciones de Áreas")
                st.info("Registre la información para cada una de las 10 áreas evaluadas.")
                # Se crean 10 sub-tabs, uno por cada área
                area_tabs = st.tabs([f"Área {i}" for i in range(1, 11)])
                areas_data = []
                for i, area_tab in enumerate(area_tabs, start=1):
                    with area_tab:
                        st.markdown(f"### Área {i}")
                        col1, col2 = st.columns(2)
                        with col1:
                            area_sector = st.text_input(f"Área o sector (Área {i})", key=f"area_{i}")
                            espec_sector = st.text_input(f"Especificación sector (Área {i})", key=f"espec_{i}")
                            tbs = st.number_input(f"Temperatura bulbo seco (°C) - Área {i}", value=0.0, step=0.1, key=f"tbs_{i}")
                            tg = st.number_input(f"Temperatura globo (°C) - Área {i}", value=0.0, step=0.1, key=f"tg_{i}")
                            hr = st.number_input(f"Humedad relativa (%) - Área {i}", min_value=0.0, max_value=100.0,
                                                 value=50.0, step=0.1, key=f"hr_{i}")
                            vel_aire = st.number_input(f"Velocidad del aire (m/s) - Área {i}", min_value=0.0,
                                                       max_value=20.0, value=0.0, step=0.1, key=f"vel_{i}")
                        with col2:
                            puesto_trabajo = st.text_input(f"Puesto de trabajo - Área {i}", key=f"puesto_{i}")
                            posicion_trabajador = st.selectbox(f"Trabajador (de pie o sentado) - Área {i}",
                                                               options=["", "De pie", "Sentado"],
                                                               key=f"pos_{i}")
                            techumbre = st.text_input(f"Techumbre - Área {i}", key=f"techumbre_{i}")
                            obs_techumbre = st.text_input(f"Observación techumbre - Área {i}", key=f"obs_techumbre_{i}")
                            paredes = st.text_input(f"Paredes - Área {i}", key=f"paredes_{i}")
                            obs_paredes = st.text_input(f"Observación paredes - Área {i}", key=f"obs_paredes_{i}")
                            ventanales = st.text_input(f"Ventanales - Área {i}", key=f"ventanales_{i}")
                            obs_ventanales = st.text_input(f"Observación ventanales - Área {i}", key=f"obs_ventanales_{i}")
                        st.markdown("##### Otros aspectos")
                        col3, col4 = st.columns(2)
                        with col3:
                            aire_acond = st.text_input(f"Aire acondicionado - Área {i}", key=f"aire_{i}")
                            obs_aire_acond = st.text_input(f"Observaciones aire acondicionado - Área {i}", key=f"obs_aire_{i}")
                            ventiladores = st.text_input(f"Ventiladores - Área {i}", key=f"venti_{i}")
                            obs_ventiladores = st.text_input(f"Observaciones ventiladores - Área {i}", key=f"obs_venti_{i}")
                            inyeccion_extrac = st.text_input(f"Inyección y/o extracción de aire - Área {i}", key=f"inye_{i}")
                            obs_inyeccion = st.text_input(f"Observaciones inyección/extracción de aire - Área {i}", key=f"obs_inye_{i}")
                        with col4:
                            ventanas = st.text_input(f"Ventanas (ventilación natural) - Área {i}", key=f"ventana_{i}")
                            obs_ventanas = st.text_input(f"Observaciones ventanas - Área {i}", key=f"obs_ventana_{i}")
                            puertas = st.text_input(f"Puertas (ventilación natural) - Área {i}", key=f"puertas_{i}")
                            obs_puertas = st.text_input(f"Observaciones puertas - Área {i}", key=f"obs_puertas_{i}")
                            condiciones_disconfort = st.text_input(f"Otras condiciones de disconfort térmico - Área {i}", key=f"cond_{i}")
                            obs_condiciones = st.text_input(f"Observaciones sobre disconfort térmico - Área {i}", key=f"obs_cond_{i}")
                        evidencia = st.text_input(f"Evidencia fotográfica (URL(s) separadas por coma) - Área {i}", key=f"evidencia_{i}")

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

            # ==============================================
            # Pestaña 5: Fin de visita
            # ==============================================
            with tabs[4]:
                #st.header("Fin de visita")
                col1 = st.columns(1)
                with col1[0]:
                    verif_tbs_final = st.text_input("Verificación TBS final")
                    verif_tbh_final = st.text_input("Verificación TBH final")
                    verif_tg_final = st.text_input("Verificación TG final")
                    comentarios_finales = st.text_area("Comentarios finales de evaluación")

            # Botón de envío del formulario
            submitted = st.form_submit_button("Generar Informe")

            if submitted:
                # Consolidar los datos en una estructura para enviar a la función generadora del Word
                datos_informe = {
                    "empresa": {
                        "Razón Social": razon_social,
                        "RUT": rut,
                        "CUV": cuv_val,
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
                st.json(datos_informe)  # Para debug: muestra el JSON generado

                # Llamada a la función que genera el informe en Word.
                informe_docx = generar_informe_en_word(df_filtrado, df_info_cuv)
                st.download_button(
                    label="Descargar Informe",
                    data=informe_docx,
                    file_name=f"informe_{st.session_state['input_cuv_str']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        # Fin del st.form("informe_form")
    else:
        st.info("Ingresa un CUV y haz clic en 'Buscar' para ver la información y generar el informe.")


if __name__ == "__main__":
    main()
