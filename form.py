import streamlit as st
import pandas as pd
from datetime import datetime, date, time
from data_access2 import get_data   # Función que obtiene el CSV principal
from data_access import insertar_visita, insert_verif_final_visita   # Función que obtiene el CSV principal
from doc_utils import generar_informe_en_word  # Función para generar el Word
from pythermalcomfort.models import pmv_ppd_iso
import zipfile
import io
from data_access import (
    get_centro,
    get_visita,
    get_mediciones,
    get_equipos,
    get_all_cuvs_with_visits  # Si también deseas agregar la generación masiva
)
from doc_utils import generar_informe_en_word
from informe import generar_informe_desde_cuv

st.set_page_config(page_title="Informes Confort Térmico", layout="wide")


def interpret_pmv(pmv_value):
    if pmv_value >= 2.5:
        return "Calurosa"
    elif pmv_value >= 1.5:
        return "Cálida"
    elif pmv_value >= 0.5:
        return "Ligeramente cálida"
    elif pmv_value > -0.5:
        return "Neutra - Confortable"
    elif pmv_value > -1.5:
        return "Ligeramente fresca"
    elif pmv_value > -2.5:
        return "Fresca"
    else:
        return "Fría"


def main():
    st.header("Informes Confort Térmico")
    st.write("Versión 3.0.20250205")
    st.write("Mensaje de login")
    st.write("Bienvenido Rodrigo... (usuario)")

    # --- Carga de CSVs ---
    csv_url_main = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vTPdZTyxM6BDLmnlqe246tfBm7H06vXBdQKruh2mPg-rhQSD8olCS30ej4BdtJ1R__3W6K-Va3hm5Ax/"
        "pub?output=csv"
    )
    df_main = get_data(csv_url_main)
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
    if "areas_data" not in st.session_state:
        st.session_state["areas_data"] = [{} for _ in range(10)]  # 10 áreas por defecto
    if "datos_generales" not in st.session_state:
        st.session_state["datos_generales"] = {}
    if "cierre" not in st.session_state:
        st.session_state["cierre"] = {}

    # --- Búsqueda por CUV ---
    input_cuv = st.text_input("Ingresa el CUV: ej. 183885")
    if st.button("Buscar"):
        st.session_state["input_cuv_str"] = input_cuv.strip()
        df_main["CUV"] = df_main["CUV"].astype(str).str.strip()
        df_cuv_info["CUV"] = df_cuv_info["CUV"].astype(str).str.strip()
        st.session_state["df_filtrado"] = df_main[df_main["CUV"] == st.session_state["input_cuv_str"]]
        st.session_state["df_info_cuv"] = df_cuv_info[df_cuv_info["CUV"] == st.session_state["input_cuv_str"]]

    df_filtrado = st.session_state["df_filtrado"]
    df_info_cuv = st.session_state["df_info_cuv"]

    if not df_filtrado.empty:
        # 1. Datos generales
        st.markdown("#### Datos generales")
        if not df_info_cuv.empty:
            cuv_info_row = df_info_cuv.iloc[0]
            razon_social = st.text_input("Razón Social", value=str(cuv_info_row.get("RAZÓN SOCIAL", "")))
            rut = st.text_input("RUT", value=str(cuv_info_row.get("RUT", "")))
            nombre_local = st.text_input("Nombre de Local", value=str(cuv_info_row.get("Nombre de Local", "")))
            direccion = st.text_input("Dirección", value=str(cuv_info_row.get("Dirección", "")))
            comuna = st.text_input("Comuna", value=str(cuv_info_row.get("Comuna", "")))
            region = st.text_input("Región", value=str(cuv_info_row.get("Región", "")))
            cuv_val = st.text_input("CUV", value=str(cuv_info_row.get("CUV", "")))
        else:
            razon_social = st.text_input("Razón Social")
            rut = st.text_input("RUT")
            nombre_local = st.text_input("Nombre de Local")
            direccion = st.text_input("Dirección")
            comuna = st.text_input("Comuna")
            region = st.text_input("Región")
            cuv_val = st.text_input("CUV")

        # 2. Formulario 1: Datos de la visita y calibración
        with st.form("form1"):
            st.markdown("#### Datos de la visita")
            fecha_visita = st.date_input("Fecha de visita", value=date.today())
            hora_medicion = st.time_input("Hora de medición", value=time(hour=9, minute=0))
            temp_max = st.number_input("Temperatura máxima del día (°C)", min_value=-50.0, max_value=60.0, value=0.0, step=0.1)
            motivo_evaluacion = st.selectbox("Motivo de evaluación",
                                             options=["Seleccione...", "Programa anual", "Solicitud empresa", "Fiscalización"],
                                             index=0)
            nombre_personal = st.text_input("Nombre del personal SMU")
            cargo = st.text_input("Cargo", value="Administador/a")
            consultor_ist = st.text_input("Consultor IST")
            st.markdown("#### Verificación de parámetros")
            cod_equipo_t = st.selectbox("Equipo temperatura",
                                        options=["Seleccione...",
                                                 "T1",
                                                 "T2",
                                                 "T3",
                                                 "T4",
                                                 "T5",
                                                 "T6",
                                                 "T7",
                                                 "T8",
                                                 "T9",
                                                 "T10",
                                                 "T11",
                                                 "T12",
                                                 "T13",
                                                 "T14",
                                                 "T15",
                                                 "T16",
                                                 "T17",
                                                 "T18",
                                                 "T19",
                                                 "T20",
                                                 "T21",
                                                 "T22",
                                                 "T23",
                                                 "T24",
                                                 "T25",
                                                 "T26",
                                                 "T27",
                                                 "T28",
                                                 "T29",
                                                 "T30",
                                                 "T31",
                                                 "T32"],
                                        index=0)
            cod_equipo_v = st.selectbox("Equipo velocidad aire",
                                        options=["V1",
                                                 "V2",
                                                 "V3",
                                                 "V4",
                                                 "V5",
                                                 "V6",
                                                 "V7",
                                                 "V8",
                                                 "V9",
                                                 "V10",
                                                 "V11",
                                                 "V12",
                                                 "V13",
                                                 "V14",
                                                 "V15",
                                                 ],
                                        index=0)
            patron_tbs = st.number_input("Patrón TBS", value=46.4, step=0.1)
            verif_tbs_inicial = st.number_input("Verificación TBS inicial", step=0.1)
            patron_tbh = st.number_input("Patrón TBH", value=12.7, step=0.1)
            verif_tbh_inicial = st.number_input("Verificación TBH inicial", step=0.1)
            patron_tg = st.number_input("Patrón TG", value=69.8, step=0.1)
            verif_tg_inicial = st.number_input("Verificación TG inicial", step=0.1)
            submit1 = st.form_submit_button("Guardar datos")
        if submit1:
            st.session_state["datos_generales"] = {
                "Fecha visita": fecha_visita,
                "Hora medición": hora_medicion,
                "Temperatura máxima del día": temp_max,
                "Motivo de evaluación": motivo_evaluacion,
                "Nombre del personal SMU": nombre_personal,
                "Cargo": cargo,
                "Consultor IST": consultor_ist,
                "Equipo temperatura": cod_equipo_t,
                "Equipo velocidad aire": cod_equipo_v,
                "Patrón TBS": patron_tbs,
                "Verificación TBS inicial": verif_tbs_inicial,
                "Patrón TBH": patron_tbh,
                "Verificación TBH inicial": verif_tbh_inicial,
                "Patrón TG": patron_tg,
                "Verificación TG inicial": verif_tg_inicial
            }

            # Guardar la visita en la base de datos
            id_visita = insertar_visita(
                st.session_state["input_cuv_str"],
                fecha_visita,
                hora_medicion,
                temp_max,
                motivo_evaluacion,
                nombre_personal,
                cargo,
                consultor_ist,
                cod_equipo_t,
                cod_equipo_v,
                patron_tbs,
                verif_tbs_inicial,
                patron_tbh,
                verif_tbh_inicial,
                patron_tg,
                verif_tg_inicial
            )

            if id_visita:
                st.session_state["id_visita"] = id_visita
                st.success(f"Visita guardada con éxito. ID de la visita: {id_visita}")
            else:
                st.error("Error al guardar la visita en la base de datos.")

        # 3. Formulario 2: Mediciones de Áreas (Formularios Independientes)
        st.subheader("Mediciones de Áreas")
        st.info("Completa y guarda cada área individualmente")

        # Asegurar que siempre hay 10 áreas en session_state
        if len(st.session_state.get("areas_data", [])) < 10:
            st.session_state.areas_data = st.session_state.get("areas_data", [{}]) + [{}] * (
                        10 - len(st.session_state.get("areas_data", [])))

        for i in range(1, 11):
            area_idx = i - 1  # Índice base 0
            default_area = st.session_state.areas_data[area_idx] if area_idx < len(st.session_state.areas_data) else {}

            with st.expander(f"Área {i} - Haz clic para expandir", expanded=False):
                with st.form(key=f"form_area_{i}"):
                    st.markdown(f"#### Identificación del Área {i}")

                    # Listas de opciones
                    options_area_sector = ["Seleccione...", "Linea de cajas", "Sala de venta", "Bodega", "Recepción"]
                    options_espec = ["Seleccione...", "Centro", "Izquierda", "Derecha"]
                    options_puesto = ["Seleccione...", "Cajera", "Reponedor", "Bodeguero", "Recepcionista"]
                    options_pos = ["Seleccione...", "De pie - 1.10 m", "Sentado - 0.60 m"]
                    options_vestimenta = ["Seleccione...", "Vestimenta habitual", "Vestimenta de invierno"]

                    # Campos del formulario
                    area_sector = st.selectbox(
                        f"Área {i}",
                        options=options_area_sector,
                        index=options_area_sector.index(default_area.get("Area o sector", "Seleccione...")),
                        key=f"area_sector_{i}"
                    )

                    espec_sector = st.selectbox(
                        f"Sector específico dentro de área {i}",
                        options=options_espec,
                        index=options_espec.index(default_area.get("Especificación sector", "Seleccione...")),
                        key=f"espec_sector_{i}"
                    )

                    puesto_trabajo = st.selectbox(
                        f"Puesto de trabajo área {i}",
                        options=options_puesto,
                        index=options_puesto.index(default_area.get("Puesto de trabajo", "Seleccione...")),
                        key=f"puesto_trabajo_{i}"
                    )

                    posicion_trabajador = st.selectbox(
                        f"Medición a trabajadores área {i}",
                        options=options_pos,
                        index=options_pos.index(default_area.get("Trabajador de pie o sentado", "Seleccione...")),
                        key=f"pos_trabajador_{i}"
                    )

                    vestimenta = st.selectbox(
                        f"Vestimenta trabajadores área {i}",
                        options=options_vestimenta,
                        index=options_vestimenta.index(default_area.get("Vestimenta", "Seleccione...")),
                        key=f"vestimenta_{i}"
                    )

                    st.markdown(f"#### Mediciones del área {i}")
                    tbs = st.number_input(
                        f"Temperatura bulbo seco (°C) - Área {i}",
                        value=float(default_area.get("Temperatura bulbo seco", 0.0)),
                        step=0.1,
                        key=f"tbs_{i}"
                    )

                    tg = st.number_input(
                        f"Temperatura globo (°C) - Área {i}",
                        value=float(default_area.get("Temperatura globo", 0.0)),
                        step=0.1,
                        key=f"tg_{i}"
                    )

                    hr = st.number_input(
                        f"Humedad relativa (%) - Área {i}",
                        min_value=0.0,
                        max_value=100.0,
                        key=f"hr_{i}"
                    )

                    vel_aire = st.number_input(
                        f"Velocidad del aire (m/s) - Área {i}",
                        min_value=0.0,
                        max_value=20.0,
                        value=float(default_area.get("Velocidad del aire", 0.0)),
                        step=0.1,
                        key=f"vel_aire_{i}"
                    )

                    # Sección condicional para observaciones
                    if tg > 30:
                        st.markdown(f"#### Observaciones y condiciones de confort del área {i}")
                        # Se reemplaza st.pills por st.radio (que acepta 'index')
                        techumbre = st.radio(
                            "¿La techumbre cuenta con materiales aislantes?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Techumbre", "Sí") == "Sí" else 1,
                            key=f"techumbre_{i}"
                        )
                        obs_techumbre = st.text_input("Observación techumbre",
                                                      value=default_area.get("Observación techumbre", ""),
                                                      key=f"obs_techumbre_{i}")

                        paredes = st.radio(
                            "¿Las paredes cuentan con material aislante?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Paredes", "Sí") == "Sí" else 1,
                            key=f"paredes_{i}"
                        )
                        obs_paredes = st.text_input("Observación paredes",
                                                    value=default_area.get("Observación paredes", ""),
                                                    key=f"obs_paredes_{i}")

                        ventanales = st.radio(
                            "¿Los ventanales tienen material aislante?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Ventanales", "Sí") == "Sí" else 1,
                            key=f"ventanales_{i}"
                        )
                        obs_ventanales = st.text_input("Observación ventanales",
                                                       value=default_area.get("Observación ventanales", ""),
                                                       key=f"obs_ventanales_{i}")

                        aire_acond = st.radio(
                            "¿El área cuenta con aire acondicionado o enfriador?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Aire acondicionado", "Sí") == "Sí" else 1,
                            key=f"aire_acond_{i}"
                        )
                        obs_aire_acond = st.text_input("Observaciones aire acondicionado",
                                                       value=default_area.get("Observaciones aire acondicionado", ""),
                                                       key=f"obs_aire_acond_{i}")

                        ventiladores = st.radio(
                            "¿El área cuenta con ventiladores?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Ventiladores", "Sí") == "Sí" else 1,
                            key=f"ventiladores_{i}"
                        )
                        obs_ventiladores = st.text_input("Observaciones ventiladores",
                                                         value=default_area.get("Observaciones ventiladores", ""),
                                                         key=f"obs_ventiladores_{i}")

                        inyeccion_extrac = st.radio(
                            "¿El área cuenta con inyección y/o extracción de aire?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Inyección y/o extracción de aire", "Sí") == "Sí" else 1,
                            key=f"inyeccion_extrac_{i}"
                        )
                        obs_inyeccion = st.text_input("Observaciones inyección/extracción de aire",
                                                      value=default_area.get("Observaciones inyección/extracción de aire",
                                                                             ""), key=f"obs_inyeccion_{i}")

                        ventanas = st.radio(
                            "¿El área cuenta con ventanas que favorecen la ventilación natural?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Ventanas (ventilación natural)", "Sí") == "Sí" else 1,
                            key=f"ventanas_{i}"
                        )
                        obs_ventanas = st.text_input("Observaciones ventanas",
                                                     value=default_area.get("Observaciones ventanas", ""),
                                                     key=f"obs_ventanas_{i}")

                        puertas = st.radio(
                            "¿El área cuenta con puertas que facilitan la ventilación natural?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Puertas (ventilación natural)", "Sí") == "Sí" else 1,
                            key=f"puertas_{i}"
                        )
                        obs_puertas = st.text_input("Observaciones puertas",
                                                    value=default_area.get("Observaciones puertas", ""),
                                                    key=f"obs_puertas_{i}")

                        condiciones_disconfort = st.radio(
                            "¿Existen otras condiciones que generen disconfort térmico?",
                            options=["Sí", "No"],
                            index=0 if default_area.get("Otras condiciones de disconfort térmico", "Sí") == "Sí" else 1,
                            key=f"condiciones_disconfort_{i}"
                        )
                        obs_condiciones = st.text_input("Observaciones sobre disconfort térmico",
                                                        value=default_area.get("Observaciones sobre disconfort térmico",
                                                                               ""), key=f"obs_condiciones_{i}")

                        st.markdown(f"#### Evidencia fotográfica del área {i}")
                        foto = st.file_uploader(f"Adjunta una foto para el Área {i}", type=["png", "jpg", "jpeg"],
                                                key=f"foto_{i}")
                    else:
                        techumbre = None
                        obs_techumbre = ""
                        paredes = None
                        obs_paredes = ""
                        ventanales = None
                        obs_ventanales = ""
                        aire_acond = None
                        obs_aire_acond = ""
                        ventiladores = None
                        obs_ventiladores = ""
                        inyeccion_extrac = None
                        obs_inyeccion = ""
                        ventanas = None
                        obs_ventanas = ""
                        puertas = None
                        obs_puertas = ""
                        condiciones_disconfort = None
                        obs_condiciones = ""
                        foto = None

                    # Botón de guardado individual
                    submit_area = st.form_submit_button(f"Guardar Área {i}")

                    if submit_area:
                    # Actualizar solo el área correspondiente
                        st.session_state.areas_data[area_idx] = {
                            "Area o sector": area_sector,
                            "Especificación sector": espec_sector,
                            "Puesto de trabajo": puesto_trabajo,
                            "Trabajador de pie o sentado": posicion_trabajador,
                            "Vestimenta": vestimenta,
                            "Temperatura bulbo seco": tbs,
                            "Temperatura globo": tg,
                            "Humedad relativa": hr,
                            "Velocidad del aire": vel_aire,
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
                            "Evidencia fotográfica": foto,
                        }
                    st.success(f"Datos del Área {i} guardados exitosamente!")


        # 4. Formulario 3: Cierre

        def number_input_con_inicial(label, key_inicial=None, default=0.0):
            if key_inicial is not None and key_inicial in st.session_state:
                valor_inicial = st.session_state[key_inicial]
                label = f"{label} - El valor inicial fue de {valor_inicial}"
                default = valor_inicial
            return st.number_input(label, value=default, step=0.1)

        with st.form("form3"):
            st.markdown("#### Verificación final")
            verif_tbs_final = number_input_con_inicial("Verificación TBS final", "verif_tbs_inicial")
            verif_tbh_final = number_input_con_inicial("Verificación TBH final", "verif_tbh_inicial")
            verif_tg_final = number_input_con_inicial("Verificación TG final", "verif_tg_inicial")
            comentarios_finales = st.text_area("Comentarios finales de evaluación")
            submit3 = st.form_submit_button("Guardar Cierre")

        if submit3:
            st.session_state["cierre"] = {
                "Verificación TBS final": verif_tbs_final,
                "Verificación TBH final": verif_tbh_final,
                "Verificación TG final": verif_tg_final,
                "Comentarios finales de evaluación": comentarios_finales
            }

            # Obtener ID de la visita desde session_state
            id_visita = st.session_state.get("id_visita", None)

            if id_visita:
                # Intentamos actualizar el registro en la base de datos
                actualizado = insert_verif_final_visita(id_visita, verif_tbs_final, verif_tbh_final, verif_tg_final,
                                                        comentarios_finales)

                if actualizado:
                    st.success("Formulario 3 guardado y visita actualizada correctamente.")

                    # Guardar el estado de actualización en session_state
                    st.session_state["visita_actualizada"] = True
                else:
                    st.error(
                        "No se pudo actualizar la visita en la base de datos. Revisa los datos e intenta nuevamente.")
                    st.session_state["visita_actualizada"] = False
            else:
                st.error(
                    "No se encontró el ID de la visita en la sesión. Verifica que la visita fue creada correctamente.")
                st.session_state["visita_actualizada"] = False

        # Fuera del `if`, mostrar la opción de generar el informe solo si la actualización fue exitosa
        if st.session_state.get("visita_actualizada", False):
            st.subheader("Generar Informe")
            st.write("Los datos han sido guardados. ¿Deseas generar el informe basado en la base de datos?")

            if st.button("Sí, generar informe automáticamente", key="generar_informe"):
                cuv = st.session_state.get("input_cuv_str", "")
                informe_docx = generar_informe_desde_cuv(cuv)

                if informe_docx:
                    st.success("Informe generado correctamente.")
                    st.download_button(
                        label="Descargar Informe",
                        data=informe_docx,
                        file_name=f"informe_{cuv}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("Hubo un problema al generar el informe.")

        # 5. Informe y Calculadora de Confort

        st.markdown("---")
        st.header("Calculadora de confort")
        st.write("Selecciona el área para ver los resultados calculados automáticamente.")

        if "areas_data" in st.session_state and st.session_state["areas_data"]:
            area_options = [
                f"Área {i + 1} - {area.get('Area o sector', 'Sin dato')}"
                for i, area in enumerate(st.session_state["areas_data"])
            ]
            opcion_area = st.selectbox("Selecciona el área para el cálculo de PMV/PPD", options=area_options)
            indice_area = int(opcion_area.split(" ")[1]) - 1  # Índice 0-based
            datos_area = st.session_state["areas_data"][indice_area]

            tdb_default = datos_area.get("Temperatura bulbo seco", 0.0)
            tr_default = datos_area.get("Temperatura globo", 0.0)
            rh_default = datos_area.get("Humedad relativa", 0.0)
            v_default = datos_area.get("Velocidad del aire", 0.8)
            puesto_default = datos_area.get("Puesto de trabajo", "Cajera")
            vestimenta_default = datos_area.get("Vestimenta", "Vestimenta habitual")
        else:
            st.warning("No hay datos de áreas en la sesión. Se usarán valores por defecto.")
            tdb_default, tr_default, rh_default, v_default = 30.0, 30.0, 32.0, 0.8
            puesto_default, vestimenta_default = "Cajera", "Vestimenta habitual"

        st.markdown("### Ajusta o verifica los valores del área seleccionada")
        met_mapping = {"Cajera": 1.1, "Reponedor": 1.2, "Bodeguero": 1.89, "Recepcionista": 1.89}
        clo_mapping = {"Vestimenta habitual": 0.5, "Vestimenta de invierno": 1.0}
        met = met_mapping.get(puesto_default, 1.2)
        clo_dynamic = clo_mapping.get(vestimenta_default, 0.5)
        st.write("Puesto de trabajo:", puesto_default, " -- ", met, " met")
        st.write("Vestimenta:", vestimenta_default, " -- clo", clo_dynamic)

        tdb = st.number_input("Temperatura de bulbo seco (°C):", value=tdb_default)
        tr = st.number_input("Temperatura radiante (°C):", value=tr_default)
        rh = st.number_input("Humedad relativa (%):", value=rh_default)
        v = st.number_input("Velocidad del aire (m/s):", value=v_default)

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

        st.subheader("Resultados")
        st.write(f"**PMV:** {results.pmv}")
        st.write(f"**PPD:** {results.ppd}%")
        interpretation = interpret_pmv(results.pmv)
        st.markdown(f"##### El valor de PMV {results.pmv} indica que la sensación térmica es: **{interpretation}**.")
    else:
        st.info("Ingresa un CUV y haz clic en 'Buscar' para ver la información y generar el informe.")


if __name__ == "__main__":
    main()
