import streamlit as st
import pandas as pd
from datetime import datetime, date, time
from data_access2 import get_data   # Función que obtiene el CSV principal
from data_access import insertar_visita, insert_verif_final_visita, insertar_medicion
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

def get_met(puesto_trabajo):
    if puesto_trabajo == "Cajera":
        return 1.1
    elif puesto_trabajo == "Reponedor":
        return 1.2
    elif puesto_trabajo in ["Bodeguero", "Recepcionista"]:
        return 1.89
    else:
        return 1.1  # Opcional: puedes retornar un valor por defecto si el puesto no es reconocido


def check_resultado_pmv(pmv):
    return "Cumple" if -1 <= pmv <= 1 else "No cumple"

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
        cuv_ingresado = input_cuv.strip()
        #st.session_state["input_cuv_str"] = input_cuv.strip()
        st.session_state["df_centro"] = get_centro(cuv_ingresado)
    df_centro = st.session_state.get("df_centro", pd.DataFrame())
    #st.write(df_centro)
    #st.write(df_centro)

    if not df_centro.empty:
        # 1. Datos generales
        st.markdown("#### Datos generales")
        if not df_centro.empty:
            centro_info = df_centro.iloc[0]
            razon_social = st.text_input("Razón Social", value=str(centro_info.get("razon_social", "")))
            rut = st.text_input("RUT", value=str(centro_info.get("rut", "")))
            nombre_local = st.text_input("Nombre de Local", value=str(centro_info.get("nombre_ct", "")))
            direccion = st.text_input("Dirección", value=str(centro_info.get("direccion_ct", "")))
            comuna = st.text_input("Comuna", value=str(centro_info.get("comuna_ct", "")))
            region = st.text_input("Región", value=str(centro_info.get("region_ct", "")))
            cuv_val = st.text_input("CUV", value=str(centro_info.get("cuv", "")))
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
            cargo_consultor = st.selectbox("Cargo del consultor", options=["Seleccione...", "Consultor en Higiene Ocupacional", "Consultor en Prevención de Riesgos"], index=0)
            zonal_consultor = st.selectbox("Zonal del consultor", options=["Seleccione...", "Gerencia Zonal Centro - Viña del Mar", "Gerencia Zonal Sur Austral", "Gerencia Zonal Metropolitana", "Gerencia Zonal Sur", "Gerencia Zonal Norte"], index=0)

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
                verif_tg_inicial,
                cargo_consultor,
                zonal_consultor
            )

            if id_visita:
                st.session_state["id_visita"] = id_visita
                st.success(f"Visita guardada con éxito. ID de la visita: {id_visita}")
            else:
                st.error("Error al guardar la visita en la base de datos.")

        # 3. Formulario 2: Mediciones de Áreas (Formularios Independientes)
        st.subheader("Mediciones de Áreas")
        st.info("Completa y guarda cada área individualmente")

        # Verificar que el ID de la visita existe antes de guardar mediciones
        id_visita = st.session_state.get("id_visita", None)

        # Inicializar un diccionario en session_state para almacenar los IDs de medición si aún no existe
        if "mediciones_ids" not in st.session_state:
            st.session_state["mediciones_ids"] = {}

        if id_visita:
            for i in range(1, 11):  # Iterar por cada área de medición
                area_idx = i - 1
                default_area = st.session_state.areas_data[area_idx] if area_idx < len(
                    st.session_state.areas_data) else {}

                with st.expander(f"Área {i} - Haz clic para expandir", expanded=False):
                    with st.form(key=f"form_area_{i}"):

                        # Captura de datos del formulario
                        nombre_area = st.selectbox(f"Área {i}",
                                                   ["Seleccione...", "Linea de cajas", "Sala de venta", "Bodega",
                                                    "Recepción"], key=f"area_sector_{i}")
                        sector_especifico = st.selectbox(f"Sector específico {i}",
                                                         ["Seleccione...", "Centro", "Izquierda", "Derecha"],
                                                         key=f"espec_sector_{i}")
                        puesto_trabajo = st.selectbox(f"Puesto de trabajo {i}",
                                                      ["Seleccione...", "Cajera", "Reponedor", "Bodeguero",
                                                       "Recepcionista"], key=f"puesto_trabajo_{i}")
                        posicion_trabajador = st.selectbox(f"Posición {i}", ["Seleccione...", "De pie", "Sentado"],
                                                           key=f"pos_trabajador_{i}")
                        vestimenta_trabajador = st.selectbox(f"Vestimenta {i}",
                                                             ["Seleccione...", "Habitual", "Invierno"],
                                                             key=f"vestimenta_{i}")

                        # Mediciones
                        t_bul_seco = st.number_input(f"Temp. bulbo seco (°C) {i}", step=0.1, key=f"tbs_{i}")
                        t_globo = st.number_input(f"Temp. globo (°C) {i}", step=0.1, key=f"tg_{i}")
                        hum_rel = st.number_input(f"Humedad relativa (%) {i}", step=0.1, key=f"hr_{i}")
                        vel_air = st.number_input(f"Velocidad del aire (m/s) {i}", step=0.1, key=f"vel_aire_{i}")

                        # Cálculo de PMV y PPD
                        met = get_met(puesto_trabajo)  # Puede depender del puesto de trabajo
                        clo = 0.5 if vestimenta_trabajador == "Habitual" else 1.0
                        resultados = pmv_ppd_iso(tdb=t_bul_seco, tr=t_globo, vr=vel_air, rh=hum_rel, met=met, clo=clo,
                                                 model="7730-2005", limit_inputs=False)
                        pmv = resultados.pmv
                        ppd = resultados.ppd
                        resultado_medicion = check_resultado_pmv(pmv)

                        # Condiciones y observaciones
                        cond_techumbre = st.radio(f"Techumbre aislante {i}", ["Sí", "No"], key=f"techumbre_{i}")
                        obs_techumbre = st.text_input(f"Obs. Techumbre {i}", key=f"obs_techumbre_{i}")
                        cond_techumbre = 1 if cond_techumbre == "Sí" else 0

                        cond_paredes = st.radio(f"Paredes aislantes {i}", ["Sí", "No"], key=f"paredes_{i}")
                        obs_paredes = st.text_input(f"Obs. Paredes {i}", key=f"obs_paredes_{i}")
                        cond_paredes = 1 if cond_paredes == "Sí" else 0

                        cond_vantanal = st.radio(f"Ventanas aislantes {i}", ["Sí", "No"], key=f"ventanales_{i}")
                        obs_ventanal = st.text_input(f"Obs. Ventanas {i}", key=f"obs_ventanales_{i}")
                        cond_vantanal = 1 if cond_vantanal == "Sí" else 0

                        cond_aire_acond = st.radio(f"Aire acondicionado {i}", ["Sí", "No"], key=f"aire_acond_{i}")
                        obs_aire_acond = st.text_input(f"Obs. Aire Acondicionado {i}", key=f"obs_aire_acond_{i}")
                        cond_aire_acond = 1 if cond_aire_acond == "Sí" else 0

                        cond_ventiladores = st.radio(f"Ventiladores {i}", ["Sí", "No"], key=f"ventiladores_{i}")
                        obs_ventiladores = st.text_input(f"Obs. Ventiladores {i}", key=f"obs_ventiladores_{i}")
                        cond_ventiladores = 1 if cond_ventiladores == "Sí" else 0

                        cond_inyeccion_extraccion = st.radio(f"Inyección/Extracción {i}", ["Sí", "No"],
                                                             key=f"inyeccion_extrac_{i}")
                        obs_inyeccion_extraccion = st.text_input(f"Obs. Inyección {i}", key=f"obs_inyeccion_{i}")
                        cond_inyeccion_extraccion = 1 if cond_inyeccion_extraccion == "Sí" else 0

                        cond_ventanas = st.radio(f"Ventanas abiertas {i}", ["Sí", "No"], key=f"ventanas_{i}")
                        obs_ventanas = st.text_input(f"Obs. Ventanas {i}", key=f"obs_ventanas_{i}")
                        cond_ventanas = 1 if cond_ventanas == "Sí" else 0

                        cond_puertas = st.radio(f"Puertas abiertas {i}", ["Sí", "No"], key=f"puertas_{i}")
                        obs_puertas = st.text_input(f"Obs. Puertas {i}", key=f"obs_puertas_{i}")
                        cond_puertas = 1 if cond_puertas == "Sí" else 0

                        cond_otras = st.radio(f"Puertas abiertas {i}", ["Sí", "No"], key=f"otras_{i}")
                        obs_otras = st.text_input(f"¿Se identifican otras condiciones que pueden considerarse como disconfort térmico? {i}", key=f"obs_otras_{i}")
                        cond_otras = 1 if cond_otras == "Sí" else 0

                        # Guardar medición
                        if st.form_submit_button(f"Guardar Área {i}"):
                            print(f"Inserción de medición - Área {i}")
                            print(f"ID Visita: {id_visita} (Tipo: {type(id_visita)})")
                            print(f"Nombre Área: {nombre_area} (Tipo: {type(nombre_area)})")
                            print(f"Sector Específico: {sector_especifico} (Tipo: {type(sector_especifico)})")
                            print(f"Puesto Trabajo: {puesto_trabajo} (Tipo: {type(puesto_trabajo)})")
                            print(f"Posición Trabajador: {posicion_trabajador} (Tipo: {type(posicion_trabajador)})")
                            print(f"Vestimenta Trabajador: {vestimenta_trabajador} (Tipo: {type(vestimenta_trabajador)})")
                            print(f"T Bulbo Seco: {t_bul_seco} (Tipo: {type(t_bul_seco)})")
                            print(f"T Globo: {t_globo} (Tipo: {type(t_globo)})")
                            print(f"Humedad Relativa: {hum_rel} (Tipo: {type(hum_rel)})")
                            print(f"Velocidad Aire: {vel_air} (Tipo: {type(vel_air)})")
                            print(f"PPD: {ppd} (Tipo: {type(ppd)})")
                            print(f"PMV: {pmv} (Tipo: {type(pmv)})")
                            print(f"Resultado Medición: {resultado_medicion} (Tipo: {type(resultado_medicion)})")
                            print(f"Condición Techumbre: {cond_techumbre} (Tipo: {type(cond_techumbre)})")
                            print(f"Observación Techumbre: {obs_techumbre} (Tipo: {type(obs_techumbre)})")
                            # Solo insertar si todos los datos están completos
                            if nombre_area != "Seleccione..." and sector_especifico != "Seleccione..." and puesto_trabajo != "Seleccione..." and posicion_trabajador != "Seleccione...":

                                id_medicion = insertar_medicion(id_visita, nombre_area, sector_especifico,
                                                                puesto_trabajo, posicion_trabajador,
                                                                vestimenta_trabajador, t_bul_seco, t_globo, hum_rel,
                                                                vel_air, ppd, pmv,
                                                                resultado_medicion, cond_techumbre, obs_techumbre,
                                                                cond_paredes, obs_paredes,
                                                                cond_vantanal, obs_ventanal, cond_aire_acond,
                                                                obs_aire_acond, cond_ventiladores,
                                                                obs_ventiladores, cond_inyeccion_extraccion,
                                                                obs_inyeccion_extraccion, cond_ventanas,
                                                                obs_ventanas, cond_puertas, obs_puertas, cond_otras, obs_otras, met, clo)

                                if id_medicion:
                                    # Almacenar el ID de la medición en session_state pareado con el número de formulario
                                    st.session_state["mediciones_ids"][f"medicion_{i}"] = id_medicion
                                    st.success(f"Área {i} guardada con éxito. ID de la medición: {id_medicion}")
                                    for key, id_medicion in st.session_state["mediciones_ids"].items():
                                        st.write(
                                            f"**{key.replace('_', ' ').capitalize()}** - ID Medición: {id_medicion}")
                                else:
                                    st.error(f"No se pudo guardar la medición para el área {i}.")
                            else:
                                st.warning(f"Completa todos los campos antes de guardar el Área {i}.")

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
                    st.session_state["informe_docx"] = informe_docx


        # Mostrar el botón de descarga solo si ya se generó el informe
        if "informe_docx" in st.session_state:
            st.success("Informe generado correctamente.")
            st.download_button(
                label="Descargar Informe",
                data=st.session_state["informe_docx"],
                file_name=f"informe_{st.session_state.get('input_cuv_str', 'documento')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="descargar_informe"
            )

        # 5. Informe y Calculadora de Confort

    else:
        st.info("Ingresa un CUV y haz clic en 'Buscar' para ver la información y generar el informe.")


if __name__ == "__main__":
    main()
