import streamlit as st
import pandas as pd
from datetime import datetime, date, time

from data_access2 import get_data  # Función que obtiene el CSV principal
from doc_utils import generar_informe_en_word  # Función para generar el Word

from pythermalcomfort.models import pmv_ppd_iso


st.set_page_config(page_title="Informes Confort Térmico", layout="wide")


def interpret_pmv(pmv_value):
    if pmv_value >= 1:
        return "NO CUMPLE"
    elif pmv_value > -1:
        return "CUMPLE"
    else:
        return "NO CUMPLE"


def main():
    st.header("Informes Confort Térmico")
    st.write("")
    st.write("Versión 3.0.20250207")
    st.write("")

    st.write("Mensaje de login")
    st.write("Bienvenido Rodrigo... (usuario)")

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
    # Se inicializa la variable para guardar las mediciones de áreas, en caso de que no exista.
    if "areas_data" not in st.session_state:
        st.session_state["areas_data"] = []

    # --- Búsqueda por CUV ---
    input_cuv = st.text_input("Ingresa el CUV: ej. 183885 ")

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


        # 1: Datos generales
        st.subheader("Datos generales")
        if not df_info_cuv.empty:
            cuv_info_row = df_info_cuv.iloc[0]
            razon_social = st.text_input("Razón Social", value=str(cuv_info_row.get("RAZÓN SOCIAL", "")), disabled=False)
            rut = st.text_input("RUT", value=str(cuv_info_row.get("RUT", "")), disabled=False)
            nombre_local = st.text_input("Nombre de Local", value=str(cuv_info_row.get("Nombre de Local", "")), disabled=False)
            direccion = st.text_input("Dirección", value=str(cuv_info_row.get("Dirección", "")), disabled=False)
            comuna = st.text_input("Comuna", value=str(cuv_info_row.get("Comuna", "")), disabled=False)
            region = st.text_input("Región", value=str(cuv_info_row.get("Región", "")), disabled=False)
            cuv_val = st.text_input("CUV", value=str(cuv_info_row.get("CUV", "")), disabled=False)
        else:
            razon_social = st.text_input("Razón Social")
            rut = st.text_input("RUT")
            nombre_local = st.text_input("Nombre de Local")
            direccion = st.text_input("Dirección")
            comuna = st.text_input("Comuna")
            region = st.text_input("Región")
            cuv_val = st.text_input("CUV")

        st.write("")

        # --- Formulario (st.form) ---

        st.markdown("---")
        st.markdown("INICIO FORMULARIO 1")
        st.markdown("---")



        with st.form("informe_form"):
            # 2: Inicio
            st.subheader("Datos de la visita")

            fecha_visita = st.date_input("Fecha de visita", value=date.today())
            hora_medicion = st.time_input("Hora de medición", value=time(hour=9, minute=0))
            temp_max = st.number_input("Temperatura máxima del día (°C) -- Agregar link", min_value=-50.0, max_value=60.0,
                                       value=0.0, step=0.1)
            motivo_evaluacion = st.selectbox("Motivo de evaluación",
                                             options=["Seleccione...", "Programa anual", "Solicitud empresa", "Fiscalización"],
                                             index=0)
            nombre_personal = st.text_input("Nombre del personal SMU")
            cargo = st.text_input("Cargo - modificar en caso necesario", value="Administador/a")
            consultor_ist = st.text_input("Consultor IST - Esto se va a completar solo mas adelante....")

            # 3: Calibración
            st.markdown("---")
            st.subheader("Calibración")
            cod_equipo_t = st.selectbox(
                "Equipo temperatura",
                options=["Seleccione...", "T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8", "T9", "T10",
                         "T11", "T12", "T13", "T14", "T15", "T16", "T17", "T18", "T19", "T20"]
            )
            cod_equipo_v = st.selectbox(
                "Equipo velocidad aire",
                options=["Seleccione...", "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8", "V9", "V10",
                         "V11", "V12", "V13", "V14", "V15"]
            )
            patron_tbs = st.number_input("Patrón TBS (Sólo modificar en caso necesario)", value=46.4, step=0.1)
            verif_tbs_inicial = st.number_input("Verificación TBS inicial", step=0.1)
            patron_tbh = st.number_input("Patrón TBH (Sólo modificar en caso necesario)", value=12.7, step=0.1)
            verif_tbh_inicial = st.number_input("Verificación TBH inicial", step=0.1)
            patron_tg = st.number_input("Patrón TG (Sólo modificar en caso necesario)", value=69.8, step=0.1)
            verif_tg_inicial = st.number_input("Verificación TG inicial", step=0.1)

            st.markdown("NOTA PARA GPT")
            st.markdown("FIN DE FORMULARIO 1")
            st.markdown("---")


            st.markdown("NOTA PARA GPT")
            st.markdown("INICIO FORMULARIO 2")
            st.markdown("Aquí se requiere generar un formulario diferente por cada area ")
            st.markdown(" y buscar una forma en que se muestren colapsados y al hacer click desplegarlo ")
            st.markdown("---")

            # 4: Mediciones de Áreas (10 áreas fijas)
            st.markdown("---")
            st.markdown("---")
            st.markdown("---")
            st.subheader("Mediciones de Áreas")
            st.info("Selecciona un área para agregar información")
            # Se crean 10 sub-tabs, uno por cada área
            area_tabs = st.tabs([f"Área {i}" for i in range(1, 11)])
            areas_data = []
            for i, area_tab in enumerate(area_tabs, start=1):
                with area_tab:
                    st.markdown(f"#### Identificación del Área {i}")
                    # Widgets de identificación del área
                    default_area = {"Area o sector": "Seleccione..."}
                    options_area_sector = ["Seleccione...", "Linea de cajas", "Sala de venta", "Bodega", "Recepción"]
                    default_area_sector = default_area.get("Area o sector", "Seleccione...")
                    index_area_sector = options_area_sector.index(
                        default_area_sector) if default_area_sector in options_area_sector else 0
                    area_sector = st.selectbox(
                        f"Área {i}",
                        options=options_area_sector,
                        index=index_area_sector,
                        key=f"area_sector_{i}"
                    )
                    espec_sector = st.selectbox(f"Sector específico dentro de área {i}",
                                                options=["Seleccione...", "Centro", "Izquierda", "Derecha"],
                                                key=f"espec_{i}")

                    puesto_trabajo = st.selectbox(
                        f"Puesto de trabajo área {i}",
                        options=["Seleccione...", "Cajera", "Reponedor", "Bodeguero", "Recepcionista"],
                        index=0,
                        key=f"puesto_{i}"
                    )

                    posicion_trabajador = st.selectbox(
                        f"Medición a trabajadores área {i}",
                        options=["Seleccione...", "De pie - 1.10 m", "Sentado - 0.60 m"],
                        index=0,
                        key=f"pos_{i}"
                    )

                    vestimenta = st.selectbox(
                        f"Vestimenta trabajadores área {i}",
                        options=["Seleccione...", "Vestimenta habitual", "Vestimenta de invierno"],
                        index=0,
                        key=f"ves_{i}"
                    )

                    st.markdown(f"#### Mediciones del área {i}")

                    tbs = st.number_input(f"Temperatura bulbo seco (°C) - Área {i}", value=0.0, step=0.1, key=f"tbs_{i}")
                    tg = st.number_input(f"Temperatura globo (°C) - Área {i}", value=0.0, step=0.1, key=f"tg_{i}")
                    hr = st.number_input(f"Humedad relativa (%) - Área {i}", min_value=0.0, max_value=100.0,
                                         value=0.0, step=0.1, key=f"hr_{i}")
                    vel_aire = st.number_input(f"Velocidad del aire (m/s) - Área {i}", min_value=0.0,
                                               max_value=20.0, value=0.0, step=0.1, key=f"vel_{i}")

                    st.markdown(f"#### Resultado calculado área {i}")
                    st.markdown(f"#### PPD: 22,88%")
                    st.markdown(f"#### PMV: 0,92")
                    st.markdown(f"#### CUMPLE")
                    st.write("")
                    st.write("")
                    st.write("Area 2")
                    st.write("")
                    st.write("")
                    st.write("En caso de CUMPLE lo siguiente no se completa")
                    st.write("")
                    st.write("")
                    st.markdown(f"#### Condiciones del área {i}")

                    techumbre = st.pills(f"**La techumbre del área evaluada, cuenta con materiales aislantes térmico tales como: Policarbonatos extendidos, lana mineral, gomas, espuma de poliuretano, entre otros.**", key=f"techumbre_{i}", options=["Sí", "No"])
                    obs_techumbre = st.text_input(f"Observación techumbre - obligatorio en caso de disconformidad", key=f"obs_techumbre_{i}")
                    st.write("")
                    paredes = st.pills(f"**En las paredes del área evaluada,donde  incida directamente el sol, cuentan con material aislante térmicos tales como: Policarbonatos extendidos, lana mineral, gomas, espuma de poliuretano, o construcción de hormigón entre otros (especificar)**", key=f"paredes_{i}", options=["Sí", "No"])
                    obs_paredes = st.text_input(f"Observación paredes - obligatorio en caso de disconformidad",
                                                key=f"obs_paredes_{i}")
                    st.write("")
                    ventanales = st.pills(f"**Los ventanales del área en los cuales incide directamente el sol, cuentan con algún tipo de material aislante, tales como laminas de protección solar, cortinas, gigantografías que proporcionen un apantallamiento.**", key=f"ventanales_{i}" , options=["Sí", "No"])
                    obs_ventanales = st.text_input(f"Observación ventanales - obligatorio en caso de disconformidad", key=f"obs_ventanales_{i}")
                    st.write("")
                    aire_acond = st.pills(f"**El área cuenta con sistema de aire acondicionado y/o enfriador de aire, que proporcione movimiento de aire frío, con el fin de mejorar el confort para los trabajadores. Indicar temperatura de funcionamiento (°C).**", key=f"aire_{i}", options=["Sí", "No"])
                    obs_aire_acond = st.text_input(f"Observaciones aire acondicionado - obligatorio especificar cantidad y operatividad.", key=f"obs_aire_{i}")
                    st.write("")
                    ventiladores = st.pills(f"**El área cuenta con ventiladores que proporcionan movimiento de aire, con el fin de mejorar el confort para los trabajadores.**", key=f"venti_{i}", options=["Sí", "No"])
                    obs_ventiladores = st.text_input(f"Observaciones ventiladores - obligatorio especificar cantidad y operatividad.", key=f"obs_venti_{i}")
                    st.write("")
                    inyeccion_extrac = st.pills(f"**El área cuenta con inyección y/o extracción de aire.**", key=f"inye_{i}", options=["Sí", "No"])
                    obs_inyeccion = st.text_input(f"Observaciones inyección/extracción de aire - obligatorio especificar cantidad y operatividad.", key=f"obs_inye_{i}")
                    st.write("")
                    ventanas = st.pills(f"**El área evaluada cuenta con ventanas que permiten la entrada de aire fresco y la salida de aire caliente, favoreciendo así la ventilación natural.**", key=f"ventana_{i}", options=["Sí", "No"])
                    obs_ventanas = st.text_input(f"Observaciones ventanas - obligatorio en caso de disconformidad", key=f"obs_ventana_{i}")
                    st.write("")
                    puertas = st.pills(f"**El área evaluada cuenta con puertas que proporcionan una vía adicional para la ventilación natural, al permitir la circulación del aire de forma controlada.**", key=f"puertas_{i}", options=["Sí", "No"])
                    obs_puertas = st.text_input(f"Observaciones puertas - obligatorio en caso de disconformidad", key=f"obs_puertas_{i}")
                    st.write("")
                    condiciones_disconfort = st.pills(f"**El area presenta otras condiciones que pueden considerarse como causantes de disconfort térmico, que no hayan sido indicadas en los puntos anteriores.**", key=f"cond_{i}", options=["Sí", "No"])
                    obs_condiciones = st.text_input(f"Observaciones sobre disconfort térmico - obligatorio en caso de presencia", key=f"obs_cond_{i}")

                    st.markdown(f"#### Evidencia fotográfica del área {i}")
                    foto = st.file_uploader(f"Adjunta una foto para el Área {i}", type=["png", "jpg", "jpeg"], key=f"foto_{i}")
                    if foto is not None:
                        st.image(foto, caption=f"Foto cargada área {i}", use_column_width=True)

                    areas_data.append({
                        "Area o sector": area_sector,
                        "Especificación sector": espec_sector,
                        "Temperatura bulbo seco": tbs,
                        "Temperatura globo": tg,
                        "Humedad relativa": hr,
                        "Velocidad del aire": vel_aire,
                        "Puesto de trabajo": puesto_trabajo,
                        "Trabajador de pie o sentado": posicion_trabajador,
                        "Vestimenta": vestimenta,
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
                    })
            st.markdown("---")

            st.markdown("NOTA PARA GPT")
            st.markdown("FIN FORMULARIOS tipo 2 (uno para cara area)")
            st.markdown("---------------")


            submitted = st.form_submit_button("Guardar información")
            # 5: Fin

            st.markdown("NOTA PARA GPT")
            st.markdown("INICIO FORMULARIOS tipo 3")
            st.markdown("---------------")

            st.subheader("Cierre")
            verif_tbs_final = st.text_input("Verificación TBS final")
            verif_tbh_final = st.text_input("Verificación TBH final")
            verif_tg_final = st.text_input("Verificación TG final")
            comentarios_finales = st.text_area("Comentarios finales de evaluación")

            # Botón de envío del formulario
            #submitted = st.form_submit_button("Guardar información")

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
                        "Consultor IST": consultor_ist,
                        "Equipo temperatura": cod_equipo_t,
                        "Equipo velocidad viento": cod_equipo_v,
                        "Verificación TBS inicial": verif_tbs_inicial,
                        "Verificación TBH inicial": verif_tbh_inicial,
                        "Verificación TG inicial": verif_tg_inicial,
                        "Verificación TBS final": verif_tbs_final,
                        "Verificación TBH final": verif_tbh_final,
                        "Verificación TG final": verif_tg_final,
                        "Patrón TBS": patron_tbs,
                        "Patrón TBH": patron_tbh,
                        "Patrón TG": patron_tg,
                        "Motivo de evaluación": motivo_evaluacion,
                        "Comentarios finales de evaluación": comentarios_finales,
                    },
                    "mediciones_areas": areas_data
                }

                st.success("¡Formulario enviado correctamente!")
                # Para debug: se puede visualizar el JSON resultante
                # st.json(datos_informe)

                # Guardar la información de las áreas en session_state para la calculadora
                st.session_state["areas_data"] = areas_data

                # Llamada a la función que genera el informe en Word.
                # informe_docx = generar_informe_en_word(df_filtrado, df_info_cuv)
                # st.download_button(
                #     label="Descargar Informe",
                #     data=informe_docx,
                #     file_name=f"informe_{st.session_state['input_cuv_str']}.docx",
                #     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                # )

        # Fin del st.form("informe_form")


        st.markdown("NOTA PARA GPT")
        st.markdown("FIN FORMULARIOS tipo 3")
        st.markdown("---------------")

        st.markdown("---")
        st.header("Informe")
        st.markdown("---")

        # --- Calculadora de PMV y PPD (Modo Mixto) ---
        st.markdown("---")
        st.header("Calculadora de confort")
        st.write("Selecciona el área ver los resultados calculados automáticamente.")

        # 1. Seleccionar el área cuyos datos se usarán para la calculadora
        if "areas_data" in st.session_state and st.session_state["areas_data"]:
            area_options = [
                f"Área {i + 1} - {area.get('Area o sector', 'Sin dato')}"
                for i, area in enumerate(st.session_state["areas_data"])
            ]
            opcion_area = st.selectbox("Selecciona el área para el cálculo de PMV/PPD", options=area_options)
            indice_area = int(opcion_area.split(" ")[1]) - 1  # Convertir a índice 0-based
            datos_area = st.session_state["areas_data"][indice_area]

            # Valores para temperatura, humedad y velocidad obtenidos del área
            tdb_default = datos_area.get("Temperatura bulbo seco", 0.0)
            tr_default = datos_area.get("Temperatura globo", 0.0)
            rh_default = datos_area.get("Humedad relativa", 0.0)
            v_default = datos_area.get("Velocidad del aire", 0.8)

            # Obtener los valores de "Puesto de trabajo" y "Vestimenta" almacenados en el área
            puesto_default = datos_area.get("Puesto de trabajo", "Cajera")
            vestimenta_default = datos_area.get("Vestimenta", "Vestimenta habitual")
        else:
            st.warning("No hay datos de áreas en la sesión. Se usarán valores por defecto.")
            tdb_default, tr_default, rh_default, v_default = 30.0, 30.0, 32.0, 0.8
            puesto_default, vestimenta_default = "Cajera", "Vestimenta habitual"

        st.markdown("### Ajusta o verifica los valores del área seleccionada")

        # --- Mostrar parámetros fijos para met y clo según el área seleccionada ---
        met_mapping = {
            "Cajera": 1.1,
            "Reponedor": 1.2,
            "Bodeguero": 1.89,
            "Recepcionista": 1.89
        }
        clo_mapping = {
            "Vestimenta habitual": 0.5,
            "Vestimenta de invierno": 1.0
        }
        met = met_mapping.get(puesto_default, 1.2)
        clo_dynamic = clo_mapping.get(vestimenta_default, 0.5)
        st.write("Puesto de trabajo:", puesto_default, " -- ", met, " met")
        #st.write("Tasa metabólica (met):", met)
        st.write("Vestimenta:", vestimenta_default, " -- clo ", clo_dynamic)
        #st.write("Aislamiento de la ropa (clo):", clo_dynamic)


        tdb = st.number_input("Temperatura de bulbo seco (°C):", value=tdb_default)
        tr = st.number_input("Temperatura radiante (°C):", value=tr_default)
        rh = st.number_input("Humedad relativa (%):", value=rh_default)
        v = st.number_input("Velocidad del aire (m/s):", value=v_default)

        # 2. Calcular PMV y PPD con los valores actuales
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
