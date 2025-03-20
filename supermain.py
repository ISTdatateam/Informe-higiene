import streamlit as st
import zipfile
import io

from data_access import (
    get_centro,
    get_visita,
    get_mediciones,
    get_equipos,
    get_all_cuvs_with_visits  # Nueva función importada
)
from doc_utils import generar_informe_en_word


def main():
    st.header("Informes Confort Térmico")
    st.write("Versión 4.0 (Generación Automática)")
    st.write("Bienvenido Rodrigo... (usuario)")


    # Inicializamos en session_state las variables que usaremos para almacenar la info consultada
    if "df_centro" not in st.session_state:
        st.session_state["df_centro"] = None
    if "df_visitas" not in st.session_state:
        st.session_state["df_visitas"] = None
    if "df_mediciones" not in st.session_state:
        st.session_state["df_mediciones"] = None
    if "df_equipos" not in st.session_state:
        st.session_state["df_equipos"] = None
    if "input_cuv" not in st.session_state:
        st.session_state["input_cuv"] = ""

    # Sección para generación manual (original)
    st.subheader("Generar Informe Individual")
    input_cuv = st.text_input("Ingresa el CUV: ej. 178050")

    if st.button("Buscar"):
        # Guardamos el CUV ingresado en session_state
        st.session_state["input_cuv"] = input_cuv.strip()

        # Consulta de la información del centro de trabajo
        df_centro = get_centro(st.session_state["input_cuv"])
        # Consulta de las visitas para el CUV (se ordenan para obtener la visita más reciente)
        df_visitas = get_visita(st.session_state["input_cuv"])
        # Si existe al menos una visita, se selecciona la más reciente y se consultan las mediciones asociadas
        if not df_visitas.empty:
            visita_id = df_visitas.iloc[0].get("id_visita")
            df_mediciones = get_mediciones(visita_id)
        else:
            df_mediciones = pd.DataFrame()
        # Se obtiene la información completa de equipos de medición
        df_equipos = get_equipos()

        # Se actualizan los valores en session_state
        st.session_state["df_centro"] = df_centro
        st.session_state["df_visitas"] = df_visitas
        st.session_state["df_mediciones"] = df_mediciones
        st.session_state["df_equipos"] = df_equipos

    # Mostramos el resumen de la información consultada
    df_centro = st.session_state["df_centro"]
    df_visitas = st.session_state["df_visitas"]
    df_mediciones = st.session_state["df_mediciones"]

    if df_centro is not None and not df_centro.empty:
        st.subheader("Resumen de Información del Centro de Trabajo")
        centro = df_centro.iloc[0]
        st.write(f"**CUV**: {centro.get('cuv', '')}")
        st.write(f"**RUT**: {centro.get('rut', '')}")
        st.write(f"**Razón Social**: {centro.get('razon_social', '')}")
        st.write(f"**Nombre de Local**: {centro.get('nombre_ct', '')}")
        st.write(f"**Dirección**: {centro.get('direccion_ct', '')}")
        st.write(f"**Comuna**: {centro.get('comuna_ct', '')}")
        st.write(f"**Región**: {centro.get('region_ct', '')}")
    else:
        st.info("No se encontró información del centro para este CUV.")

    if df_visitas is not None and not df_visitas.empty:
        st.subheader("Información de la Visita más Reciente")
        visita = df_visitas.iloc[0]
        st.write(f"**Fecha de Visita**: {visita.get('fecha_visita', '')}")
        st.write(f"**Hora de Visita**: {visita.get('hora_visita', '')}")
        st.write(f"**Motivo de Evaluación**: {visita.get('motivo_evaluacion', '')}")
        st.write(f"**Personal de Visita**: {visita.get('nombre_personal_visita', '')} - {visita.get('cargo_personal_visita', '')}")
    else:
        st.info("No se encontró información de visitas para este CUV.")

    if df_mediciones is not None and not df_mediciones.empty:
        st.subheader("Mediciones Asociadas a la Visita")
        st.dataframe(df_mediciones)
    else:
        st.info("No se encontraron mediciones para este CUV.")

    # Botón para generar el informe en Word
    if st.button("Generar Informe en Word"):
        if (st.session_state["df_centro"] is not None and not st.session_state["df_centro"].empty) and \
           (st.session_state["df_visitas"] is not None and not st.session_state["df_visitas"].empty):
            # Se llama a la función generadora pasando los dataframes obtenidos
            informe_docx = generar_informe_en_word(
                st.session_state["df_centro"],
                st.session_state["df_visitas"],
                st.session_state["df_mediciones"],
                st.session_state["df_equipos"]
            )
            st.session_state["nombre_ct"] = df_centro.get('nombre_ct', '')


            st.download_button(
                label="Descargar Informe",
                data=informe_docx,
                file_name=f"informe_termico_{st.session_state['input_cuv']}_{st.session_state['nombre_ct']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("No hay información suficiente para generar el informe. Verifica el CUV ingresado.")

    # Nueva sección para generación automática
    st.subheader("Generación Automática de Informes")
    if st.button("Generar Informes para Todos los CUVs"):
        cuvs = get_all_cuvs_with_visits()
        total = len(cuvs)

        if total == 0:
            st.warning("No hay CUVs con visitas registradas")
            return

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            progress_bar = st.progress(0)

            for i, cuv in enumerate(cuvs):
                try:
                    # Obtener datos del centro
                    df_centro = get_centro(cuv)
                    if df_centro.empty:
                        continue

                    # Obtener última visita
                    df_visitas = get_visita(cuv)
                    if df_visitas.empty:
                        continue

                    # Obtener mediciones
                    visita_id = df_visitas.iloc[0]["id_visita"]
                    df_mediciones = get_mediciones(visita_id)
                    df_equipos = get_equipos()

                    # Generar informe
                    doc_bytes = generar_informe_en_word(
                        df_centro,
                        df_visitas,
                        df_mediciones,
                        df_equipos
                    )

                    centro = df_centro.iloc[0]
                    # Agregar al ZIP
                    zip_file.writestr(
                        f"informe_termico_{cuv}_{centro.get('nombre_ct', '')}.docx",
                        doc_bytes.getvalue()  # Asumiendo que devuelve BytesIO
                    )

                except Exception as e:
                    st.error(f"Error con CUV {cuv}: {str(e)}")

                # Actualizar progreso
                progress_bar.progress((i + 1) / total)

        # Preparar descarga del ZIP
        zip_buffer.seek(0)
        st.download_button(
            label="Descargar Todos los Informes",
            data=zip_buffer,
            file_name="informes_confort_termico.zip",
            mime="application/zip"
        )


if __name__ == "__main__":
    main()