import streamlit as st
import pandas as pd
import io
import zipfile
from data_access import (
    get_centro,
    get_visita,
    get_mediciones,
    get_equipos,
    get_all_cuvs_with_visits
)
from doc_utils import generar_informe_en_word


def generar_informe_desde_cuv(cuv):
    """Genera un informe basado en el CUV y devuelve el archivo en formato BytesIO."""
    df_centro = get_centro(cuv)
    df_visitas = get_visita(cuv)
    df_mediciones = get_mediciones(df_visitas.iloc[0].get("id_visita")) if not df_visitas.empty else pd.DataFrame()
    df_equipos = get_equipos()

    if df_centro.empty or df_visitas.empty:
        st.error(f"No se encontró suficiente información para generar el informe del CUV {cuv}.")
        return None

    informe_docx = generar_informe_en_word(df_centro, df_visitas, df_mediciones, df_equipos)
    return informe_docx


def generar_informes_masivos():
    """Genera informes para todos los CUVs con visitas registradas y los empaqueta en un archivo ZIP."""
    cuvs = get_all_cuvs_with_visits()
    total = len(cuvs)

    if total == 0:
        st.warning("No hay CUVs con visitas registradas.")
        return None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_bar = st.progress(0)

        for i, cuv in enumerate(cuvs):
            try:
                df_centro = get_centro(cuv)
                if df_centro.empty:
                    continue

                df_visitas = get_visita(cuv)
                if df_visitas.empty:
                    continue

                visita_id = df_visitas.iloc[0]["id_visita"]
                df_mediciones = get_mediciones(visita_id)
                df_equipos = get_equipos()

                doc_bytes = generar_informe_en_word(df_centro, df_visitas, df_mediciones, df_equipos)

                # Agregar el informe al archivo ZIP
                zip_file.writestr(f"informe_{cuv}.docx", doc_bytes.getvalue())

            except Exception as e:
                st.error(f"Error generando informe para CUV {cuv}: {str(e)}")

            progress_bar.progress((i + 1) / total)

    zip_buffer.seek(0)
    return zip_buffer


def main():
    st.header("Informes Confort Térmico")
    st.write("Versión 4.0 (Generación Automática)")
    st.write("Bienvenido Rodrigo... (usuario)")

    # Sección para generación manual
    st.subheader("Generar Informe Individual")
    input_cuv = st.text_input("Ingresa el CUV: ej. 178050")

    if st.button("Buscar y Generar Informe"):
        informe = generar_informe_desde_cuv(input_cuv)
        if informe:
            st.success("Informe generado correctamente.")
            st.download_button("Descargar Informe", data=informe, file_name=f"informe_{input_cuv}.docx")

    # Sección para generación automática masiva
    st.subheader("Generación Automática de Informes")
    if st.button("Generar Informes para Todos los CUVs"):
        zip_file = generar_informes_masivos()
        if zip_file:
            st.success("Informes generados correctamente.")
            st.download_button(
                label="Descargar Todos los Informes",
                data=zip_file,
                file_name="informes_confort_termico.zip",
                mime="application/zip"
            )


# Permite que `supermain.py` se pueda ejecutar como script independiente
if __name__ == "__main__":
    main()
