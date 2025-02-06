import streamlit as st
import pandas as pd

from data_access import get_data  # Tu función que obtiene el CSV principal
from doc_utils import generar_informe_en_word  # Ajusta para usar la versión modificada

def main():
    st.title("Informes confort térmico")
    st.write("")
    st.write("Version 0.4.20250203")
    st.write("")

    # CSV 1: Datos generales / mediciones
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

    # Mostramos columnas (opcional, para debug)
    #st.write("Columnas en df_main:", df_main.columns.tolist())
    #st.write("Columnas en df_cuv_info:", df_cuv_info.columns.tolist())

    # Inicializamos en session_state las variables que usaremos
    if "df_filtrado" not in st.session_state:
        st.session_state["df_filtrado"] = pd.DataFrame()
    if "df_info_cuv" not in st.session_state:
        st.session_state["df_info_cuv"] = pd.DataFrame()
    if "input_cuv_str" not in st.session_state:
        st.session_state["input_cuv_str"] = ""

    # Campo para ingresar el CUV
    input_cuv = st.text_input("Ingresa el CUV: ej. 183885")

    # Botón "Buscar"
    if st.button("Buscar"):
        # Guardamos el CUV en session_state para uso posterior
        st.session_state["input_cuv_str"] = input_cuv.strip()

        # Filtramos en el df_main y en df_cuv_info
        df_main["CUV"] = df_main["CUV"].astype(str).str.strip()
        df_cuv_info["CUV"] = df_cuv_info["CUV"].astype(str).str.strip()

        # Filtrado principal
        st.session_state["df_filtrado"] = df_main[df_main["CUV"] == st.session_state["input_cuv_str"]]

        # Filtrado info CUV (RUT, Razón Social, etc.)
        st.session_state["df_info_cuv"] = df_cuv_info[df_cuv_info["CUV"] == st.session_state["input_cuv_str"]]

    # Mostramos resultados
    df_filtrado = st.session_state["df_filtrado"]
    df_info_cuv = st.session_state["df_info_cuv"]

    if not df_filtrado.empty:
        st.subheader("Resumen de Datos Generales")

        # Si hay datos de info_cuv, mostramos en resumen
        if not df_info_cuv.empty:
            # Tomamos la primera fila
            cuv_info_row = df_info_cuv.iloc[0]
            st.write(f"**CUV**: {cuv_info_row.get('CUV', '')}")
            st.write(f"**RUT**: {cuv_info_row.get('RUT', '')}")
            st.write(f"**Razón Social**: {cuv_info_row.get('RAZÓN SOCIAL', '')}")
            st.write(f"**Nombre de Local**: {cuv_info_row.get('Nombre de Local', '')}")
            st.write(f"**Dirección**: {cuv_info_row.get('Dirección', '')}")
            st.write(f"**Comuna**: {cuv_info_row.get('Comuna', '')}")
            st.write(f"**Región**: {cuv_info_row.get('Región', '')}")
        else:
            st.warning("No se encontró información de RUT/Razón Social/etc. para este CUV.")

        # Muestra un resumen de datos generales y mediciones
        df_datos_generales = df_filtrado[df_filtrado["Que seccion quieres completar"] == "Datos generales"]
        if not df_datos_generales.empty:
            st.write("**Datos Generales hallados en CSV principal**")
            #st.dataframe(df_datos_generales)
        else:
            st.write("No hay 'Datos Generales' en el CSV principal para este CUV.")

        df_mediciones = df_filtrado[df_filtrado["Que seccion quieres completar"] == "Medición de un area"]
        if not df_mediciones.empty:
            st.write("**Mediciones halladas**")
            #st.dataframe(df_mediciones)
        else:
            st.write("No hay 'Medición de un area' en el CSV principal para este CUV.")

        # Botón para generar informe en Word
        if st.button("Generar Informe en Word"):
            informe_docx = generar_informe_en_word(df_filtrado, df_info_cuv)
            st.download_button(
                label="Descargar Informe",
                data=informe_docx,
                file_name=f"informe_{st.session_state['input_cuv_str']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.info("Ingresa un CUV y haz clic en 'Buscar' para ver la información y generar el informe.")

if __name__ == "__main__":
    main()
