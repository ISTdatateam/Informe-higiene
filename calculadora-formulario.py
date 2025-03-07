"""st.markdown("---")
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
        st.markdown(f"##### El valor de PMV {results.pmv} indica que la sensación térmica es: **{interpretation}**.")"""