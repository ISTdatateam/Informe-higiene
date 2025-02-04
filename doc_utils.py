import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from datetime import datetime
import requests
import re
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def get_or_add_lang(rPr):
    # Busca el elemento 'w:lang'
    lang = rPr.find(qn('w:lang'))
    if lang is None:
        # Si no existe, lo creamos y lo añadimos a rPr
        lang = OxmlElement('w:lang')
        rPr.append(lang)
    return lang





def descargar_imagen_gdrive(url_foto: str) -> BytesIO or None:
    """
    Dada una URL de Google Drive (tipo 'https://drive.google.com/open?id=ABC123'),
    intenta descargar el archivo y devolverlo como BytesIO.
    Retorna None si no se puede descargar o no tenga permisos adecuados.
    """
    match = re.search(r"id=(.*)", url_foto)
    if not match:
        return None

    file_id = match.group(1)
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    try:
        response = requests.get(download_url)
        if response.status_code == 200:
            return BytesIO(response.content)
        else:
            return None
    except Exception as e:
        return None

def generar_informe_en_word(df_filtrado: pd.DataFrame,
                            df_info_cuv: pd.DataFrame) -> BytesIO:
    """
    Genera un informe en Word (DOCX) combinando:
      - df_info_cuv: Info del CUV (RUT, RAZÓN SOCIAL, etc.)
      - df_filtrado: Datos Generales / Mediciones del CSV principal
    con el siguiente formato:
      1) Fuente Calibri
      2) Tabla de 2 columnas para datos iniciales (A, B, C)
      3) Resumen en forma de tabla (6 columnas) para todas las mediciones
      4) Desarrollo detallado por cada área, también en tablas de 2 columnas
      5) Fotos embebidas
    """

    # Crea el documento Word
    doc = Document()

    # --- Ajustar la fuente a Calibri por defecto ---
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ============================
    #   ESTABLECER IDIOMA: es-CL
    # ============================
    # Obtenemos la referencia a las propiedades de fuente del estilo Normal.
    rPr = style.element.get_or_add_rPr()
    lang = get_or_add_lang(rPr)
    # Ajustamos el atributo w:val con 'es-CL' para indicar Español de Chile.
    lang.set(qn('w:val'), 'es-CL')
    lang.set(qn('w:eastAsia'), 'es-CL')
    lang.set(qn('w:bidi'), 'es-CL')

    # Portada / Encabezado principal
    doc.add_heading("EVALUACIÓN CONFORT TÉRMICO", 0)
    doc.add_paragraph(
        f"Fecha de generación del informe: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    )

    # ============================
    #   1) DATOS INICIALES
    # ============================
    doc.add_heading("Identificación de empresa y centro de trabajo", level=1)

    # -- A. Información empresa / B. Información centro --
    # Creamos UNA tabla de 2 columnas para agrupar la info A, B y la parte C.
    table_inicial = doc.add_table(rows=0, cols=2)
    table_inicial.style = 'Table Grid'  # o el estilo que prefieras

    def add_row(table, label, value=""):
        """
        Agrega una fila a la tabla.
        Si 'value' es vacío se asume que se trata de un título,
        se fusionan ambas celdas y se aplica fondo morado y letras blancas.
        """
        row_cells = table.add_row().cells

        # Si value es vacío, tratamos la fila como título.
        if value == "":
            # Fusionar ambas celdas.
            merged_cell = row_cells[0].merge(row_cells[1])
            merged_cell.text = label

            # Aplicar fondo morado (hex "800080") al contenido de la celda.
            shading_elm = parse_xml(r'<w:shd {} w:fill="800080"/>'.format(nsdecls('w')))
            merged_cell._tc.get_or_add_tcPr().append(shading_elm)

            # Cambiar el color de fuente a blanco, negrita y aumentar el tamaño
            for paragraph in merged_cell.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True
                    run.font.size = Pt(12)
        else:
            # Para filas normales, se coloca la etiqueta en la primera columna y el valor en la segunda.
            row_cells[0].text = label
            # Convertir a string en caso de que sea float o NaN
            value_str = "" if pd.isna(value) else str(value)
            row_cells[1].text = value_str

    # --- Sección A y B: Información empresa y centro de trabajo ---
    if not df_info_cuv.empty:
        row_cuv_info = df_info_cuv.iloc[0]

        # A. Información empresa (título)
        add_row(table_inicial, "A. Información empresa")
        add_row(table_inicial, "Razón Social", row_cuv_info.get('RAZÓN SOCIAL', ''))
        add_row(table_inicial, "RUT", row_cuv_info.get('RUT', ''))

        # B. Información centro de trabajo (título)
        add_row(table_inicial, "B. Información centro de trabajo")
        add_row(table_inicial, "CUV", row_cuv_info.get('CUV', ''))
        add_row(table_inicial, "Nombre de Local", row_cuv_info.get('Nombre de Local', ''))
        add_row(table_inicial, "Dirección", row_cuv_info.get('Dirección', ''))
        add_row(table_inicial, "Comuna", row_cuv_info.get('Comuna', ''))
        add_row(table_inicial, "Región", row_cuv_info.get('Región', ''))
    else:
        add_row(table_inicial, "Información de empresa/centro", "No se encontró información para este CUV.")

    # --- Separamos los registros del DF principal ---
    datos_generales = df_filtrado[df_filtrado["Que seccion quieres completar"] == "Datos generales"]
    mediciones = df_filtrado[df_filtrado["Que seccion quieres completar"] == "Medición de un area"]

    # -- C. Antecedentes generales (en la misma tabla de 2 columnas) --
    if not datos_generales.empty:
        # Tomamos la primera fila (o podrías iterar si hay varias)
        row_gen = datos_generales.iloc[0]

        # Título de la sección C. Se fusionan las celdas y se aplica el formato.
        add_row(table_inicial, "C. Antecedentes generales de la evaluación")

        # Nuevos / actualizados campos:
        add_row(table_inicial, "Fecha visita", row_gen.get('Fecha visita', ''))
        add_row(table_inicial, "Hora medición", row_gen.get('Hora medicion', ''))
        temp_max = row_gen.get('Temperatura máxima del día', '')
        add_row(table_inicial, "Temperatura máxima del día", f"{temp_max} °C" if temp_max else "")
        add_row(table_inicial, "Nombre del personal SMU", row_gen.get('Nombre del personal SMU', ''))
        add_row(table_inicial, "Cargo", row_gen.get('Cargo', ''))
        add_row(table_inicial, "Profesional IST (Correo)", row_gen.get('Dirección de correo electrónico', ''))

        # Agregamos columnas nuevas relacionadas con equipo y verificación
        add_row(table_inicial, "Código equipo temperatura", row_gen.get('Código equipo temperatura', ''))
        add_row(table_inicial, "Código equipo 2", row_gen.get('Codigo equipo 2', ''))
        add_row(table_inicial, "Verificación TBS inicial", row_gen.get('Verificación TBS inicial', ''))
        add_row(table_inicial, "Verificación TBH inicial", row_gen.get('Verificación TBH inicial', ''))
        add_row(table_inicial, "Verificación TG inicial", row_gen.get('Verificación TG inicial', ''))
        add_row(table_inicial, "Verificación TBS final", row_gen.get('Verificación TBS final', ''))
        add_row(table_inicial, "Verificación TBH final", row_gen.get('Verificación TBH final', ''))
        add_row(table_inicial, "Verificación TG final", row_gen.get('Verificación TG final', ''))

        # Más columnas nuevas
        add_row(table_inicial, "Patrón utilizado para calibrar", row_gen.get('Patrón utilizado para calibrar', ''))
        add_row(table_inicial, "Patrón TBS", row_gen.get('Patrón TBS', ''))
        add_row(table_inicial, "Patrón TBH", row_gen.get('Patrón TBH', ''))
        add_row(table_inicial, "Patrón TG", row_gen.get('Patrón TG', ''))
        add_row(table_inicial, "Tipo de vestimenta utilizada", row_gen.get('Tipo de vestimenta utilizada', ''))
        add_row(table_inicial, "Motivo de evaluación", row_gen.get('Motivo de evaluación', ''))
        add_row(table_inicial, "Comentarios finales de evaluación",
                row_gen.get('Comentarios finales de evaluación', ''))

        # Info adicional:
        add_row(table_inicial,
                "... A partir del registro de correo se buscan datos del profesional IST...",
                "")
    else:
        add_row(table_inicial, "C. Antecedentes generales de la evaluación",
                "No se han encontrado filas de 'Datos Generales' para este CUV.")

    doc.add_paragraph("")  # espacio extra

    # ============================
    #   2) RESUMEN DE MEDICIONES
    # ============================

    doc.add_heading("Resumen de Mediciones por Área.", level=1)
    doc.add_paragraph()
    doc.add_paragraph()

    doc.add_heading("Metodología de evaluación y parámetros utilizados.", level=1)
    doc.add_paragraph()
    doc.add_paragraph()

    doc.add_heading("Resultado de mediciones y evaluaciones por área", level=1)
    doc.add_paragraph()
    doc.add_paragraph()

    if not mediciones.empty:
        # Ajustar con los nombres exactos de tu CSV
        cols_resumen = [
            "Area o sector",
            "Especificación sector",
            "Temperatura bulbo seco",
            "Temperatura globo",
            "Humedad relativa",
            "Velocidad del aire"
        ]

        # Creamos la tabla con 1 fila extra para la cabecera
        tabla_resumen = doc.add_table(rows=1, cols=len(cols_resumen))
        tabla_resumen.style = 'Table Grid'
        hdr_cells = tabla_resumen.rows[0].cells

        # Cabeceras
        hdr_cells[0].text = "Área"
        hdr_cells[1].text = "Especificación sector"
        hdr_cells[2].text = "TBS (°C)"
        hdr_cells[3].text = "TG (°C)"
        hdr_cells[4].text = "HR (%)"
        hdr_cells[5].text = "Vel. aire (m/s)"

        # Rellenamos las filas
        for _, row_med in mediciones.iterrows():
            row_cells = tabla_resumen.add_row().cells
            row_cells[0].text = str(row_med.get("Area o sector", ""))
            row_cells[1].text = str(row_med.get("Especificación sector", ""))
            row_cells[2].text = str(row_med.get("Temperatura bulbo seco", ""))
            row_cells[3].text = str(row_med.get("Temperatura globo", ""))
            row_cells[4].text = str(row_med.get("Humedad relativa", ""))
            row_cells[5].text = str(row_med.get("Velocidad del aire", ""))

        doc.add_paragraph("")  # Espacio extra
    else:
        doc.add_paragraph("No se han encontrado filas de 'Medición de un area' para este CUV.")

    # ============================
    #   3) DESARROLLO POR ÁREA
    # ============================
    doc.add_page_break()
    doc.add_heading("Anexo de condiciones observadas en áreas de medición", level=1)
    if not mediciones.empty:
        for idx, row_med in mediciones.iterrows():
            doc.add_heading(f"Área: {row_med.get('Area o sector', '')} - {row_med.get('Especificación sector', '')}",
                            level=2)

            # Tabla de 2 columnas con la info adicional
            tabla_area = doc.add_table(rows=0, cols=2)
            tabla_area.style = 'Table Grid'

            def add_area_row(label, value):
                r = tabla_area.add_row().cells
                # Convertir todo a string, evitando error con floats o NaN
                value_str = "" if pd.isna(value) else str(value)
                r[0].text = label
                r[1].text = value_str

            # Info de techumbre, paredes, ventanales, etc.
            add_area_row("Puesto de trabajo", row_med.get("Puesto de trabajo", ""))
            add_area_row("Trabajador de pie o sentado", row_med.get("Trabajador de pie o sentado", ""))
            add_area_row("Techumbre", row_med.get("Techumbre", ""))
            add_area_row("Observación techumbre", row_med.get("Observacion techumbre - Indique tipo de material", ""))
            add_area_row("Paredes", row_med.get("Paredes", ""))
            add_area_row("Observación paredes", row_med.get("Observacion paredes", ""))
            add_area_row("Ventanales", row_med.get("Ventanales", ""))
            add_area_row("Observación ventanales", row_med.get("Observacion ventanales", ""))
            add_area_row("Aire acondicionado", row_med.get("Aire acondicionado", ""))
            add_area_row("Observaciones aire acondicionado", row_med.get("Observaciones aire acondicionado", ""))
            add_area_row("Ventiladores", row_med.get("Ventiladores", ""))
            add_area_row("Observaciones ventiladores", row_med.get("Observaciones ventiladores", ""))
            add_area_row("Inyección y/o extracción de aire", row_med.get("Inyección y/o extracción de aire", ""))
            add_area_row("Observaciones inyección y/o extracción de aire",
                         row_med.get("Observaciones inyeccion y/o extracción de aire", ""))
            add_area_row("Ventanas (Ventilación natural)", row_med.get("Ventanas (Ventilación natural)", ""))
            add_area_row("Observaciones ventanas (ventilación natural)",
                         row_med.get("Observaciones ventanas (ventilación natural)", ""))
            add_area_row("Puertas (ventilación natural)", row_med.get("Puertas (ventilación natural)", ""))
            add_area_row("Observaciones puertas (ventilación natural)",
                         row_med.get("Observaciones puertas (ventilación natural)", ""))
            add_area_row("Otras condiciones de disconfort térmico",
                         row_med.get("Otras condiciones de disconfort termico", ""))
            add_area_row("Observaciones sobre otras condiciones de disconfort térmico",
                         row_med.get("Observaciones sobre otras condiciones de disconfort térmico", ""))

            # Fotografías
            evidencia = row_med.get("Evidencia fotografica", "")
            if pd.notnull(evidencia) and isinstance(evidencia, str) and evidencia.strip():
                doc.add_paragraph("Fotografías asociadas:")
                for foto_url in [f.strip() for f in evidencia.split(",") if f.strip()]:
                    imagen = descargar_imagen_gdrive(foto_url)
                    if imagen is not None:
                        doc.add_picture(imagen, width=Inches(4))
                    else:
                        doc.add_paragraph(f"No se pudo descargar la imagen: {foto_url}")

            doc.add_paragraph("")  # Espacio extra
    else:
        doc.add_paragraph("No se encontró información de mediciones para detallar.")

    # --- Pie de informe ---
    doc.add_paragraph("Informe versión prueba")

    # Convertimos a BytesIO para descargar
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer