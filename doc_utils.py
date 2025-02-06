import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from datetime import datetime
import requests
import re
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION


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


def look_informe(doc):
    """
    Configura el documento en orientación horizontal (apaisado) y establece
    los márgenes y estilos para todo el documento.
    """
    for section in doc.sections:
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(2)
        section.right_margin = Cm(1.5)

    # Estilo Normal
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(9)  # Tamaño de fuente opcional; ajusta según prefieras

    # Estilo destacado
    destacado = doc.styles.add_style('destacado', 1)  # 1 para párrafos
    destacado_font = destacado.font
    destacado_font.name = 'Calibri'
    destacado_font.size = Pt(12)  # Tamaño de la fuente en puntos

    # Idioma Español de Chile
    rpr = style.element.get_or_add_rPr()
    lang = get_or_add_lang(rpr)
    lang.set(qn('w:val'), 'es-CL')
    lang.set(qn('w:eastAsia'), 'es-CL')
    lang.set(qn('w:bidi'), 'es-CL')


def set_vertical_alignment(doc, section_index=0, alignment='top'):
    """
    Ajusta la alineación vertical de la sección especificada.

    Parámetros:
    -----------
    doc : Document
        Objeto Document de python-docx.
    section_index : int
        Índice de la sección que se desea modificar. Por defecto 0 (la primera sección).
    alignment : str
        Valor de alineación vertical.
        Puede ser 'top' (arriba), 'center' (centrado) o 'both'.
    """
    # Obtenemos la sección
    section = doc.sections[section_index]
    # Obtenemos o creamos la propiedad sectPr
    sectPr = section._sectPr

    # Buscamos el elemento vAlign
    vAlign = sectPr.find(qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        sectPr.append(vAlign)

    # Ajustamos el atributo w:val según el parámetro alignment
    alignment = alignment.lower()
    if alignment not in ['top', 'center', 'both']:
        alignment = 'top'  # Valor por defecto si no es válido

    vAlign.set(qn('w:val'), alignment)


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

    '''
    Código comentado (por ejemplo, para pruebas de otros formatos):
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    # Agregar imagen del logo
    doc.add_picture('IST.jpg', width=Inches(2))
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Título principal y subtítulos
    titulo = doc.add_heading('INFORME TÉCNICO', level=1)
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    subtitulo = doc.add_heading('PRESCRIPCIÓN DE MEDIDAS PARA PROTOCOLO DE VIGILANCIA', level=2)
    subtitulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    subtitulo = doc.add_heading('DE RIESGOS PSICOSOCIALES EN EL TRABAJO', level=2)
    subtitulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph()
    doc.add_paragraph()

    # Información general (ejemplo)
    p = doc.add_paragraph()
    p.add_run('Razón Social: ').bold = True
    p.add_run(f"{safe_get(datos, 'Nombre_Empresa')}\n")
    # ... más información ...
    p.paragraph_format.left_indent = Cm(1.5)
    '''

    # ============================
    # Crea el documento Word
    # ============================
    doc = Document()
    look_informe(doc)

    # --- Sección 0 (portada) ya existe por defecto ---
    doc.add_picture('IST.jpg', width=Inches(2))
    doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    titulo = doc.add_heading('INFORME TÉCNICO', level=2)
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    subtitulo = doc.add_heading('EVALUACIÓN DE CONFORT TÉRMICO', level=2)
    subtitulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    if not df_info_cuv.empty:
        portada_info = df_info_cuv.iloc[0]
        n_local = doc.add_heading(portada_info.get('Nombre de Local'), level=2)
        n_local.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        n_local = doc.add_heading(portada_info.get('RAZÓN SOCIAL'), level=2)
        n_local.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    set_vertical_alignment(doc, section_index=0, alignment='center')

    # --- Crear NUEVA SECCIÓN para el resto del contenido ---
    doc.add_section(WD_SECTION.NEW_PAGE)
    set_vertical_alignment(doc, section_index=1, alignment='top')

    # Encabezado principal del contenido
    doc.add_heading("EVALUACIÓN CONFORT TÉRMICO", level=1)
    doc.add_paragraph()
        # (Se puede agregar fecha o más información aquí)

    # ============================
    #   1) DATOS INICIALES
    # ============================
    # Creamos una tabla de 2 columnas para agrupar la información (secciones A, B y C)
    table_inicial = doc.add_table(rows=0, cols=2)
    table_inicial.style = 'Table Grid'

    def add_row(table, label, value=""):
        """
        Agrega una fila a la tabla.
        Si 'value' es vacío se asume que se trata de un título; se fusionan ambas celdas
        y se aplica fondo morado con letras blancas.
        """
        row_cells = table.add_row().cells

        if value == "":
            # Fusionar ambas celdas para títulos
            merged_cell = row_cells[0].merge(row_cells[1])
            merged_cell.text = label

            # Aplicar fondo morado (hex "800080")
            shading_elm = parse_xml(r'<w:shd {} w:fill="800080"/>'.format(nsdecls('w')))
            merged_cell._tc.get_or_add_tcPr().append(shading_elm)

            # Cambiar el color de fuente a blanco, negrita y aumentar el tamaño
            for paragraph in merged_cell.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True
                    run.font.size = Pt(12)
        else:
            row_cells[0].text = label
            value_str = "" if pd.isna(value) else str(value)
            row_cells[1].text = value_str

    # --- Sección A y B: Información empresa y centro de trabajo ---
    if not df_info_cuv.empty:
        row_cuv_info = df_info_cuv.iloc[0]
        # A. Información empresa
        add_row(table_inicial, "A. Información empresa")
        add_row(table_inicial, "Razón Social", row_cuv_info.get('RAZÓN SOCIAL', ''))
        add_row(table_inicial, "RUT", row_cuv_info.get('RUT', ''))

        # B. Información centro de trabajo
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

    # -- Sección C: Antecedentes generales (en la misma tabla de 2 columnas) --
    if not datos_generales.empty:
        row_gen = datos_generales.iloc[0]
        add_row(table_inicial, "C. Antecedentes generales de la evaluación")
        add_row(table_inicial, "Fecha visita", row_gen.get('Fecha visita', ''))
        add_row(table_inicial, "Hora medición", row_gen.get('Hora medicion', ''))
        temp_max = row_gen.get('Temperatura máxima del día', '')
        add_row(table_inicial, "Temperatura máxima del día", f"{temp_max} °C" if temp_max else "")
        add_row(table_inicial, "Nombre del personal empresa", row_gen.get('Nombre del personal SMU', ''))
        add_row(table_inicial, "Cargo del personal empresa", row_gen.get('Cargo', ''))
        add_row(table_inicial, "Destinatario empresa del informe", "--Pendiente generar en formulario--")
        add_row(table_inicial, "Nombre consultor IST", "-- Pendiente vincular con usuario login--")
        add_row(table_inicial, "Nombre revisor higiene IST", "-- Pendiente definir cuando se completa --")

    # ============================
    #   2) RESUMEN DE MEDICIONES
    # ============================
    if not mediciones.empty:
        cols_resumen = [
            "Area o sector",
            "Especificación sector",
            "Temperatura bulbo seco",
            "Temperatura globo",
            "Humedad relativa",
            "Velocidad del aire"
        ]

        doc.add_heading("Resumen de Mediciones por Área.", level=1)
        doc.add_paragraph()

        tabla_resumen = doc.add_table(rows=1, cols=len(cols_resumen))
        tabla_resumen.style = 'Table Grid'
        hdr_cells = tabla_resumen.rows[0].cells

        hdr_cells[0].text = "Área"
        hdr_cells[1].text = "Especificación sector"
        hdr_cells[2].text = "TBS (°C)"
        hdr_cells[3].text = "TG (°C)"
        hdr_cells[4].text = "HR (%)"
        hdr_cells[5].text = "Vel. aire (m/s)"

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

    doc.add_heading("Metodología de evaluación y parámetros utilizados.", level=1)
    doc.add_paragraph()

    table_calib = doc.add_table(rows=0, cols=2)
    table_calib.style = 'Table Grid'

    # Título para la sección de detalles de equipos y calibración
    add_row(table_calib, "D. Detalles de equipos y calibración")
    add_row(table_calib, "Equipo temperatura", row_gen.get('Código equipo temperatura', ''))
    add_row(table_calib, "Equipo velocidad viento", row_gen.get('Codigo equipo 2', ''))

    # Se añaden más detalles relacionados
    add_row(table_calib, "Tipo de vestimenta utilizada", row_gen.get('Tipo de vestimenta utilizada', ''))
    add_row(table_calib, "Motivo de evaluación", row_gen.get('Motivo de evaluación', ''))
    add_row(table_calib, "Comentarios finales de evaluación", row_gen.get('Comentarios finales de evaluación', ''))
    doc.add_page_break()

    # ============================
    #   3) DESARROLLO POR ÁREA
    # ============================

    if not mediciones.empty:
        for idx, row_med in mediciones.iterrows():
            heading_anexoareas = doc.add_heading("Anexo de condiciones observadas en áreas de medición", level=2)
            heading_anexoareas.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            doc.add_heading(
                f"Área: {row_med.get('Area o sector', '')} - {row_med.get('Especificación sector', '')}",
                level=2
            )

            tabla_area = doc.add_table(rows=0, cols=2)
            tabla_area.style = 'Table Grid'

            def add_area_row(label, value):
                r = tabla_area.add_row().cells
                value_str = "" if pd.isna(value) else str(value)
                r[0].text = label
                r[1].text = value_str

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

            # Fotografías asociadas
            doc.add_paragraph()
            doc.add_paragraph()

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
            doc.add_page_break()
    else:
        doc.add_paragraph("No se encontró información de mediciones para detallar.")

    # ============================
    #   4) APARTADO DETALLES TÉCNICOS DE LA EVALUACIÓN
    # ============================

    heading_anexotec = doc.add_heading("Anexo de detalles técnicos de evaluación", level=2)
    heading_anexotec.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    if not datos_generales.empty:
        row_gen = datos_generales.iloc[0]
        # Se crea una tabla de 4 columnas para el nuevo formato:
        table_tec = doc.add_table(rows=0, cols=4)
        table_tec.style = 'Table Grid'
        # Fila de cabecera
        hdr_cells = table_tec.add_row().cells
        hdr_cells[0].text = "Verificación"
        hdr_cells[1].text = "Patrón equipo"
        hdr_cells[2].text = "Inicio medición"
        hdr_cells[3].text = "Final medición"
        # Fila para Valor TBS
        row = table_tec.add_row().cells
        row[0].text = "Valor TBS"
        row[1].text = str(row_gen.get('Patrón TBS', ''))
        row[2].text = str(row_gen.get('Verificación TBS inicial', ''))
        row[3].text = str(row_gen.get('Verificación TBS final', ''))
        # Fila para Valor TBH
        row = table_tec.add_row().cells
        row[0].text = "Valor TBH"
        row[1].text = str(row_gen.get('Patrón TBH', ''))
        row[2].text = str(row_gen.get('Verificación TBH inicial', ''))
        row[3].text = str(row_gen.get('Verificación TBH final', ''))
        # Fila para Valor TG
        row = table_tec.add_row().cells
        row[0].text = "Valor TG"
        row[1].text = str(row_gen.get('Patrón TG', ''))
        row[2].text = str(row_gen.get('Verificación TG inicial', ''))
        row[3].text = str(row_gen.get('Verificación TG final', ''))
    else:
        # En caso de no encontrar datos generales, se crea una tabla informativa
        table_tec = doc.add_table(rows=0, cols=4)
        table_tec.style = 'Table Grid'
        hdr_cells = table_tec.add_row().cells
        hdr_cells[0].text = "Verificación"
        hdr_cells[1].text = "Patrón equipo"
        hdr_cells[2].text = "Inicio medición"
        hdr_cells[3].text = "Final medición"
        row = table_tec.add_row().cells
        row[0].text = ""
        row[1].text = ""
        row[2].text = "No se han encontrado información para este CUV."
        row[3].text = ""

    # Convertir el documento a BytesIO para descargar
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
