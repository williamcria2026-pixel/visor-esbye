import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re
import io
from PIL import Image as PILImage

# ReportLab para la generación del Acta PDF en formato horizontal (Landscape)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# Configuración de la página
st.set_page_config(page_title="Visor ESBYE", page_icon="📦", layout="centered")

# Estilos CSS avanzados para mantener el buscador y la alineación responsiva
st.markdown("""
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 1.0rem;
            padding-left: 0.8rem;
            padding-right: 0.8rem;
        }
        @media (max-width: 768px) {
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                align-items: flex-end !important;
                gap: 0.4rem !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
                width: auto !important;
                flex: 1 1 auto !important;
                min-width: 0px !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child {
                flex: 0 0 55px !important;
            }
        }
        div[data-testid="stTextInput"] label {
            display: none !important;
        }
        h2 {
            margin-top: 0.5rem !important;
            margin-bottom: -0.3rem !important;
        }
        p {
            margin-bottom: 0.3rem !important;
        }
        hr {
            margin: 0.6rem 0rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# Función para cargar el Excel completo asegurando lectura limpia sin alterar índices
def cargar_excel(origen):
    for engine in ['calamine', 'openpyxl', None]:
        try:
            if engine:
                df_temp = pd.read_excel(origen, engine=engine)
            else:
                df_temp = pd.read_excel(origen)
            if df_temp is not None and not df_temp.empty:
                return df_temp
        except Exception:
            continue
    return None

# Detección robusta con ruta absoluta para garantizar que el PDF siempre encuentre el escudo
def obtener_ruta_escudo():
    ruta_candidata = os.path.abspath("escudo.png")
    if os.path.exists(ruta_candidata):
        return ruta_candidata
    for f in os.listdir('.'):
        if f.lower().startswith('escudo') and f.lower().endswith(('png', 'jpg', 'jpeg')):
            return os.path.abspath(f)
    return None

# Barra lateral con opciones y gestión opcional del escudo
with st.sidebar:
    st.header("⚙️ Opciones")
    archivo_subido = st.file_uploader("Subir nuevo reporte (.xlsx)", type=["xlsx", "xls"])
    
    if archivo_subido is not None:
        try:
            with open("reporte_activo.xlsx", "wb") as f:
                f.write(archivo_subido.getbuffer())
            with open("nombre_reporte.txt", "w", encoding="utf-8") as f:
                f.write(archivo_subido.name)
            st.success(f"Reporte '{archivo_subido.name}' cargado con éxito.")
        except Exception as e:
            st.error(f"Error al guardar archivo: {e}")

    st.divider()
    with st.expander("🛡️ Configuración del Escudo"):
        escudo_subido = st.file_uploader("Actualizar archivo escudo.png", type=["png", "jpg", "jpeg"], key="uploader_escudo")
        if escudo_subido is not None:
            try:
                with open("escudo.png", "wb") as f:
                    f.write(escudo_subido.getbuffer())
                st.success("¡Escudo actualizado permanentemente!")
            except Exception as e:
                st.error(f"Error al guardar el escudo: {e}")

# Encabezado corporativo
st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>BIENES Y BODEGAS</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #1D4ED8; font-style: italic; font-weight: bold; font-size: 0.85rem;'>Desarrollado por: Ing. William Cisneros R.</p>", unsafe_allow_html=True)
st.divider()

# Recuperación segura del archivo con indicador de carga profesional (Spinner)
df = None
display_reporte = "Ninguno"

if os.path.exists("reporte_activo.xlsx"):
    with st.spinner("📂 Cargando y procesando reporte institucional..."):
        df = cargar_excel("reporte_activo.xlsx")
    if df is not None:
        if os.path.exists("nombre_reporte.txt"):
            try:
                with open("nombre_reporte.txt", "r", encoding="utf-8") as f:
                    display_reporte = f.read().strip()
            except:
                display_reporte = "reporte_activo.xlsx"
        else:
            display_reporte = "reporte_activo.xlsx"
    else:
        try:
            os.remove("reporte_activo.xlsx")
            if os.path.exists("nombre_reporte.txt"):
                os.remove("nombre_reporte.txt")
        except:
            pass
else:
    if os.path.exists("informe.xlsx"):
        with st.spinner("📂 Cargando reporte predeterminado..."):
            df = cargar_excel("informe.xlsx")
        display_reporte = "informe.xlsx"
    else:
        archivos_excel = [f for f in os.listdir('.') if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('escudo') and f != 'nombre_reporte.txt']
        if archivos_excel:
            with st.spinner("📂 Cargando reporte detectado..."):
                df = cargar_excel(archivos_excel[0])
            display_reporte = archivos_excel[0]

fecha_acta_actual = datetime.now().strftime("%d/%m/%Y")

# Funciones de limpieza de datos
def limpiar_val_fecha(val):
    if pd.isna(val) or str(val).strip().lower() in ["nan", "none", "", "nat", "null"]:
        return "-"
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%d/%m/%Y")
    texto = str(val).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}', texto):
        try:
            dt = datetime.strptime(texto.split()[0], "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except:
            pass
    if re.match(r'^\d{2}/\d{2}/\d{4}', texto):
        return texto.split()[0]
    return texto

def limpiar_val(val):
    if pd.isna(val) or str(val).strip().lower() in ["nan", "none", "", "nat", "null"]:
        return "-"
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%d/%m/%Y")
    texto = str(val).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}', texto):
        try:
            dt = datetime.strptime(texto.split()[0], "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except:
            pass
    return texto.replace("DISENIO", "DISEÑO").replace("disenio", "diseño").replace("Disenio", "Diseño")

# Función para generar el PDF oficial con espacios entre el signo $ y el valor
def generar_acta_pdf(row, fecha_acta_str, col_map, val_f_ingreso_pdf):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=0.4*inch,
        leftMargin=0.4*inch,
        topMargin=0.4*inch,
        bottomMargin=0.4*inch
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=11, leading=13, alignment=1, fontName="Helvetica-Bold")
    subtitle_style_center = ParagraphStyle('SubTitleStyleCenter', parent=styles['Normal'], fontSize=9, leading=12, alignment=1)
    acta_title = ParagraphStyle('ActaTitle', parent=styles['Heading2'], fontSize=10, leading=12, alignment=1, fontName="Helvetica-Bold")
    
    normal_left = ParagraphStyle('NormalLeft', parent=styles['Normal'], fontSize=7, leading=8, alignment=0)
    normal_center = ParagraphStyle('NormalCenter', parent=styles['Normal'], fontSize=7, leading=8, alignment=1)
    totales_right = ParagraphStyle('TotalesRight', parent=styles['Normal'], fontSize=7, leading=8, fontName="Helvetica-Bold", alignment=2)
    header_table = ParagraphStyle('HeaderTable', parent=styles['Normal'], fontSize=7, leading=8, fontName="Helvetica-Bold", alignment=1, textColor=colors.white)

    escudo_path = obtener_ruta_escudo()
    img_flowable = Paragraph("", normal_center)
    if escudo_path and os.path.exists(escudo_path):
        try:
            img_escudo = RLImage(escudo_path, width=0.65*inch, height=0.65*inch)
            img_escudo.hAlign = 'LEFT'
            img_flowable = img_escudo
        except Exception as e:
            print(f"Error procesando escudo: {e}")

    header_text_elements = [
        Paragraph("MINISTERIO DE ECONOMÍA Y FINANZAS", title_style),
        Paragraph("176-9999-0000-UNIVERSIDAD TÉCNICA DE AMBATO - PLANTA CENTRAL", subtitle_style_center),
        Spacer(1, 3),
        Paragraph("ACTA DE ASIGNACIÓN - REASIGNACIÓN DE BIENES", acta_title)
    ]

    header_table_data = [[img_flowable, header_text_elements]]
    header_table_obj = Table(header_table_data, colWidths=[1.0*inch, 9.2*inch])
    header_table_obj.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table_obj)
    story.append(Spacer(1, 4))

    story.append(Paragraph(f"<b>FECHA:</b> {fecha_acta_str}", normal_left))
    story.append(Spacer(1, 4))

    col_widths = [55, 55, 120, 65, 55, 75, 90, 40, 50, 50, 30, 55]
    headers = [
        Paragraph("Código Bien", header_table),
        Paragraph("Código Anterior", header_table),
        Paragraph("Nombre Bien", header_table),
        Paragraph("Serie", header_table),
        Paragraph("Modelo", header_table),
        Paragraph("Marca/Raza/Otros", header_table),
        Paragraph("Bodega", header_table),
        Paragraph("Tipo Reg.", header_table),
        Paragraph("Fecha Ingreso", header_table),
        Paragraph("Fecha Asignación", header_table),
        Paragraph("Cant.", header_table),
        Paragraph("Valor Contable", header_table)
    ]

    val_contable = row.get(col_map.get('valor'), '-') if col_map.get('valor') else '-'
    val_contable_clean = limpiar_val(val_contable)
    if val_contable_clean not in ["-", "."]:
        try:
            val_contable_num = float(val_contable_clean.replace('$', '').replace(',', ''))
            # SE APLICÓ EL ESPACIADO CON DOS ESPACIOS PROTEGIDOS (&nbsp;&nbsp;)
            val_contable_str = f"$&nbsp;&nbsp;{val_contable_num:,.2f}"
        except:
            val_contable_str = val_contable_clean
    else:
        val_contable_str = "-"

    data_row = [
        Paragraph(limpiar_val(row.get(col_map.get('codigo'), '-')), normal_center),
        Paragraph(limpiar_val(row.get(col_map.get('anterior'), '-')), normal_center),
        Paragraph(limpiar_val(row.get(col_map.get('nombre'), '-')), normal_center),
        Paragraph(limpiar_val(row.get(col_map.get('serie'), '-')), normal_center),
        Paragraph(limpiar_val(row.get(col_map.get('modelo'), '-')), normal_center),
        Paragraph(limpiar_val(row.get(col_map.get('marca'), '-')), normal_center),
        Paragraph(limpiar_val(row.get(col_map.get('bodega'), '-')), normal_center),
        Paragraph("BLD", normal_center),
        Paragraph(val_f_ingreso_pdf, normal_center),
        Paragraph(fecha_acta_str, normal_center),
        Paragraph("1", normal_center),
        Paragraph(val_contable_str, normal_center)
    ]

    total_row = [
        Paragraph("<b>TOTAL:</b>", totales_right),
        "", "", "", "", "", "", "", "", "",
        Paragraph("<b>1</b>", normal_center),
        Paragraph(f"<b>{val_contable_str}</b>", normal_center)
    ]

    table_data = [headers, data_row, total_row]
    t = Table(table_data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('SPAN', (0, 2), (9, 2)),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#F3F4F6')),
    ]))
    story.append(t)
    story.append(Spacer(1, 25))

    custodio_dinamico = limpiar_val(row.get(col_map.get('custodio'), '-')) if col_map.get('custodio') else '-'

    cell_entrega = [
        Paragraph("________________________________________", subtitle_style_center),
        Spacer(1, 4),
        Paragraph(f"{custodio_dinamico}", subtitle_style_center),
        Paragraph("<b>QUIEN ENTREGA</b>", subtitle_style_center)
    ]

    cell_recibe = [
        Paragraph("________________________________________", subtitle_style_center),
        Spacer(1, 4),
        Paragraph("Nombre: ...................................................", subtitle_style_center),
        Paragraph("C.I.: ...........................................................", subtitle_style_center),
        Spacer(1, 4),
        Paragraph("<b>QUIEN RECIBE</b>", subtitle_style_center)
    ]

    firmas_data = [[cell_entrega, cell_recibe]]
    firmas_table = Table(firmas_data, colWidths=[370, 370])
    firmas_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(firmas_table)

    doc.build(story)
    buffer.seek(0)
    return buffer

if df is not None:
    df.columns = df.columns.astype(str).str.strip()
    cols = list(df.columns)
    
    col_map = {}
    col_fecha_ingreso_real = None

    if len(cols) > 0: col_map['codigo'] = cols[0]
    if len(cols) > 1: col_map['anterior'] = cols[1]
    if len(cols) > 5: col_map['nombre'] = cols[5]
    if len(cols) > 6: col_map['serie'] = cols[6]
    if len(cols) > 7: col_map['modelo'] = cols[7]
    if len(cols) > 8: col_map['marca'] = cols[8]
    if len(cols) > 20: col_map['bodega'] = cols[20]
    if len(cols) > 23: col_map['cedula'] = cols[23]
    if len(cols) > 24: col_map['custodio'] = cols[24]
    if len(cols) > 37: col_fecha_ingreso_real = cols[37]
    if len(cols) > 41: col_map['valor'] = cols[41]

    c_codigo = col_map.get('codigo', cols[0])
    c_serie = col_map.get('serie', cols[6] if len(cols) > 6 else cols[0])
    c_anterior = col_map.get('anterior', c_codigo)

    st.markdown(
        f"<div style='background-color: #1E293B; padding: 8px 12px; border-radius: 6px; margin-bottom: 8px; text-align: center;'>"
        f"<span style='color: #CBD5E1; font-size: 0.88rem;'>📄 Reporte Activo: <b style='color: #93C5FD;'>{display_reporte}</b></span>"
        f"</div>", 
        unsafe_allow_html=True
    )
    st.markdown("<p style='color: #FFFFFF; font-weight: bold; font-size: 0.95rem; margin-bottom: 2px;'>🔍 Buscar Código / Serie:</p>", unsafe_allow_html=True)

    if "busqueda_val" not in st.session_state:
        st.session_state.busqueda_val = ""

    col_input, col_btn = st.columns([4.2, 0.8])
    
    def limpiar_callback():
        st.session_state.busqueda_val = ""
        st.session_state.mi_buscador = ""

    with col_input:
        codigo = st.text_input(
            "",
            value=st.session_state.busqueda_val,
            key="mi_buscador",
            placeholder="Ingrese código o serie..."
        ).strip()
        
        if codigo:
            st.session_state.busqueda_val = codigo

    with col_btn:
        st.button("🗑️", on_click=limpiar_callback, help="Limpiar búsqueda", use_container_width=True)

    termino_busqueda = st.session_state.busqueda_val.strip()

    if termino_busqueda:
        s_codigo_str = df[c_codigo].astype(str).str.strip() if c_codigo in df.columns else pd.Series([""] * len(df))
        s_serie_str = df[c_serie].astype(str).str.strip() if c_serie in df.columns else pd.Series([""] * len(df))
        s_anterior_str = df[c_anterior].astype(str).str.strip() if c_anterior in df.columns else pd.Series([""] * len(df))

        res = df[
            (s_codigo_str.str.lower() == termino_busqueda.lower()) |
            (s_serie_str.str.lower() == termino_busqueda.lower()) |
            (s_anterior_str.str.lower() == termino_busqueda.lower())
        ]
        
        if res.empty:
            res = df[
                (s_codigo_str.str.contains(termino_busqueda, case=False, na=False)) |
                (s_serie_str.str.contains(termino_busqueda, case=False, na=False))
            ]
            
        if not res.empty:
            row = res.iloc[0]
            
            val_bien = limpiar_val(row.get(col_map.get('nombre'), '-')) if col_map.get('nombre') else '-'
            val_bodega = limpiar_val(row.get(col_map.get('bodega'), '-')) if col_map.get('bodega') else '-'
            
            val_fecha_ingreso = "-"
            if col_fecha_ingreso_real and col_fecha_ingreso_real in row:
                val_fecha_ingreso = limpiar_val_fecha(row.get(col_fecha_ingreso_real))
            elif len(row) > 37:
                val_fecha_ingreso = limpiar_val_fecha(row.iloc[37])
            
            with st.container(border=True):
                st.success(f"📌 **BIEN:** {val_bien}")
                st.info(f"🏢 **BODEGA:** {val_bodega}")
                
                val_contable = row.get(col_map.get('valor'), '-') if col_map.get('valor') else '-'
                val_contable_clean = limpiar_val(val_contable)
                if val_contable_clean not in ["-", "."]:
                    try:
                        val_contable_fmt = f"${float(val_contable_clean.replace('$', '').replace(',', '')):,.2f}"
                    except:
                        val_contable_fmt = val_contable_clean
                else:
                    val_contable_fmt = "-"

                tab1, tab2 = st.tabs(["🏷️ Datos del Bien", "👤 Datos del Custodio"])
                
                with tab1:
                    st.write(f"**Código del Bien:** {limpiar_val(row.get(c_codigo, '-'))}")
                    st.write(f"**Código Anterior:** {limpiar_val(row.get(c_anterior, '-'))}")
                    st.write(f"**Marca:** {limpiar_val(row.get(col_map.get('marca'), '-'))}")
                    st.write(f"**Modelo:** {limpiar_val(row.get(col_map.get('modelo'), '-'))}")
                    st.write(f"**Serie:** {limpiar_val(row.get(c_serie, '-'))}")
                    st.write(f"**Fecha de Ingreso del Bien:** {val_fecha_ingreso}")
                    st.write(f"**Valor Contable:** {val_contable_fmt}")
                    
                with tab2:
                    st.write(f"**Custodio Actual:** {limpiar_val(row.get(col_map.get('custodio'), '-'))}")
                    st.write(f"**Número de Cédula / RUC:** {limpiar_val(row.get(col_map.get('cedula'), '-'))}")

                st.divider()
                
                with st.spinner("📄 Generando documento PDF oficial..."):
                    pdf_bytes = generar_acta_pdf(row, fecha_acta_actual, col_map, val_fecha_ingreso)
                
                cod_bien_clean = str(row.get(c_codigo, 'Bien')).replace('/', '_')
                fecha_nom = datetime.now().strftime("%d-%m-%Y")
                nombre_archivo_pdf = f"Acta_{cod_bien_clean}_{fecha_nom}.pdf"

                st.download_button(
                    label="📄 Generar e Imprimir Acta PDF Oficial",
                    data=pdf_bytes,
                    file_name=nombre_archivo_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.error(f"❌ No se encontró ningún bien con el código: {termino_busqueda}")
else:
    st.warning("⚠️ No se encontró ningún archivo Excel válido o no se pudo cargar. Por favor, suba su reporte en formato .xlsx mediante el menú de la izquierda.")
