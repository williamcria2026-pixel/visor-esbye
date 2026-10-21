import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
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
            padding-bottom: 1rem;
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

# Barra lateral para opciones y archivos
with st.sidebar:
    st.header("⚙️ Opciones")
    archivo_subido = st.file_uploader("Subir nuevo reporte (.xlsx)", type=["xlsx", "xls"])
    
    st.divider()
    st.subheader("🛡️ Escudo Institucional")
    escudo_subido = st.file_uploader("Subir imagen del escudo", type=["png", "jpg", "jpeg"], key="uploader_escudo")
    
    if escudo_subido is not None:
        try:
            nombre_escudo_guardado = f"escudo_{escudo_subido.name}"
            for f in os.listdir('.'):
                if f.lower().startswith('escudo'):
                    try: os.remove(f)
                    except: pass
            with open(nombre_escudo_guardado, "wb") as f:
                f.write(escudo_subido.getbuffer())
        except Exception:
            pass

# Encabezado corporativo
st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>BIENES Y BODEGAS</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #1D4ED8; font-style: italic; font-weight: bold; font-size: 0.85rem;'>Desarrollado por: Ing. William Cisneros R.</p>", unsafe_allow_html=True)
st.divider()

# Función para limpiar valores nulos o 'nan'
def limpiar_val(val):
    if pd.isna(val) or str(val).strip().lower() in ["nan", "none", "", "nat", "null"]:
        return "-"
    texto = str(val).strip()
    return texto.replace("DISENIO", "DISEÑO").replace("disenio", "diseño").replace("Disenio", "Diseño")

# Función para cargar Excel
def cargar_excel(origen):
    for engine in ['calamine', 'openpyxl', None]:
        try:
            if engine:
                return pd.read_excel(origen, engine=engine)
            else:
                return pd.read_excel(origen)
        except Exception:
            continue
    st.error("No se pudo procesar el archivo Excel con ningún motor disponible.")
    return None

# Función inteligente para detectar el nombre del archivo de escudo activo
def obtener_ruta_escudo():
    for f in os.listdir('.'):
        if f.lower().startswith('escudo'):
            return f
    return None

# Función para generar el PDF del Acta con la Fecha de Ingreso del Bien antes del Valor Contable
def generar_acta_pdf(row, fecha_acta_str, col_map):
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
    normal_small = ParagraphStyle('NormalSmall', parent=styles['Normal'], fontSize=7, leading=8)
    header_table = ParagraphStyle('HeaderTable', parent=styles['Normal'], fontSize=7, leading=8, fontName="Helvetica-Bold", alignment=1, textColor=colors.white)

    escudo_path = obtener_ruta_escudo()
    img_flowable = Paragraph("", normal_small)
    if escudo_path and os.path.exists(escudo_path):
        try:
            with PILImage.open(escudo_path) as pil_img:
                bg = PILImage.new("RGB", pil_img.size, (255, 255, 255))
                if pil_img.mode == "RGBA":
                    bg.paste(pil_img, mask=pil_img.split()[3])
                else:
                    bg.paste(pil_img)
                
                img_buffer = io.BytesIO()
                bg.save(img_buffer, format='PNG')
                img_buffer.seek(0)
            
            img_escudo = RLImage(img_buffer, width=0.65*inch, height=0.65*inch)
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

    story.append(Paragraph(f"<b>FECHA:</b> {fecha_acta_str}", normal_small))
    story.append(Spacer(1, 4))

    col_widths = [55, 55, 115, 60, 50, 70, 85, 40, 45, 45, 30, 55, 55]
    headers = [
        Paragraph("Código Bien", header_table),
        Paragraph("Código Anterior", header_table),
        Paragraph("Nombre Bien", header_table),
        Paragraph("Serie", header_table),
        Paragraph("Modelo", header_table),
        Paragraph("Marca/Otros", header_table),
        Paragraph("Bodega", header_table),
        Paragraph("Tipo Reg.", header_table),
        Paragraph("F. Ingreso", header_table),
        Paragraph("F. Asignación", header_table),
        Paragraph("Cant.", header_table),
        Paragraph("Fecha Ingreso del Bien", header_table),
        Paragraph("Valor Contable", header_table)
    ]

    val_contable = row.get(col_map.get('valor'), '-')
    val_contable_clean = limpiar_val(val_contable)
    if val_contable_clean not in ["-", "."]:
        try:
            val_contable_num = float(val_contable_clean)
            val_contable_str = f"${val_contable_num:,.2f}"
        except:
            val_contable_str = val_contable_clean
    else:
        val_contable_str = "-"

    val_f_ingreso = limpiar_val(row.get(col_map.get('fecha_ingreso'), '-'))

    data_row = [
        Paragraph(limpiar_val(row.get(col_map.get('codigo'))), normal_small),
        Paragraph(limpiar_val(row.get(col_map.get('anterior'))), normal_small),
        Paragraph(limpiar_val(row.get(col_map.get('nombre'))), normal_small),
        Paragraph(limpiar_val(row.get(col_map.get('serie'))), normal_small),
        Paragraph(limpiar_val(row.get(col_map.get('modelo'))), normal_small),
        Paragraph(limpiar_val(row.get(col_map.get('marca'))), normal_small),
        Paragraph(limpiar_val(row.get(col_map.get('bodega'))), normal_small),
        Paragraph("BLD", normal_small),
        Paragraph("-", normal_small),
        Paragraph(fecha_acta_str, normal_small),
        Paragraph("1", normal_small),
        Paragraph(val_f_ingreso, normal_small),
        Paragraph(val_contable_str, normal_small)
    ]

    total_row = [
        Paragraph("<b>TOTALES:</b>", normal_small),
        "", "", "", "", "", "", "", "", "", "",
        Paragraph("<b>1</b>", normal_small),
        Paragraph(f"<b>{val_contable_str}</b>", normal_small)
    ]

    table_data = [headers, data_row, total_row]
    t = Table(table_data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('SPAN', (0, 2), (10, 2)),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#F3F4F6')),
    ]))
    story.append(t)
    story.append(Spacer(1, 25))

    custodio_dinamico = limpiar_val(row.get(col_map.get('custodio')))

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

archivos_excel = [f for f in os.listdir('.') if f.endswith('.xlsx') or f.endswith('.xls')]

df = None
nombre_excel_activo = "Ninguno"
fecha_acta_actual = datetime.now().strftime("%d/%m/%Y")

if archivo_subido is not None:
    df = cargar_excel(archivo_subido)
    nombre_excel_activo = archivo_subido.name
elif archivos_excel:
    archivo_local = archivos_excel[0]
    df = cargar_excel(archivo_local)
    nombre_excel_activo = archivo_local

ruta_escudo_activa = obtener_ruta_escudo()
if ruta_escudo_activa:
    nombre_escudo_activo = ruta_escudo_activa.replace("escudo_", "")
else:
    nombre_escudo_activo = "Sin escudo"

if df is not None:
    df.columns = df.columns.astype(str).str.strip()
    
    col_map = {}
    for col in df.columns:
        col_lower = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
        if 'codigo' in col_lower and 'ant' not in col_lower:
            col_map['codigo'] = col
        elif 'serie' in col_lower or 'identificacion' in col_lower:
            col_map['serie'] = col
        elif 'anterior' in col_lower:
            col_map['anterior'] = col
        elif 'bien' in col_lower and 'codigo' not in col_lower:
            col_map['nombre'] = col
        elif 'bodega' in col_lower:
            col_map['bodega'] = col
        elif 'custodio' in col_lower:
            col_map['custodio'] = col
        elif 'cedula' in col_lower or 'ruc' in col_lower:
            col_map['cedula'] = col
        elif ('fecha' in col_lower or 'ingreso' in col_lower or 'alta' in col_lower) and 'asigna' not in col_lower:
            col_map['fecha_ingreso'] = col
        elif 'valor' in col_lower or 'contable' in col_lower:
            col_map['valor'] = col
        elif 'marca' in col_lower or 'raza' in col_lower or 'otros' in col_lower:
            col_map['marca'] = col
        elif 'modelo' in col_lower or 'caracteristicas' in col_lower:
            col_map['modelo'] = col

    c_codigo = col_map.get('codigo', df.columns[0])
    c_serie = col_map.get('serie', df.columns[1] if len(df.columns) > 1 else df.columns[0])
    c_anterior = col_map.get('anterior', df.columns[0])

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    st.markdown(
        f"<p style='color: #CBD5E1; font-size: 0.82rem; margin-bottom: 6px; line-height: 1.4;'>"
        f"📄 Reporte: <b style='color: #93C5FD;'>{nombre_excel_activo}</b> &nbsp;|&nbsp; "
        f"🛡️ Escudo: <b style='color: #93C5FD;'>{nombre_escudo_activo}</b>"
        f"</p>", 
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
        res = df[
            (df[c_codigo].str.lower() == termino_busqueda.lower()) |
            (df[c_serie].str.lower() == termino_busqueda.lower()) |
            (df[c_anterior].str.lower() == termino_busqueda.lower())
        ]
        
        if res.empty:
            res = df[
                (df[c_codigo].str.contains(termino_busqueda, case=False, na=False)) |
                (df[c_serie].str.contains(termino_busqueda, case=False, na=False))
            ]
            
        if not res.empty:
            row = res.iloc[0]
            
            val_bien = limpiar_val(row.get(col_map.get('nombre'), '-'))
            val_bodega = limpiar_val(row.get(col_map.get('bodega'), '-'))
            val_fecha_ingreso = limpiar_val(row.get(col_map.get('fecha_ingreso'), '-'))
            
            st.success(f"📌 **BIEN:** {val_bien}")
            st.info(f"🏢 **BODEGA:** {val_bodega}")
            
            val_contable = row.get(col_map.get('valor'), '-')
            val_contable_clean = limpiar_val(val_contable)
            if val_contable_clean not in ["-", "."]:
                try:
                    val_contable_fmt = f"${float(val_contable_clean):,.2f}"
                except:
                    val_contable_fmt = val_contable_clean
            else:
                val_contable_fmt = "-"

            tab1, tab2 = st.tabs(["🏷️ Datos del Bien", "👤 Datos del Custodio"])
            
            with tab1:
                st.write(f"**Código del Bien:** {limpiar_val(row.get(c_codigo))}")
                st.write(f"**Código Anterior:** {limpiar_val(row.get(c_anterior))}")
                st.write(f"**Marca:** {limpiar_val(row.get(col_map.get('marca'), '-'))}")
                st.write(f"**Modelo:** {limpiar_val(row.get(col_map.get('modelo'), '-'))}")
                st.write(f"**Serie:** {limpiar_val(row.get(c_serie))}")
                st.write(f"**Fecha de Ingreso del Bien:** {val_fecha_ingreso}")
                st.write(f"**Valor Contable:** {val_contable_fmt}")
                
            with tab2:
                st.write(f"**Custodio Actual:** {limpiar_val(row.get(col_map.get('custodio'), '-'))}")
                st.write(f"**Número de Cédula / RUC:** {limpiar_val(row.get(col_map.get('cedula'), '-'))}")

            st.divider()
            
            pdf_bytes = generar_acta_pdf(row, fecha_acta_actual, col_map)
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
    st.warning("⚠️ No se encontró ningún archivo Excel válido o no se pudo cargar.")
