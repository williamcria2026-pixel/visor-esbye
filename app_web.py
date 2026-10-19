iimport streamlit as st
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

# Estilos CSS avanzados para mantener el buscador y el botón perfectamente alineados
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

# Encabezado corporativo optimizado
st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>BIENES Y BODEGAS</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #1D4ED8; font-style: italic; font-weight: bold; font-size: 0.85rem;'>Desarrollado por: Ing. William Cisneros R.</p>", unsafe_allow_html=True)
st.divider()

# Función para limpiar valores nulos o 'nan' y corregir DISENIO a DISEÑO
def limpiar_val(val):
    if pd.isna(val) or str(val).strip().lower() in ["nan", "none", "", "nat", "null"]:
        return "-"
    texto = str(val).strip()
    texto = texto.replace("DISENIO", "DISEÑO").replace("disenio", "diseño").replace("Disenio", "Diseño")
    return texto

# Función para cargar archivos Excel con múltiples motores
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

# Función inteligente para detectar cualquier archivo de escudo (png, jpg, etc.)
def obtener_ruta_escudo():
    for f in os.listdir('.'):
        if f.lower().startswith('escudo'):
            return f
    return None

# Función para generar el PDF del Acta con el Escudo institucional centrado
def generar_acta_pdf(row, fecha_acta_str):
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

    # Estilos de texto CENTRADOS
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=11, leading=13, alignment=1, fontName="Helvetica-Bold")
    subtitle_style_center = ParagraphStyle('SubTitleStyleCenter', parent=styles['Normal'], fontSize=9, leading=12, alignment=1)
    acta_title = ParagraphStyle('ActaTitle', parent=styles['Heading2'], fontSize=10, leading=12, alignment=1, fontName="Helvetica-Bold")
    normal_small = ParagraphStyle('NormalSmall', parent=styles['Normal'], fontSize=7, leading=8)
    header_table = ParagraphStyle('HeaderTable', parent=styles['Normal'], fontSize=7, leading=8, fontName="Helvetica-Bold", alignment=1, textColor=colors.white)

    # Inserción automática del escudo detectado
    escudo_path = obtener_ruta_escudo()
    if escudo_path and os.path.exists(escudo_path):
        try:
            with PILImage.open(escudo_path) as img_test:
                img_test.verify()
            img_escudo = RLImage(escudo_path, width=0.6*inch, height=0.6*inch)
            img_escudo.hAlign = 'CENTER'
            story.append(img_escudo)
            story.append(Spacer(1, 4))
        except Exception:
            pass 

    # 1. Encabezado Institucional (TODO CENTRADO)
    story.append(Paragraph("MINISTERIO DE ECONOMÍA Y FINANZAS", title_style))
    story.append(Paragraph("176-9999-0000-UNIVERSIDAD TÉCNICA DE AMBATO - PLANTA CENTRAL", subtitle_style_center))
    story.append(Spacer(1, 4))
    story.append(Paragraph("ACTA DE ASIGNACIÓN - REASIGNACIÓN DE BIENES", acta_title))
    story.append(Spacer(1, 6))

    story.append(Paragraph(f"<b>FECHA:</b> {fecha_acta_str}", normal_small))
    story.append(Spacer(1, 4))

    # 2. Tabla de Datos Dinámica
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

    val_contable = row.get('Valor Contable', '-')
    val_contable_clean = limpiar_val(val_contable)
    if val_contable_clean not in ["-", "."]:
        try:
            val_contable_num = float(val_contable_clean)
            val_contable_str = f"${val_contable_num:,.2f}"
        except:
            val_contable_str = val_contable_clean
    else:
        val_contable_str = "-"

    data_row = [
        Paragraph(limpiar_val(row.get('Código del Bien')), normal_small),
        Paragraph(limpiar_val(row.get('Código Anterior')), normal_small),
        Paragraph(limpiar_val(row.get('Bien')), normal_small),
        Paragraph(limpiar_val(row.get('Serie/ Identificación')), normal_small),
        Paragraph(limpiar_val(row.get('Modelo/ Características')), normal_small),
        Paragraph(limpiar_val(row.get('Marca/ Otros')), normal_small),
        Paragraph(limpiar_val(row.get('Ubicación de Bodega')), normal_small),
        Paragraph("BLD", normal_small),
        Paragraph("-", normal_small),
        Paragraph(fecha_acta_str, normal_small),
        Paragraph("1", normal_small),
        Paragraph(val_contable_str, normal_small)
    ]

    total_row = [
        Paragraph("<b>TOTALES:</b>", normal_small),
        "", "", "", "", "", "", "", "", "",
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
        ('SPAN', (0, 2), (9, 2)),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#F3F4F6')),
    ]))
    story.append(t)
    story.append(Spacer(1, 25))

    # 3. Sección de Firmas (LÍNEAS Y TEXTOS 100% CENTRADOS)
    custodio_dinamico = limpiar_val(row.get('Custodio Actual'))

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

with st.sidebar:
    st.header("⚙️ Opciones")
    archivo_subido = st.file_uploader("Subir nuevo reporte (.xlsx)", type=["xlsx", "xls"])

df = None
fecha_archivo_str = datetime.now().strftime("%d/%m/%Y")

if archivo_subido is not None:
    df = cargar_excel(archivo_subido)
elif archivos_excel:
    archivo_local = archivos_excel[0]
    df = cargar_excel(archivo_local)
    try:
        timestamp = os.path.getmtime(archivo_local)
        fecha_archivo_str = time.strftime('%d/%m/%Y', time.localtime(timestamp))
    except:
        pass

if df is not None:
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    # Bloque de Fecha y Título del buscador
    st.markdown(f"<p style='color: #CBD5E1; font-size: 0.85rem; margin-bottom: 4px;'>📅 Base actualizada: <b>{fecha_archivo_str}</b></p>", unsafe_allow_html=True)
    st.markdown("<p style='color: #FFFFFF; font-weight: bold; font-size: 0.95rem; margin-bottom: 2px;'>🔍 Buscar Código / Serie:</p>", unsafe_allow_html=True)

    if "busqueda_val" not in st.session_state:
        st.session_state.busqueda_val = ""

    # Fila única perfecta lado a lado en móvil y PC
    col_input, col_btn = st.columns([4.2, 0.8])
    with col_input:
        codigo = st.text_input(
            "",
            value=st.session_state.busqueda_val,
            key="input_codigo",
            placeholder="Ingrese código o serie..."
        ).strip()
    
    with col_btn:
        if st.button("🗑️", help="Limpiar búsqueda", use_container_width=True):
            st.session_state.busqueda_val = ""
            st.rerun()

    if codigo:
        res = df[
            (df["Código del Bien"].str.lower() == codigo.lower()) |
            (df["Serie/ Identificación"].str.lower() == codigo.lower()) |
            (df["Código Anterior"].str.lower() == codigo.lower())
        ]
        
        if res.empty:
            res = df[
                (df["Código del Bien"].str.contains(codigo, case=False, na=False)) |
                (df["Serie/ Identificación"].str.contains(codigo, case=False, na=False))
            ]
            
        if not res.empty:
            row = res.iloc[0]
            
            st.success(f"📌 **BIEN:** {limpiar_val(row.get('Bien'))}")
            
            ubicacion_bodega = limpiar_val(row.get('Ubicación de Bodega'))
            st.info(f"🏢 **BODEGA:** {ubicacion_bodega}")
            
            val_contable = row.get('Valor Contable', '-')
            val_contable_clean = limpiar_val(val_contable)
            if val_contable_clean not in ["-", "."]:
                try:
                    val_contable_fmt = f"${float(val_contable_clean):,.2f}"
                except:
                    val_contable_fmt = val_contable_clean
            else:
                val_contable_fmt = "-"

            # Pestañas compactas
            tab1, tab2 = st.tabs(["🏷️ Datos del Bien", "👤 Datos del Custodio"])
            
            with tab1:
                st.write(f"**Código del Bien:** {limpiar_val(row.get('Código del Bien'))}")
                st.write(f"**Código Anterior:** {limpiar_val(row.get('Código Anterior'))}")
                st.write(f"**Marca:** {limpiar_val(row.get('Marca/ Otros'))}")
                st.write(f"**Modelo:** {limpiar_val(row.get('Modelo/ Características'))}")
                st.write(f"**Serie:** {limpiar_val(row.get('Serie/ Identificación'))}")
                st.write(f"**Valor Contable:** {val_contable_fmt}")
                
            with tab2:
                st.write(f"**Custodio Actual:** {limpiar_val(row.get('Custodio Actual'))}")
                st.write(f"**Número de Cédula / RUC:** {limpiar_val(row.get('Nro de Cédula/ RUC'))}")

            st.divider()
            
            pdf_bytes = generar_acta_pdf(row, fecha_archivo_str)
            cod_bien_clean = str(row.get('Código del Bien', 'Bien')).replace('/', '_')
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
            st.error(f"❌ No se encontró ningún bien con el código: {codigo}")
else:
    st.warning("⚠️ No se encontró ningún archivo Excel válido o no se pudo cargar.")
