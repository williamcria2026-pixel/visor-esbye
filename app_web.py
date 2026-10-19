import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
import io

# ReportLab para la generación del Acta PDF en formato horizontal (Landscape)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# Configuración de la página
st.set_page_config(page_title="Visor ESBYE", page_icon="📦", layout="centered")

# Encabezado corporativo
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>CONSULTA RÁPIDA DE BIENES Y BODEGAS</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #1D4ED8; font-style: italic; font-weight: bold;'>Desarrollado por: Ing. William Cisneros R.</p>", unsafe_allow_html=True)
st.divider()

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

# Función para generar el PDF del Acta con formato totalmente dinámico
def generar_acta_pdf(row):
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

    # Estilos de texto
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=11, leading=13, alignment=1, fontName="Helvetica-Bold")
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=9, leading=11, alignment=1)
    acta_title = ParagraphStyle('ActaTitle', parent=styles['Heading2'], fontSize=10, leading=12, alignment=1, fontName="Helvetica-Bold")
    normal_small = ParagraphStyle('NormalSmall', parent=styles['Normal'], fontSize=7, leading=8)
    header_table = ParagraphStyle('HeaderTable', parent=styles['Normal'], fontSize=7, leading=8, fontName="Helvetica-Bold", alignment=1, textColor=colors.white)

    # 1. Encabezado Institucional (FIJO)
    story.append(Paragraph("MINISTERIO DE ECONOMÍA Y FINANZAS", title_style))
    story.append(Paragraph("176-9999-0000-UNIVERSIDAD TÉCNICA DE AMBATO - PLANTA CENTRAL", subtitle_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("ACTA DE ASIGNACIÓN - REASIGNACIÓN DE BIENES", acta_title))
    story.append(Spacer(1, 8))

    # Fecha actual
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    story.append(Paragraph(f"<b>FECHA:</b> {fecha_actual}", normal_small))
    story.append(Spacer(1, 6))

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
    if val_contable not in ["nan", "None", "-", "."]:
        try:
            val_contable_num = float(val_contable)
            val_contable_str = f"${val_contable_num:,.2f}"
        except:
            val_contable_str = str(val_contable)
    else:
        val_contable_str = "-"

    data_row = [
        Paragraph(str(row.get('Código del Bien', '-')), normal_small),
        Paragraph(str(row.get('Código Anterior', '-')), normal_small),
        Paragraph(str(row.get('Bien', '-')), normal_small),
        Paragraph(str(row.get('Serie/ Identificación', '-')), normal_small),
        Paragraph(str(row.get('Modelo/ Características', '-')), normal_small),
        Paragraph(str(row.get('Marca/ Otros', '-')), normal_small),
        Paragraph(str(row.get('Ubicación de Bodega', '-')), normal_small),
        Paragraph("BLD", normal_small),
        Paragraph("-", normal_small),
        Paragraph(fecha_actual, normal_small),
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
    story.append(Spacer(1, 35))

    # 3. Sección de Firmas Dinámica (Custodio de la Búsqueda + Bodega)
    custodio_dinamico = str(row.get('Custodio Actual', 'CUSTODIO NO ESPECIFICADO')).strip()
    bodega_dinamica = str(row.get('Ubicación de Bodega', 'BODEGA NO ESPECIFICADA')).strip()

    firma_entrega = Paragraph(
        f"<br/>________________________________________<br/>"
        f"<b>{custodio_dinamico}</b><br/>"
        f"QUIEN ENTREGA<br/>"
        f"{bodega_dinamica}", 
        subtitle_style
    )
    
    firma_recibe = Paragraph(
        f"<br/>________________________________________<br/>"
        f"<b>CUSTODIO RECEPCIONA</b><br/>"
        f"FIRMA Y FECHA", 
        subtitle_style
    )

    firmas_data = [[firma_entrega, firma_recibe]]
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
nombre_fuente = None
fecha_mod = ""

if archivo_subido is not None:
    df = cargar_excel(archivo_subido)
    nombre_fuente = archivo_subido.name
    fecha_mod = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime())
elif archivos_excel:
    archivo_local = archivos_excel[0]
    df = cargar_excel(archivo_local)
    nombre_fuente = archivo_local
    timestamp = os.path.getmtime(archivo_local)
    fecha_mod = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime(timestamp))

if df is not None:
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    st.caption(f"🟢 **Reporte:** {nombre_fuente} ({len(df):,} registros) | 📅 **Actualizado:** {fecha_mod}")

    # Control de estado de la búsqueda
    if "busqueda_val" not in st.session_state:
        st.session_state.busqueda_val = ""

    col_input, col_btn = st.columns([4, 1])
    with col_input:
        codigo = st.text_input(
            "🔍 Ingrese o escanee el Código del Bien / Serie:",
            value=st.session_state.busqueda_val,
            key="input_codigo"
        ).strip()
    
    with col_btn:
        st.write(" ")
        st.write(" ")
        if st.button("🗑️ Limpiar", use_container_width=True):
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
            
            # Título del bien
            st.success(f"📌 **BIEN ENCONTRADO:** {row.get('Bien', '-')}")
            
            # Ubicación de Bodega sin saltos sobrantes
            ubicacion_bodega = str(row.get('Ubicación de Bodega', '-')).strip()
            st.info(f"🏢 **UBICACIÓN DE BODEGA:** {ubicacion_bodega}")
            
            # Formateo de valor contable
            val_contable = row.get('Valor Contable', '-')
            if val_contable not in ["nan", "None", "-", "."]:
                try:
                    val_contable = f"${float(val_contable):,.2f}"
                except:
                    pass

            # Pestañas de información
            tab1, tab2 = st.tabs(["🏷️ Datos del Bien", "👤 Datos del Custodio"])
            
            with tab1:
                st.write(f"**Marca:** {row.get('Marca/ Otros', '-')}")
                st.write(f"**Modelo:** {row.get('Modelo/ Características', '-')}")
                st.write(f"**Serie:** {row.get('Serie/ Identificación', '-')}")
                st.write(f"**Valor Contable:** {val_contable}")
                st.write(f"**Código del Bien:** {row.get('Código del Bien', '-')}")
                st.write(f"**Código Anterior:** {row.get('Código Anterior', '-')}")
                
            with tab2:
                st.write(f"**Custodio Actual:** {row.get('Custodio Actual', '-')}")
                st.write(f"**Número de Cédula / RUC:** {row.get('Nro de Cédula/ RUC', '-')}")

            st.divider()
            
            # Botón para Generar y Descargar el Acta PDF Oficial
            pdf_bytes = generar_acta_pdf(row)
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
