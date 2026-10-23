import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re
import io
from xml.sax.saxutils import escape
from PIL import Image as PILImage

# ReportLab para la generación del Acta PDF en formato horizontal (Landscape)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# Configuración de la página
st.set_page_config(page_title="Visor ESBYE", page_icon="📦", layout="centered")

# CSS: diseño móvil limpio, sin solapamientos y con espacio superior para que se vea el título
st.markdown("""
    <style>
        .block-container {
            padding-top: 4.5rem !important;
            padding-bottom: 6.0rem !important;
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }

        div.element-container {
            margin-bottom: 0px !important;
        }
        div[data-testid="stVerticalBlock"] {
            gap: 0.2rem !important;
        }

        @media (max-width: 768px) {
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                align-items: flex-end !important;
                gap: 0.3rem !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
                width: auto !important;
                flex: 1 1 auto !important;
                min-width: 0px !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child {
                flex: 0 0 50px !important;
            }
        }

        div[data-testid="stTextInput"] {
            margin-top: 4px !important;
        }
        div[data-testid="stTextInput"] label {
            display: none !important;
        }

        hr {
            margin: 0.3rem 0rem !important;
        }

        .info-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-top: 4px;
            margin-bottom: 4px;
            box-sizing: border-box;
        }
        .info-card {
            background: #1E293B;
            border: 1px solid #334155;
            border-radius: 5px;
            padding: 6px 8px;
            width: calc(50% - 3px);
            box-sizing: border-box;
        }
        .info-card-full {
            background: #1E293B;
            border: 1px solid #334155;
            border-radius: 5px;
            padding: 6px 8px;
            width: 100%;
            box-sizing: border-box;
        }
        .info-label {
            color: #94A3B8;
            font-size: 0.65rem;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            margin-bottom: 2px;
        }
        .info-val {
            color: #F8FAFC;
            font-size: 0.82rem;
            font-weight: 600;
            word-wrap: break-word;
        }
        .info-val-highlight {
            color: #38BDF8;
            font-size: 0.86rem;
            font-weight: bold;
            word-wrap: break-word;
        }
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# Funciones de apoyo
# ----------------------------------------------------------------------
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


def obtener_ruta_escudo():
    if os.path.exists("escudo.png"):
        return "escudo.png"
    for f in os.listdir('.'):
        if f.lower().startswith('escudo') and f.lower().endswith(('png', 'jpg', 'jpeg')):
            return f
    return None


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
        except Exception:
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
        except Exception:
            pass
    return texto.replace("DISENIO", "DISEÑO").replace("disenio", "diseño").replace("Disenio", "Diseño")


def valor_numerico(val):
    """Devuelve el valor como número, o None si no se puede convertir."""
    limpio = limpiar_val(val)
    if limpio in ["-", "."]:
        return None
    try:
        return float(limpio.replace('$', '').replace(',', ''))
    except Exception:
        return None


def formato_dinero(num):
    return f"${num:,.2f}"


def p(texto):
    """Texto seguro para Paragraph de ReportLab (evita errores con & < >)."""
    return escape(str(texto))


# ----------------------------------------------------------------------
# Generación del PDF (uno o varios bienes)
# ----------------------------------------------------------------------
def generar_acta_pdf(rows, fecha_acta_str, col_map):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.4 * inch
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

    # Escudo
    escudo_path = obtener_ruta_escudo()
    img_flowable = Paragraph("", normal_center)
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
            img_escudo = RLImage(img_buffer, width=0.65 * inch, height=0.65 * inch)
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

    header_table_obj = Table([[img_flowable, header_text_elements]], colWidths=[1.0 * inch, 9.2 * inch])
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

    col_widths = [55, 55, 114, 65, 55, 75, 90, 40, 50, 50, 30, 55]
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

    def dato(row, clave, defecto='-'):
        columna = col_map.get(clave)
        if not columna:
            return defecto
        return row.get(columna, defecto)

    table_data = [headers]
    total_valor = 0.0
    custodios = []

    for row in rows:
        num = valor_numerico(dato(row, 'valor'))
        if num is not None:
            total_valor += num
            val_str = f"$  {num:,.2f}"
        else:
            val_str = limpiar_val(dato(row, 'valor'))

        custodio = limpiar_val(dato(row, 'custodio'))
        if custodio != "-" and custodio not in custodios:
            custodios.append(custodio)

        table_data.append([
            Paragraph(p(limpiar_val(dato(row, 'codigo'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'anterior'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'nombre'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'serie'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'modelo'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'marca'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'bodega'))), normal_center),
            Paragraph(p(limpiar_val(dato(row, 'tipo_reg', 'BLD'))), normal_center),
            Paragraph(limpiar_val_fecha(dato(row, 'fecha_ingreso')), normal_center),
            Paragraph(fecha_acta_str, normal_center),
            Paragraph("1", normal_center),
            Paragraph(p(val_str), normal_center)
        ])

    total_str = f"$  {total_valor:,.2f}"
    table_data.append([
        Paragraph("<b>TOTAL:</b>", totales_right),
        "", "", "", "", "", "", "", "", "",
        Paragraph(f"<b>{len(rows)}</b>", normal_center),
        Paragraph(f"<b>{total_str}</b>", normal_center)
    ])

    ultima = len(table_data) - 1
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('SPAN', (0, ultima), (9, ultima)),
        ('BACKGROUND', (0, ultima), (-1, ultima), colors.HexColor('#F3F4F6')),
    ]))
    story.append(t)
    story.append(Spacer(1, 25))

    # Firmas: si hay varios custodios, se listan todos
    custodio_texto = "<br/>".join(p(c) for c in custodios) if custodios else "-"

    cell_entrega = [
        Paragraph("________________________________________", subtitle_style_center),
        Spacer(1, 4),
        Paragraph(custodio_texto, subtitle_style_center),
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
    firmas_table = Table([[cell_entrega, cell_recibe]], colWidths=[370, 370])
    firmas_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(KeepTogether([firmas_table]))

    doc.build(story)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------
# Estado de la sesión: bienes agregados a la acta
# ----------------------------------------------------------------------
if "acta_items" not in st.session_state:
    st.session_state.acta_items = []


def agregar_a_acta(row_dict, codigo):
    codigos = [i["codigo"] for i in st.session_state.acta_items]
    if codigo not in codigos:
        st.session_state.acta_items.append({"codigo": codigo, "row": row_dict})


def quitar_de_acta(indice):
    if 0 <= indice < len(st.session_state.acta_items):
        st.session_state.acta_items.pop(indice)


def vaciar_acta():
    st.session_state.acta_items = []


# ----------------------------------------------------------------------
# Barra lateral
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Opciones")
    archivo_subido = st.file_uploader("Subir nuevo reporte (.xlsx)", type=["xlsx", "xls"])

    if archivo_subido is not None:
        firma_archivo = f"{archivo_subido.name}-{archivo_subido.size}"
        # Solo se guarda una vez por archivo; así no se borra la acta en cada clic
        if st.session_state.get("ultimo_archivo") != firma_archivo:
            guardado = False
            try:
                with open("reporte_activo.xlsx", "wb") as f:
                    f.write(archivo_subido.getbuffer())
                with open("nombre_reporte.txt", "w", encoding="utf-8") as f:
                    f.write(archivo_subido.name)
                guardado = True
            except Exception as e:
                st.error(f"Error al guardar archivo: {e}")
            if guardado:
                st.session_state.ultimo_archivo = firma_archivo
                st.session_state.acta_items = []  # un reporte nuevo reinicia la acta
                st.rerun()
        else:
            st.success(f"Reporte '{archivo_subido.name}' cargado correctamente.")

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

# ----------------------------------------------------------------------
# Carga de datos
# ----------------------------------------------------------------------
df = None
display_reporte = "Ninguno"

if os.path.exists("reporte_activo.xlsx"):
    df = cargar_excel("reporte_activo.xlsx")
    if df is not None and os.path.exists("nombre_reporte.txt"):
        try:
            with open("nombre_reporte.txt", "r", encoding="utf-8") as f:
                display_reporte = f.read().strip()
        except Exception:
            display_reporte = "reporte_activo.xlsx"

fecha_acta_actual = datetime.now().strftime("%d/%m/%Y")

if df is not None:
    df.columns = df.columns.astype(str).str.strip()
    cols = list(df.columns)

    col_map = {}
    if len(cols) > 0: col_map['codigo'] = cols[0]
    if len(cols) > 1: col_map['anterior'] = cols[1]
    if len(cols) > 2: col_map['tipo_reg'] = cols[2]
    if len(cols) > 3: col_map['nombre'] = cols[3]
    if len(cols) > 4: col_map['serie'] = cols[4]
    if len(cols) > 5: col_map['modelo'] = cols[5]
    if len(cols) > 6: col_map['marca'] = cols[6]
    if len(cols) > 7: col_map['bodega'] = cols[7]
    if len(cols) > 8: col_map['cedula'] = cols[8]
    if len(cols) > 9: col_map['custodio'] = cols[9]
    if len(cols) > 10: col_map['fecha_ingreso'] = cols[10]
    if len(cols) > 11: col_map['valor'] = cols[11]

    c_codigo = col_map.get('codigo', cols[0])
    c_serie = col_map.get('serie', cols[4] if len(cols) > 4 else cols[0])
    c_anterior = col_map.get('anterior', c_codigo)

    # Encabezado visual
    st.markdown(f"""
        <div style="text-align: center; margin-top: 2px; margin-bottom: 6px;">
            <div style="color: #38BDF8; font-size: 1.25rem; font-weight: bold; line-height: 1.1; margin: 0;">BIENES Y BODEGAS</div>
            <div style="color: #60A5FA; font-style: italic; font-weight: bold; font-size: 0.78rem; margin: 2px 0 4px 0;">Desarrollado por: Ing. William Cisneros R.</div>
            <div style="border-top: 1px solid #334155; margin: 4px 0;"></div>
            <div style="background-color: #0F172A; border: 1px solid #1E293B; padding: 4px 8px; border-radius: 5px; font-size: 0.78rem; color: #94A3B8;">
                📄 Reporte Activo: <b style="color: #38BDF8;">{display_reporte}</b>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Etiqueta de búsqueda con espacio para que no quede tapada
    st.markdown("<p style='color: #F8FAFC; font-weight: bold; font-size: 0.85rem; margin: 12px 0px 12px 0px;'>🔍 Buscar Código / Serie:</p>", unsafe_allow_html=True)

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
            val_fecha_ingreso = limpiar_val_fecha(row.get(col_map.get('fecha_ingreso'), '-'))

            num_valor = valor_numerico(row.get(col_map.get('valor'), '-')) if col_map.get('valor') else None
            val_contable_fmt = formato_dinero(num_valor) if num_valor is not None else limpiar_val(row.get(col_map.get('valor'), '-'))

            st.markdown(f"""
                <div style="background-color: #064E3B; border-left: 4px solid #10B981; padding: 6px 10px; border-radius: 5px; color: #ECFDF5; font-size: 0.82rem; font-weight: bold; margin-top: 6px; margin-bottom: 4px;">📌 Bien: {val_bien}</div>
                <div style="background-color: #1E3A8A; border-left: 4px solid #3B82F6; padding: 6px 10px; border-radius: 5px; color: #EFF6FF; font-size: 0.80rem; font-weight: bold; margin-bottom: 6px;">🏢 Bodega: {val_bodega}</div>
            """, unsafe_allow_html=True)

            tab1, tab2 = st.tabs(["🏷️ Datos del Bien", "👤 Datos del Custodio"])

            with tab1:
                st.markdown(f"""
                    <div class="info-grid">
                        <div class="info-card">
                            <div class="info-label">Código del Bien</div>
                            <div class="info-val-highlight">{limpiar_val(row.get(c_codigo, '-'))}</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Código Anterior</div>
                            <div class="info-val">{limpiar_val(row.get(c_anterior, '-'))}</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Marca/ Otros</div>
                            <div class="info-val">{limpiar_val(row.get(col_map.get('marca'), '-'))}</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Modelo/ Características</div>
                            <div class="info-val">{limpiar_val(row.get(col_map.get('modelo'), '-'))}</div>
                        </div>
                        <div class="info-card-full">
                            <div class="info-label">Serie/ Identificación</div>
                            <div class="info-val">{limpiar_val(row.get(c_serie, '-'))}</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Fecha de Ingreso</div>
                            <div class="info-val">{val_fecha_ingreso}</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Valor Contable</div>
                            <div class="info-val-highlight">{val_contable_fmt}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

            with tab2:
                st.markdown(f"""
                    <div class="info-card-full" style="margin-top: 4px; margin-bottom: 5px;">
                        <div class="info-label">Custodio Actual</div>
                        <div class="info-val">{limpiar_val(row.get(col_map.get('custodio'), '-'))}</div>
                    </div>
                    <div class="info-card-full">
                        <div class="info-label">Nro de Cédula/ RUC</div>
                        <div class="info-val">{limpiar_val(row.get(col_map.get('cedula'), '-'))}</div>
                    </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
            st.divider()

            codigo_bien = limpiar_val(row.get(c_codigo, '-'))
            ya_agregado = codigo_bien in [i["codigo"] for i in st.session_state.acta_items]

            # Botón para juntar este bien con otros en una sola acta
            if ya_agregado:
                st.button("✅ Este bien ya está en la lista", disabled=True, use_container_width=True, key="btn_ya_agregado")
            else:
                st.button(
                    "➕ Agregar a la lista de varios bienes",
                    on_click=agregar_a_acta,
                    args=(row.to_dict(), codigo_bien),
                    use_container_width=True,
                    key="btn_agregar"
                )

            # Acta de solo este bien
            pdf_individual = generar_acta_pdf([row.to_dict()], fecha_acta_actual, col_map)
            cod_bien_clean = str(row.get(c_codigo, 'Bien')).replace('/', '_')
            fecha_nom = datetime.now().strftime("%d-%m-%Y")

            st.download_button(
                label="📄 Generar e Imprimir Acta PDF Oficial",
                data=pdf_individual,
                file_name=f"Acta_{cod_bien_clean}_{fecha_nom}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dl_individual"
            )
        else:
            st.error(f"❌ No se encontró ningún bien con el código: {termino_busqueda}")

    # ------------------------------------------------------------------
    # Acta en preparación (varios bienes)
    # ------------------------------------------------------------------
    items = st.session_state.acta_items
    if items:
        total_acta = 0.0
        for it in items:
            n = valor_numerico(it["row"].get(col_map.get('valor'), '-')) if col_map.get('valor') else None
            if n is not None:
                total_acta += n

        st.divider()
        lineas = ""
        for i, it in enumerate(items):
            r = it["row"]
            nombre_it = limpiar_val(r.get(col_map.get('nombre'), '-')) if col_map.get('nombre') else '-'
            lineas += f"<div style='margin-top:4px;'><b style='color:#38BDF8;'>{i + 1}. {it['codigo']}</b> - {nombre_it}</div>"
        st.markdown(
            "<div class='info-card-full'>"
            f"<div class='info-label'>Lista para acta: {len(items)} bien(es) | Total {formato_dinero(total_acta)}</div>"
            f"<div class='info-val'>{lineas}</div>"
            "</div>",
            unsafe_allow_html=True
        )
        with st.expander("❌ Quitar un bien de la lista"):
            for i, it in enumerate(items):
                st.button(f"Quitar {it['codigo']}", key=f"quitar_{i}_{it['codigo']}", on_click=quitar_de_acta, args=(i,), use_container_width=True)

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        pdf_grupal = generar_acta_pdf([it["row"] for it in items], fecha_acta_actual, col_map)
        st.download_button(
            label=f"📄 Generar Acta PDF con los {len(items)} bienes de la lista",
            data=pdf_grupal,
            file_name=f"Acta_Grupal_{len(items)}_bienes_{datetime.now().strftime('%d-%m-%Y')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_grupal"
        )
        st.button("🧹 Vaciar lista", on_click=vaciar_acta, use_container_width=True, key="btn_vaciar")

    # Espacio libre amplio en la parte inferior para celulares
    st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)

else:
    st.warning("⚠️ No se encontró ningún archivo Excel activo en el servidor. Por favor, suba su reporte institucional en formato .xlsx mediante el menú de la barra lateral izquierda.")
