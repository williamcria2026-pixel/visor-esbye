import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
from datetime import datetime
import re
import io
import base64
import json
import time
import hmac
import hashlib
import unicodedata
from html import escape as hesc
from xml.sax.saxutils import escape
from PIL import Image as PILImage

# ReportLab para la generación del Acta PDF en formato horizontal (Landscape)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# Configuración de la página
st.set_page_config(page_title="Visor ESBYE", page_icon="📦", layout="centered")

# ----------------------------------------------------------------------
# Estilo visual inspirado en el sistema ESBYE (fondo claro, azul y morado)
# ----------------------------------------------------------------------
st.markdown("""
    <style>
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
            background: #FFFFFF !important;
            color: #222222 !important;
            font-family: Arial, Helvetica, "Liberation Sans", sans-serif !important;
        }
        .stApp p, .stApp label { color: #222222; }
        .block-container {
            padding-top: 4rem !important;
            padding-bottom: 6.0rem !important;
            padding-left: 0.6rem !important;
            padding-right: 0.6rem !important;
        }
        div.element-container { margin-bottom: 0px !important; }
        div[data-testid="stVerticalBlock"] { gap: 0.2rem !important; }

        /* Filas de botones en una sola línea en el celular */
        @media (max-width: 768px) {
            div[data-testid="stHorizontalBlock"] {
                display: flex !important;
                flex-direction: row !important;
                flex-wrap: nowrap !important;
                align-items: center !important;
                gap: 0.3rem !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
                width: auto !important;
                flex: 1 1 auto !important;
                min-width: 0px !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child {
                flex: 0 0 56px !important;
            }
        }

        /* Cuadro de texto */
        div[data-testid="stTextInput"] label { display: none !important; }
        div[data-testid="stTextInput"] { margin-top: 2px !important; }
        div[data-baseweb="input"], div[data-baseweb="base-input"] {
            background: #EDEEF0 !important;
            border-radius: 3px !important;
        }
        div[data-baseweb="input"] { border: 1px solid #767676 !important; }
        div[data-baseweb="input"] input {
            color: #222222 !important;
            -webkit-text-fill-color: #222222 !important;
            background: #EDEEF0 !important;
            font-size: 16px !important;
            height: 42px;
        }

        /* Botones grises estilo ESBYE */
        div[data-testid="stButton"] button,
        div[data-testid="stDownloadButton"] button,
        div[data-testid="stFormSubmitButton"] button {
            background: #EFEFEF !important;
            border: 1px solid #767676 !important;
            border-radius: 3px !important;
            color: #000000 !important;
            min-height: 46px !important;
            font-size: 16px !important;
        }
        div[data-testid="stButton"] button p,
        div[data-testid="stDownloadButton"] button p { color: inherit !important; }
        div[data-testid="stButton"] button:disabled { color: #555555 !important; opacity: 0.8; }
        .st-key-dl_individual button, .st-key-dl_grupal button {
            background: #2B1D6B !important;
            border: 1px solid #1B1250 !important;
            color: #FFFFFF !important;
            font-weight: bold !important;
        }
        .st-key-dl_individual button:hover:not(:disabled), .st-key-dl_grupal button:hover:not(:disabled) {
            background: #55E6FF !important;
            border-color: #55E6FF !important;
            color: #1B1250 !important;
        }

        /* Pestañas */
        div[data-baseweb="tab-list"] { gap: 2px !important; }
        div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] { display: none !important; }
        button[data-baseweb="tab"] {
            flex: 1 !important;
            background: #2B1D6B !important;
            color: #FFFFFF !important;
            border-radius: 0 !important;
            padding: 10px 4px !important;
            height: auto !important;
            justify-content: center !important;
        }
        button[data-baseweb="tab"] p { color: inherit !important; font-size: 14px !important; }
        button[data-baseweb="tab"][aria-selected="true"] {
            background: #55E6FF !important;
            color: #1B1250 !important;
            font-weight: bold !important;
        }
        div[data-baseweb="tab-panel"] { padding-top: 2px !important; }

        /* Barra lateral */
        section[data-testid="stSidebar"] { background: #EEF1F6 !important; }
        section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] small { color: #1B2A78 !important; }

        hr { margin: 0.3rem 0rem !important; }

        /* Componentes de diseño */
        .nav {
            background: #2B1D6B; border-bottom: 3px solid #55E6FF; color: #FFFFFF; height: 54px;
            display: flex; align-items: center; justify-content: center;
            font-weight: bold; font-size: 1.1rem; letter-spacing: 1px;
        }
        .dev { text-align: center; font-size: 0.75rem; font-style: italic; color: #1B2A78; padding: 8px 0 2px; }
        .red { border-top: 1px solid #E02020; margin: 6px 0 0; }
        .bar {
            display: block; background-color: #16478F;
            background-image: radial-gradient(rgba(255,255,255,.14) 1px, transparent 1px);
            background-size: 4px 4px; color: #FFFFFF; text-align: center;
            font-weight: bold; font-size: 0.95rem; padding: 11px 8px; margin-top: 8px;
        }
        .row { display: grid; grid-template-columns: 40% 60%; gap: 2px; margin-top: 2px; }
        .lb {
            background: #DDE3EA; color: #1B2A78; font-size: 0.74rem; text-transform: uppercase;
            text-align: center; padding: 10px 6px; display: flex; align-items: center; justify-content: center;
        }
        .vl { background: #DDE3EA; padding: 6px; min-width: 0; }
        .vl span {
            display: block; overflow-wrap: anywhere; background: #EDEEF0; border: 1px solid #C3C8D0;
            border-radius: 3px; padding: 8px 10px; font-size: 0.92rem; color: #333333;
        }
        .vl span.hl { color: #16478F; font-weight: bold; }
        .acta-item {
            background: #EDEEF0; border: 1px solid #C3C8D0; border-radius: 3px;
            padding: 11px 10px; font-size: 0.92rem; color: #222222;
        }
        .acta-item b { color: #16478F; }
        .aviso {
            background: #FFF4CE; border: 1px solid #C99700; border-radius: 3px;
            padding: 10px; font-size: 0.8rem; color: #4A3600; margin-top: 8px;
        }
        .t { font-size: 0.75rem; color: #1B2A78; margin: 8px 0 2px; }
        .err {
            background: #FDE7E9; border: 1px solid #C62828; border-radius: 3px;
            padding: 10px; margin-top: 8px; font-size: 0.88rem; color: #7A1010;
        }
        div[data-testid="stRadio"] label p { color: #222222 !important; font-size: 0.9rem !important; }
        .st-key-dl_individual button:disabled, .st-key-dl_grupal button:disabled {
            background: #9AA7BD !important; border-color: #9AA7BD !important; color: #FFFFFF !important;
        }
        div[data-testid="stForm"] {
            border: 1px solid #8A8A8A !important; border-radius: 0 !important;
            padding: 24px 16px 28px !important; background: #FFFFFF !important; margin-top: 110px;
        }
        .login-t { text-align: center; font-weight: bold; font-size: 0.9rem; color: #2B1D6B; margin-bottom: 26px; }
        .login-l { text-align: center; font-weight: bold; font-size: 0.8rem; color: #2B1D6B; margin: 0 0 14px; }
        div[data-testid="stForm"] div[data-testid="stTextInput"],
        div[data-testid="stForm"] div[data-testid="stFormSubmitButton"] {
            width: 55% !important; min-width: 170px; margin-left: auto !important; margin-right: auto !important;
        }
        div[data-testid="stForm"] div[data-testid="stFormSubmitButton"] { margin-top: 16px !important; }

        /* Pestañas: selector con aspecto de pestañas */
        .st-key-pestana { margin-top: 4px !important; margin-bottom: 4px !important; }
        .st-key-pestana div[role="radiogroup"] { gap: 2px !important; flex-wrap: nowrap !important; width: 100% !important; }
        /* Pestañas: una fila, esquinas redondeadas, sin círculos (el color indica la activa) */
        .st-key-pestana [data-testid="stWidgetLabel"] { display: none !important; }
        .st-key-pestana div[role="radiogroup"] { gap: 4px !important; }
        .st-key-pestana label:not([data-testid="stWidgetLabel"]) {
            flex: 0 1 auto !important; min-width: 0 !important;
            display: flex !important; align-items: center !important; justify-content: center !important;
            background: #2B1D6B !important; margin: 0 !important;
            padding: 6px 9px !important; min-height: 34px !important; border-radius: 10px !important; cursor: pointer;
        }
        .st-key-pestana label :is(div, span, svg):not(:has([data-testid="stMarkdownContainer"])):not([data-testid="stMarkdownContainer"]):not([data-testid="stMarkdownContainer"] *) { display: none !important; }
        .st-key-pestana div[role="radiogroup"] { display: flex !important; flex-direction: row !important; flex-wrap: nowrap !important; width: 100% !important; }
        .st-key-pestana div[role="radiogroup"] > * { flex: 0 1 auto !important; min-width: 0 !important; }
        .st-key-pestana label:not([data-testid="stWidgetLabel"]) { width: auto !important; box-sizing: border-box !important; }
        .st-key-pestana label:not([data-testid="stWidgetLabel"]) p {
            color: #FFFFFF !important; font-size: 13px !important; margin: 0 !important; white-space: nowrap;
        }
        .st-key-pestana label:not([data-testid="stWidgetLabel"]):has(input:checked) { background: #55E6FF !important; }
        .st-key-pestana label:not([data-testid="stWidgetLabel"]):has(input:checked) p { color: #1B1250 !important; font-weight: bold !important; }

        /* Corrige la superposición: Streamlit resta 1rem al final de cada bloque markdown */
        div[data-testid="stMarkdown"], div[data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
        div[data-testid="stMarkdownContainer"] > div { margin-bottom: 0 !important; }
        /* Lista de coincidencias ("Elija una"): bloques sin círculo, el color indica la elegida */
        .st-key-receptor_radio div[role="radiogroup"] { flex-direction: column !important; gap: 2px !important; width: 100% !important; }
        .st-key-receptor_radio [data-testid="stWidgetLabel"] { display: none !important; }
        .st-key-receptor_radio label:not([data-testid="stWidgetLabel"]) {
            display: flex !important; flex-direction: row-reverse !important; justify-content: space-between !important;
            align-items: center !important; width: 100% !important; margin: 0 !important;
            padding: 11px 10px !important; border-radius: 3px !important;
            background: #EDEEF0 !important; border: 1px solid #C3C8D0 !important; cursor: pointer;
        }
        .st-key-receptor_radio label p { color: #222222 !important; font-size: 0.92rem !important; margin: 0 !important; }
        .st-key-receptor_radio label:has(input:checked) { background: #DDE9FA !important; border-color: #16478F !important; }
        .st-key-receptor_radio label:has(input:checked) p { color: #16478F !important; font-weight: bold !important; }
        .st-key-receptor_radio label :is(div, span, svg):not(:has([data-testid="stMarkdownContainer"])):not([data-testid="stMarkdownContainer"]):not([data-testid="stMarkdownContainer"] *) { display: none !important; }

        /* Botones ❌ para quitar bienes de la lista: misma altura que el renglón, en rojo */
        [class*="st-key-quitar_"] button {
            background: #FDE7E9 !important; border: 1px solid #C62828 !important; color: #C62828 !important;
            min-height: 46px !important;
        }
        .acta-item { min-height: 46px; display: flex; align-items: center; box-sizing: border-box; }
        .acta-item b { margin-right: 6px; }

        .nota { text-align: center; font-size: 0.75rem; color: #1B2A78; margin: 6px 0; }

        /* Textura punteada (estilo ESBYE) en todas las zonas de color */
        .nav, .st-key-dl_individual button, .st-key-dl_grupal button,
        .st-key-pestana label:not([data-testid="stWidgetLabel"]) {
            background-image: radial-gradient(rgba(255,255,255,.16) 1px, transparent 1px) !important;
            background-size: 4px 4px !important;
        }
        .st-key-pestana label:not([data-testid="stWidgetLabel"]):has(input:checked) {
            background-image: radial-gradient(rgba(255,255,255,.55) 1px, transparent 1px) !important;
            background-size: 4px 4px !important;
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


def fila(etiqueta, valor, resaltar=False):
    """Fila de datos estilo ESBYE (etiqueta + valor)."""
    clase = ' class="hl"' if resaltar else ''
    return f'<div class="row"><div class="lb">{hesc(etiqueta)}</div><div class="vl"><span{clase}>{hesc(str(valor))}</span></div></div>'


def limpiar_nombre_archivo(texto):
    """Quita caracteres no permitidos en nombres de archivo y cambia espacios por _."""
    texto = re.sub(r'[\\/:*?"<>|]+', '_', str(texto).strip())
    texto = re.sub(r'\s+', '_', texto)
    return texto.strip('_')


def nombre_pdf(items):
    """Un bien: Acta_codigo_dd_mm_aaaa.pdf | Varios: Acta_texto_dd_mm_aaaa.pdf"""
    fecha = datetime.now().strftime("%d_%m_%Y")
    if len(items) == 1:
        base = limpiar_nombre_archivo(items[0]["codigo"]) or "Bien"
    else:
        base = limpiar_nombre_archivo(st.session_state.get("nombre_acta", ""))
        if not base:
            base = f"{len(items)}_bienes"
    return f"Acta_{base}_{fecha}.pdf"


def normalizar_cedula(valor):
    """Deja solo dígitos y completa con ceros al inicio hasta 10 dígitos (RUC de 13 no se toca)."""
    s = "" if valor is None else str(valor).strip()
    if s.lower() in ["nan", "none", "", "nat", "null", "-"]:
        return ""
    if re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    d = re.sub(r'\D', '', s)
    if not d:
        return ""
    if len(d) < 10:
        d = d.zfill(10)
    return d


def sin_tildes(texto):
    t = unicodedata.normalize('NFD', str(texto))
    return ''.join(c for c in t if unicodedata.category(c) != 'Mn').upper()


def lista_custodios(df_, col_custodio, col_cedula):
    """Lista única de custodios {nombre, cedula} tomada del reporte."""
    if not col_custodio or col_custodio not in df_.columns:
        return []
    columnas = [col_custodio] + ([col_cedula] if col_cedula and col_cedula in df_.columns else [])
    unicos = df_[columnas].drop_duplicates()
    salida, vistos = [], set()
    for _, fila_ in unicos.iterrows():
        nombre = limpiar_val(fila_[col_custodio])
        if nombre == "-":
            continue
        cedula = normalizar_cedula(fila_[col_cedula]) if len(columnas) > 1 else ""
        clave = (nombre, cedula)
        if clave in vistos:
            continue
        vistos.add(clave)
        salida.append({"nombre": nombre, "cedula": cedula})
    return salida


def buscar_custodios(custodios, termino):
    """Busca por cédula (completando el 0 inicial) o por apellidos/nombres (sin tildes)."""
    t = str(termino).strip()
    if not t:
        return []
    if re.fullmatch(r'[\d\s.\-]+', t):
        d = normalizar_cedula(t)
        res = [c for c in custodios if c["cedula"] and c["cedula"] == d]
        crudo = re.sub(r'\D', '', t)
        if not res and len(crudo) >= 5:
            res = [c for c in custodios if crudo in c["cedula"]]
    else:
        palabras = sin_tildes(t).split()
        res = [c for c in custodios if all(w in sin_tildes(c["nombre"]) for w in palabras)]
    return sorted(res, key=lambda c: c["nombre"])[:15]


ARCHIVO_CLAVE = "clave_config.json"


def clave_secreta():
    """Clave original guardada en Secrets de Streamlit (clave = "1234")."""
    try:
        valor = str(st.secrets["clave"]).strip()
    except Exception:
        return None
    return valor or None


def clave_base():
    return clave_secreta() or "1234"


def _huella(texto):
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def _hash_clave(clave, sal_hex):
    return hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(sal_hex), 120000).hex()


def verificar_clave(ingresada):
    """La clave cambiada desde la aplicación vale mientras no se modifique la de Secrets."""
    ingresada = str(ingresada)
    base = clave_base()
    try:
        if os.path.exists(ARCHIVO_CLAVE):
            with open(ARCHIVO_CLAVE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            if cfg.get("origen") == _huella(base):
                return hmac.compare_digest(_hash_clave(ingresada, cfg["sal"]), cfg["hash"])
    except Exception:
        pass
    return hmac.compare_digest(ingresada.encode("utf-8"), base.encode("utf-8"))


def guardar_clave_nueva(nueva):
    """Guarda la clave cifrada (no en texto visible)."""
    try:
        sal = os.urandom(16).hex()
        cfg = {"sal": sal, "hash": _hash_clave(nueva, sal), "origen": _huella(clave_base())}
        with open(ARCHIVO_CLAVE, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
        return True
    except Exception:
        return False


# ----------------------------------------------------------------------
# PDF con numeración de páginas centrada ("Página X de Y")
# ----------------------------------------------------------------------
class NumberedCanvas(rl_canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._paginas_guardadas = []

    def showPage(self):
        self._paginas_guardadas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._paginas_guardadas)
        for estado in self._paginas_guardadas:
            self.__dict__.update(estado)
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#333333"))
            self.drawCentredString(self._pagesize[0] / 2.0, 0.22 * inch, f"Página {self._pageNumber} de {total}")
            super().showPage()
        super().save()


def generar_acta_pdf(rows, fecha_acta_str, col_map, receptor=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.5 * inch
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
        Paragraph(p(receptor["nombre"]) if receptor else "...................................................", subtitle_style_center),
        Paragraph(f"C.I.: {p(receptor['cedula'])}" if receptor and receptor.get("cedula") else "C.I.: ...........................................................", subtitle_style_center),
        Spacer(1, 4),
        Paragraph("<b>QUIEN RECIBE</b>", subtitle_style_center)
    ]
    firmas_table = Table([[cell_entrega, cell_recibe]], colWidths=[370, 370])
    firmas_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(KeepTogether([firmas_table]))

    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------
# Estado de la sesión: bienes agregados a la lista
# ----------------------------------------------------------------------
if "acta_items" not in st.session_state:
    st.session_state.acta_items = []
for _clave, _defecto in [("receptor", None), ("receptor_res", []), ("receptor_busqueda", ""), ("receptor_msg", "")]:
    if _clave not in st.session_state:
        st.session_state[_clave] = _defecto


def reiniciar_receptor():
    st.session_state.receptor = None
    st.session_state.receptor_res = []
    st.session_state.receptor_busqueda = ""
    st.session_state.receptor_msg = ""
    st.session_state.pop("receptor_radio", None)
    for k in ["q_receptor", "rec_nombre_manual", "rec_cedula_manual"]:
        if k in st.session_state:
            st.session_state[k] = ""


def agregar_a_acta(row_dict, codigo):
    codigos = [i["codigo"] for i in st.session_state.acta_items]
    if codigo not in codigos:
        st.session_state.acta_items.append({"codigo": codigo, "row": row_dict})


def quitar_de_acta(indice):
    if 0 <= indice < len(st.session_state.acta_items):
        st.session_state.acta_items.pop(indice)


def vaciar_acta():
    st.session_state.acta_items = []
    st.session_state.nombre_acta = ""
    st.session_state.nombre_acta_w = ""
    reiniciar_receptor()


# ----------------------------------------------------------------------
# Ingreso con clave (estilo ESBYE)
# ----------------------------------------------------------------------
for _clave, _defecto in [("autenticado", False), ("intentos", 0), ("bloqueo_hasta", 0.0)]:
    if _clave not in st.session_state:
        st.session_state[_clave] = _defecto

ENCABEZADO = (
    '<div class="nav">BIENES Y BODEGAS</div>'
    '<div class="dev">Desarrollado por: Ing. William Cisneros R.</div>'
    '<div class="red"></div>'
)

# Script que traduce al español los textos fijos en inglés de Streamlit
# (ayuda del campo de clave y cargador de archivos), conservando el contador (0/4) y el formato.
# Se vuelve a aplicar cada vez que Streamlit actualiza la pantalla.
TRADUCIR_AYUDA_JS = """
<script>
(function () {
    var doc = window.parent.document;
    var PARES = [
        ["Press Enter to submit form", "Presione Enter para ingresar"],
        ["Drag and drop files here", "Arrastre el archivo aquí"],
        ["Drag and drop file here", "Arrastre el archivo aquí"],
        ["Browse files", "Seleccionar archivo"],
        ["Browse file", "Seleccionar archivo"],
        ["Limit ", "Límite "],
        [" per file", " por archivo"]
    ];
    var SEL = '[data-testid="InputInstructions"], [data-testid="stFileUploaderDropzone"]';

    function traducir() {
        doc.querySelectorAll(SEL).forEach(function (nodo) {
            var walker = doc.createTreeWalker(nodo, NodeFilter.SHOW_TEXT, null);
            var t;
            while ((t = walker.nextNode())) {
                var v = t.nodeValue;
                if (!v) continue;
                var nuevo = v;
                PARES.forEach(function (par) { nuevo = nuevo.split(par[0]).join(par[1]); });
                if (nuevo !== v) t.nodeValue = nuevo;
            }
        });
    }

    traducir();
    if (!window.parent.__traductorEnter) {
        window.parent.__traductorEnter = new MutationObserver(traducir);
        window.parent.__traductorEnter.observe(doc.body, {childList: true, subtree: true, characterData: true});
    }
})();
</script>
"""

if not st.session_state.autenticado:
    st.markdown(ENCABEZADO, unsafe_allow_html=True)
    with st.form("form_login", clear_on_submit=True):
        st.markdown('<div class="login-t">Datos del usuario</div><div class="login-l">Clave:</div>', unsafe_allow_html=True)
        clave_ingresada = st.text_input("Clave", type="password", max_chars=4, placeholder="••••", label_visibility="collapsed")
        entrar = st.form_submit_button("Ingresar", use_container_width=True)
    st.markdown('<div class="nota">Clave numérica de 4 dígitos</div>', unsafe_allow_html=True)
    components.html(TRADUCIR_AYUDA_JS, height=0)

    if entrar:
        ahora = time.time()
        if ahora < st.session_state.bloqueo_hasta:
            falta = int(st.session_state.bloqueo_hasta - ahora) + 1
            st.error(f"Demasiados intentos. Espere {falta} segundos.")
        elif re.fullmatch(r"\d{4}", clave_ingresada or "") and verificar_clave(clave_ingresada):
            st.session_state.autenticado = True
            st.session_state.intentos = 0
            st.rerun()
        else:
            st.session_state.intentos += 1
            if st.session_state.intentos >= 5:
                st.session_state.bloqueo_hasta = ahora + 60
                st.session_state.intentos = 0
                st.error("Clave incorrecta. Demasiados intentos: espere 60 segundos.")
            else:
                st.error("Clave incorrecta.")
    st.stop()


def cerrar_sesion_cb():
    st.session_state.autenticado = False
    st.session_state.acta_items = []
    reiniciar_receptor()


def cambiar_clave_cb():
    actual = st.session_state.get("cc_actual", "")
    nueva = st.session_state.get("cc_nueva", "")
    repetida = st.session_state.get("cc_rep", "")
    if not verificar_clave(actual):
        msg = ("error", "La clave actual no es correcta.")
    elif not re.fullmatch(r"\d{4}", nueva):
        msg = ("error", "La clave nueva debe tener exactamente 4 dígitos numéricos.")
    elif nueva != repetida:
        msg = ("error", "Las dos claves nuevas no coinciden.")
    elif nueva == actual:
        msg = ("error", "La clave nueva debe ser diferente a la actual.")
    elif guardar_clave_nueva(nueva):
        msg = ("ok", "Clave cambiada correctamente.")
    else:
        msg = ("error", "No se pudo guardar la clave nueva.")
    for k in ["cc_actual", "cc_nueva", "cc_rep"]:
        st.session_state[k] = ""
    st.session_state.msg_clave = msg


# ----------------------------------------------------------------------
# Barra lateral
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Opciones")
    archivo_subido = st.file_uploader("Subir nuevo reporte (.xlsx)", type=["xlsx", "xls"])

    if archivo_subido is not None:
        firma_archivo = f"{archivo_subido.name}-{archivo_subido.size}"
        # Solo se guarda una vez por archivo; así no se borra la lista en cada clic
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
                st.session_state.acta_items = []  # un reporte nuevo reinicia la lista
                reiniciar_receptor()
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

    with st.expander("🔑 Cambiar clave"):
        st.text_input("Clave actual", type="password", max_chars=4, key="cc_actual", placeholder="Clave actual", label_visibility="collapsed")
        st.text_input("Clave nueva", type="password", max_chars=4, key="cc_nueva", placeholder="Clave nueva (4 dígitos)", label_visibility="collapsed")
        st.text_input("Repita la clave nueva", type="password", max_chars=4, key="cc_rep", placeholder="Repita la clave nueva", label_visibility="collapsed")
        st.button("Cambiar clave", on_click=cambiar_clave_cb, use_container_width=True, key="btn_cambiar_clave")
        msg_clave = st.session_state.pop("msg_clave", None)
        if msg_clave:
            (st.success if msg_clave[0] == "ok" else st.error)(msg_clave[1])
        if clave_secreta() is None:
            st.warning("Se está usando la clave temporal. Configure 'clave' en Secrets de Streamlit.")

    st.button("🚪 Cerrar sesión", on_click=cerrar_sesion_cb, use_container_width=True, key="btn_salir")

# ----------------------------------------------------------------------
# Carga de datos
# ----------------------------------------------------------------------
df = None
display_reporte = "Ninguno"

# Orden de búsqueda del reporte:
# 1) reporte_activo.xlsx  -> el que se sube desde la barra lateral (se pierde si el servidor se reinicia)
# 2) reporte_base.xlsx    -> copia de respaldo guardada en el repositorio de GitHub (no se pierde)
for _ruta in ["reporte_activo.xlsx", "reporte_base.xlsx"]:
    if os.path.exists(_ruta):
        _df = cargar_excel(_ruta)
        if _df is not None:
            df = _df
            display_reporte = _ruta
            if _ruta == "reporte_activo.xlsx" and os.path.exists("nombre_reporte.txt"):
                try:
                    with open("nombre_reporte.txt", "r", encoding="utf-8") as f:
                        display_reporte = f.read().strip() or _ruta
                except Exception:
                    pass
            else:
                display_reporte = "reporte_base.xlsx (respaldo)"
            break

fecha_acta_actual = datetime.now().strftime("%d/%m/%Y")
components.html(TRADUCIR_AYUDA_JS, height=0)   # traduce también el cargador de archivos tras el ingreso

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
    custodios = lista_custodios(df, col_map.get('custodio'), col_map.get('cedula'))

    s_codigo_str = df[c_codigo].astype(str).str.strip() if c_codigo in df.columns else pd.Series([""] * len(df))
    s_serie_str = df[c_serie].astype(str).str.strip() if c_serie in df.columns else pd.Series([""] * len(df))
    s_anterior_str = df[c_anterior].astype(str).str.strip() if c_anterior in df.columns else pd.Series([""] * len(df))

    def encontrar_bien(termino):
        """Devuelve la fila del bien (exacta primero, parcial después) o None."""
        t = str(termino).strip()
        if not t:
            return None
        res_ = df[
            (s_codigo_str.str.lower() == t.lower()) |
            (s_serie_str.str.lower() == t.lower()) |
            (s_anterior_str.str.lower() == t.lower())
        ]
        if res_.empty:
            res_ = df[
                (s_codigo_str.str.contains(t, case=False, na=False, regex=False)) |
                (s_serie_str.str.contains(t, case=False, na=False, regex=False))
            ]
        return None if res_.empty else res_.iloc[0]

    if "busqueda_val" not in st.session_state:
        st.session_state.busqueda_val = ""

    def al_buscar_cb():
        termino = st.session_state.get("mi_buscador", "").strip()
        st.session_state.busqueda_val = termino
        if termino and encontrar_bien(termino) is not None:
            st.session_state.pestana = "Bien"   # salta sola a la pestaña Bien

    def limpiar_callback():
        st.session_state.busqueda_val = ""
        st.session_state.mi_buscador = ""
        st.session_state.foto_msg = None

    # ---- Búsqueda por foto (lee el código escrito/impreso en el bien con Claude) ----
    MODELO_OCR = "claude-sonnet-5-5"
    PROMPT_OCR = (
        "En la foto hay un código de bien (inventario) escrito a mano o impreso sobre una superficie, "
        "por ejemplo un mueble o equipo. Lea el código numérico completo. "
        "Si hay varias secuencias de números, devuélvalas todas separadas por comas, la más visible primero. "
        "Responda solo con los números, sin explicación. Si no se ve ningún código, responda NINGUNO."
    )

    def leer_codigo_foto(archivo):
        """Devuelve (lista_de_codigos, mensaje_de_error)."""
        try:
            clave_api = st.secrets["ANTHROPIC_API_KEY"]
        except Exception:
            return [], "Falta configurar ANTHROPIC_API_KEY en Secrets de Streamlit."
        try:
            import anthropic
            img = PILImage.open(archivo).convert("RGB")
            img.thumbnail((1600, 1600))
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=88)
            b64 = base64.b64encode(buf.getvalue()).decode()
            cliente = anthropic.Anthropic(api_key=clave_api)
            resp = cliente.messages.create(
                model=MODELO_OCR,
                max_tokens=100,
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                    {"type": "text", "text": PROMPT_OCR},
                ]}],
            )
            texto = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
            return re.findall(r"\d{5,}", texto), None
        except Exception as e:
            return [], f"No se pudo leer la foto: {str(e)[:120]}"

    def leer_foto_cb():
        n = st.session_state.get("foto_n", 0)
        archivo = st.session_state.get(f"foto_busqueda_{n}")
        st.session_state.foto_n = n + 1          # nuevo cargador vacío para la próxima foto
        if archivo is None:
            return
        codigos, error = leer_codigo_foto(archivo)
        if error:
            st.session_state.foto_msg = ("err", error)
            return
        if not codigos:
            st.session_state.foto_msg = ("err", "No se detectó ningún código en la foto. Acérquese, enfoque bien y evite reflejos.")
            return
        for c in codigos:
            if encontrar_bien(c) is not None:
                st.session_state.mi_buscador = c
                st.session_state.busqueda_val = c
                st.session_state.pestana = "Bien"
                st.session_state.foto_msg = ("ok", f"Código leído de la foto: {c}")
                return
        st.session_state.mi_buscador = codigos[0]
        st.session_state.busqueda_val = codigos[0]
        st.session_state.foto_msg = ("err", f"Se leyó {codigos[0]}, pero no existe en el reporte. Si un dígito está mal, corríjalo en el cuadro y pulse Enter.")

    termino_busqueda = st.session_state.busqueda_val.strip()
    fila_bien = encontrar_bien(termino_busqueda) if termino_busqueda else None

    bien_actual = None
    if fila_bien is not None:
        bien_actual = {
            "row": fila_bien.to_dict(),
            "codigo": limpiar_val(fila_bien.get(c_codigo, '-')),
            "nombre": limpiar_val(fila_bien.get(col_map.get('nombre'), '-')) if col_map.get('nombre') else '-',
            "bodega": limpiar_val(fila_bien.get(col_map.get('bodega'), '-')) if col_map.get('bodega') else '-',
        }

    items = st.session_state.acta_items

    # Encabezado y pestañas
    st.markdown(ENCABEZADO, unsafe_allow_html=True)
    if "pestana" not in st.session_state:
        st.session_state.pestana = "Buscar"
    st.radio(
        "Secciones", ["Buscar", "Bien", "Custodio", "PDF"], key="pestana", horizontal=True,
        label_visibility="collapsed",
        format_func=lambda k: (f"Asignar / Reporte ({len(items)})" if items else "Asignar / Reporte") if k == "PDF" else k
    )
    pestana = st.session_state.pestana

    def aviso_vacio(texto):
        st.markdown(f'<div class="nota" style="margin-top:14px">{hesc(texto)}</div>', unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # Pestaña 1: Buscar
    # ------------------------------------------------------------------
    if pestana == "Buscar":
        st.markdown(fila("Reporte activo", display_reporte) + '<div class="bar">Código a buscar</div>', unsafe_allow_html=True)

        if "mi_buscador" not in st.session_state:
            st.session_state.mi_buscador = st.session_state.busqueda_val

        col_input, col_btn = st.columns([4.2, 0.8])
        with col_input:
            st.text_input("Código o serie", key="mi_buscador", placeholder="Ingrese código o serie...",
                          label_visibility="collapsed", on_change=al_buscar_cb)
        with col_btn:
            st.button("🗑️", on_click=limpiar_callback, help="Limpiar búsqueda", use_container_width=True)

        st.markdown('<div class="t">O lea el código con una foto (elija la cámara del celular):</div>', unsafe_allow_html=True)
        st.file_uploader("Foto del código", type=["jpg", "jpeg", "png"],
                         key=f"foto_busqueda_{st.session_state.get('foto_n', 0)}",
                         on_change=leer_foto_cb, label_visibility="collapsed")
        _fm = st.session_state.get("foto_msg")
        if _fm:
            if _fm[0] == "ok":
                st.markdown(f'<div class="nota"><b>{hesc(_fm[1])}</b></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="err">{hesc(_fm[1])}</div>', unsafe_allow_html=True)

        if bien_actual:
            st.markdown(
                '<div style="height:6px"></div>' + fila("Bien", bien_actual["nombre"]) + fila("Bodega", bien_actual["bodega"])
                + '<div class="nota">Vea los datos completos en la pestaña "Bien".</div>',
                unsafe_allow_html=True
            )
        elif termino_busqueda:
            st.error(f"❌ No se encontró ningún bien con el código: {termino_busqueda}")
        else:
            aviso_vacio("Escriba el código o la serie del bien y pulse Enter.")

    # ------------------------------------------------------------------
    # Pestaña 2: Bien
    # ------------------------------------------------------------------
    elif pestana == "Bien":
        if not bien_actual:
            aviso_vacio('Primero busque un bien en la pestaña "Buscar".')
        else:
            row = fila_bien
            val_fecha_ingreso = limpiar_val_fecha(row.get(col_map.get('fecha_ingreso'), '-'))
            num_valor = valor_numerico(row.get(col_map.get('valor'), '-')) if col_map.get('valor') else None
            val_contable_fmt = formato_dinero(num_valor) if num_valor is not None else limpiar_val(row.get(col_map.get('valor'), '-'))

            st.markdown(
                '<div style="height:6px"></div>'
                + fila("Bien", bien_actual["nombre"])
                + fila("Bodega", bien_actual["bodega"])
                + fila("Código del bien", bien_actual["codigo"], True)
                + fila("Código anterior", limpiar_val(row.get(c_anterior, '-')))
                + fila("Marca / otros", limpiar_val(row.get(col_map.get('marca'), '-')))
                + fila("Modelo / características", limpiar_val(row.get(col_map.get('modelo'), '-')))
                + fila("Serie / identificación", limpiar_val(row.get(c_serie, '-')))
                + fila("Fecha de ingreso", val_fecha_ingreso)
                + fila("Valor contable", val_contable_fmt, True)
                + '<div class="red" style="margin-bottom:10px"></div>',
                unsafe_allow_html=True
            )

            ya_agregado = bien_actual["codigo"] in [i["codigo"] for i in items]
            if ya_agregado:
                st.button("✅ Este bien ya está en la lista", disabled=True, use_container_width=True, key="btn_ya_agregado")
            else:
                st.button(
                    "➕ Agregar a la lista de varios bienes",
                    on_click=agregar_a_acta,
                    args=(bien_actual["row"], bien_actual["codigo"]),
                    use_container_width=True,
                    key="btn_agregar"
                )
            if items:
                st.markdown(f'<div class="nota">En la lista: {len(items)} bien(es) (se ven en la pestaña "Asignar / Reporte")</div>', unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # Pestaña 3: Custodio
    # ------------------------------------------------------------------
    elif pestana == "Custodio":
        if not bien_actual:
            aviso_vacio('Primero busque un bien en la pestaña "Buscar".')
        else:
            row = fila_bien
            st.markdown(
                '<div style="height:6px"></div>'
                + fila("Custodio actual", limpiar_val(row.get(col_map.get('custodio'), '-')))
                + fila("Nro de cédula / RUC", normalizar_cedula(row.get(col_map.get('cedula'), '')) or '-'),
                unsafe_allow_html=True
            )

    # ------------------------------------------------------------------
    # Pestaña 4: PDF
    # ------------------------------------------------------------------
    else:
        if not bien_actual and not items:
            aviso_vacio('Busque un bien o agregue bienes a la lista para generar el acta.')
        else:
            if bien_actual:
                st.markdown(
                    '<div style="height:6px"></div>' + fila("Bien consultado", f'{bien_actual["codigo"]} - {bien_actual["nombre"]}'),
                    unsafe_allow_html=True
                )

            # Lista de varios bienes
            if items:
                st.markdown('<div class="bar" style="margin-top:12px">Lista para acta</div>', unsafe_allow_html=True)
                for i, it in enumerate(items):
                    c_item, c_del = st.columns([4.2, 0.8])
                    with c_item:
                        st.markdown(f"<div class='acta-item'><b>{i + 1}.</b> {hesc(str(it['codigo']))}</div>", unsafe_allow_html=True)
                    with c_del:
                        st.button("❌", key=f"quitar_{i}_{it['codigo']}", on_click=quitar_de_acta, args=(i,), help="Quitar de la lista", use_container_width=True)

                conteo_bodegas = {}
                for it in items:
                    b = limpiar_val(it["row"].get(col_map.get('bodega'), '-')) if col_map.get('bodega') else '-'
                    conteo_bodegas[b] = conteo_bodegas.get(b, 0) + 1
                if len(conteo_bodegas) > 1:
                    detalle = "".join(f"<br>{hesc(b)}: {n} bien(es)" for b, n in conteo_bodegas.items())
                    st.markdown(
                        f"<div class='aviso'><b>Atención:</b> la lista tiene bienes de diferentes bodegas.{detalle}</div>",
                        unsafe_allow_html=True
                    )

            # Quien recibe
            def buscar_receptor_cb():
                termino = st.session_state.get("q_receptor", "").strip()
                st.session_state.receptor = None
                st.session_state.receptor_msg = ""
                st.session_state.pop("receptor_radio", None)
                encontrados = buscar_custodios(custodios, termino) if termino else []
                st.session_state.receptor_res = encontrados
                st.session_state.receptor_busqueda = termino
                if len(encontrados) == 1:
                    st.session_state.receptor = encontrados[0]

            def usar_manual_cb():
                nombre = " ".join(st.session_state.get("rec_nombre_manual", "").split()).upper()
                if not nombre:
                    st.session_state.receptor_msg = "Escriba los apellidos y nombres."
                    return
                st.session_state.receptor_msg = ""
                st.session_state.receptor = {
                    "nombre": nombre,
                    "cedula": normalizar_cedula(st.session_state.get("rec_cedula_manual", ""))
                }

            receptor = st.session_state.receptor

            if receptor:
                st.markdown('<div class="bar" style="margin-top:14px">Quien recibe</div>', unsafe_allow_html=True)
                st.markdown(fila("Nombre", receptor["nombre"]) + fila("C.I.", receptor["cedula"] or "-"), unsafe_allow_html=True)
                st.button("Cambiar", on_click=reiniciar_receptor, use_container_width=True, key="btn_cambiar_rec")
            else:
                st.markdown('<div class="bar" style="margin-top:14px">¿A quién se asigna?</div>', unsafe_allow_html=True)
                st.markdown('<div class="t">Escriba número de cédula o apellido:</div>', unsafe_allow_html=True)
                c_q, c_b = st.columns([4.2, 0.8])
                with c_q:
                    st.text_input("Cédula o apellido", key="q_receptor", placeholder="Cédula o apellido...",
                                  on_change=buscar_receptor_cb, label_visibility="collapsed")
                with c_b:
                    st.button("🔍", on_click=buscar_receptor_cb, help="Buscar", use_container_width=True, key="btn_buscar_rec")

                busqueda = st.session_state.receptor_busqueda
                encontrados = st.session_state.receptor_res
                if busqueda and encontrados:
                    st.markdown(f'<div class="t">{len(encontrados)} coincidencias. Elija una:</div>', unsafe_allow_html=True)
                    etiquetas = [f'{c["nombre"]} - {c["cedula"] or "sin cédula"}' for c in encontrados]
                    elegido = st.radio("Coincidencias", etiquetas, index=None, key="receptor_radio", label_visibility="collapsed")
                    if elegido is not None:
                        st.session_state.receptor = encontrados[etiquetas.index(elegido)]
                        st.rerun()
                elif busqueda:
                    st.markdown(
                        f'<div class="err"><b>Usuario no existe.</b><br>No se encontró "{hesc(busqueda)}" en el reporte.</div>',
                        unsafe_allow_html=True
                    )
                    st.button("Buscar nuevamente", on_click=reiniciar_receptor, use_container_width=True, key="btn_buscar_nuevo")
                    st.markdown('<div class="bar" style="margin-top:14px">O digite los datos manualmente</div>', unsafe_allow_html=True)
                    st.markdown('<div class="t">Apellidos y nombres (obligatorio):</div>', unsafe_allow_html=True)
                    st.text_input("Apellidos y nombres", key="rec_nombre_manual", placeholder="Ej: ZAMBRANO CEDEÑO LUIS", label_visibility="collapsed")
                    st.markdown('<div class="t">Cédula (opcional; si escribe 9 dígitos se agrega el 0 inicial):</div>', unsafe_allow_html=True)
                    st.text_input("Cédula", key="rec_cedula_manual", placeholder="Ej: 0912345678", label_visibility="collapsed")
                    if st.session_state.receptor_msg:
                        st.markdown(f'<div class="err">{hesc(st.session_state.receptor_msg)}</div>', unsafe_allow_html=True)
                    st.button("Usar estos datos", on_click=usar_manual_cb, use_container_width=True, key="btn_manual")

            receptor = st.session_state.receptor
            if not receptor:
                st.markdown('<div class="nota">Indique a quién se asigna para habilitar el PDF.</div>', unsafe_allow_html=True)

            # PDF de este bien: Acta_codigo_dd_mm_aaaa.pdf
            if bien_actual:
                nombre_individual = nombre_pdf([{"codigo": bien_actual["codigo"]}])
                datos_ind = generar_acta_pdf([bien_actual["row"]], fecha_acta_actual, col_map, receptor) if receptor else b""
                st.download_button(
                    label="📄 Generar e imprimir Acta PDF oficial" if not items else "📄 Acta PDF solo del bien consultado",
                    data=datos_ind,
                    file_name=nombre_individual,
                    mime="application/pdf",
                    use_container_width=True,
                    disabled=not receptor,
                    key="dl_individual"
                )
                st.markdown(f'<div class="nota">Se guarda como: {hesc(nombre_individual)}</div>', unsafe_allow_html=True)

            # PDF de varios bienes: Acta_texto_dd_mm_aaaa.pdf
            if items:
                if "nombre_acta" not in st.session_state:
                    st.session_state.nombre_acta = ""

                if len(items) >= 2:
                    def guardar_nombre_acta_cb():
                        st.session_state.nombre_acta = st.session_state.get("nombre_acta_w", "")

                    if "nombre_acta_w" not in st.session_state:
                        st.session_state.nombre_acta_w = st.session_state.nombre_acta
                    st.markdown('<div class="bar" style="margin-top:14px">Nombre del archivo PDF</div>', unsafe_allow_html=True)
                    st.markdown('<div class="nota" style="text-align:left">Escriba el nombre del acta y pulse Enter (se guardará como Acta_<b>su texto</b>_fecha):</div>', unsafe_allow_html=True)
                    st.text_input("Nombre del acta", key="nombre_acta_w", placeholder="Ej: Entrega_Bodega_Diseño",
                                  on_change=guardar_nombre_acta_cb, label_visibility="collapsed")

                nombre_grupal = nombre_pdf(items)
                st.markdown(f'<div class="nota">Se guarda como: {hesc(nombre_grupal)}</div>', unsafe_allow_html=True)

                datos_grupal = generar_acta_pdf([it["row"] for it in items], fecha_acta_actual, col_map, receptor) if receptor else b""
                st.download_button(
                    label=f"📄 Generar Acta PDF con los {len(items)} bienes de la lista" if len(items) > 1 else "📄 Generar Acta PDF con el bien de la lista",
                    data=datos_grupal,
                    file_name=nombre_grupal,
                    mime="application/pdf",
                    use_container_width=True,
                    disabled=not receptor,
                    key="dl_grupal"
                )
                st.button("🧹 Vaciar lista", on_click=vaciar_acta, use_container_width=True, key="btn_vaciar")

    # Espacio libre amplio en la parte inferior para celulares
    st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)

else:
    st.warning("⚠️ No se encontró ningún archivo Excel activo en el servidor. Por favor, suba su reporte institucional en formato .xlsx mediante el menú de la barra lateral izquierda.")



