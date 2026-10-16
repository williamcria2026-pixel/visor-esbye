import streamlit as st
import pandas as pd
import os
import time

st.set_page_config(page_title="Visor ESBYE", page_icon="📦", layout="centered")

# Encabezado
st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>CONSULTA RÁPIDA DE BIENES Y BODEGAS</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #1D4ED8; font-style: italic;'>Desarrollado para: Ing. William Cisneros</p>", unsafe_allow_html=True)
st.divider()

# Panel Lateral para Actualizar el Archivo Excel
with st.sidebar:
    st.header("⚙️ Opciones del Sistema")
    st.write("Actualice el reporte de ESBYE cargando un archivo nuevo aquí:")
    
    archivo_subido = st.file_uploader("Subir nuevo reporte (.xlsx)", type=["xlsx", "xls"])
    
    st.divider()
    st.markdown("**Nota:** Si no sube ningún archivo nuevo, el sistema usará por defecto el último reporte guardado en la carpeta.")

# Definir la fuente de los datos (Prioriza el archivo subido por el usuario, luego busca el local)
if archivo_subido is not None:
    df = pd.read_excel(archivo_subido)
    nombre_fuente = archivo_subido.name
    # Fecha actual o simulación de actualización
    fecha_mod = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime())
    st.sidebar.success("🟢 Archivo actualizado desde la web con éxito.")
elif os.path.exists("reporte.xlsx"):
    df = pd.read_excel("reporte.xlsx")
    nombre_fuente = "reporte.xlsx"
    timestamp = os.path.getmtime("reporte.xlsx")
    fecha_mod = time.strftime('%d/%m/%Y %H:%M:%S', time.localtime(timestamp))
else:
    df = None
    nombre_fuente = None

if df is not None:
    # Limpiar espacios en columnas y datos
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    st.caption(f"🟢 **Reporte Activo:** {nombre_fuente} ({len(df):,} registros) | 📅 **Actualizado:** {fecha_mod}")
    
    # Buscador principal
    codigo = st.text_input("🔍 Ingrese o escanee el Código del Bien / Serie:", "").strip()
    
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
            st.success(f"📌 **BIEN ENCONTRADO:** {row.get('Bien', '-')}")
            
            # Tarjeta de resultados
            st.info(f"🏢 **UBICACIÓN DE BODEGA:**\n### {row.get('Ubicación de Bodega', '-')}")
            
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Custodio:** {row.get('Custodio Actual', '-')}")
                st.write(f"**Código bien:** {row.get('Código del Bien', '-')}")
                st.write(f"**Serie:** {row.get('Serie/ Identificación', '-')}")
                st.write(f"**Marca:** {row.get('Marca/ Otros', '-')}")
            with col2:
                st.write(f"**Cédula:** {row.get('Nro de Cédula/ RUC', '-')}")
                st.write(f"**Código Ant.:** {row.get('Código Anterior', '-')}")
                st.write(f"**Modelo:** {row.get('Modelo/ Características', '-')}")
                
                val_contable = row.get('Valor Contable', '-')
                if val_contable not in ["nan", "None", "-", "."]:
                    try:
                        val_contable = f"${float(val_contable):,.2f}"
                    except:
                        pass
                st.write(f"**Valor Contable:** {val_contable}")
        else:
            st.error(f"❌ No se encontró ningún bien con el código o serie: {codigo}")
else:
    st.warning("⚠️ No se encontró ningún archivo de reporte. Utilice el panel lateral izquierdo para subir su archivo .xlsx.")