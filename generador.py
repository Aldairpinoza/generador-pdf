import streamlit as st
import pandas as pd
import os
import base64
from weasyprint import HTML
import datetime
import io
import openpyxl

# --- 1. CONFIGURACIÓN DE CLIENTES ---
CONFIG_CLIENTES = {
    "Carl's Jr": {
        "logo": "logos/CJ.png", 
        "color_principal": "#E31837", 
        "color_secundario": "#FFC82C", 
        "logo_height": "140px" 
    },
    "Dalton": {
        "logo": "logos/dalton.png", 
        "color_principal": "#515151", 
        "color_secundario": "#666666",
        "logo_height": "140px" 
    },
    "Omnilife": {
        "logo": "logos/omnilife.png", 
        "color_principal": "#7C3D70", 
        "color_secundario": "#F29400",
        "logo_height": "140px" 
    },
    "Radial Llantas": {
        "logo": "logos/radial.png", 
        "color_principal": "#FFD200", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#000000", 
        "logo_height": "140px", 
        "texto_oscuro": True, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #FFD200 0%, #FFFFFF 100%)"
    },
    "OMNIDATA": {
        "logo": "logos/omnidata.png", 
        "color_principal": "#0056b3", 
        "color_secundario": "#003d82",
        "logo_height": "140px" 
    },
    "Andares": {
        "logo": "logos/andares.png", 
        "color_principal": "#8c8c8c", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#1a1a1a", 
        "logo_height": "140px",
        "texto_oscuro": True, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #8c8c8c 0%, #FFFFFF 100%)"
    },
    "Liverpool": {
        "logo": "logos/liverpool.png", 
        "color_principal": "#E2007A", 
        "color_secundario": "#FFFFFF", 
        "logo_height": "140px", 
        "fecha_oscura": True 
    },
    "Popeyes": {
        "logo": "logos/popeyes.png", 
        "color_principal": "#FF7D00", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#FF7D00", 
        "logo_height": "140px", 
        "texto_oscuro": True, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #FF7D00 0%, #FFFFFF 100%)"
    },
    "Bluepoint Solution": {
        "logo": "logos/bluepoint.png", 
        "color_principal": "#007BFF", 
        "color_secundario": "#0056b3",
        "logo_height": "140px" 
    },
    "TLapps": {
        "logo": "logos/tlapps.png", 
        "color_principal": "#6f42c1", 
        "color_secundario": "#5a32a3",
        "logo_height": "140px" 
    },
    "BAIC": {
        "logo": "logos/baic.png", 
        "color_principal": "#DF0026", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#DF0026", 
        "logo_height": "140px", 
        "texto_oscuro": True, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #DF0026 0%, #FFFFFF 100%)"
    },
    "Suzuki": {
        "logo": "logos/suzuki.png", 
        "color_principal": "#0028B3", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#0028B3", 
        "logo_height": "140px", 
        "texto_oscuro": True, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #0028B3 0%, #FFFFFF 100%)"
    },
    "Forthing": {
        "logo": "logos/forthing.png", 
        "color_principal": "#000000", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#000000", 
        "logo_height": "140px", 
        "texto_oscuro": False, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #000000 0%, #FFFFFF 100%)"
    },
    "Taco Bell": {
        "logo": "logos/tacobell.png", 
        "color_principal": "#702082", 
        "color_secundario": "#FFFFFF", 
        "color_tabla": "#702082", 
        "logo_height": "140px", 
        "texto_oscuro": True, 
        "fecha_oscura": True, 
        "header_bg": "linear-gradient(135deg, #702082 0%, #FFFFFF 100%)"
    },
    # ACTUALIZADO: Perfil para PH con Amarillo y Café
    "PH": {
        "logo": "logos/ph.png", 
        "color_principal": "#FFCC00", # Amarillo
        "color_secundario": "#4A2E15", # Café oscuro
        "color_tabla": "#4A2E15", # Tabla en café para contraste
        "logo_height": "140px", 
        "texto_oscuro": True, # Texto oscuro sobre la zona amarilla
        "fecha_oscura": False, # Texto blanco/claro sobre la zona café
        "header_bg": "linear-gradient(135deg, #FFCC00 0%, #4A2E15 100%)"
    }
}

def get_base64_image(file_path):
    if os.path.exists(file_path):
        with open(file_path, "rb") as img_file:
            ext = file_path.split('.')[-1].lower()
            if ext == 'svg':
                mime = "image/svg+xml"
            elif ext in ['jpg', 'jpeg']:
                mime = "image/jpeg"
            else:
                mime = f"image/{ext}"
                
            b64_str = base64.b64encode(img_file.read()).decode('utf-8')
            return f"data:{mime};base64,{b64_str}"
    return ""

st.set_page_config(page_title="Generador de PDFs", page_icon="📄")
st.title("📄 Convertidor a PDF Profesional")
st.markdown("---")

opciones_menu = ["-- Selecciona un cliente --"] + list(CONFIG_CLIENTES.keys())
cliente_seleccionado = st.selectbox("1. Selecciona el cliente:", opciones_menu)

archivo_subido = st.file_uploader("2. Sube el archivo Excel aquí:", type=["xlsx", "xls"])

if st.button("Procesar y Generar PDF", type="primary"):
    if cliente_seleccionado == "-- Selecciona un cliente --":
        st.error("⚠️ Por favor, selecciona un cliente del menú desplegable antes de continuar.")
    elif archivo_subido is None:
        st.error("⚠️ Por favor, sube un archivo Excel.")
    else:
        if archivo_subido.name.lower().endswith('.xls'):
            st.warning("⚠️ **AVISO IMPORTANTE:** Subiste un archivo con formato antiguo (`.xls`). El sistema no podrá detectar los textos de colores. Guarda el archivo como 'Libro de Excel (.xlsx)' para conservar los colores.")
            
        with st.spinner("Aplicando diseño corporativo y leyendo el archivo..."):
            
            datos_cliente = CONFIG_CLIENTES[cliente_seleccionado]
            logo_cliente = get_base64_image(datos_cliente["logo"])
            logo_admira = get_base64_image("logos/admira.svg") 
            
            color_header = datos_cliente["color_principal"]
            color_secundario = datos_cliente["color_secundario"]
            fondo_custom = datos_cliente.get("header_bg", f"linear-gradient(90deg, {color_header} 0%, {color_secundario} 100%)")
            color_tabla = datos_cliente.get("color_tabla", color_header) 
            alto_logo = datos_cliente.get("logo_height", "140px") 
            
            if color_tabla == "#000000" or color_tabla == "#1a1a1a" or color_tabla == "#515151":
                color_alerta_final = "#D97706" 
            else:
                color_alerta_final = color_header 
            
            img_tag_cliente = f'<img src="{logo_cliente}" style="height: {alto_logo}; background: transparent !important; display: block; margin: 0 auto;">' if logo_cliente else ''
            img_tag_admira = f'<img src="{logo_admira}" style="height: 55px; margin-bottom: 2px; background: transparent !important; display: block; margin-left: auto;">' if logo_admira else ''

            file_bytes = archivo_subido.read()
            
            alert_map = set()
            if archivo_subido.name.lower().endswith('.xlsx'):
                try:
                    wb_load = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
                    ws_load = wb_load.active
                    
                    theme_colors = {
                        4: "#4472C4", 5: "#ED7D31", 6: "#A5A5A5", 
                        7: "#FFC000", 8: "#5B9BD5", 9: "#70AD47"  
                    }
                    
                    for r_idx, row in enumerate(ws_load.iter_rows()):
                        for c_idx, cell in enumerate(row):
                            if cell.font and cell.font.color:
                                color_obj = cell.font.color
                                is_alert = False
                                
                                if color_obj.type == 'rgb' and type(color_obj.rgb) == str:
                                    rgb = color_obj.rgb
                                    hex_val = None
                                    if len(rgb) == 8: hex_val = "#" + rgb[2:]
                                    elif len(rgb) == 6: hex_val = "#" + rgb
                                    
                                    if hex_val and hex_val.upper() not in ["#000000", "#FFFFFF", "#00000000"]:
                                        is_alert = True
                                        
                                elif color_obj.type == 'theme':
                                    if color_obj.theme not in [0, 1]:
                                        is_alert = True
                                
                                if is_alert:
                                    alert_map.add((r_idx, c_idx))
                except Exception as e:
                    pass
            
            try:
                df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
                engine_usado = 'xlrd'
            except Exception:
                try:
                    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
                    engine_usado = None
                except Exception:
                    df_raw = pd.read_html(io.BytesIO(file_bytes))[0]
                    engine_usado = 'html'

            header_idx = 0
            for i, row in df_raw.head(20).iterrows():
                if row.astype(str).str.contains('player|nombre|serie', case=False, na=False).any():
                    header_idx = i
                    break
            
            if engine_usado == 'xlrd':
                df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_idx, engine='xlrd', engine_kwargs={'ignore_workbook_corruption': True})
            elif engine_usado == 'html':
                df = pd.read_html(io.BytesIO(file_bytes), header=header_idx)[0]
            else:
                df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_idx)
            
            columnas_originales = list(df.columns)
            df = df.dropna(how='all', axis=1)
            total_pantallas = len(df)
            indices_columnas_reales = [columnas_originales.index(col) for col in df.columns]
            
            encabezados_html = ""
            for col in df.columns:
                if "Unnamed" not in str(col):
                    encabezados_html += f"<th>{str(col).upper()}</th>"
            
            filas_html = ""
            for offset, (df_row_idx, row) in enumerate(df.iterrows()):
                filas_html += "<tr>"
                real_excel_row = header_idx + 1 + offset
                
                for col_loop_idx, col_name in enumerate(df.columns):
                    if "Unnamed" not in str(col_name):
                        valor = str(row.iloc[col_loop_idx]) if pd.notna(row.iloc[col_loop_idx]) else ""
                        real_excel_col = indices_columnas_reales[col_loop_idx]
                        
                        if (real_excel_row, real_excel_col) in alert_map:
                            filas_html += f'<td><span style="color: {color_alerta_final}; font-weight: 700;">{valor}</span></td>'
                        else:
                            filas_html += f"<td>{valor}</td>"
                filas_html += "</tr>"
            
            texto_oscuro = datos_cliente.get("texto_oscuro", False)
            fecha_oscura = datos_cliente.get("fecha_oscura", False)
            
            titulo_color = "#111111" if texto_oscuro else "rgba(255, 255, 255, 0.9)"
            tabla_texto_color = "white"
            
            badge_bg = "rgba(0, 0, 0, 0.06)" if texto_oscuro else "rgba(255, 255, 255, 0.15)"
            badge_border = "rgba(0, 0, 0, 0.15)" if texto_oscuro else "rgba(255, 255, 255, 0.3)"
            
            color_fecha = "#333333" if fecha_oscura else "rgba(255, 255, 255, 0.9)"
            shadow_fecha = "none" if fecha_oscura else "1px 1px 2px rgba(0,0,0,0.5)"
            fecha_actual = datetime.date.today().strftime('%d/%m/%Y')
            
            html_template = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');
                    
                    @page {{ size: A4 landscape; margin: 0; }}
                    
                    body {{ font-family: 'Montserrat', sans-serif; margin: 0; padding: 0; color: #333; background-color: #FAFAFA; }}
                    
                    .header-container {{ 
                        background: {fondo_custom};
                        width: 100%; 
                        padding: 6mm 15mm; 
                        box-sizing: border-box; 
                        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                    }}
                    .header-table {{ width: 100%; border: none; }}
                    .header-table td {{ border: none; padding: 0; color: white; vertical-align: middle; }}
                    
                    .left-block {{
                        display: inline-block;
                        text-align: center;
                    }}
                    
                    .badge-total {{
                        background-color: {badge_bg}; 
                        color: {titulo_color};
                        padding: 4px 12px;
                        border-radius: 12px;
                        font-size: 8.5pt;
                        font-weight: 600;
                        display: inline-block;
                        border: 1px solid {badge_border};
                        letter-spacing: 0.5px;
                    }}
                    
                    .content {{ padding: 6mm 15mm; }}
                    
                    .table-wrapper {{
                        border-radius: 8px;
                        overflow: hidden;
                        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                        background-color: white;
                    }}
                    
                    .data-table {{ 
                        width: 100%; 
                        border-collapse: collapse; 
                        border-style: hidden; 
                    }}
                    
                    .data-table th {{ 
                        background-color: {color_tabla}; 
                        color: {tabla_texto_color}; 
                        padding: 8px 6px; 
                        font-size: 7.5pt; 
                        text-align: left; 
                        font-weight: 700;
                        border-right: 1px solid rgba(255,255,255,0.2); 
                    }}
                    .data-table th:last-child {{ border-right: none; }}
                    
                    .data-table td {{ 
                        padding: 6px; 
                        font-size: 7.5pt; 
                        border: 1px solid #E5E7EB; 
                        color: #4b5563; 
                        font-weight: 400;
                    }}
                    
                    .data-table tr:nth-child(even) td {{ background-color: #F9FAFB; }}
                    .data-table tr:nth-child(odd) td {{ background-color: #FFFFFF; }}
                </style>
            </head>
            <body>
                <div class="header-container">
                    <table class="header-table">
                        <tr>
                            <td style="width: 70%; text-align: left;">
                                <div class="left-block">
                                    {img_tag_cliente}
                                    <div style="margin-top: 10px;"><span class="badge-total">TOTAL DE PANTALLAS: {total_pantallas}</span></div>
                                </div>
                            </td>
                            <td style="width: 30%; text-align: right; vertical-align: top;">
                                {img_tag_admira}
                                <div style="font-size: 9pt; font-weight: 500; color: {color_fecha}; margin-top: 3px; text-shadow: {shadow_fecha};">Fecha de informe: {fecha_actual}</div>
                            </td>
                        </tr>
                    </table>
                </div>
                
                <div class="content">
                    <div class="table-wrapper">
                        <table class="data-table">
                            <thead>
                                <tr>{encabezados_html}</tr>
                            </thead>
                            <tbody>
                                {filas_html}
                            </tbody>
                        </table>
                    </div>
                </div>
            </body>
            </html>
            """
            
            pdf_bytes = io.BytesIO()
            HTML(string=html_template).write_pdf(pdf_bytes)
            
        st.success(f"¡PDF generado con éxito! Se procesaron {total_pantallas} registros.")
        st.download_button(
            label="⬇️ Descargar Reporte PDF",
            data=pdf_bytes.getvalue(),
            file_name=f"Reporte_{cliente_seleccionado}.pdf",
            mime="application/pdf"
        )