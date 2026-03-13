import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
from datetime import datetime
from fpdf import FPDF
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Sophos VPN Reporter", page_icon="🛡️")
st.title("🛡️ Generador de Reportes VPN")

# --- CLASE PDF (ESTÉTICA) ---
class CustomPDF(FPDF):
    def __init__(self, metadata):
        super().__init__()
        self.metadata = metadata

    def header(self):
        # Logo: Intentar cargar logo.jpg
        if os.path.exists("logo.jpg"):
            try:
                self.image("logo.jpg", x=10, y=8, w=30)
            except:
                pass
        
        self.set_y(20)
        self.set_font("helvetica", "B", 26)
        self.cell(0, 15, "System events", align="L", new_x="LMARGIN", new_y="NEXT")
        self.set_font("helvetica", "", 12)
        
        start = str(self.metadata.get('Start Date', ''))
        end = str(self.metadata.get('End Date', ''))
        self.cell(0, 10, f"{start} - {end}", align="L", new_x="LMARGIN", new_y="NEXT")
        self.ln(5)

    def footer(self):
        self.set_y(-25)
        self.set_font("helvetica", "I", 8)
        server_t = str(self.metadata.get('Server Time', ''))
        # Limpieza para evitar errores de caracteres
        clean_footer = f"Server time: {server_t}".encode('ascii', 'ignore').decode('ascii')
        self.cell(0, 10, clean_footer, align='R')

# --- LÓGICA DE PROCESAMIENTO ---
def procesar_datos(uploaded_file):
    try:
        content = uploaded_file.getvalue().decode("utf-8")
    except:
        content = uploaded_file.getvalue().decode("latin-1")
        
    lines = content.splitlines()
    metadata = {}
    
    # Extraer metadatos (Logo/Info superior)
    for line in lines[:15]:
        parts = [p.strip().replace('"', '') for p in line.split(',')]
        if len(parts) >= 2:
            key, val = parts[0], parts[1]
            if "Start Date" in key: metadata["Start Date"] = val
            elif "End Date" in key: metadata["End Date"] = val
            elif "Server Time" in key: metadata["Server Time"] = val
            elif "Appliance" in key and "Key" not in key: metadata["Appliance"] = val
            elif "Firmware Version" in key: metadata["Firmware Version"] = val
            elif "Device Serial Number" in key: metadata["Appliance Key"] = val
            elif "Criteria" in key or (len(parts) > 1 and "Event Type is" in parts[1]): 
                metadata["Criteria"] = val if "Criteria" not in key else parts[1]

    # Encontrar inicio de tabla
    data_start = next((i for i, line in enumerate(lines) if "Time,Event Type,Severity,Message" in line), None)
    if data_start is None: return None, None, None

    df = pd.read_csv(StringIO('\n'.join(lines[data_start:])))
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
    df = df[df['Message'].str.contains("SSL VPN User", na=False)]

    def extraer_usuario_accion(msg):
        match = re.search(r"SSL VPN User '([^']+)' (connected|disconnected)", msg)
        return match.groups() if match else (None, None)

    df[['Usuario', 'Accion']] = df['Message'].apply(extraer_usuario_accion).apply(pd.Series)
    df = df.dropna(subset=['Usuario', 'Accion']).sort_values(by=['Usuario', 'Time'])

    # Lógica de sesiones (Excel y PDF)
    conexiones, abiertas = [], []
    server_time_str = metadata.get("Server Time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    st_dt = pd.to_datetime(server_time_str, errors='coerce')

    for usuario, grupo in df.groupby('Usuario'):
        pila = []
        for _, fila in grupo.iterrows():
            if fila['Accion'] == 'connected':
                pila.append(fila['Time'])
            elif fila['Accion'] == 'disconnected' and pila:
                inicio = pila.pop(0)
                conexiones.append({
                    'Usuario': usuario, 'Inicio': inicio, 'Fin': fila['Time'], 
                    'Duración': str(fila['Time'] - inicio).split('.')[0]
                })
        for t in pila:
            abiertas.append({'Usuario': usuario, 'Inicio': t, 'Estado': 'Sesión sin desconexión'})

    return pd.DataFrame(conexiones), pd.DataFrame(abiertas), metadata

# --- INTERFAZ DE USUARIO ---
archivo = st.file_uploader("Subir CSV de Sophos (System Events)", type="csv")

if archivo:
    df_f, df_a, meta = procesar_datos(archivo)
    
    if meta:
        st.success("✅ Reporte generado exitosamente")
        
        # 1. BOTÓN EXCEL
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_f.to_excel(writer, index=False, sheet_name='Conexiones')
            df_a.to_excel(writer, index=False, sheet_name='Sesiones Abiertas')
        
        # 2. GENERACIÓN DE PDF
        try:
            pdf = CustomPDF(meta)
            pdf.add_page()
            
            # Info Superior
            pdf.set_font("helvetica", "", 10)
            pdf.cell(0, 8, f"Appliance: {meta.get('Appliance', 'N/A')}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 8, f"Appliance key: {meta.get('Appliance Key', 'N/A')}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 8, f"Firmware Version: {meta.get('Firmware Version', 'N/A')}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 8, f"Filter(s) applied: {meta.get('Criteria', 'N/A')}", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(5)

            # Tabla Conexiones
            pdf.set_font("helvetica", "B", 12)
            pdf.cell(0, 10, "Conexiones completadas", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("helvetica", "B", 9)
            headers = [("Usuario", 45), ("Inicio", 45), ("Fin", 45), ("Duración", 45)]
            for h, w in headers: pdf.cell(w, 8, h, border=1, align="C")
            pdf.ln()
            
            pdf.set_font("helvetica", "", 8)
            for _, r in df_f.iterrows():
                u_clean = str(r["Usuario"]).encode('ascii', 'ignore').decode('ascii')
                pdf.cell(45, 7, u_clean, border=1)
                pdf.cell(45, 7, r["Inicio"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(45, 7, r["Fin"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(45, 7, str(r["Duración"]), border=1, align="C")
                pdf.ln()

            # IMPORTANTE: Convertir a bytes de forma segura para Streamlit
            pdf_data = pdf.output()
            
            # --- MOSTRAR BOTONES ---
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📥 Descargar en Excel", output_excel.getvalue(), "Reporte_VPN.xlsx")
            with col2:
                st.download_button("📄 Descargar en PDF", bytes(pdf_data), "Reporte_VPN.pdf")

        except Exception as e:
            st.error(f"Error al crear el PDF: {e}")
            st.download_button("📥 Descargar solo Excel", output_excel.getvalue(), "Reporte_VPN.xlsx")
