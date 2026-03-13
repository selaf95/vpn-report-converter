import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
from datetime import datetime
from fpdf import FPDF
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import os

# --- CLASE PDF PERSONALIZADA ---
class CustomPDF(FPDF):
    def __init__(self, metadata):
        super().__init__()
        self.metadata = metadata

    def header(self):
        # El logo debe llamarse logo.jpg en tu GitHub
        if os.path.exists("logo.jpg"):
            try:
                self.image("logo.jpg", x=10, y=8, w=30)
            except:
                pass
        
        self.set_y(20)
        self.set_font("Arial", "B", 24)
        # Limpiamos caracteres extraños para evitar errores en PDF
        self.cell(0, 15, "System events", ln=True, align="L")
        self.set_font("Arial", "", 12)
        
        start = str(self.metadata.get('Start Date', ''))
        end = str(self.metadata.get('End Date', ''))
        self.cell(0, 10, f"{start} - {end}", ln=True, align="L")
        self.ln(5)

    def footer(self):
        self.set_y(-25)
        self.set_font("Arial", "I", 8)
        server_t = str(self.metadata.get('Server Time', ''))
        # Limpieza ASCII para el footer
        clean_footer = f"Server time: {server_t}".encode('ascii', 'ignore').decode('ascii')
        self.cell(0, 10, clean_footer, 0, 0, 'R')

def procesar_datos(uploaded_file):
    try:
        content = uploaded_file.getvalue().decode("utf-8")
    except UnicodeDecodeError:
        content = uploaded_file.getvalue().decode("latin-1")
        
    stringio = StringIO(content)
    lines = stringio.readlines()

    metadata = {}
    for line in lines[:15]:
        line_clean = line.strip().replace('"', "")
        if "," in line_clean:
            parts = line_clean.split(",", 1)
            key, value = parts[0].strip(), parts[1].strip()
            if "Start Date" in key: metadata["Start Date"] = value
            elif "End Date" in key: metadata["End Date"] = value
            elif "Server Time" in key: metadata["Server Time"] = value
            elif "Appliance" in key: metadata["Appliance"] = value
            elif "Firmware Version" in key: metadata["Firmware Version"] = value
            elif "Device Serial Number" in key: metadata["Appliance Key"] = value
            elif "Criteria" in key: metadata["Criteria"] = value

    data_start = next((i for i, line in enumerate(lines) if "Time,Event Type,Severity,Message" in line), None)
    if data_start is None: return None, None, None

    data_content = StringIO(''.join(lines[data_start:]))
    df = pd.read_csv(data_content)
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
    df = df[df['Message'].str.contains("SSL VPN User", na=False)]

    def extraer_usuario_accion(msg):
        match = re.search(r"SSL VPN User '([^']+)' (connected|disconnected)", msg)
        return match.groups() if match else (None, None)

    df[['Usuario', 'Accion']] = df['Message'].apply(extraer_usuario_accion).apply(pd.Series)
    df = df.dropna(subset=['Usuario', 'Accion']).sort_values(by=['Usuario', 'Time'])

    server_time = pd.to_datetime(metadata.get("Server Time"), errors="coerce")
    if pd.isna(server_time): server_time = datetime.now()

    fusionadas, abiertas = [], []
    for usuario, grupo in df.groupby('Usuario'):
        grupo = grupo.reset_index(drop=True)
        pila, eventos = [], []
        for _, fila in grupo.iterrows():
            if fila['Accion'] == 'connected': pila.append(fila['Time'])
            elif fila['Accion'] == 'disconnected' and pila:
                eventos.append({'Inicio': pila.pop(0), 'Fin': fila['Time']})
        
        for t in pila:
            abiertas.append({'Usuario': usuario, 'Inicio': t, 'Fin': server_time, 'Duracion': str(server_time - t).split('.')[0], 'Estado': 'Abierta'})

        if eventos:
            eventos.sort(key=lambda x: x['Inicio'])
            temp_fusion = [eventos[0]]
            for e in eventos[1:]:
                if e['Inicio'] <= temp_fusion[-1]['Fin']:
                    temp_fusion[-1]['Fin'] = max(temp_fusion[-1]['Fin'], e['Fin'])
                else:
                    temp_fusion.append(e)
            for s in temp_fusion:
                fusionadas.append({'Usuario': usuario, 'Inicio': s['Inicio'], 'Fin': s['Fin'], 'Duracion': str(s['Fin'] - s['Inicio']).split('.')[0]})

    return pd.DataFrame(fusionadas), pd.DataFrame(abiertas), metadata

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Sophos VPN Reporter", layout="wide")
st.title("🛡️ Sophos VPN Report Generator")

uploaded_file = st.file_uploader("Subir archivo CSV de Sophos", type="csv")

if uploaded_file:
    with st.spinner('Procesando datos...'):
        df_f, df_a, meta = procesar_datos(uploaded_file)
    
    if meta:
        # 1. GENERACIÓN DE EXCEL
        excel_out = BytesIO()
        with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
            df_f.to_excel(writer, sheet_name="Completadas", index=False)
            df_a.to_excel(writer, sheet_name="Abiertas", index=False)
        excel_out.seek(0)

        # 2. GENERACIÓN DE PDF (Seguro contra caracteres especiales)
        try:
            pdf = CustomPDF(meta)
            pdf.add_page()
            pdf.set_font("Arial", size=10)
            
            # Datos del equipo (Limpios de caracteres no-ASCII)
            for k in ['Appliance', 'Appliance Key', 'Firmware Version', 'Criteria']:
                clean_val = str(meta.get(k, 'N/A')).encode('ascii', 'ignore').decode('ascii')
                pdf.cell(0, 7, f"{k}: {clean_val}", ln=True)
            
            pdf.ln(5)
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, "Conexiones completadas", ln=True)
            
            # Encabezados tabla
            pdf.set_font("Arial", "B", 9)
            pdf.cell(45, 8, "Usuario", border=1, align="C")
            pdf.cell(45, 8, "Inicio", border=1, align="C")
            pdf.cell(45, 8, "Fin", border=1, align="C")
            pdf.cell(45, 8, "Duracion", border=1, align="C")
            pdf.ln()
            
            # Filas de la tabla
            pdf.set_font("Arial", "", 8)
            for _, r in df_f.iterrows():
                # Forzamos nombre de usuario a ASCII para evitar errores de FPDF
                u_clean = str(r["Usuario"]).encode('ascii', 'ignore').decode('ascii')
                pdf.cell(45, 7, u_clean, border=1)
                pdf.cell(45, 7, r["Inicio"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(45, 7, r["Fin"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(45, 7, str(r["Duracion"]), border=1, align="C")
                pdf.ln()

            # Salida del PDF a bytes
            pdf_bytes = pdf.output(dest='S')
            if isinstance(pdf_bytes, str):
                pdf_bytes = pdf_bytes.encode('latin-1', 'replace')

            # --- LÓGICA DE NOMBRE DINÁMICO ---
            serial = meta.get("Appliance Key", "Serial")
            try:
                s_date = pd.to_datetime(meta.get("Start Date")).strftime("%Y-%m-%d")
                e_date = pd.to_datetime(meta.get("End Date")).strftime("%Y-%m-%d")
            except:
                s_date = "Fecha"
                e_date = "Fecha"
                
            name_base = f"{serial}_{s_date}" if s_date == e_date else f"{serial}_{s_date}_{e_date}"

            # --- BOTONES ---
            st.success(f"✅ Reporte generado para el Serial: {serial}")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📥 Descargar Excel", excel_out, f"Reporte_VPN_{name_base}.xlsx")
            with c2:
                st.download_button("📄 Descargar PDF Estético", pdf_bytes, f"Reporte_VPN_{name_base}.pdf")
                
        except Exception as e:
            st.error(f"Error técnico en PDF: {e}")
            st.info("Puedes descargar el Excel mientras corregimos el PDF.")
            st.download_button("📥 Descargar Excel (Alternativo)", excel_out, "Reporte_VPN.xlsx")
    else:
        st.error("No se pudo extraer información del archivo. Verifica que sea un CSV válido de Sophos.")
