import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
from datetime import datetime
from fpdf import FPDF
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl import load_workbook

# --- CLASE PDF PERSONALIZADA ---
class CustomPDF(FPDF):
    def __init__(self, metadata):
        super().__init__()
        self.metadata = metadata

    def header(self):
        # Intenta cargar el logo si existe en el repo
        try:
            self.image("logo.jpg", x=10, y=8, w=30)
        except:
            pass 
        
        self.set_y(20)
        self.set_font("Arial", "B", 26)
        self.cell(0, 15, "System events", ln=True, align="L")
        self.set_font("Arial", "", 12)
        start = self.metadata.get('Start Date', '')
        end = self.metadata.get('End Date', '')
        self.cell(0, 10, f"{start} - {end}", ln=True, align="L")
        self.ln(5)

    def footer(self):
        self.set_y(-25)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Server time: {self.metadata.get('Server Time')}", 0, 0, 'R')

def procesar_datos(uploaded_file):
    stringio = StringIO(uploaded_file.getvalue().decode("utf-8"))
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

    # --- LÓGICA DE FUSIÓN DE SESIONES ---
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
            abiertas.append({'Usuario': usuario, 'Inicio': t, 'Fin': server_time, 'Duración': str(server_time - t), 'Estado': 'Sesión abierta'})

        if eventos:
            eventos.sort(key=lambda x: x['Inicio'])
            temp_fusion = [eventos[0]]
            for e in eventos[1:]:
                if e['Inicio'] <= temp_fusion[-1]['Fin']:
                    temp_fusion[-1]['Fin'] = max(temp_fusion[-1]['Fin'], e['Fin'])
                else:
                    temp_fusion.append(e)
            for s in temp_fusion:
                fusionadas.append({'Usuario': usuario, 'Inicio': s['Inicio'], 'Fin': s['Fin'], 'Duración': str(s['Fin'] - s['Inicio'])})

    return pd.DataFrame(fusionadas), pd.DataFrame(abiertas), metadata

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Sophos VPN Reporter", layout="wide")
st.title("📊 Sophos VPN Report Generator")

uploaded_file = st.file_uploader("Subir CSV de System Events", type="csv")

if uploaded_file:
    df_f, df_a, meta = procesar_datos(uploaded_file)
    
    if meta:
        # 1. GENERAR EXCEL
        excel_out = BytesIO()
        with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
            df_f.to_excel(writer, sheet_name="Completadas", index=False)
            df_a.to_excel(writer, sheet_name="Abiertas", index=False)
        excel_out.seek(0)

        # 2. GENERAR PDF ESTÉTICO
        pdf = CustomPDF(meta)
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        for k in ['Appliance', 'Appliance Key', 'Firmware Version', 'Criteria']:
            pdf.cell(0, 7, f"{k}: {meta.get(k, 'N/A')}", ln=True)
        
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Conexiones completadas", ln=True)
        pdf.set_font("Arial", "B", 9)
        cols = [("Usuario", 40), ("Inicio", 45), ("Fin", 45), ("Duración", 55)]
        for c, w in cols: pdf.cell(w, 8, c, border=1, align="C")
        pdf.ln()
        pdf.set_font("Arial", "", 8)
        for _, r in df_f.iterrows():
            pdf.cell(40, 7, str(r["Usuario"]), border=1)
            pdf.cell(45, 7, r["Inicio"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
            pdf.cell(45, 7, r["Fin"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
            pdf.cell(55, 7, str(r["Duración"]), border=1, align="C")
            pdf.ln()

        pdf_out = pdf.output(dest='S').encode('latin-1')

        # --- LÓGICA DE NOMBRE ---
        serial = meta.get("Appliance Key", "Serial")
        s_date = pd.to_datetime(meta.get("Start Date")).strftime("%Y-%m-%d")
        e_date = pd.to_datetime(meta.get("End Date")).strftime("%Y-%m-%d")
        name_base = f"{serial}_{s_date}" if s_date == e_date else f"{serial}_{s_date}_{e_date}"

        # --- BOTONES DE DESCARGA ---
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 Descargar Excel", excel_out, f"Reporte_VPN_{name_base}.xlsx")
        with col2:
            st.download_button("📄 Descargar PDF Estético", pdf_out, f"Reporte_VPN_{name_base}.pdf")
        
        st.success(f"Reporte listo para: {name_base}")
