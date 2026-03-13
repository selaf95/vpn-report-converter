import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
from datetime import datetime
from fpdf import FPDF
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment # Para centrar en Excel
import os

# --- CLASE PDF MEJORADA ---
class CustomPDF(FPDF):
    def __init__(self, metadata):
        super().__init__()
        self.metadata = metadata

    def header(self):
        if os.path.exists("logo.jpg"):
            try:
                self.image("logo.jpg", x=10, y=8, w=30)
            except: pass
        self.set_y(20)
        self.set_font("helvetica", "B", 26)
        self.cell(0, 15, "System events", align="L", new_x="LMARGIN", new_y="NEXT")
        self.set_font("helvetica", "", 12)
        start, end = str(self.metadata.get('Start Date', '')), str(self.metadata.get('End Date', ''))
        self.cell(0, 10, f"{start} - {end}", align="L", new_x="LMARGIN", new_y="NEXT")
        self.ln(5)

    def footer(self):
        self.set_y(-25)
        self.set_font("helvetica", "I", 8)
        server_t = str(self.metadata.get('Server Time', ''))
        clean_f = f"Server time: {server_t}".encode('ascii', 'ignore').decode('ascii')
        self.cell(0, 10, clean_f, align='R')

# --- PROCESAMIENTO ---
def procesar_datos(uploaded_file):
    try:
        content = uploaded_file.getvalue().decode("utf-8")
    except:
        content = uploaded_file.getvalue().decode("latin-1")
    lines = content.splitlines()
    metadata = {}
    for line in lines[:15]:
        parts = [p.strip().replace('"', '') for p in line.split(',')]
        if len(parts) >= 2:
            k, v = parts[0], parts[1]
            if "Start Date" in k: metadata["Start Date"] = v
            elif "End Date" in k: metadata["End Date"] = v
            elif "Server Time" in k: metadata["Server Time"] = v
            elif "Appliance" in k and "Key" not in k: metadata["Appliance"] = v
            elif "Firmware Version" in k: metadata["Firmware Version"] = v
            elif "Device Serial Number" in k: metadata["Appliance Key"] = v
            elif "Criteria" in k or (len(parts) > 1 and "Event Type is" in parts[1]): 
                metadata["Criteria"] = v if "Criteria" not in k else parts[1]

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

    conex, actv = [], []
    for usuario, grupo in df.groupby('Usuario'):
        pila = []
        for _, fila in grupo.iterrows():
            if fila['Accion'] == 'connected': pila.append(fila['Time'])
            elif fila['Accion'] == 'disconnected' and pila:
                ini = pila.pop(0)
                conex.append({'Usuario': usuario, 'Inicio': ini, 'Fin': fila['Time'], 'Duración': str(fila['Time'] - ini).split('.')[0]})
        for t in pila: actv.append({'Usuario': usuario, 'Inicio': t, 'Estado': 'Conectado'})
    return pd.DataFrame(conex), pd.DataFrame(actv), metadata

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Sophos VPN Reporter", page_icon="🛡️")
st.title("🛡️ Generador de Reportes VPN")
archivo = st.file_uploader("Subir CSV de Sophos", type="csv")

if archivo:
    df_f, df_a, meta = procesar_datos(archivo)
    if meta:
        serial = meta.get("Appliance Key", "SERIAL")
        try:
            d_s = pd.to_datetime(meta.get("Start Date")).strftime("%Y-%m-%d")
            d_e = pd.to_datetime(meta.get("End Date")).strftime("%Y-%m-%d")
        except: d_s, d_e = "FECHA", "FECHA"
        nombre = f"Reporte_VPN_{serial}_{d_s}" if d_s == d_e else f"Reporte_VPN_{serial}_{d_s}_{d_e}"

        # --- EXCEL (AUTOAJUSTE + CENTRADO) ---
        out_xl = BytesIO()
        with pd.ExcelWriter(out_xl, engine='openpyxl') as writer:
            df_f.to_excel(writer, index=False, sheet_name='Completadas')
            df_a.to_excel(writer, index=False, sheet_name='Conexiones Activas')
            for sheet in writer.sheets:
                ws = writer.sheets[sheet]
                for col in ws.columns:
                    max_len = 0
                    for cell in col:
                        cell.alignment = Alignment(horizontal='center', vertical='center') # CENTRADO
                        try:
                            if len(str(cell.value)) > max_len: max_len = len(str(cell.value))
                        except: pass
                    ws.column_dimensions[col[0].column_letter].width = max_len + 5

        # --- PDF (ANCHO ADAPTABLE) ---
        try:
            pdf = CustomPDF(meta)
            pdf.add_page()
            pdf.set_font("helvetica", "", 10)
            for k, label in [('Appliance', 'Appliance'), ('Appliance Key', 'Appliance key'), ('Firmware Version', 'Firmware Version'), ('Criteria', 'Filter(s) applied')]:
                pdf.cell(0, 7, f"{label}: {meta.get(k if k != 'Appliance Key' else 'Appliance Key', 'N/A')}", new_x="LMARGIN", new_y="NEXT")
            
            pdf.ln(5)
            pdf.set_font("helvetica", "B", 12); pdf.cell(0, 10, "Conexiones completadas", new_x="LMARGIN", new_y="NEXT")
            
            # Encabezados PDF con Usuario más ancho (60 en vez de 45)
            pdf.set_font("helvetica", "B", 9)
            pdf.cell(60, 8, "Usuario", border=1, align="C")
            pdf.cell(40, 8, "Inicio", border=1, align="C")
            pdf.cell(40, 8, "Fin", border=1, align="C")
            pdf.cell(50, 8, "Duración", border=1, align="C")
            pdf.ln()
            
            pdf.set_font("helvetica", "", 8)
            for _, r in df_f.iterrows():
                u = str(r["Usuario"]).encode('ascii', 'ignore').decode('ascii')
                # Usamos multi_cell para el usuario si es muy largo
                x, y = pdf.get_x(), pdf.get_y()
                pdf.multi_cell(60, 7, u, border=1, align="L")
                pdf.set_xy(x + 60, y)
                pdf.cell(40, 7, r["Inicio"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(40, 7, r["Fin"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(50, 7, str(r["Duración"]), border=1, align="C")
                pdf.ln()

            if not df_a.empty:
                pdf.ln(10); pdf.set_font("helvetica", "B", 12); pdf.cell(0, 10, "Conexiones activas", new_x="LMARGIN", new_y="NEXT")
                pdf.set_font("helvetica", "B", 9)
                pdf.cell(70, 8, "Usuario", border=1, align="C"); pdf.cell(60, 8, "Inicio de Sesión", border=1, align="C"); pdf.cell(60, 8, "Estado", border=1, align="C"); pdf.ln()
                pdf.set_font("helvetica", "", 8)
                for _, r in df_a.iterrows():
                    u = str(r["Usuario"]).encode('ascii', 'ignore').decode('ascii')
                    pdf.cell(70, 7, u, border=1); pdf.cell(60, 7, r["Inicio"].strftime("%Y-%m-%d %H:%M"), border=1, align="C"); pdf.cell(60, 7, "Conectado", border=1, align="C"); pdf.ln()

            col1, col2 = st.columns(2)
            with col1: st.download_button("📥 Excel", out_xl.getvalue(), f"{nombre}.xlsx")
            with col2: st.download_button("📄 PDF", bytes(pdf.output()), f"{nombre}.pdf")
        except Exception as e:
            st.error(f"Error PDF: {e}")
            st.download_button("📥 Descargar Excel", out_xl.getvalue(), f"{nombre}.xlsx")
