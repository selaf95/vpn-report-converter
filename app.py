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

# --- CLASE PDF CONFIGURADA PARA FPDF2 ---
class CustomPDF(FPDF):
    def __init__(self, metadata):
        super().__init__()
        self.metadata = metadata

    def header(self):
        # Si el logo da problemas, el bloque try evita que la app muera
        if os.path.exists("logo.jpg"):
            try:
                self.image("logo.jpg", x=10, y=8, w=30)
            except Exception:
                pass
        
        self.set_y(20)
        self.set_font("helvetica", "B", 24)
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
        self.cell(0, 10, f"Server time: {server_t}", align='R')

def procesar_datos(uploaded_file):
    try:
        content = uploaded_file.getvalue().decode("utf-8")
    except:
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

# --- INTERFAZ ---
st.set_page_config(page_title="Sophos VPN Reporter", page_icon="🛡️")
st.title("🛡️ Generador de Reportes VPN")

archivo = st.file_uploader("Sube tu CSV de Sophos", type="csv")

if archivo:
    with st.spinner("Procesando datos..."):
        df_f, df_a, meta = procesar_datos(archivo)
    
    if meta:
        # EXCEL
        excel_out = BytesIO()
        with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
            df_f.to_excel(writer, sheet_name="Completadas", index=False)
            df_a.to_excel(writer, sheet_name="Abiertas", index=False)
        excel_out.seek(0)

        # PDF con FPDF2
        pdf_bytes = None
        try:
            pdf = CustomPDF(meta)
            pdf.add_page()
            pdf.set_font("helvetica", size=10)
            for k in ['Appliance', 'Appliance Key', 'Firmware Version', 'Criteria']:
                val = str(meta.get(k, 'N/A')).encode('ascii', 'ignore').decode('ascii')
                pdf.cell(0, 7, f"{k}: {val}", new_x="LMARGIN", new_y="NEXT")
            
            pdf.ln(5)
            pdf.set_font("helvetica", "B", 12)
            pdf.cell(0, 10, "Conexiones completadas", new_x="LMARGIN", new_y="NEXT")
            
            pdf.set_font("helvetica", "B", 9)
            cols = [("Usuario", 45), ("Inicio", 45), ("Fin", 45), ("Duracion", 45)]
            for header, width in cols:
                pdf.cell(width, 8, header, border=1, align="C")
            pdf.ln()
            
            pdf.set_font("helvetica", "", 8)
            for _, r in df_f.iterrows():
                u = str(r["Usuario"]).encode('ascii', 'ignore').decode('ascii')
                pdf.cell(45, 7, u, border=1)
                pdf.cell(45, 7, r["Inicio"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(45, 7, r["Fin"].strftime("%Y-%m-%d %H:%M"), border=1, align="C")
                pdf.cell(45, 7, str(r["Duracion"]), border=1, align="C")
                pdf.ln()

            pdf_bytes = pdf.output()
        except Exception as e:
            st.warning(f"Nota: El PDF no se pudo generar con el diseño completo ({e}). El Excel está disponible.")

        # Lógica de nombre
        serial = meta.get("Appliance Key", "Serial")
        s_date = str(meta.get("Start Date", "Inicio")).split(" ")[0]
        e_date = str(meta.get("End Date", "Fin")).split(" ")[0]
        name = f"{serial}_{s_date}" if s_date == e_date else f"{serial}_{s_date}_{e_date}"

        st.success("✅ Reporte listo")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Descargar Excel", excel_out, f"Reporte_{name}.xlsx")
        if pdf_bytes:
            with c2:
                st.download_button("📄 Descargar PDF", pdf_bytes, f"Reporte_{name}.pdf")
        st.error("No se pudo extraer información del archivo. Verifica que sea un CSV válido de Sophos.")
