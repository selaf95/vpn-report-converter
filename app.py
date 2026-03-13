import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl import load_workbook

def procesar_csv_web(uploaded_file):
    # Leer el contenido del archivo subido
    stringio = StringIO(uploaded_file.getvalue().decode("utf-8"))
    lines = stringio.readlines()

    # --- Lógica de metadatos ---
    metadata = {}
    for line in lines[:15]: 
        line_clean = line.strip().replace('"', "")
        if "," in line_clean:
            parts = line_clean.split(",", 1)
            key, value = parts[0].strip(), parts[1].strip()
            if "Start Date" in key: metadata["Start Date"] = value
            elif "End Date" in key: metadata["End Date"] = value # Capturamos End Date
            elif "Server Time" in key: metadata["Server Time"] = value
            elif "Device Serial Number" in key: metadata["Appliance Key"] = value

    # --- Encontrar encabezado de datos ---
    data_start = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Time,Event Type,Severity,Message"):
            data_start = i
            break
    
    if data_start is None:
        return None, "No se encontró el encabezado 'Time,Event Type...'"

    # --- Cargar datos ---
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

    # --- Cálculo de sesiones ---
    sesiones_fusionadas = []
    sesiones_abiertas = []

    for usuario, grupo in df.groupby('Usuario'):
        grupo = grupo.reset_index(drop=True)
        pila = []
        eventos = []

        for _, fila in grupo.iterrows():
            if fila['Accion'] == 'connected':
                pila.append(fila['Time'])
            elif fila['Accion'] == 'disconnected' and pila:
                inicio = pila.pop(0)
                eventos.append({'Inicio': inicio, 'Fin': fila['Time']})

        for inicio_sin_fin in pila:
            sesiones_abiertas.append({
                'Usuario': usuario, 'Inicio': inicio_sin_fin, 'Fin': server_time,
                'Duración': str(server_time - inicio_sin_fin), 'Estado': 'Sesión abierta'
            })

        if eventos:
            eventos = sorted(eventos, key=lambda x: x['Inicio'])
            fusion = [eventos[0]]
            for e in eventos[1:]:
                if e['Inicio'] <= fusion[-1]['Fin']:
                    fusion[-1]['Fin'] = max(fusion[-1]['Fin'], e['Fin'])
                else:
                    fusion.append(e)
            for s in fusion:
                sesiones_fusionadas.append({
                    'Usuario': usuario, 'Inicio': s['Inicio'], 'Fin': s['Fin'],
                    'Duración': str(s['Fin'] - s['Inicio'])
                })

    # --- Crear Excel en memoria ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(sesiones_fusionadas).to_excel(writer, sheet_name="Conexiones completadas", index=False)
        pd.DataFrame(sesiones_abiertas).to_excel(writer, sheet_name="Conexiones abiertas", index=False)

    # Ajustes de formato
    output.seek(0)
    wb = load_workbook(output)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for col in ws.columns:
            max_length = 0
            for cell in col:
                if cell.value: max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2
    
    final_output = BytesIO()
    wb.save(final_output)
    return final_output.getvalue(), metadata

# --- INTERFAZ DE STREAMLIT ---
st.set_page_config(page_title="Analizador VPN Sophos", page_icon="🛡️")
st.title("🛡️ Procesador de Reportes VPN Sophos")

uploaded_file = st.file_uploader("Sube tu archivo CSV de Sophos", type="csv")

if uploaded_file is not None:
    with st.spinner('Procesando...'):
        excel_data, meta = procesar_csv_web(uploaded_file)
        
        if excel_data:
            serial = meta.get("Appliance Key", "SIN_SERIAL")
            
            # Formatear fechas
            start_raw = meta.get("Start Date", "")
            end_raw = meta.get("End Date", "")
            
            try:
                start_dt = pd.to_datetime(start_raw)
                start_fmt = start_dt.strftime("%Y-%m-%d")
            except:
                start_fmt = str(start_raw).split(" ")[0].replace("/", "-")
                
            try:
                end_dt = pd.to_datetime(end_raw)
                end_fmt = end_dt.strftime("%Y-%m-%d")
            except:
                end_fmt = str(end_raw).split(" ")[0].replace("/", "-")

            # --- Lógica de nombre de archivo ---
            if start_fmt == end_fmt:
                filename = f"Reporte_VPN_{serial}_{start_fmt}.xlsx"
            else:
                filename = f"Reporte_VPN_{serial}_{start_fmt}_{end_fmt}.xlsx"

            st.success("✅ Archivo procesado")
            st.info(f"**Archivo generado:** `{filename}`")
            
            st.download_button(
                label="📥 Descargar Reporte Excel",
                data=excel_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
