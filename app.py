import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import math
import os
import requests
import base64
from fpdf import FPDF
import tempfile
from PIL import Image
from io import BytesIO

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestión de Servicios Adicionales", page_icon="🛡️", layout="wide")

# URL DE TU APLICACIÓN WEB DE APPS SCRIPT CONFIGURADA
URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxqVp5T0fVGtPncuwkocbXbU5TSHLgPqgmD0A5pQ6CmuEU3C-UrGx2eZHxg2HTLUjU1/exec"

@st.cache_resource
def conectar_google():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        if "gcp_service_account" in st.secrets:
            service_account_info = dict(st.secrets["gcp_service_account"])
            credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
        else:
            ruta_json = os.path.join(os.path.dirname(__file__), "credenciales.json")
            if not os.path.exists(ruta_json):
                ruta_json = "credenciales.json"
                
            if not os.path.exists(ruta_json):
                st.error("🚨 **ARCHIVO DE CREDENCIALES NO ENCONTRADO**")
                st.warning("No se encontró `credenciales.json` ni los secretos configurados.")
                st.stop()
                
            credentials = Credentials.from_service_account_file(ruta_json, scopes=scopes)
            
        client = gspread.authorize(credentials)
        return client
    except Exception as e:
        st.error(f"Error al conectar con Google: {e}")
        st.stop()

# --- INICIALIZACIÓN DE SESIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_dni" not in st.session_state:
    st.session_state.usuario_dni = None
if "dependencia" not in st.session_state:
    st.session_state.dependencia = None
if "rol" not in st.session_state:
    st.session_state.rol = None
if "cambio_clave_requerido" not in st.session_state:
    st.session_state.cambio_clave_requerido = False

def pantalla_cambio_clave():
    st.title("🔑 Definir Nueva Contraseña Personal")
    st.warning("⚠️ Primer ingreso detectado. Debes cambiar tu contraseña temporal.")
    nueva_clave = st.text_input("Nueva Contraseña", type="password")
    confirmar_clave = st.text_input("Confirmar Nueva Contraseña", type="password")
    if st.button("Guardar Nueva Contraseña", type="primary"):
        if nueva_clave == confirmar_clave and nueva_clave != "":
            try:
                client = conectar_google()
                sheet = client.open("Gestion_Adicionales").worksheet("Usuarios")
                cell = sheet.find(str(st.session_state.dni_temp))
                sheet.update_cell(cell.row, 3, nueva_clave)
                sheet.update_cell(cell.row, 4, "")
                st.success("¡Contraseña actualizada con éxito! Inicia sesión nuevamente.")
                st.session_state.cambio_clave_requerido = False
                st.session_state.dni_temp = None
                st.rerun()
            except Exception as e:
                st.error(f"Error al actualizar: {e}")
        else:
            st.error("Las contraseñas no coinciden.")

def login():
    st.title("🛡️ Acceso a Servicios Adicionales")
    dni = st.text_input("DNI / Usuario")
    clave = st.text_input("Contraseña", type="password")
    if st.button("Ingresar", type="primary"):
        client = conectar_google()
        if client:
            try:
                sheet = client.open("Gestion_Adicionales").worksheet("Usuarios")
                data = pd.DataFrame(sheet.get_all_records())
                user = data[data['dni'].astype(str) == str(dni)]
                if not user.empty:
                    user_data = user.iloc[0]
                    clave_pers = str(user_data['clave_personal']).strip()
                    clave_temp = str(user_data['clave_temporal']).strip()
                    if clave_pers == "" or clave_pers == str(dni) or (clave_temp != "" and clave == clave_temp):
                        if clave == clave_temp or clave == str(dni) or clave == clave_pers:
                            st.session_state.cambio_clave_requerido = True
                            st.session_state.dni_temp = dni
                            st.rerun()
                        else:
                            st.error("Contraseña temporal incorrecta.")
                    else:
                        if clave == clave_pers:
                            st.session_state.autenticado = True
                            st.session_state.usuario_dni = dni
                            st.session_state.dependencia = user_data['dependencia']
                            st.session_state.rol = str(user_data['rol']).strip()
                            st.rerun()
                        else:
                            st.error("Contraseña incorrecta.")
                else:
                    st.error("DNI no registrado.")
            except Exception as db_err:
                st.error(f"Error al consultar la base de datos: {db_err}")

# --- CONVERTIDOR DE LINK DE DRIVE A DESCARGA DIRECTA DE IMAGEN ---
def obtener_imagen_desde_drive(url_drive):
    try:
        if not url_drive or not str(url_drive).startswith("http"):
            return None
        
        file_id = None
        if "id=" in url_drive:
            file_id = url_drive.split("id=")[1].split("&")[0]
        elif "/d/" in url_drive:
            file_id = url_drive.split("/d/")[1].split("/")[0]
            
        if not file_id:
            return None
            
        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(download_url, timeout=10)
        
        if response.status_code == 200 and len(response.content) > 1000:
            image = Image.open(BytesIO(response.content))
            if image.mode in ("RGBA", "P"):
                image = image.convert("RGB")
            
            tmp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            image.save(tmp_img.name, "JPEG")
            return tmp_img.name
    except Exception:
        pass
    return None

# --- GENERADOR DE PDF COMPLETO CON CONSTANCIAS ---
def generar_pdf_servicio(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Encabezado
    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, "INFORME DE SERVICIO ADICIONAL", ln=True, align="center")
    pdf.set_font("helvetica", "I", 10)
    pdf.cell(0, 6, f"ID de Solicitud: {row.get('ID Servicio', '')}", ln=True, align="center")
    pdf.ln(5)
    
    # Datos Principales
    pdf.set_font("helvetica", "B", 12)
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(0, 8, " 1. DATOS GENERALES DEL EVENTO", ln=True, fill=True)
    pdf.set_font("helvetica", "", 10)
    
    pdf.cell(95, 6, f"Dependencia: {row.get('Dependencia', '')}", ln=0)
    pdf.cell(95, 6, f"Fecha: {row.get('Fecha Evento', '')}", ln=1)
    pdf.cell(95, 6, f"Tipo de Evento: {row.get('Tipo de Evento', '')}", ln=0)
    pdf.cell(95, 6, f"Horario: {row.get('Horario', '')}", ln=1)
    pdf.cell(0, 6, f"Ubicación / Predio: {row.get('Ubicación Predio', '')}", ln=1)
    pdf.cell(0, 6, f"Organizador: {row.get('Organizador', '')}", ln=1)
    pdf.ln(3)
    
    # Personal y Módulos
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, " 2. RECURSOS Y MÓDULOS ASIGNADOS", ln=True, fill=True)
    pdf.set_font("helvetica", "", 10)
    pdf.cell(95, 6, f"Cantidad de Jefes: {row.get('Cant. Jefes', '')}", ln=0)
    pdf.cell(95, 6, f"Cantidad de Subalternos: {row.get('Cant. Subalternos', '')}", ln=1)
    pdf.cell(95, 6, f"Módulos (Turnos): {row.get('Módulos', '')}", ln=0)
    pdf.cell(95, 6, f"Monto Total: ${row.get('Monto Total', 0):,.2f}", ln=1)
    pdf.ln(3)
    
    # Facturación
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, " 3. DATOS DE FACTURACIÓN Y PAGO", ln=True, fill=True)
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 6, f"Pagado por: {row.get('Pagador', '')}", ln=1)
    pdf.cell(0, 6, f"CUIL / CUIT Facturación: {row.get('CUIL Pagador', '')}", ln=1)
    pdf.ln(3)
    
    # Enlaces de Drive como referencia textual
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, " 4. ENLACES DIRECTOS A GOOGLE DRIVE", ln=True, fill=True)
    pdf.set_font("helvetica", "", 10)
    
    def agregar_enlace(titulo, url):
        if url and str(url).startswith("http"):
            pdf.cell(0, 6, f"- {titulo}: {url}", ln=True)
        else:
            pdf.cell(0, 6, f"- {titulo}: No disponible", ln=True)

    agregar_enlace("Habilitación", row.get('Enlace Habilitación', ''))
    agregar_enlace("Comprobante", row.get('Enlace Comprobante', ''))
    agregar_enlace("Contrato", row.get('Enlace Contrato', ''))
    agregar_enlace("Asistencia", row.get('Enlace Asistencia', ''))
    
    # --- ANEXAR CONSTANCIAS COMO IMÁGENES ---
    archivos_a_anexar = [
        ("CONSTANCIA: HABILITACIÓN MUNICIPAL", row.get('Enlace Habilitación', '')),
        ("CONSTANCIA: COMPROBANTE DE PAGO", row.get('Enlace Comprobante', '')),
        ("CONSTANCIA: CONTRATO DE ADICIONAL", row.get('Enlace Contrato', '')),
        ("CONSTANCIA: PLANILLA DE ASISTENCIA", row.get('Enlace Asistencia', ''))
    ]
    
    for titulo_anexo, url_archivo in archivos_a_anexar:
        ruta_img = obtener_imagen_desde_drive(url_archivo)
        if ruta_img:
            pdf.add_page()
            pdf.set_font("helvetica", "B", 12)
            pdf.cell(0, 8, titulo_anexo, ln=True, align="center")
            pdf.ln(5)
            try:
                pdf.image(ruta_img, x=15, y=30, w=180)
            except Exception:
                pdf.cell(0, 6, "(No se pudo renderizar la imagen del archivo)", ln=True, align="center")
    
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_file.name)
    return tmp_file.name

def vista_dependencia():
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Colaborador")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()
        
    tab1, tab2, tab3 = st.tabs(["🟢 Cargar Nuevo Servicio", "🔵 Pendientes de Asistencia", "📜 Historial de Servicios"])
    
    with tab1:
        st.header(f"Solicitud de Servicio Adicional - {st.session_state.dependencia}")
        fecha = st.date_input("Fecha del Evento", min_value=datetime.date.today())
        
        st.subheader("📋 Datos del Evento")
        tipo_evento = st.text_input("Tipo de Evento (Ej: Partido de fútbol, Baile, Recital, etc.)")
        organizador = st.text_input("Nombre Completo y DNI del Organizador")
        ubicacion_predio = st.text_input("Ubicación / Predio del Evento")
        
        col_h1, col_h2 = st.columns(2)
        with col_h1: hora_inicio = st.time_input("Hora de Inicio", value=datetime.time(20, 0))
        with col_h2: hora_fin = st.time_input("Hora de Fin", value=datetime.time(4, 0))
            
        dt_inicio = datetime.datetime.combine(datetime.date.today(), hora_inicio)
        dt_fin = datetime.datetime.combine(datetime.date.today(), hora_fin)
        if dt_fin <= dt_inicio: dt_fin += datetime.timedelta(days=1)
            
        diferencia_horas = (dt_fin - dt_inicio).total_seconds() / 3600.0
        max_modulos = max(1, min(4, math.ceil(diferencia_horas / 4.0)))
        st.info(f"⏱️ Duración: **{diferencia_horas:.1f} horas** (Hasta **{max_modulos} módulo(s)**)")

        jefes_mod1, subs_mod1 = 0, 0
        jefes_mod2, subs_mod2 = 0, 0
        jefes_mod3, subs_mod3 = 0, 0
        jefes_mod4, subs_mod4 = 0, 0

        if max_modulos >= 1:
            st.markdown("**Turno Simple (1 Módulo - 4 horas)**")
            c1, c2 = st.columns(2)
            with c1: jefes_mod1 = st.number_input("Jefes (Simple)", 0, step=1, key="j1")
            with c2: subs_mod1 = st.number_input("Subalternos (Simple)", 0, step=1, key="s1")
                
        if max_modulos >= 2:
            st.markdown("**Turno Doble (2 Módulos - 8 horas)**")
            c3, c4 = st.columns(2)
            with c3: jefes_mod2 = st.number_input("Jefes (Doble)", 0, step=1, key="j2")
            with c4: subs_mod2 = st.number_input("Subalternos (Doble)", 0, step=1, key="s2")

        if max_modulos >= 3:
            st.markdown("**Turno Triple (3 Módulos - 12 horas)**")
            c5, c6 = st.columns(2)
            with c5: jefes_mod3 = st.number_input("Jefes (Triple)", 0, step=1, key="j3")
            with c6: subs_mod3 = st.number_input("Subalternos (Triple)", 0, step=1, key="s3")

        if max_modulos >= 4:
            st.markdown("**Turno Cuádruple (4 Módulos - 16 horas)**")
            c7, c8 = st.columns(2)
            with c7: jefes_mod4 = st.number_input("Jefes (Cuádruple)", 0, step=1, key="j4")
            with c8: subs_mod4 = st.number_input("Subalternos (Cuádruple)", 0, step=1, key="s4")

        total_jefes = jefes_mod1 + jefes_mod2 + jefes_mod3 + jefes_mod4
        total_subalternos = subs_mod1 + subs_mod2 + subs_mod3 + subs_mod4

        monto_total = (((jefes_mod1 * 33000) + (subs_mod1 * 28000)) * 1 +
                       ((jefes_mod2 * 33000) + (subs_mod2 * 28000)) * 2 +
                       ((jefes_mod3 * 33000) + (subs_mod3 * 28000)) * 3 +
                       ((jefes_mod4 * 33000) + (subs_mod4 * 28000)) * 4)

        st.markdown("---")
        st.success(f"💰 **Monto Total a Pagar: ${monto_total:,.2f}**")
        
        with st.form("form_envio"):
            st.subheader("💳 Datos de Facturación y Pago")
            mismo_organizador = st.checkbox("¿El que paga es el mismo organizador?", value=True)
            
            if mismo_organizador:
                pagado_por = organizador
                cuil_pagado_por = st.text_input("CUIL / CUIT de Facturación (del Organizador)")
            else:
                col_p1, col_p2 = st.columns(2)
                with col_p1: pagado_por = st.text_input("Nombre y DNI de quien paga")
                with col_p2: cuil_pagado_por = st.text_input("CUIL / CUIT de Facturación")
                
            st.markdown("---")
            st.subheader("📎 Documentación (Archivos)")
            file_hab = st.file_uploader("Habilitación Municipal (Obligatorio)", type=["png", "jpg", "jpeg", "pdf"])
            file_comp = st.file_uploader("Comprobante de Pago (Obligatorio)", type=["png", "jpg", "jpeg", "pdf"])
            file_cont = st.file_uploader("Contrato de Adicional (Opcional)", type=["png", "jpg", "jpeg", "pdf"])
            
            submitted = st.form_submit_button("Enviar Solicitud")
            if submitted:
                if total_jefes == 0 and total_subalternos == 0:
                    st.error("⚠️ Debes asignar al menos un efectivo (Jefe o Subalterno) al servicio.")
                elif file_hab and file_comp and organizador and tipo_evento and pagado_por and cuil_pagado_por and ubicacion_predio:
                    try:
                        with st.spinner("Subiendo archivos a la nube y registrando solicitud..."):
                            id_servicio = f"SERV-{int(datetime.datetime.now().timestamp())}"
                            
                            payload = {
                                "id_servicio": id_servicio,
                                "dependencia": str(st.session_state.dependencia),
                                "fecha_evento": str(fecha),
                                "tipo_evento": str(tipo_evento),
                                "organizador": str(organizador),
                                "ubicacion": str(ubicacion_predio),
                                "horario": f"{hora_inicio.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                                "cant_jefes": int(total_jefes),
                                "cant_subalternos": int(total_subalternos),
                                "modulos": int(max_modulos),
                                "monto_total": float(monto_total),
                                "pagado_por": str(pagado_por),
                                "cuil_pagado_por": str(cuil_pagado_por),
                                "file_hab_base64": base64.b64encode(file_hab.getvalue()).decode('utf-8'),
                                "name_hab": f"{id_servicio}_habilitacion.{file_hab.name.split('.')[-1]}",
                                "type_hab": file_hab.type,
                                "file_comp_base64": base64.b64encode(file_comp.getvalue()).decode('utf-8'),
                                "name_comp": f"{id_servicio}_comprobante.{file_comp.name.split('.')[-1]}",
                                "type_comp": file_comp.type,
                                "file_cont_base64": base64.b64encode(file_cont.getvalue()).decode('utf-8') if file_cont else "",
                                "name_cont": f"{id_servicio}_contrato.{file_cont.name.split('.')[-1]}" if file_cont else "",
                                "type_cont": file_cont.type if file_cont else ""
                            }
                            
                            response = requests.post(URL_APPS_SCRIPT, json=payload)
                            res_json = response.json()
                            
                            if res_json.get("status") == "success":
                                st.success("¡Solicitud y archivos enviados con éxito a tu Google Drive y Sheet!")
                            else:
                                st.error(f"Error en Apps Script: {res_json.get('message')}")
                    except Exception as e:
                        st.error(f"Error al enviar: {e}")
                else:
                    st.error("Completá todos los campos obligatorios (incluyendo CUIL y tipo de evento) y adjuntá los archivos.")

    with tab2:
        st.header("Servicios Pendientes de Planilla de Asistencia")
        st.info(f"Mostrando servicios de: **{st.session_state.dependencia}**")
        
        try:
            client = conectar_google()
            sheet = client.open("Gestion_Adicionales").worksheet("Servicios") if "Servicios" in [w.title for w in client.open("Gestion_Adicionales").worksheets()] else client.open("Gestion_Adicionales").sheet1
            data = pd.DataFrame(sheet.get_all_records())
            
            if not data.empty and 'Dependencia' in data.columns:
                df_dep = data[data['Dependencia'].astype(str) == str(st.session_state.dependencia)]
                col_asist = 'Enlace Asistencia' if 'Enlace Asistencia' in df_dep.columns else df_dep.columns[16] if len(df_dep.columns) > 16 else None
                
                if col_asist:
                    df_pendientes = df_dep[df_dep[col_asist].astype(str).str.strip() == ""]
                    
                    if df_pendientes.empty:
                        st.success("🎉 ¡No tenés servicios pendientes de asistencia!")
                    else:
                        st.warning(f"Tenés **{len(df_pendientes)}** servicio(s) pendiente(s) de adjuntar planilla.")
                        
                        for idx, row in df_pendientes.iterrows():
                            with st.expander(f"📅 {row['Fecha Evento']} - {row['Tipo de Evento']} ({row['Ubicación Predio']}) [ID: {row['ID Servicio']}]"):
                                st.write(f"**Organizador:** {row['Organizador']}")
                                st.write(f"**Horario:** {row['Horario']} | **Módulos:** {row['Módulos']}")
                                st.write(f"**Personal:** {row['Cant. Jefes']} Jefes / {row['Cant. Subalternos']} Subalternos")
                                
                                with st.form(f"form_asist_{row['ID Servicio']}"):
                                    file_asist = st.file_uploader("Adjuntar Planilla de Asistencia Firmada", type=["png", "jpg", "jpeg", "pdf"], key=f"file_{row['ID Servicio']}")
                                    sub_asist = st.form_submit_button("Subir Planilla de Asistencia")
                                    
                                    if sub_asist:
                                        if file_asist:
                                            with st.spinner("Subiendo planilla..."):
                                                payload_asist = {
                                                    "action": "subir_asistencia",
                                                    "id_servicio": str(row['ID Servicio']),
                                                    "file_asist_base64": base64.b64encode(file_asist.getvalue()).decode('utf-8'),
                                                    "name_asist": f"{row['ID Servicio']}_asistencia.{file_asist.name.split('.')[-1]}",
                                                    "type_asist": file_asist.type
                                                }
                                                resp = requests.post(URL_APPS_SCRIPT, json=payload_asist)
                                                res_j = resp.json()
                                                if res_j.get("status") == "success":
                                                    st.success("¡Planilla subida y registrada con éxito!")
                                                    st.rerun()
                                                else:
                                                    st.error(f"Error: {res_j.get('message')}")
                                        else:
                                            st.error("Seleccioná un archivo antes de enviar.")
                else:
                    st.info("No se encontró la columna de asistencia en la hoja.")
            else:
                st.info("Aún no hay servicios registrados en la base de datos.")
        except Exception as e:
            st.error(f"Error al cargar los pendientes: {e}")

    with tab3:
        st.header("Historial General de Servicios")
        st.info(f"Listado completo de servicios registrados por **{st.session_state.dependencia}**.")
        try:
            client = conectar_google()
            sheet = client.open("Gestion_Adicionales").worksheet("Servicios") if "Servicios" in [w.title for w in client.open("Gestion_Adicionales").worksheets()] else client.open("Gestion_Adicionales").sheet1
            data = pd.DataFrame(sheet.get_all_records())
            if not data.empty and 'Dependencia' in data.columns:
                df_dep = data[data['Dependencia'].astype(str) == str(st.session_state.dependencia)]
                if not df_dep.empty:
                    st.dataframe(df_dep, use_container_width=True)
                else:
                    st.info("No hay registros en el historial.")
            else:
                st.info("Aún no hay datos cargados.")
        except Exception as e:
            st.error(f"Error al cargar historial: {e}")

def vista_admin():
    st.sidebar.write("👑 **Administrador General**")
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Administrador")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()
        
    st.title("📋 Panel de Control General - Administrador")
    
    try:
        client = conectar_google()
        sheet = client.open("Gestion_Adicionales").worksheet("Servicios") if "Servicios" in [w.title for w in client.open("Gestion_Adicionales").worksheets()] else client.open("Gestion_Adicionales").sheet1
        data = pd.DataFrame(sheet.get_all_records())
        
        if not data.empty:
            total_servicios = len(data)
            monto_recaudado = data['Monto Total'].sum() if 'Monto Total' in data.columns else 0
            
            col_m1, col_m2 = st.columns(2)
            with col_m1: st.metric("Total de Servicios Registrados", total_servicios)
            with col_m2: st.metric("Recaudación Total Consolidada", f"${monto_recaudado:,.2f}")
            
            st.markdown("---")
            st.subheader("Gestión y Generación de Informes en PDF con Constancias")
            
            dependencias_disponibles = ["Todas"] + list(data['Dependencia'].unique()) if 'Dependencia' in data.columns else ["Todas"]
            dep_seleccionada = st.selectbox("Filtrar por Dependencia", options=dependencias_disponibles)
            
            if dep_seleccionada != "Todas":
                data_filtrada = data[data['Dependencia'].astype(str) == str(dep_seleccionada)]
            else:
                data_filtrada = data
                
            st.dataframe(data_filtrada, use_container_width=True)
            
            st.markdown("---")
            st.subheader("📥 Descargar o Compartir Informe PDF (Con imágenes de Drive)")
            
            if not data_filtrada.empty:
                opciones_servicios = [f"{row['ID Servicio']} - {row['Dependencia']} ({row['Fecha Evento']} - {row['Tipo de Evento']})" for idx, row in data_filtrada.iterrows()]
                servicio_elegido = st.selectbox("Seleccioná un servicio para generar el PDF completo", options=opciones_servicios)
                
                if servicio_elegido:
                    id_seleccionado = servicio_elegido.split(" - ")[0]
                    row_data = data_filtrada[data_filtrada['ID Servicio'].astype(str) == id_seleccionado].iloc[0]
                    
                    if st.button("🚀 Generar PDF con Constancias Adjuntas", type="primary"):
                        with st.spinner("Descargando archivos desde Google Drive y armando PDF..."):
                            pdf_path = generar_pdf_servicio(row_data)
                            with open(pdf_path, "rb") as pdf_file:
                                pdf_bytes = pdf_file.read()
                                
                        st.download_button(
                            label="📥 Descargar PDF Final (Listo para enviar por WhatsApp)",
                            data=pdf_bytes,
                            file_name=f"{id_seleccionado}_informe_completo.pdf",
                            mime="application/pdf",
                            type="primary"
                        )
        else:
            st.info("No hay servicios cargados en el sistema actualmente.")
            
    except Exception as e:
        st.error(f"Error al cargar el panel de administración: {e}")

if st.session_state.get("cambio_clave_requerido", False):
    pantalla_cambio_clave()
elif not st.session_state.autenticado:
    login()
else:
    if st.session_state.rol == "Administrador":
        vista_admin()
    else:
        vista_dependencia()
