import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import math
import traceback
import os
import requests
import base64

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestión de Servicios Adicionales", page_icon="🛡️", layout="centered")

# URL DE TU APLICACIÓN WEB DE APPS SCRIPT CONFIGURADA
URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxqVp5T0fVGtPncuwkocbXbU5TSHLgPqgmD0A5pQ6CmuEU3C-UrGx2eZHxg2HTLUjU1/exec"

@st.cache_resource
def conectar_google():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        ruta_json = os.path.join(os.path.dirname(__file__), "credenciales.json")
        if not os.path.exists(ruta_json):
            ruta_json = "credenciales.json"
            
        if not os.path.exists(ruta_json):
            st.error("🚨 **ARCHIVO DE CREDENCIALES NO ENCONTRADO**")
            st.warning(f"No se encontró `credenciales.json` en la carpeta del proyecto.")
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

def vista_dependencia():
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Colaborador")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()
        
    tab1, tab2 = st.tabs(["🟢 Cargar Nuevo Servicio", "🔵 Pendientes de Asistencia"])
    
    with tab1:
        st.header(f"Solicitud de Servicio Adicional - {st.session_state.dependencia}")
        fecha = st.date_input("Fecha del Evento", min_value=datetime.date.today())
        
        eventos_habituales = ["Partido de Fútbol", "Baile / Recital", "Torneo Deportivo", "Evento Privado / Salón", "Acto Público", "Reunión Comercial"]
        tipo_evento = st.selectbox("Tipo de Evento", options=eventos_habituales, index=0)
        organizador = st.text_input("Nombre / Razón Social del Organizador")
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
            pagado_por = st.text_input("Nombre y DNI de quien hizo el Pago")
            st.markdown("---")
            st.subheader("📎 Documentación (Archivos)")
            file_hab = st.file_uploader("Habilitación Municipal (Obligatorio)", type=["png", "jpg", "jpeg", "pdf"])
            file_comp = st.file_uploader("Comprobante de Pago (Obligatorio)", type=["png", "jpg", "jpeg", "pdf"])
            file_cont = st.file_uploader("Contrato de Adicional (Opcional)", type=["png", "jpg", "jpeg", "pdf"])
            
            submitted = st.form_submit_button("Enviar Solicitud")
            if submitted:
                if file_hab and file_comp and organizador and pagado_por and ubicacion_predio:
                    try:
                        with st.spinner("Subiendo archivos a la nube y registrando solicitud..."):
                            id_servicio = f"SERV-{int(datetime.datetime.now().timestamp())}"
                            
                            payload = {
                                "id_servicio": id_servicio,
                                "dependencia": str(st.session_state.dependencia),
                                "fecha_evento": str(fecha),
                                "organizador": f"{tipo_evento} - {ubicacion_predio}",
                                "horario": f"{hora_inicio.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                                "cant_jefes": int(total_jefes),
                                "cant_subalternos": int(total_subalternos),
                                "modulos": int(max_modulos),
                                "monto_total": float(monto_total),
                                "pagado_por": str(pagado_por),
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
                    st.error("Completá todos los campos obligatorios y adjuntá los archivos.")

    with tab2:
        st.header("Servicios Pendientes de Planilla de Asistencia")
        st.info("Aquí aparecerán los servicios para adjuntar la asistencia.")

def vista_admin():
    st.sidebar.write("👑 **Administrador General**")
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Administrador")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()
    st.title("📋 Panel de Control General")
    st.info("Vista global y consolidada.")

if st.session_state.get("cambio_clave_requerido", False):
    pantalla_cambio_clave()
elif not st.session_state.autenticado:
    login()
else:
    if st.session_state.rol == "Administrador":
        vista_admin()
    else:
        vista_dependencia()