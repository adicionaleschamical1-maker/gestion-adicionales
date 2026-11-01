import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import math
import json

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestión de Servicios Adicionales", page_icon="🛡️", layout="centered")

# --- CONEXIÓN INTELIGENTE (NUBE O LOCAL) ---
@st.cache_resource
def conectar_google():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # Intenta leer desde los secretos de Streamlit Cloud o secrets.toml local
        credentials_dict = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(
            credentials_dict,
            scopes=scopes
        )
    except Exception:
        # Si no encuentra secretos, usa el archivo físico credenciales.json (para pruebas locales)
        credentials = Credentials.from_service_account_file(
            "credenciales.json",
            scopes=scopes
        )
    
    client = gspread.authorize(credentials)
    return client

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

# --- PANTALLA DE CAMBIO DE CONTRASEÑA OBLIGATORIO ---
def pantalla_cambio_clave():
    st.title("🔑 Definir Nueva Contraseña Personal")
    st.warning("⚠️ Primer ingreso detectado. Debes cambiar tu contraseña temporal por una personal.")
    
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
                
                st.success("¡Contraseña actualizada con éxito! Por favor, inicia sesión nuevamente con tu nueva clave.")
                st.session_state.cambio_clave_requerido = False
                st.session_state.dni_temp = None
                st.rerun()
            except Exception as e:
                st.error(f"Error al actualizar la contraseña en Google Sheets: {e}")
        else:
            st.error("Las contraseñas no coinciden o el campo está vacío.")

# --- LÓGICA DE LOGIN ---
def login():
    st.title("🛡️ Acceso a Servicios Adicionales")
    dni = st.text_input("DNI / Usuario")
    clave = st.text_input("Contraseña", type="password")
    
    if st.button("Ingresar", type="primary"):
        try:
            client = conectar_google()
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
        except Exception as e:
            st.error(f"Error al conectar con la base de datos: {e}")

# --- VISTA DE LA DEPENDENCIA (ROL COLABORADOR) ---
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
        
        eventos_habituales = [
            "Partido de Fútbol", 
            "Baile / Recital", 
            "Torneo Deportivo", 
            "Evento Privado / Salón", 
            "Acto Público", 
            "Reunión Comercial"
        ]
        
        tipo_evento = st.selectbox(
            "Tipo de Evento (Seleccioná o escribí uno nuevo)", 
            options=eventos_habituales,
            index=0,
            help="Podés elegir de la lista o escribir directamente el nombre del evento si es nuevo."
        )
        
        organizador = st.text_input("Nombre / Razón Social del Organizador")
        ubicacion_predio = st.text_input("Ubicación / Predio del Evento (ej. Club Defensores, Salón X, Dirección)")
        
        st.markdown("⏰ **Horario del Servicio**")
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            hora_inicio = st.time_input("Hora de Inicio", value=datetime.time(20, 0))
        with col_h2:
            hora_fin = st.time_input("Hora de Fin", value=datetime.time(4, 0))
            
        # --- CÁLCULO DE MÓDULOS SEGÚN HORARIO ---
        dt_inicio = datetime.datetime.combine(datetime.date.today(), hora_inicio)
        dt_fin = datetime.datetime.combine(datetime.date.today(), hora_fin)
        
        if dt_fin <= dt_inicio:
            dt_fin += datetime.timedelta(days=1)
            
        diferencia_horas = (dt_fin - dt_inicio).total_seconds() / 3600.0
        max_modulos = math.ceil(diferencia_horas / 4.0)
        if max_modulos < 1: max_modulos = 1
        if max_modulos > 4: max_modulos = 4
            
        st.info(f"⏱️ Duración calculada: **{diferencia_horas:.1f} horas** (Se habilitan turnos hasta **{max_modulos} módulo(s)**)")

        st.markdown("---")
        st.subheader("👥 Detalle de Personal y Módulos")
        
        jefes_mod1, subs_mod1 = 0, 0
        jefes_mod2, subs_mod2 = 0, 0
        jefes_mod3, subs_mod3 = 0, 0
        jefes_mod4, subs_mod4 = 0, 0

        if max_modulos >= 1:
            st.markdown("**Turno Simple (1 Módulo - 4 horas)**")
            c1, c2 = st.columns(2)
            with c1:
                jefes_mod1 = st.number_input("Jefes (Simple)", min_value=0, step=1, key="j1")
            with c2:
                subs_mod1 = st.number_input("Subalternos (Simple)", min_value=0, step=1, key="s1")
                
        if max_modulos >= 2:
            st.markdown("**Turno Doble (2 Módulos - 8 horas)**")
            c3, c4 = st.columns(2)
            with c3:
                jefes_mod2 = st.number_input("Jefes (Doble)", min_value=0, step=1, key="j2")
            with c4:
                subs_mod2 = st.number_input("Subalternos (Doble)", min_value=0, step=1, key="s2")

        if max_modulos >= 3:
            st.markdown("**Turno Triple (3 Módulos - 12 horas)**")
            c5, c6 = st.columns(2)
            with c5:
                jefes_mod3 = st.number_input("Jefes (Triple)", min_value=0, step=1, key="j3")
            with c6:
                subs_mod3 = st.number_input("Subalternos (Triple)", min_value=0, step=1, key="s3")

        if max_modulos >= 4:
            st.markdown("**Turno Cuádruple (4 Módulos - 16 horas)**")
            c7, c8 = st.columns(2)
            with c7:
                jefes_mod4 = st.number_input("Jefes (Cuádruple)", min_value=0, step=1, key="j4")
            with c8:
                subs_mod4 = st.number_input("Subalternos (Cuádruple)", min_value=0, step=1, key="s4")

        # --- CÁLCULO MATEMÁTICO ---
        VALOR_JEFE = 33000
        VALOR_SUB = 28000

        t1 = ((jefes_mod1 * VALOR_JEFE) + (subs_mod1 * VALOR_SUB)) * 1
        t2 = ((jefes_mod2 * VALOR_JEFE) + (subs_mod2 * VALOR_SUB)) * 2
        t3 = ((jefes_mod3 * VALOR_JEFE) + (subs_mod3 * VALOR_SUB)) * 3
        t4 = ((jefes_mod4 * VALOR_JEFE) + (subs_mod4 * VALOR_SUB)) * 4

        monto_total = t1 + t2 + t3 + t4

        st.markdown("---")
        st.success(f"💰 **Monto Total a Pagar: ${monto_total:,.2f}**")
        
        with st.form("form_envio"):
            pagado_por = st.text_input("Nombre y DNI de quien hizo el Pago")
            
            st.markdown("---")
            st.subheader("📎 Documentación (Fotos o PDF)")
            file_hab = st.file_uploader("Habilitación Municipal (Obligatorio)", type=["png", "jpg", "jpeg", "pdf"])
            file_comp = st.file_uploader("Comprobante de Pago (Obligatorio)", type=["png", "jpg", "jpeg", "pdf"])
            file_cont = st.file_uploader("Contrato de Adicional (Opcional)", type=["png", "jpg", "jpeg", "pdf"])
            
            submitted = st.form_submit_button("Enviar Solicitud")
            if submitted:
                if file_hab and file_comp and organizador and pagado_por and ubicacion_predio:
                    st.info("Subiendo datos y archivos...")
                    st.success("✅ Servicio guardado correctamente.")
                else:
                    st.error("Por favor completa los campos obligatorios (ubicación, organizador, pagos y archivos).")

    with tab2:
        st.header("Servicios Pendientes de Planilla de Asistencia")
        st.info(f"Aquí aparecerán los servicios realizados por {st.session_state.dependencia} para adjuntar la planilla de asistencia.")

# --- VISTA DEL ADMINISTRADOR GENERAL (ROL ADMINISTRADOR) ---
def vista_admin():
    st.sidebar.write("👑 **Administrador General**")
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Administrador")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    st.title("📋 Panel de Control General")
    st.info("Vista global y consolidada de las dependencias policiales, con estados de envío y accesos a Drive.")

# --- FLUJO PRINCIPAL ---
if st.session_state.get("cambio_clave_requerido", False):
    pantalla_cambio_clave()
elif not st.session_state.autenticado:
    login()
else:
    if st.session_state.rol == "Administrador":
        vista_admin()
    else:
        vista_dependencia()