import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import datetime
import math
import os
import re
import tempfile
import logging
import requests
import base64
from fpdf import FPDF
from PIL import Image
from io import BytesIO

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="URV V - Adicionales", layout="centered", page_icon="👮")

# --- ESTILOS VISUALES EXACTOS DE LA CAPTURA ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background: #F5F5F5; }
    
    /* Encabezado Institucional Azul con Borde Rojo */
    .header-container {
        background: linear-gradient(135deg, #002060, #002a80);
        color: white;
        padding: 2rem 1rem;
        border-radius: 20px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px rgba(0,32,96,0.3);
        border-bottom: 6px solid #8B1E1E;
    }
    
    /* Botón rojo ovalado "CHAMICAL" */
    .badge-chamical {
        background: #8B1E1E;
        display: inline-block;
        padding: 8px 30px;
        border-radius: 50px;
        font-size: 1rem;
        font-weight: 800;
        color: white;
        box-shadow: 0 4px 15px rgba(139,30,30,0.5);
        margin-top: 10px;
        letter-spacing: 1px;
    }
    
    /* Tarjetas de contenedores y formularios */
    div.stExpander, div.stForm {
        background: #FFFFFF;
        border-radius: 20px;
        padding: 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border-left: 6px solid #002060;
        border-top: 1px solid #E0E0E0;
        border-right: 1px solid #E0E0E0;
        border-bottom: 1px solid #E0E0E0;
    }

    /* Etiquetas legibles */
    label, .stTextInput label, .stDateInput label, .stTimeInput label, .stSelectbox label, .stFileUploader label {
        color: #1a1a1a !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    
    /* Pestañas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: white;
        padding: 5px;
        border-radius: 30px;
        border: 1px solid #E0E0E0;
        margin-bottom: 20px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 30px;
        padding: 8px 20px;
        font-weight: 500;
        color: #666666;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #002060, #002a80) !important;
        color: white !important;
        box-shadow: 0 2px 8px rgba(0,32,96,0.3);
    }
    
    /* Botones de acción azul oscuro institucionales */
    .stButton > button, div.stFormSubmitButton > button {
        border-radius: 30px !important;
        font-weight: 700 !important;
        background: linear-gradient(135deg, #002060, #002a80) !important;
        color: white !important;
        border: none !important;
        transition: all 0.2s !important;
        height: 3.2em !important;
        width: 100% !important;
        box-shadow: 0 4px 12px rgba(0,32,96,0.3);
    }
    
    .stButton > button:hover, div.stFormSubmitButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(0,32,96,0.4) !important;
    }
    
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #F8F9FA, #EFEFEF);
    }
    
    .user-card {
        background: white;
        border-radius: 20px;
        padding: 20px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border: 1px solid #E0E0E0;
    }
    
    .user-avatar {
        width: 56px;
        height: 56px;
        background: linear-gradient(135deg, #002060, #002a80);
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 12px;
        font-size: 1.8rem;
        color: white;
        box-shadow: 0 4px 12px rgba(0,32,96,0.3);
    }
    </style>

    <div class="header-container">
        <div style="font-size: 3.5rem; margin-bottom: 5px;">👮‍♂️</div>
        <div style="font-size: 1.3rem; font-weight: 800; letter-spacing: 1px;">POLICÍA DE LA PROVINCIA DE LA RIOJA</div>
        <div style="font-size: 1.1rem; font-weight: 700; color: #ffd700; margin: 5px 0;">UNIDAD REGIONAL QUINTA - CHAMICAL</div>
        <div style="font-size: 0.95rem; font-weight: 600; margin-bottom: 10px; opacity: 0.95;">OFICINA DE ADICIONALES</div>
        <div class="badge-chamical">CHAMICAL</div>
    </div>
""", unsafe_allow_html=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Constantes de negocio
SHEET_NAME = "Gestion_Adicionales"
HOJA_SERVICIOS = "Servicios"
HOJA_USUARIOS = "Usuarios"
PRECIO_JEFE = 33000
PRECIO_SUB = 28000
HORAS_POR_MODULO = 4
TIMEOUT_HTTP = 30

# URL de Apps Script
try:
    URL_APPS_SCRIPT = st.secrets["URL_APPS_SCRIPT"]
except (KeyError, FileNotFoundError):
    URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxqVp5T0fVGtPncuwkocbXbU5TSHLgPqgmD0A5pQ6CmuEU3C-UrGx2eZHxg2HTLUjU1/exec"


# --- CONEXIÓN GOOGLE ---
@st.cache_resource
def conectar_google():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        ruta = os.path.join(os.path.dirname(__file__), "credenciales.json")
        if not os.path.exists(ruta):
            ruta = "credenciales.json"
        if not os.path.exists(ruta):
            raise FileNotFoundError("No se encontró credenciales.json ni st.secrets['gcp_service_account']")
        credentials = Credentials.from_service_account_file(ruta, scopes=scopes)

    client = gspread.authorize(credentials)
    drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    return client, drive_service


def abrir_hoja(nombre: str):
    client, _ = conectar_google()
    ss = client.open(SHEET_NAME)
    try:
        return ss.worksheet(nombre)
    except gspread.WorksheetNotFound:
        return ss.sheet1


@st.cache_data(ttl=60, show_spinner=False)
def cargar_servicios_df() -> pd.DataFrame:
    hoja = abrir_hoja(HOJA_SERVICIOS)
    return pd.DataFrame(hoja.get_all_records())


@st.cache_data(ttl=60, show_spinner=False)
def cargar_usuarios_df() -> pd.DataFrame:
    hoja = abrir_hoja(HOJA_USUARIOS)
    return pd.DataFrame(hoja.get_all_records())


def boton_refrescar():
    if st.sidebar.button("🔄 Refrescar datos"):
        st.cache_data.clear()
        st.rerun()


# --- ESTADO DE SESIÓN ---
_defaults = {
    "autenticado": False,
    "usuario_dni": None,
    "dependencia": None,
    "rol": None,
    "cambio_clave_requerido": False,
    "form_counter": 0,
}
for k, v in _defaults.items():
    st.session_state.setdefault(k, v)


# --- HELPERS ---
def validar_cuil(valor: str) -> bool:
    return bool(re.fullmatch(r"\d{11}", str(valor).strip()))


def calcular_monto_y_modulos(hora_inicio: datetime.time, hora_fin: datetime.time):
    dt_ini = datetime.datetime.combine(datetime.date.today(), hora_inicio)
    dt_fin = datetime.datetime.combine(datetime.date.today(), hora_fin)
    if dt_fin <= dt_ini:
        dt_fin += datetime.timedelta(days=1)
    horas = (dt_fin - dt_ini).total_seconds() / 3600.0
    max_mod = max(1, min(4, math.ceil(horas / HORAS_POR_MODULO)))
    return horas, max_mod


def calcular_monto_total(jefes_por_mod: list, subs_por_mod: list) -> float:
    total = 0
    for i, (j, s) in enumerate(zip(jefes_por_mod, subs_por_mod), start=1):
        total += ((j * PRECIO_JEFE) + (s * PRECIO_SUB)) * i
    return total


# --- LOGIN ---
def pantalla_cambio_clave():
    st.title("🔑 Definir Nueva Contraseña Personal")
    st.warning("⚠️ Primer ingreso detectado. Debés cambiar tu contraseña temporal.")
    nueva = st.text_input("Nueva Contraseña", type="password")
    confirm = st.text_input("Confirmar Nueva Contraseña", type="password")
    if st.button("Guardar Nueva Contraseña"):
        if not nueva or nueva != confirm:
            st.error("Las contraseñas no coinciden o están vacías.")
            return
        try:
            hoja = abrir_hoja(HOJA_USUARIOS)
            cell = hoja.find(str(st.session_state.dni_temp))
            hoja.update_cell(cell.row, 3, nueva)
            hoja.update_cell(cell.row, 4, "")
            st.success("¡Contraseña actualizada! Iniciá sesión nuevamente.")
            st.session_state.cambio_clave_requerido = False
            st.session_state.dni_temp = None
            st.rerun()
        except Exception as e:
            logger.exception("Error actualizando clave")
            st.error(f"Error al actualizar: {e}")


def login():
    st.subheader("🔐 Acceso al Sistema")
    dni = st.text_input("DNI / Usuario")
    clave = st.text_input("Contraseña", type="password")
    if not st.button("INGRESAR"):
        return
    try:
        data = cargar_usuarios_df()
    except Exception as e:
        st.error(f"Error al consultar la base de datos: {e}")
        return

    user = data[data["dni"].astype(str) == str(dni)]
    if user.empty:
        st.error("DNI no registrado.")
        return

    u = user.iloc[0]
    clave_pers = str(u["clave_personal"]).strip()
    clave_temp = str(u["clave_temporal"]).strip()

    requiere_cambio = (
        clave_pers == ""
        or clave_pers == str(dni)
        or (clave_temp != "" and clave == clave_temp)
    )
    if requiere_cambio:
        if clave == clave_temp or clave == str(dni) or clave == clave_pers:
            st.session_state.cambio_clave_requerido = True
            st.session_state.dni_temp = dni
            st.rerun()
        else:
            st.error("Contraseña temporal incorrecta.")
        return

    if clave == clave_pers:
        st.session_state.autenticado = True
        st.session_state.usuario_dni = dni
        st.session_state.dependencia = u["dependencia"]
        st.session_state.rol = str(u["rol"]).strip()
        st.rerun()
    else:
        st.error("Contraseña incorrecta.")


# --- DRIVE: DESCARGA DE IMAGEN VÍA API ---
def obtener_imagen_desde_drive(url_drive: str):
    file_id = None
    if url_drive:
        s = str(url_drive)
        if "id=" in s:
            file_id = s.split("id=")[1].split("&")[0]
        elif "/d/" in s:
            file_id = s.split("/d/")[1].split("/")[0]

    if not file_id:
        return None

    tmp_path = None
    try:
        _, drive = conectar_google()
        content = drive.files().get_media(fileId=file_id).execute()
        if len(content) < 1000:
            return None

        image = Image.open(BytesIO(content))
        if image.mode in ("RGBA", "P"):
            image = image.convert("RGB")

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        tmp_path = tmp.name
        image.save(tmp_path, "JPEG")
        return tmp_path
    except Exception as e:
        logger.warning(f"Fallo descargando imagen Drive {file_id}: {e}")
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return None


# --- GENERADOR DE PDF ---
def generar_pdf_servicio(row: pd.Series) -> str:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, "INFORME DE SERVICIO ADICIONAL", ln=True, align="center")
    pdf.set_font("helvetica", "I", 10)
    pdf.cell(0, 6, f"ID de Solicitud: {row.get('ID Servicio', '')}", ln=True, align="center")
    pdf.ln(5)

    def seccion(titulo):
        pdf.set_font("helvetica", "B", 12)
        pdf.set_fill_color(230, 230, 230)
        pdf.cell(0, 8, f" {titulo}", ln=True, fill=True)
        pdf.set_font("helvetica", "", 10)

    seccion("1. DATOS GENERALES DEL EVENTO")
    pdf.cell(95, 6, f"Dependencia: {row.get('Dependencia', '')}", ln=0)
    pdf.cell(95, 6, f"Fecha: {row.get('Fecha Evento', '')}", ln=1)
    pdf.cell(95, 6, f"Tipo de Evento: {row.get('Tipo de Evento', '')}", ln=0)
    pdf.cell(95, 6, f"Horario: {row.get('Horario', '')}", ln=1)
    pdf.cell(0, 6, f"Ubicación / Predio: {row.get('Ubicación Predio', '')}", ln=1)
    pdf.cell(0, 6, f"Organizador: {row.get('Organizador', '')}", ln=1)
    pdf.ln(3)

    seccion("2. RECURSOS Y MÓDULOS ASIGNADOS")
    pdf.cell(95, 6, f"Cantidad de Jefes: {row.get('Cant. Jefes', '')}", ln=0)
    pdf.cell(95, 6, f"Cantidad de Subalternos: {row.get('Cant. Subalternos', '')}", ln=1)
    pdf.cell(95, 6, f"Módulos (Turnos): {row.get('Módulos', '')}", ln=0)
    pdf.cell(95, 6, f"Monto Total: ${row.get('Monto Total', 0):,.2f}", ln=1)
    pdf.ln(3)

    seccion("3. DATOS DE FACTURACIÓN Y PAGO")
    pdf.cell(0, 6, f"Pagado por: {row.get('Pagador', '')}", ln=1)
    pdf.cell(0, 6, f"CUIL / CUIT Facturación: {row.get('CUIL Pagador', '')}", ln=1)
    pdf.ln(3)

    seccion("4. ENLACES DIRECTOS A GOOGLE DRIVE")
    for titulo, key in [
        ("Habilitación", "Enlace Habilitación"),
        ("Comprobante", "Enlace Comprobante"),
        ("Contrato", "Enlace Contrato"),
        ("Asistencia", "Enlace Asistencia"),
    ]:
        url = row.get(key, "")
        val = url if url and str(url).startswith("http") else "No disponible"
        pdf.cell(0, 6, f"- {titulo}: {val}", ln=True)

    anexos = [
        ("CONSTANCIA: HABILITACIÓN MUNICIPAL", row.get("Enlace Habilitación", "")),
        ("CONSTANCIA: COMPROBANTE DE PAGO", row.get("Enlace Comprobante", "")),
        ("CONSTANCIA: CONTRATO DE ADICIONAL", row.get("Enlace Contrato", "")),
        ("CONSTANCIA: PLANILLA DE ASISTENCIA", row.get("Enlace Asistencia", "")),
    ]
    for titulo, url in anexos:
        ruta_img = obtener_imagen_desde_drive(url)
        if not ruta_img:
            continue
        try:
            pdf.add_page()
            pdf.set_font("helvetica", "B", 12)
            pdf.cell(0, 8, titulo, ln=True, align="center")
            pdf.ln(5)
            with Image.open(ruta_img) as im:
                w, h = im.size
            ratio = min(180 / w, 220 / h)
            pdf.image(ruta_img, x=15, y=30, w=w * ratio, h=h * ratio)
        except Exception as e:
            logger.warning(f"PDF: no se pudo embeber {titulo}: {e}")
            pdf.cell(0, 6, "(No se pudo renderizar la imagen del archivo)", ln=True, align="center")
        finally:
            if os.path.exists(ruta_img):
                os.unlink(ruta_img)

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name


# --- VISTA DEPENDENCIA ---
def _render_modulos(max_modulos: int, counter: int):
    labels = ["Simple (4h)", "Doble (8h)", "Triple (12h)", "Cuádruple (16h)"]
    jefes, subs = [], []
    for i in range(max_modulos):
        st.markdown(f"**Turno {labels[i]}**")
        c1, c2 = st.columns(2)
        with c1:
            j = st.number_input(f"Jefes ({labels[i]})", 0, step=1, key=f"j{i+1}_{counter}")
        with c2:
            s = st.number_input(f"Subalternos ({labels[i]})", 0, step=1, key=f"s{i+1}_{counter}")
        jefes.append(int(j))
        subs.append(int(s))
    return jefes + [0] * (4 - len(jefes)), subs + [0] * (4 - len(subs))


def _post_apps_script(payload: dict) -> dict:
    try:
        r = requests.post(URL_APPS_SCRIPT, json=payload, timeout=TIMEOUT_HTTP)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.Timeout:
        return {"status": "error", "message": "Timeout al contactar Apps Script."}
    except requests.exceptions.RequestException as e:
        return {"status": "error", "message": f"Error de red: {e}"}
    except ValueError:
        return {"status": "error", "message": "Respuesta inválida del servidor."}


def vista_dependencia():
    st.sidebar.markdown(f"""
        <div class="user-card">
            <div class="user-avatar">👮</div>
            <div style="font-weight:700; font-size: 1rem;">{st.session_state.dependencia}</div>
            <div style="font-size:0.75rem; margin-top:8px; background: #E8F0FE; padding:4px 12px; border-radius:30px; display:inline-block;">
                👤 COLABORADOR
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    boton_refrescar()
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    tab1, tab2, tab3 = st.tabs(
        ["🟢 Cargar Nuevo Servicio", "🔵 Pendientes de Asistencia", "📜 Historial de Servicios"]
    )

    counter = st.session_state.form_counter

    # ---------- TAB 1 ----------
    with tab1:
        st.header(f"Solicitud de Servicio Adicional")

        fecha = st.date_input("Fecha del Evento", min_value=datetime.date.today(), key=f"fecha_{counter}")
        st.subheader("📋 Datos del Evento")
        tipo_evento = st.text_input(
            "Tipo de Evento (Ej: Partido de fútbol, Baile, Recital, etc.)", key=f"tipo_{counter}"
        )
        
        col_org1, col_org2 = st.columns(2)
        with col_org1:
            nombre_organizador = st.text_input("Nombre y Apellido del Organizador", key=f"nom_org_{counter}")
        with col_org2:
            dni_organizador = st.text_input("DNI del Organizador", key=f"dni_org_{counter}")
            
        cuil_organizador = st.text_input("CUIL / CUIT del Organizador", key=f"cuil_org_{counter}")
        ubicacion_predio = st.text_input("Ubicación / Predio del Evento", key=f"ubi_{counter}")

        col_h1, col_h2 = st.columns(2)
        with col_h1:
            hora_inicio = st.time_input("Hora de Inicio", value=datetime.time(20, 0), key=f"h_ini_{counter}")
        with col_h2:
            hora_fin = st.time_input("Hora de Fin", value=datetime.time(4, 0), key=f"h_fin_{counter}")

        diferencia_horas, max_modulos = calcular_monto_y_modulos(hora_inicio, hora_fin)
        st.info(f"⏱️ Duración: **{diferencia_horas:.1f} horas** (Hasta **{max_modulos} módulo(s)**)")

        jefes_por_mod, subs_por_mod = _render_modulos(max_modulos, counter)
        total_jefes = sum(jefes_por_mod)
        total_subalternos = sum(subs_por_mod)
        monto_total = calcular_monto_total(jefes_por_mod, subs_por_mod)

        st.markdown("---")
        st.success(f"💰 **Monto Total a Pagar: ${monto_total:,.2f}**")

        with st.form(f"form_envio_{counter}", clear_on_submit=False):
            st.subheader("💳 Datos de Facturación y Pago")
            mismo_organizador = st.checkbox("¿El que paga es el mismo organizador?", value=True)
            if mismo_organizador:
                pagado_por = f"{nombre_organizador} (DNI: {dni_organizador})" if nombre_organizador else ""
                cuil_pagado_por = cuil_organizador
            else:
                cp1, cp2 = st.columns(2)
                with cp1:
                    pagado_por = st.text_input("Nombre y DNI de quien paga", key=f"pag_por_{counter}")
                with cp2:
                    cuil_pagado_por = st.text_input("CUIL / CUIT de Facturación", key=f"cuil_pag_{counter}")

            st.markdown("---")
            st.subheader("📎 Documentación (Archivos)")
            file_hab = st.file_uploader(
                "Habilitación Municipal (Obligatorio)",
                type=["png", "jpg", "jpeg", "pdf"],
                key=f"hab_{counter}",
            )
            file_comp = st.file_uploader(
                "Comprobante de Pago (Obligatorio)",
                type=["png", "jpg", "jpeg", "pdf"],
                key=f"comp_{counter}",
            )
            file_cont = st.file_uploader(
                "Contrato de Adicional (Opcional)",
                type=["png", "jpg", "jpeg", "pdf"],
                key=f"cont_{counter}",
            )

            submitted = st.form_submit_button("Enviar Solicitud")

        if submitted:
            organizador_completo = f"{nombre_organizador} (DNI: {dni_organizador})"
            if total_jefes == 0 and total_subalternos == 0:
                st.error("⚠️ Debés asignar al menos un efectivo (Jefe o Subalterno).")
            elif not (
                file_hab
                and file_comp
                and nombre_organizador
                and dni_organizador
                and tipo_evento
                and pagado_por
                and cuil_pagado_por
                and ubicacion_predio
            ):
                st.error("Completá todos los campos obligatorios y adjuntá los archivos.")
            elif not validar_cuil(cuil_pagado_por):
                st.error("CUIL/CUIT inválido: debe tener 11 dígitos numéricos.")
            else:
                try:
                    with st.spinner("Subiendo archivos y registrando solicitud..."):
                        id_servicio = f"SERV-{int(datetime.datetime.now().timestamp())}"
                        payload = {
                            "id_servicio": id_servicio,
                            "dependencia": str(st.session_state.dependencia),
                            "fecha_evento": str(fecha),
                            "tipo_evento": str(tipo_evento),
                            "organizador": str(organizador_completo),
                            "ubicacion": str(ubicacion_predio),
                            "horario": f"{hora_inicio.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                            "cant_jefes": int(total_jefes),
                            "cant_subalternos": int(total_subalternos),
                            "modulos": int(max_modulos),
                            "monto_total": float(monto_total),
                            "pagado_por": str(pagado_por),
                            "cuil_pagado_por": str(cuil_pagado_por),
                            "file_hab_base64": base64.b64encode(file_hab.getvalue()).decode("utf-8"),
                            "name_hab": f"{id_servicio}_habilitacion.{file_hab.name.split('.')[-1]}",
                            "type_hab": file_hab.type,
                            "file_comp_base64": base64.b64encode(file_comp.getvalue()).decode("utf-8"),
                            "name_comp": f"{id_servicio}_comprobante.{file_comp.name.split('.')[-1]}",
                            "type_comp": file_comp.type,
                            "file_cont_base64": base64.b64encode(file_cont.getvalue()).decode("utf-8") if file_cont else "",
                            "name_cont": f"{id_servicio}_contrato.{file_cont.name.split('.')[-1]}" if file_cont else "",
                            "type_cont": file_cont.type if file_cont else "",
                        }
                        res = _post_apps_script(payload)
                        if res.get("status") == "success":
                            st.success("¡Solicitud enviada con éxito!")
                            st.session_state.form_counter += 1
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error(f"Error en Apps Script: {res.get('message')}")
                except Exception as e:
                    logger.exception("Error enviando solicitud")
                    st.error(f"Error al enviar: {e}")

    # ---------- TAB 2: PENDIENTES DE ASISTENCIA ----------
    with tab2:
        st.header("Servicios Pendientes de Planilla de Asistencia")
        st.info(f"Mostrando servicios de: **{st.session_state.dependencia}**")
        try:
            data = cargar_servicios_df()
            if data.empty or "Dependencia" not in data.columns:
                st.info("Aún no hay servicios registrados.")
            else:
                df_dep = data[data["Dependencia"].astype(str) == str(st.session_state.dependencia)]
                
                col_asist = None
                for c in df_dep.columns:
                    if "asistencia" in str(c).lower():
                        col_asist = c
                        break
                if not col_asist and len(df_dep.columns) > 16:
                    col_asist = df_dep.columns[16]

                if not col_asist:
                    st.warning("⚠️ No se encontró la columna de asistencia en la hoja.")
                else:
                    df_pend = df_dep[df_dep[col_asist].astype(str).str.strip() == ""]
                    if df_pend.empty:
                        st.success("🎉 ¡No tenés servicios pendientes de asistencia!")
                    else:
                        st.warning(f"Tenés **{len(df_pend)}** servicio(s) pendiente(s) de asistencia.")
                        for _, row in df_pend.iterrows():
                            id_serv = row.get('ID Servicio') or row.get('id_servicio') or row.iloc[0]
                            tipo_ev = row.get('Tipo de
