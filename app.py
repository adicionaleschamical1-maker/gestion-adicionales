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

# --- ESTILOS VISUALES MEJORADOS (ALTA VISIBILIDAD DE INPUTS) ---
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

    /* Cuadros de texto con fondo blanco puro y borde bien visible */
    input, textarea, select, div[data-baseweb="input"] > div, div[data-baseweb="base-input"] {
        background-color: #FFFFFF !important;
        border: 1.5px solid #B0BEC5 !important;
        border-radius: 8px !important;
        color: #1a1a1a !important;
    }
    
    div[data-baseweb="input"]:focus-within, div[data-baseweb="base-input"]:focus-within {
        border-color: #002060 !important;
        box-shadow: 0 0 0 2px rgba(0,32,96,0.2) !important;
    }

    /* Etiquetas altamente legibles y oscuras */
    label, .stTextInput label, .stDateInput label, .stTimeInput label, .stSelectbox label, .stFileUploader label {
        color: #0f172a !important;
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
    "nombre_apellido": None,
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


# --- SALUDO DINÁMICO (EXCLUSIVAMENTE POR EL NOMBRE DE PILA) ---
def obtener_saludo(nombre_completo: str) -> str:
    hora_actual = datetime.datetime.now().hour
    partes = str(nombre_completo).strip().split()
    
    if len(partes) > 0:
        jerarquias_comunes = ["comisario", "inspector", "oficial", "suboficial", "cabo", "sargento", "agente", "subcomisario", "principal"]
        if partes[0].lower() in jerarquias_comunes and len(partes) > 1:
            nombre_pila = partes[1]
        else:
            nombre_pila = partes[-1] if len(partes) > 1 else partes[0]
    else:
        nombre_pila = "Efectivo"
    
    if 6 <= hora_actual < 12:
        return f"🌅 ¡Buen día, {nombre_pila}!"
    elif 12 <= hora_actual < 20:
        return f"☀️ ¡Buenas tardes, {nombre_pila}!"
    else:
        return f"🌙 ¡Buenas noches, {nombre_pila}!"


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
            hoja.update_cell(cell.row, 6, nueva)
            hoja.update_cell(cell.row, 5, "")
            st.success("¡Contraseña actualizada! Iniciá sesión nuevamente.")
            st.session_state.cambio_clave_requerido = False
            st.session_state.dni_temp = None
            st.rerun()
        except Exception as e:
            logger.exception("Error actualizando clave")
            st.error(f"Error al actualizar: {e}")


def login():
    st.markdown("### 🔐 Acceso al Sistema de Adicionales")
    with st.form("form_login"):
        dni = st.text_input("DNI / Usuario")
        clave = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("INGRESAR AL SISTEMA")
        
    if not submitted:
        return

    try:
        data = cargar_usuarios_df()
    except Exception as e:
        st.error(f"Error al consultar la base de datos: {e}")
        return

    data.columns = [str(c).strip().lower() for c in data.columns]
    
    user = data[data["dni"].astype(str) == str(dni)]
    if user.empty:
        st.error("DNI no registrado.")
        return

    u = user.iloc[0]
    clave_pers = str(u.get("clave personal", u.get("clave_personal", ""))).strip()
    clave_temp = str(u.get("clave", u.get("clave_temporal", ""))).strip()

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
        st.session_state.dependencia = str(u.get("dependencia", ""))
        st.session_state.rol = str(u.get("rol", "")).strip()
        st.session_state.nombre_apellido = str(u.get("nombre y apellido", u.get("nombre_apellido", f"DNI: {dni}"))).strip()
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


# --- GENERADOR DE PDF CON TABLAS Y DISEÑO PROFESIONAL ---
def generar_pdf_servicio(row: pd.Series) -> str:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_fill_color(0, 32, 96)
    pdf.rect(0, 0, 210, 20, 'F')
    
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 10)
    pdf.set_xy(5, 6)
    pdf.cell(200, 6, "POLICÍA DE LA PROVINCIA DE LA RIOJA - UNIDAD REGIONAL QUINTA - CHAMICAL", ln=True, align="center")
    
    pdf.set_text_color(0, 0, 0)
    pdf.ln(12)

    pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 7, "INFORME DE SERVICIO ADICIONAL", ln=True, align="center")
    pdf.set_font("helvetica", "I", 9)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(0, 5, f"ID de Solicitud: {row.get('ID Servicio', '')}", ln=True, align="center")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)

    def fila_tabla(lbl1, val1, lbl2, val2, w1=35, w2=60, w3=35, w4=60):
        pdf.set_font("helvetica", "B", 9)
        pdf.set_fill_color(245, 247, 250)
        pdf.cell(w1, 6, f" {lbl1}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9)
        pdf.cell(w2, 6, f" {val1}", 1, 0, 'L')
        pdf.set_font("helvetica", "B", 9)
        pdf.set_fill_color(245, 247, 250)
        pdf.cell(w3, 6, f" {lbl2}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9)
        pdf.cell(w4, 6, f" {val2}", 1, 1, 'L')

    def fila_completa(lbl, val, w_lbl=45, w_val=145):
        pdf.set_font("helvetica", "B", 9)
        pdf.set_fill_color(245, 247, 250)
        pdf.cell(w_lbl, 6, f" {lbl}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9)
        pdf.cell(w_val, 6, f" {val}", 1, 1, 'L')

    def seccion_titulo(titulo):
        pdf.ln(2)
        pdf.set_font("helvetica", "B", 10)
        pdf.set_fill_color(0, 32, 96)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(190, 6, f"  {titulo}", 1, 1, 'L', True)
        pdf.set_text_color(0, 0, 0)

    seccion_titulo("1. DATOS GENERALES DEL EVENTO")
    fila_tabla("Dependencia", str(row.get('Dependencia', '')), "Fecha", str(row.get('Fecha Evento', '')))
    fila_tabla("Tipo de Evento", str(row.get('Tipo de Evento', '')), "Horario", str(row.get('Horario', '')))
    fila_completa("Ubicación / Predio", str(row.get('Ubicación Predio', '')))
    fila_completa("Organizador", str(row.get('Organizador', '')))

    seccion_titulo("2. RECURSOS Y MÓDULOS ASIGNADOS")
    monto_val = row.get('Monto Total', 0)
    monto_str = f"${float(monto_val):,.2f}" if monto_val else "$0.00"
    fila_tabla("Cant. Jefes", str(row.get('Cant. Jefes', '')), "Cant. Subalternos", str(row.get('Cant. Subalternos', '')))
    fila_tabla("Módulos (Turnos)", str(row.get('Módulos', '')), "Monto Total", monto_str)

    seccion_titulo("3. DATOS DE FACTURACIÓN Y PAGO")
    fila_completa("Pagado por", str(row.get('Pagador', '')))
    fila_completa("CUIL / CUIT Facturación", str(row.get('CUIL Pagador', '')))

    seccion_titulo("4. DOCUMENTACIÓN Y ENLACES (GOOGLE DRIVE)")
    enlaces = [
        ("Habilitación Municipal", row.get("Enlace Habilitación", "")),
        ("Comprobante de Pago", row.get("Enlace Comprobante", "")),
        ("Contrato de Adicional", row.get("Enlace Contrato", "")),
        ("Planilla de Asistencia", row.get("Enlace Asistencia", ""))
    ]
    
    pdf.set_font("helvetica", "B", 9)
    for etiqueta, url in enlaces:
        pdf.set_fill_color(245, 247, 250)
        pdf.cell(50, 6, f" {etiqueta}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9)
        if url and str(url).startswith("http"):
            pdf.set_text_color(0, 102, 204)
            pdf.cell(140, 6, " [ Enlace Disponible / Ver Documento ]", 1, 1, 'L', link=str(url))
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.set_text_color(120, 120, 120)
            pdf.cell(140, 6, " No disponible", 1, 1, 'L')
            pdf.set_text_color(0, 0, 0)
        pdf.set_font("helvetica", "B", 9)

    anexos = [
        ("ANEXO: HABILITACIÓN MUNICIPAL", row.get("Enlace Habilitación", "")),
        ("ANEXO: COMPROBANTE DE PAGO", row.get("Enlace Comprobante", "")),
        ("ANEXO: CONTRATO DE ADICIONAL", row.get("Enlace Contrato", "")),
        ("ANEXO: PLANILLA DE ASISTENCIA", row.get("Enlace Asistencia", ""))
    ]
    
    for titulo, url in anexos:
        ruta_img = obtener_imagen_desde_drive(url)
        if not ruta_img:
            continue
        try:
            pdf.add_page()
            pdf.set_font("helvetica", "B", 11)
            pdf.set_text_color(0, 32, 96)
            pdf.cell(0, 8, titulo, ln=True, align="center")
            pdf.set_text_color(0, 0, 0)
            pdf.ln(3)
            
            with Image.open(ruta_img) as im:
                w, h = im.size
            ratio = min(170 / w, 220 / h)
            pdf.image(ruta_img, x=20, y=28, w=w * ratio, h=h * ratio)
        except Exception as e:
            logger.warning(f"PDF: no se pudo embeber {titulo}: {e}")
        finally:
            if os.path.exists(ruta_img):
                os.unlink
