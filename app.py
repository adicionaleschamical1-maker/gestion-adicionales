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

# --- ESTILOS VISUALES ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background: #F5F5F5; }
    
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

    label, .stTextInput label, .stDateInput label, .stTimeInput label, .stSelectbox label, .stFileUploader label {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    
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
    try:
        hoja = abrir_hoja(HOJA_SERVICIOS)
        return pd.DataFrame(hoja.get_all_records())
    except Exception as e:
        st.error(f"Error al conectar con la hoja de Servicios: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60, show_spinner=False)
def cargar_usuarios_df() -> pd.DataFrame:
    try:
        hoja = abrir_hoja(HOJA_USUARIOS)
        return pd.DataFrame(hoja.get_all_records())
    except Exception as e:
        st.error(f"Error al conectar con la hoja de Usuarios: {e}")
        return pd.DataFrame()

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

# --- SALUDO DINÁMICO (USA EL NOMBRE COMPLETO DIRECTO DE LA COLUMNA B) ---
def obtener_saludo(nombre_completo: str) -> str:
    hora_actual = datetime.datetime.now().hour
    nombre_str = str(nombre_completo).strip() if nombre_completo else "Efectivo"
    
    if 6 <= hora_actual < 13:
        return f"🌅 ¡Buen día, {nombre_str}!"
    elif 13 <= hora_actual < 20:
        return f"☀️ ¡Buenas tardes, {nombre_str}!"
    else:
        return f"🌙 ¡Buenas noches, {nombre_str}!"

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
        
    if submitted:
        if not dni or not clave:
            st.error("Por favor complete el DNI y la contraseña.")
            return
            
        try:
            with st.spinner("Verificando credenciales..."):
                data = cargar_usuarios_df()
                if data.empty:
                    st.error("No se pudo cargar la base de datos de usuarios.")
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
                    
                    # CAPTURA EXACTA DEL NOMBRE COMPLETO DE LA COLUMNA B (nombre y apellido)
                    col_nombre = next((c for c in u.index if "nombre" in str(c).lower()), data.columns[1] if len(data.columns) > 1 else "Usuario")
                    st.session_state.nombre_apellido = str(u.get(col_nombre, f"DNI: {dni}")).strip()
                    
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
        except Exception as e:
            st.error(f"Error de conexión con la base de datos de usuarios: {e}")

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
                os.unlink(ruta_img)

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name

# --- VISTA DEPENDENCIA ---
def _render_modulos(max_modulos: int, counter: int):
    labels_base = ["Simple (4h)", "Doble (8h)", "Triple (12h)", "Cuádruple (16h)"]
    labels_ordenados = labels_base[:max_modulos][::-1]
    
    jefes_dict = {}
    subs_dict = {}
    
    for label in labels_ordenados:
        st.markdown(f"**Turno {label}**")
        c1, c2 = st.columns(2)
        with c1:
            j = st.number_input(f"Jefes ({label})", 0, step=1, key=f"j_{label}_{counter}")
        with c2:
            s = st.number_input(f"Subalternos ({label})", 0, step=1, key=f"s_{label}_{counter}")
        
        if "Simple" in label:
            jefes_dict[0] = int(j)
            subs_dict[0] = int(s)
        elif "Doble" in label:
            jefes_dict[1] = int(j)
            subs_dict[1] = int(s)
        elif "Triple" in label:
            jefes_dict[2] = int(j)
            subs_dict[2] = int(s)
        elif "Cuádruple" in label:
            jefes_dict[3] = int(j)
            subs_dict[3] = int(s)

    jefes_final = [jefes_dict.get(i, 0) for i in range(4)]
    subs_final = [subs_dict.get(i, 0) for i in range(4)]
    
    return jefes_final, subs_final

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
    nombre_efectivo = st.session_state.get("nombre_apellido", "Personal Policial")
    saludo = obtener_saludo(nombre_efectivo)
    
    st.sidebar.markdown(
        f"""
        <div class="user-card">
            <div class="user-avatar">👮</div>
            <div style="font-weight:800; font-size: 1.05rem; margin-top:2px;">{nombre_efectivo}</div>
            <div style="font-size:0.8rem; color: #555; margin-top:4px;">Dep: {st.session_state.dependencia}</div>
            <div style="font-size:0.75rem; margin-top:8px; background: #E8F0FE; padding:4px 12px; border-radius:30px; display:inline-block; font-weight:600;">
                👤 COLABORADOR
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    boton_refrescar()
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.markdown(f"### {saludo}")
    st.markdown(f"<p style='color: #555; margin-top: -10px;'>Dependencia: <b>{st.session_state.dependencia}</b></p>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(
        ["🟢 Cargar Nuevo Servicio", "🔵 Pendientes de Asistencia", "📜 Historial de Servicios"]
    )

    counter = st.session_state.form_counter

    with tab1:
        st.header("Solicitud de Servicio Adicional")
        st.info("💡 Complete los datos paso a paso. Cada sección se habilitará a medida que seleccione o ingrese la información.")

        if f"fecha_elegida_{counter}" not in st.session_state:
            st.session_state[f"fecha_elegida_{counter}"] = False

        st.subheader("📌 Paso 1: Fecha del Evento")
        
        with st.form(f"form_fecha_{counter}"):
            fecha = st.date_input("Seleccione la fecha del servicio", value=datetime.date.today(), min_value=datetime.date.today())
            btn_confirmar_fecha = st.form_submit_button("Confirmar Fecha y Continuar")
            
        if btn_confirmar_fecha:
            st.session_state[f"fecha_elegida_{counter}"] = True

        if st.session_state[f"fecha_elegida_{counter}"]:
            st.success(f"✔ Fecha registrada: **{fecha}**")
            st.markdown("---")
            
            st.subheader("📌 Paso 2: Tipo de Evento")
            tipo_evento = st.text_input(
                "Ingrese el tipo de evento (Ej: Partido de fútbol, Baile, Recital, etc.)", key=f"tipo_{counter}"
            )
            
            if tipo_evento.strip():
                st.success("✔ Tipo de evento registrado.")
                st.markdown("---")
                
                st.subheader("📌 Paso 3 y 4: Datos del Organizador")
                col_org1, col_org2 = st.columns(2)
                with col_org1:
                    nombre_organizador = st.text_input("Nombre y Apellido del Organizador", key=f"nom_org_{counter}")
                with col_org2:
                    dni_organizador = st.text_input("DNI del Organizador", key=f"dni_org_{counter}")
                
                if nombre_organizador.strip() and dni_organizador.strip():
                    st.success("✔ Datos del organizador registrados.")
                    st.markdown("---")
                    
                    st.subheader("📌 Paso 5: Facturación")
                    cuil_organizador = st.text_input("CUIL / CUIT del Organizador (11 dígitos)", key=f"cuil_org_{counter}")
                    
                    if cuil_organizador.strip():
                        st.success("✔ CUIL/CUIT registrado.")
                        st.markdown("---")
                        
                        st.subheader("📌 Paso 6: Ubicación del Evento")
                        ubicacion_predio = st.text_input("Ubicación / Predio donde se realizará el servicio", key=f"ubi_{counter}")
                        
                        if ubicacion_predio.strip():
                            st.success("✔ Ubicación registrada.")
                            st.markdown("---")
                            
                            st.subheader("📌 Paso 7: Horarios y Turnos")
                            col_h1, col_h2 = st.columns(2)
                            with col_h1:
                                hora_inicio = st.time_input("Hora de Inicio", value=datetime.time(20, 0), key=f"h_ini_{counter}")
                            with col_h2:
                                hora_fin = st.time_input("Hora de Fin", value=datetime.time(4, 0), key=f"h_fin_{counter}")

                            diferencia_horas, max_modulos = calcular_monto_y_modulos(hora_inicio, hora_fin)
                            st.info(f"⏱️ Duración calculada: **{diferencia_horas:.1f} horas** (Hasta **{max_modulos} módulo(s)**)")

                            jefes_por_mod, subs_por_mod = _render_modulos(max_modulos, counter)
                            total_jefes = sum(jefes_por_mod)
                            total_subalternos = sum(subs_por_mod)
                            monto_total = calcular_monto_total(jefes_por_mod, subs_por_mod)

                            st.markdown("---")
                            st.success(f"💰 **Monto Total a Pagar: ${monto_total:,.2f}**")

                            st.markdown("---")
                            st.subheader("📌 Paso 8: Facturación y Documentación Adjunta")
                            with st.form(f"form_envio_{counter}", clear_on_submit=False):
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
                                st.subheader("📎 Archivos Obligatorios y Opcionales")
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

                                submitted = st.form_submit_button("🚀 Enviar Solicitud Definitiva")

                            if submitted:
                                organizador_completo = f"{nombre_organizador} (DNI: {dni_organizador})"
                                if total_jefes == 0 and total_subalternos == 0:
                                    st.error("⚠️ Debés asignar al menos un efectivo (Jefe o Subalterno).")
                                elif not (file_hab and file_comp):
                                    st.error("Faltan los archivos obligatorios (Habilitación y Comprobante).")
                                elif not validar_cuil(cuil_pagado_por):
                                    st.error("CUIL/CUIT inválido: debe tener exactamente 11 dígitos numéricos.")
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
                                                st.session_state[f"fecha_elegida_{counter}"] = False
                                                st.session_state.form_counter += 1
                                                st.cache_data.clear()
                                                st.rerun()
                                            else:
                                                st.error(f"Error en Apps Script: {res.get('message')}")
                                    except Exception as e:
                                        logger.exception("Error enviando solicitud")
                                        st.error(f"Error al enviar: {e}")

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
                            tipo_ev = row.get('Tipo de Evento') or row.get('tipo_evento') or 'Evento'
                            fec_ev = row.get('Fecha Evento') or row.get('fecha_evento') or ''
                            ubi_ev = row.get('Ubicación Predio') or row.get('ubicacion') or ''
                            
                            with st.expander(f"📅 {fec_ev} - {tipo_ev} ({ubi_ev}) [ID: {id_serv}]"):
                                st.write(f"**Organizador:** {row.get('Organizador', '')}")
                                st.write(f"**Horario:** {row.get('Horario', '')} | **Módulos:** {row.get('Módulos', '')}")
                                st.write(f"**Personal:** {row.get('Cant. Jefes', 0)} Jefes / {row.get('Cant. Subalternos', 0)} Subalternos")
                                
                                with st.form(f"form_asist_{id_serv}"):
                                    file_asist = st.file_uploader(
                                        "Planilla de Asistencia Firmada",
                                        type=["png", "jpg", "jpeg", "pdf"],
                                        key=f"file_{id_serv}",
                                    )
                                    sub_asist = st.form_submit_button("Subir Planilla")
                                    
                                if sub_asist:
                                    if not file_asist:
                                        st.error("Seleccioná un archivo antes de enviar.")
                                    else:
                                        with st.spinner("Subiendo planilla..."):
                                            payload = {
                                                "action": "subir_asistencia",
                                                "id_servicio": str(id_serv),
                                                "file_asist_base64": base64.b64encode(
                                                    file_asist.getvalue()
                                                ).decode("utf-8"),
                                                "name_asist": (
                                                    f"{id_serv}_asistencia."
                                                    f"{file_asist.name.split('.')[-1]}"
                                                ),
                                                "type_asist": file_asist.type,
                                            }
                                            res = _post_apps_script(payload)
                                        if res.get("status") == "success":
                                            st.success("¡Planilla subida con éxito!")
                                            st.cache_data.clear()
                                            st.rerun()
                                        else:
                                            st.error(f"Error: {res.get('message')}")
        except Exception as e:
            logger.exception("Error tab pendientes")
            st.error(f"Error al cargar los pendientes: {e}")

    with tab3:
        st.header("Historial General de Servicios")
        st.info(f"Listado completo de servicios de **{st.session_state.dependencia}**.")
        try:
            data = cargar_servicios_df()
            if data.empty or "Dependencia" not in data.columns:
                st.info("Aún no hay datos cargados.")
            else:
                df_dep = data[data["Dependencia"].astype(str) == str(st.session_state.dependencia)]
                if df_dep.empty:
                    st.info("No hay registros en el historial.")
                else:
                    st.dataframe(df_dep, use_container_width=True)
        except Exception as e:
            logger.exception("Error historial")
            st.error(f"Error al cargar historial: {e}")

# --- VISTA ADMIN ---
def vista_admin():
    nombre_efectivo = st.session_state.get("nombre_apellido", "Administrador General")
    saludo = obtener_saludo(nombre_efectivo)

    st.sidebar.markdown(
        f"""
        <div class="user-card">
            <div class="user-avatar">👑</div>
            <div style="font-weight:800; font-size: 1.05rem; margin-top:2px;">{nombre_efectivo}</div>
            <div style="font-size:0.75rem; margin-top:8px; background: #E8F0FE; padding:4px 12px; border-radius:30px; display:inline-block; font-weight:600;">
                👑 ADMIN GENERAL
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    boton_refrescar()
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

    st.markdown(f"### {saludo}")
    st.title("📋 Panel de Control General - Administrador")
    try:
        data = cargar_servicios_df()
        if data.empty:
            st.info("No hay servicios cargados.")
            return

        c1, c2 = st.columns(2)
        with c1:
            st.metric("Total de Servicios", len(data))
        with c2:
            monto = data["Monto Total"].sum() if "Monto Total" in data.columns else 0
            st.metric("Recaudación Total", f"${monto:,.2f}")

        st.markdown("---")
        st.subheader("Gestión y Generación de Informes PDF")
        deps = (
            ["Todas"] + sorted(data["Dependencia"].dropna().unique().tolist())
            if "Dependencia" in data.columns
            else ["Todas"]
        )
        dep_sel = st.selectbox("Filtrar por Dependencia", options=deps)
        data_f = data if dep_sel == "Todas" else data[data["Dependencia"].astype(str) == str(dep_sel)]
        st.dataframe(data_f, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Descargar Informe PDF con Constancias")
        if data_f.empty:
            return

        opciones = [
            f"{r['ID Servicio']} - {r['Dependencia']} ({r['Fecha Evento']} - {r['Tipo de Evento']})"
            for _, r in data_f.iterrows()
        ]
        elegido = st.selectbox("Seleccioná un servicio", options=opciones)
        if not elegido:
            return

        id_sel = elegido.split(" - ")[0]
        row_data = data_f[data_f["ID Servicio"].astype(str) == id_sel].iloc[0]

        if st.button("🚀 Generar PDF con Constancias"):
            with st.spinner("Descargando archivos y armando PDF..."):
                pdf_path = generar_pdf_servicio(row_data)
                try:
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
                finally:
                    if os.path.exists(pdf_path):
                        os.unlink(pdf_path)
            st.download_button(
                label="📥 Descargar PDF Final",
                data=pdf_bytes,
                file_name=f"{id_sel}_informe_completo.pdf",
                mime="application/pdf",
            )
    except Exception as e:
        logger.exception("Error panel admin")
        st.error(f"Error al cargar el panel: {e}")

# --- ROUTER PRINCIPAL ---
if st.session_state.get("cambio_clave_requerido", False):
    pantalla_cambio_clave()
elif not st.session_state.autenticado:
    login()
elif st.session_state.rol == "Administrador":
    vista_admin()
else:
    vista_dependencia()
