import streamlit as str_lit
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
str_lit.set_page_config(page_title="URV V - Adicionales", layout="centered", page_icon="👮")

# --- PARÁMETROS VÍA URL PARA EL FICHAJE ---
query_params = str_lit.query_params
pin_servicio_qr = query_params.get("id_servicio", None)

# --- ESTILOS VISUALES ---
str_lit.markdown("""
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
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #002060, #002a80) !important;
        color: white !important;
    }
    
    .stButton > button, div.stFormSubmitButton > button {
        border-radius: 30px !important;
        font-weight: 700 !important;
        background: linear-gradient(135deg, #002060, #002a80) !important;
        color: white !important;
        border: none !important;
        height: 3.2em !important;
        width: 100% !important;
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

SHEET_NAME = "Gestion_Adicionales"
HOJA_SERVICIOS = "Servicios"
HOJA_USUARIOS = "Usuarios"
HOJA_PERSONAL = "Personal_Base"
HOJA_ASISTENCIA = "Asistencia_Digital_Global"
HOJA_CONFIG = "Config"

PRECIO_JEFE = 33000
PRECIO_SUB = 28000
HORAS_POR_MODULO = 4
MAX_MODULOS_PERMITIDOS = 3
TIMEOUT_HTTP = 30

try:
    URL_APPS_SCRIPT = str_lit.secrets["URL_APPS_SCRIPT"]
except (KeyError, FileNotFoundError):
    URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxqVp5T0fVGtPncuwkocbXbU5TSHLgPqgmD0A5pQ6CmuEU3C-UrGx2eZHxg2HTLUjU1/exec"

# --- CONEXIÓN A GOOGLE SHEETS ---
@str_lit.cache_resource
def conectar_google():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if "gcp_service_account" in str_lit.secrets:
        info = dict(str_lit.secrets["gcp_service_account"])
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

@str_lit.cache_data(ttl=60, show_spinner=False)
def cargar_servicios_df() -> pd.DataFrame:
    try:
        hoja = abrir_hoja(HOJA_SERVICIOS)
        return pd.DataFrame(hoja.get_all_records())
    except Exception as e:
        return pd.DataFrame()

@str_lit.cache_data(ttl=60, show_spinner=False)
def cargar_usuarios_df() -> pd.DataFrame:
    try:
        hoja = abrir_hoja(HOJA_USUARIOS)
        return pd.DataFrame(hoja.get_all_records())
    except Exception as e:
        return pd.DataFrame()

@str_lit.cache_data(ttl=60, show_spinner=False)
def cargar_personal_base_df() -> pd.DataFrame:
    try:
        hoja = abrir_hoja(HOJA_PERSONAL)
        return pd.DataFrame(hoja.get_all_records())
    except Exception as e:
        return pd.DataFrame()

@str_lit.cache_data(ttl=30, show_spinner=False)
def cargar_asistencia_digital_df() -> pd.DataFrame:
    try:
        hoja = abrir_hoja(HOJA_ASISTENCIA)
        return pd.DataFrame(hoja.get_all_records())
    except Exception as e:
        return pd.DataFrame()

@str_lit.cache_data(ttl=5, show_spinner=False)
def verificar_modo_mantenimiento() -> bool:
    try:
        hoja = abrir_hoja(HOJA_CONFIG)
        datos = hoja.get_all_records()
        for row in datos:
            if str(row.get("clave", "")).strip() == "modo_mantenimiento":
                return str(row.get("valor", "")).strip().upper() == "ACTIVO"
        return False
    except Exception:
        return False

def boton_refrescar():
    if str_lit.sidebar.button("🔄 Refrescar datos"):
        str_lit.cache_data.clear()
        str_lit.rerun()

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
    str_lit.session_state.setdefault(k, v)

# --- HELPERS ---
def validar_cuil(valor: str) -> bool:
    return bool(re.fullmatch(r"\d{11}", str(valor).strip()))

def calcular_monto_y_modulos(hora_inicio: datetime.time, hora_fin: datetime.time):
    dt_ini = datetime.datetime.combine(datetime.date.today(), hora_inicio)
    dt_fin = datetime.datetime.combine(datetime.date.today(), hora_fin)
    if dt_fin <= dt_ini:
        dt_fin += datetime.timedelta(days=1)
    horas = (dt_fin - dt_ini).total_seconds() / 3600.0
    max_mod = max(1, min(MAX_MODULOS_PERMITIDOS, math.ceil(horas / HORAS_POR_MODULO)))
    return horas, max_mod

def calcular_monto_total(jefes_por_mod: list, subs_por_mod: list) -> float:
    total = 0
    for i, (j, s) in enumerate(zip(jefes_por_mod, subs_por_mod), start=1):
        total += ((j * PRECIO_JEFE) + (s * PRECIO_SUB)) * i
    return total

def obtener_saludo(nombre_completo: str) -> str:
    hora_actual = datetime.datetime.now().hour
    nombre_str = str(nombre_completo).strip() if nombre_completo else "Efectivo"
    if 6 <= hora_actual < 13:
        return f"🌅 ¡Buen día, {nombre_str}!"
    elif 13 <= hora_actual < 20:
        return f"☀️ ¡Buenas tardes, {nombre_str}!"
    else:
        return f"🌙 ¡Buenas noches, {nombre_str}!"

def _post_apps_script(payload: dict) -> dict:
    try:
        r = requests.post(URL_APPS_SCRIPT, json=payload, timeout=TIMEOUT_HTTP)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        return {"status": "error", "message": str(e)}

# --- PANTALLA PÚBLICA DE FICHAJE (CON AUTOCOMPLETADO POR DNI) ---
def pantalla_fichaje_qr(id_servicio_defecto=None):
    str_lit.subheader("📱 Registro / Fichaje de Asistencia (Jefe o Terreno)")
    str_lit.info("Ingrese el DNI para autocompletar la información del efectivo desde el padrón.")

    id_serv_input = str_lit.text_input("ID / PIN del Servicio", value=id_servicio_defecto or "")
    dni_input = str_lit.text_input("DNI del Efectivo")

    nombre_auto = ""
    jerarquia_auto = ""
    dep_auto = ""
    encontrado = False

    if dni_input.strip():
        df_personal = cargar_personal_base_df()
        if not df_personal.empty and "DNI" in df_personal.columns:
            res_p = df_personal[df_personal["DNI"].astype(str).str.strip() == dni_input.strip()]
            if not res_p.empty:
                fila_p = res_p.iloc[0]
                nombre_auto = str(fila_p.get("Nombre y Apellido", fila_p.get("nombre_apellido", "")))
                jerarquia_auto = str(fila_p.get("Jerarquia", fila_p.get("jerarquia", "")))
                dep_auto = str(fila_p.get("Dependencia", fila_p.get("dependencia", "")))
                encontrado = True
                str_lit.success(f"🟢 Efectivo en Padrón: **{jerarquia_auto} {nombre_auto}** ({dep_auto})")

    with str_lit.form("form_fichaje_terreno"):
        jerarquia_val = str_lit.text_input("Jerarquía / Rango", value=jerarquia_auto)
        nombre_val = str_lit.text_input("Nombre y Apellido Completo", value=nombre_auto)
        dep_val = str_lit.text_input("Dependencia de Origen", value=dep_auto)
        
        jornada_sel = str_lit.selectbox("Turno / Tanda", [
            "Módulo Triple (12 hs)",
            "Módulo Doble (8 hs)",
            "Módulo Simple (4 hs)",
            "Día 1",
            "Día 2"
        ])
        
        btn_fichar = str_lit.form_submit_button("✅ Confirmar Asistencia")
        
    if btn_fichar:
        if not id_serv_input or not dni_input or not nombre_val:
            str_lit.error("Por favor complete los campos obligatorios.")
        else:
            payload = {
                "action": "fichaje_qr",
                "id_servicio": str(id_serv_input).strip(),
                "dni": str(dni_input).strip(),
                "jerarquia": str(jerarquia_val).strip(),
                "nombre_apellido": str(nombre_val).strip(),
                "dependencia": str(dep_val).strip(),
                "jornada": jornada_sel
            }
            res = _post_apps_script(payload)
            if res.get("status") == "success":
                str_lit.balloons()
                str_lit.success("¡Asistencia registrada e impactada con éxito!")
            else:
                str_lit.error(f"Error al registrar: {res.get('message')}")

# --- LOGIN Y CAMBIO DE CLAVE ---
def pantalla_cambio_clave():
    str_lit.title("🔑 Definir Nueva Contraseña Personal")
    str_lit.warning("⚠️ Primer ingreso detectado. Debés cambiar tu contraseña temporal.")
    nueva = str_lit.text_input("Nueva Contraseña", type="password")
    confirm = str_lit.text_input("Confirmar Nueva Contraseña", type="password")
    if str_lit.button("Guardar Nueva Contraseña"):
        if not nueva or nueva != confirm:
            str_lit.error("Las contraseñas no coinciden o están vacías.")
            return
        try:
            hoja = abrir_hoja(HOJA_USUARIOS)
            cell = hoja.find(str(str_lit.session_state.dni_temp))
            hoja.update_cell(cell.row, 6, nueva)
            hoja.update_cell(cell.row, 5, "")
            str_lit.success("¡Contraseña actualizada! Iniciá sesión nuevamente.")
            str_lit.session_state.cambio_clave_requerido = False
            str_lit.session_state.dni_temp = None
            str_lit.rerun()
        except Exception as e:
            logger.exception("Error actualizando clave")
            str_lit.error(f"Error al actualizar: {e}")

def login():
    str_lit.markdown("### 🔐 Acceso al Sistema de Adicionales")
    
    t_login, t_qr = str_lit.tabs(["🔑 Acceso Personal / Administración", "📱 Fichar Asistencia por DNI / PIN"])
    
    with t_login:
        with str_lit.form("form_login"):
            dni = str_lit.text_input("DNI / Usuario")
            clave = str_lit.text_input("Contraseña", type="password")
            submitted = str_lit.form_submit_button("INGRESAR AL SISTEMA")
            
        if submitted:
            if not dni or not clave:
                str_lit.error("Por favor complete el DNI y la contraseña.")
                return
            try:
                with str_lit.spinner("Verificando credenciales..."):
                    data = cargar_usuarios_df()
                    if data.empty:
                        str_lit.error("No se pudo cargar la base de datos de usuarios.")
                        return
                    data.columns = [str(c).strip().lower() for c in data.columns]
                    user = data[data["dni"].astype(str) == str(dni)]
                    if user.empty:
                        str_lit.error("DNI no registrado.")
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
                            str_lit.session_state.cambio_clave_requerido = True
                            str_lit.session_state.dni_temp = dni
                            str_lit.rerun()
                        else:
                            str_lit.error("Contraseña temporal incorrecta.")
                        return

                    if clave == clave_pers:
                        str_lit.session_state.autenticado = True
                        str_lit.session_state.usuario_dni = dni
                        str_lit.session_state.dependencia = str(u.get("dependencia", ""))
                        str_lit.session_state.rol = str(u.get("rol", "")).strip()
                        col_nombre = next((c for c in u.index if "nombre" in str(c).lower()), data.columns[1] if len(data.columns) > 1 else "Usuario")
                        str_lit.session_state.nombre_apellido = str(u.get(col_nombre, f"DNI: {dni}")).strip()
                        str_lit.rerun()
                    else:
                        str_lit.error("Contraseña incorrecta.")
            except Exception as e:
                str_lit.error(f"Error de conexión con la base de usuarios: {e}")

    with t_qr:
        pantalla_fichaje_qr()

# --- GENERACIÓN DE IMÁGENES DESDE DRIVE ---
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
        logger.warning(f"Error descargando imagen Drive {file_id}: {e}")
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return None

# --- GENERADOR DE REPORTES PDF ---
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
    fecha_ev_str = str(row.get('Fecha Evento', ''))
    fila_tabla("Dependencia", str(row.get('Dependencia', '')), "Fecha(s)", fecha_ev_str)
    fila_tabla("Tipo de Evento", str(row.get('Tipo de Evento', '')), "Horario", str(row.get('Horario', '')))
    fila_completa("Ubicación / Predio", str(row.get('Ubicación Predio', '')))
    fila_completa("Organizador", str(row.get('Organizador', '')))

    seccion_titulo("2. RECURSOS Y MÓDULOS ASIGNADOS")
    monto_val = row.get('Monto Total', 0)
    monto_str = f"${float(monto_val):,.2f}" if monto_val else "$0.00"
    
    total_jefes_val = row.get('Cant. Jefes', 0)
    total_subs_val = row.get('Cant. Subalternos', 0)
    modulos_val = row.get('Módulos', 1)

    fila_tabla("Total Jefes", str(total_jefes_val), "Total Subalternos", str(total_subs_val))
    fila_tabla("Módulos (Max)", str(modulos_val), "Monto Total", monto_str)

    seccion_titulo("3. DATOS DE FACTURACIÓN Y PAGO")
    fila_completa("Pagado por", str(row.get('Pagador', '')))
    fila_completa("CUIL / CUIT Facturación", str(row.get('CUIL Pagador', '')))

    seccion_titulo("4. DOCUMENTACIÓN Y ENLACES (GOOGLE DRIVE)")
    
    asist_1 = row.get("Enlace Asistencia Día 1", row.get("Enlace Asistencia", ""))
    asist_2 = row.get("Enlace Asistencia Día 2", "")

    es_dos_dias = " y " in fecha_ev_str
    fechas_partes = fecha_ev_str.split(" y ") if es_dos_dias else [fecha_ev_str]

    enlaces = [
        ("Habilitación Municipal", row.get("Enlace Habilitación", "")),
        ("Comprobante de Pago", row.get("Enlace Comprobante", "")),
        ("Contrato de Adicional", row.get("Enlace Contrato", ""))
    ]

    if es_dos_dias:
        enlaces.append((f"Asistencia Día 1 ({fechas_partes[0]})", asist_1))
        enlaces.append((f"Asistencia Día 2 ({fechas_partes[1]})", asist_2))
    else:
        enlaces.append(("Planilla de Asistencia", asist_1))

    pdf.set_font("helvetica", "B", 9)
    for etiqueta, url in enlaces:
        pdf.set_fill_color(245, 247, 250)
        pdf.cell(55, 6, f" {etiqueta}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9)
        if url and str(url).startswith("http"):
            pdf.set_text_color(0, 102, 204)
            pdf.cell(135, 6, " [ Enlace Disponible / Ver Documento ]", 1, 1, 'L', link=str(url))
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.set_text_color(120, 120, 120)
            pdf.cell(135, 6, " No disponible / Pendiente", 1, 1, 'L')
            pdf.set_text_color(0, 0, 0)
        pdf.set_font("helvetica", "B", 9)

    anexos = [
        ("ANEXO: HABILITACIÓN MUNICIPAL", row.get("Enlace Habilitación", "")),
        ("ANEXO: COMPROBANTE DE PAGO", row.get("Enlace Comprobante", "")),
        ("ANEXO: CONTRATO DE ADICIONAL", row.get("Enlace Contrato", ""))
    ]

    if es_dos_dias:
        anexos.append((f"ANEXO: PLANILLA ASISTENCIA DÍA 1 ({fechas_partes[0]})", asist_1))
        anexos.append((f"ANEXO: PLANILLA ASISTENCIA DÍA 2 ({fechas_partes[1]})", asist_2))
    else:
        anexos.append(("ANEXO: PLANILLA DE ASISTENCIA", asist_1))

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
            logger.warning(f"PDF: error al embeber {titulo}: {e}")
        finally:
            if os.path.exists(ruta_img):
                os.unlink(ruta_img)

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name

# --- PANEL DE DEPENDENCIA ---
def _render_modulos(max_modulos: int, counter: int):
    labels_base = ["Simple (4h)", "Doble (8h)", "Triple (12h)"]
    labels_ordenados = labels_base[:max_modulos][::-1]
    jefes_dict = {}
    subs_dict = {}
    for label in labels_ordenados:
        str_lit.markdown(f"**Turno {label}**")
        c1, c2 = str_lit.columns(2)
        with c1:
            j = str_lit.number_input(f"Jefes ({label})", 0, step=1, key=f"j_{label}_{counter}")
        with c2:
            s = str_lit.number_input(f"Subalternos ({label})", 0, step=1, key=f"s_{label}_{counter}")
        if "Simple" in label:
            jefes_dict[0] = int(j); subs_dict[0] = int(s)
        elif "Doble" in label:
            jefes_dict[1] = int(j); subs_dict[1] = int(s)
        elif "Triple" in label:
            jefes_dict[2] = int(j); subs_dict[2] = int(s)

    jefes_final = [jefes_dict.get(i, 0) for i in range(3)]
    subs_final = [subs_dict.get(i, 0) for i in range(3)]
    return jefes_final, subs_final

def vista_dependencia():
    nombre_efectivo = str_lit.session_state.get("nombre_apellido", "Personal Policial")
    saludo = obtener_saludo(nombre_efectivo)
    
    str_lit.sidebar.markdown(f"**{nombre_efectivo}** ({str_lit.session_state.dependencia})")
    boton_refrescar()
    if str_lit.sidebar.button("🚪 Cerrar Sesión"):
        str_lit.session_state.autenticado = False
        str_lit.rerun()

    str_lit.markdown(f"### {saludo}")
    
    tab1, tab2, tab3 = str_lit.tabs(["🟢 Cargar Nuevo Servicio", "⏳ Pendientes de Asistencia", "📜 Historial de Servicios"])

    counter = str_lit.session_state.form_counter

    with tab1:
        str_lit.header("Solicitud de Servicio Adicional")
        if f"fecha_elegida_{counter}" not in str_lit.session_state:
            str_lit.session_state[f"fecha_elegida_{counter}"] = False

        with str_lit.form(f"form_fecha_{counter}"):
            tipo_cant_dias = str_lit.radio("Modalidad de días:", ["1 Día único", "2 Días específicos (Ej: Viernes y Sábado)"], horizontal=True, key=f"tipo_dias_{counter}")
            if tipo_cant_dias == "1 Día único":
                fec_inicio = str_lit.date_input("Fecha del Evento", value=datetime.date.today(), key=f"fec_ini_{counter}")
                fec_fin = fec_inicio
            else:
                col_f1, col_f2 = str_lit.columns(2)
                with col_f1: fec_inicio = str_lit.date_input("Fecha Día 1", value=datetime.date.today(), key=f"fec_ini_{counter}")
                with col_f2: fec_fin = str_lit.date_input("Fecha Día 2", value=datetime.date.today(), key=f"fec_fin_{counter}")
            btn_confirmar_fecha = str_lit.form_submit_button("Confirmar Fechas y Continuar")
            
        if btn_confirmar_fecha:
            str_lit.session_state[f"fecha_elegida_{counter}"] = True

        if str_lit.session_state.get(f"fecha_elegida_{counter}", False):
            fecha_texto = f"{fec_inicio.strftime('%d/%m/%Y')}" if (tipo_cant_dias == "1 Día único" or fec_inicio == fec_fin) else f"{fec_inicio.strftime('%d/%m/%Y')} y {fec_fin.strftime('%d/%m/%Y')}"
            cant_dias = 1 if (tipo_cant_dias == "1 Día único" or fec_inicio == fec_fin) else 2

            tipo_evento = str_lit.text_input("Tipo de evento", key=f"tipo_{counter}")
            if tipo_evento.strip():
                col_org1, col_org2 = str_lit.columns(2)
                with col_org1: nombre_organizador = str_lit.text_input("Nombre Organizador", key=f"nom_org_{counter}")
                with col_org2: dni_organizador = str_lit.text_input("DNI Organizador", key=f"dni_org_{counter}")
                
                if nombre_organizador.strip() and dni_organizador.strip():
                    cuil_organizador = str_lit.text_input("CUIL / CUIT (11 dígitos)", key=f"cuil_org_{counter}")
                    if cuil_organizador.strip():
                        ubicacion_predio = str_lit.text_input("Ubicación / Predio", key=f"ubi_{counter}")
                        if ubicacion_predio.strip():
                            col_h1, col_h2 = str_lit.columns(2)
                            with col_h1: hora_inicio = str_lit.time_input("Hora Inicio", value=datetime.time(15, 0), key=f"h_ini_{counter}")
                            with col_h2: hora_fin = str_lit.time_input("Hora Fin", value=datetime.time(19, 0), key=f"h_fin_{counter}")

                            diferencia_horas, max_modulos = calcular_monto_y_modulos(hora_inicio, hora_fin)
                            jefes_por_mod, subs_por_mod = _render_modulos(max_modulos, counter)
                            total_jefes = sum(jefes_por_mod) * cant_dias
                            total_subalternos = sum(subs_por_mod) * cant_dias
                            monto_total = calcular_monto_total(jefes_por_mod, subs_por_mod) * cant_dias

                            pagado_por = f"{nombre_organizador} (DNI: {dni_organizador})"
                            cuil_pagado_por = cuil_organizador

                            with str_lit.form(f"form_envio_{counter}"):
                                file_hab = str_lit.file_uploader("Habilitación Municipal", type=["png", "jpg", "jpeg", "pdf"], key=f"hab_{counter}")
                                file_comp = str_lit.file_uploader("Comprobante de Pago Único", type=["png", "jpg", "jpeg", "pdf"], key=f"comp_{counter}")
                                file_cont = str_lit.file_uploader("Contrato (Opcional)", type=["png", "jpg", "jpeg", "pdf"], key=f"cont_{counter}")
                                submitted = str_lit.form_submit_button("🚀 Enviar Solicitud Definitiva")

                            if submitted:
                                if total_jefes == 0 and total_subalternos == 0:
                                    str_lit.error("Asigne al menos un efectivo.")
                                elif not (file_hab and file_comp):
                                    str_lit.error("Faltan archivos obligatorios.")
                                elif not validar_cuil(cuil_pagado_por):
                                    str_lit.error("CUIL inválido.")
                                else:
                                    id_servicio = f"SERV-{int(datetime.datetime.now().timestamp())}"
                                    payload = {
                                        "id_servicio": id_servicio,
                                        "dependencia": str(str_lit.session_state.dependencia),
                                        "fecha_evento": fecha_texto,
                                        "tipo_evento": str(tipo_evento),
                                        "organizador": f"{nombre_organizador} (DNI: {dni_organizador})",
                                        "ubicacion": str(ubicacion_predio),
                                        "horario": f"{hora_inicio.strftime('%H:%M')} a {hora_fin.strftime('%H:%M')}",
                                        "cant_jefes": int(total_jefes),
                                        "cant_subalternos": int(total_subalternos),
                                        "modulos": int(max_modulos),
                                        "monto_total": float(monto_total),
                                        "jefes_simples": int(jefes_por_mod[0]),
                                        "subalternos_simples": int(subs_por_mod[0]),
                                        "jefes_dobles": int(jefes_por_mod[1]),
                                        "subalternos_dobles": int(subs_por_mod[1]),
                                        "jefes_triples": int(jefes_por_mod[2]),
                                        "subalternos_triples": int(subs_por_mod[2]),
                                        "pagado_por": str(pagado_por),
                                        "cuil_pagado_por": str(cuil_pagado_por),
                                        "estado_central": "🟡 Pendiente",
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
                                        str_lit.success("¡Solicitud registrada!")
                                        str_lit.session_state[f"fecha_elegida_{counter}"] = False
                                        str_lit.session_state.form_counter += 1
                                        str_lit.cache_data.clear()
                                        str_lit.rerun()

    with tab2:
        str_lit.header("Servicios Pendientes de Planilla de Asistencia")
        data = cargar_servicios_df()
        if not data.empty and "Dependencia" in data.columns:
            df_dep = data[data["Dependencia"].astype(str) == str(str_lit.session_state.dependencia)]
            col_a1 = next((c for c in df_dep.columns if "asistencia día 1" in str(c).lower() or "asistencia 1" in str(c).lower()), None)
            col_a2 = next((c for c in df_dep.columns if "asistencia día 2" in str(c).lower() or "asistencia 2" in str(c).lower()), None)
            col_gen = next((c for c in df_dep.columns if "enlace asistencia" in str(c).lower()), None)

            def esta_pendiente(r):
                fec = str(r.get("Fecha Evento", ""))
                a1_val = str(r.get(col_a1, r.get(col_gen, ""))).strip()
                a2_val = str(r.get(col_a2, "")).strip()
                return (a1_val == "" or a2_val == "") if " y " in fec else (a1_val == "")

            df_pend = df_dep[df_dep.apply(esta_pendiente, axis=1)]
            if df_pend.empty:
                str_lit.success("🎉 ¡Sin pendientes!")
            else:
                for _, row in df_pend.iterrows():
                    id_serv = row.get('ID Servicio') or row.get('id_servicio') or row.iloc[0]
                    fec_ev = str(row.get('Fecha Evento') or '')
                    with str_lit.expander(f"📁 EXPEDIENTE: {id_serv} | 📅 {fec_ev}"):
                        str_lit.markdown(f"**PIN de Fichaje:** `{id_serv}`")
                        with str_lit.form(f"form_asist_1_{id_serv}"):
                            file_a1 = str_lit.file_uploader("Planilla Asistencia (Escaneada / Foto)", type=["png", "jpg", "jpeg", "pdf"], key=f"f1_{id_serv}")
                            sub_a1 = str_lit.form_submit_button("Subir Archivo al Expediente")
                        if sub_a1 and file_a1:
                            payload = {
                                "action": "subir_asistencia",
                                "id_servicio": str(id_serv),
                                "dia_num": 1,
                                "file_asist_base64": base64.b64encode(file_a1.getvalue()).decode("utf-8"),
                                "name_asist": f"{id_serv}_asistencia.{file_a1.name.split('.')[-1]}",
                                "type_asist": file_a1.type,
                            }
                            res = _post_apps_script(payload)
                            if res.get("status") == "success":
                                str_lit.success("¡Planilla adjuntada!")
                                str_lit.cache_data.clear()
                                str_lit.rerun()

    with tab3:
        str_lit.header("Historial General")
        data = cargar_servicios_df()
        if not data.empty and "Dependencia" in data.columns:
            df_dep = data[data["Dependencia"].astype(str) == str(str_lit.session_state.dependencia)]
            str_lit.dataframe(df_dep, use_container_width=True)

# --- PANEL DE ADMINISTRACIÓN CON FILTROS POR ESTADO Y MODO MANTENIMIENTO NUBE ---
def vista_admin():
    nombre_efectivo = str_lit.session_state.get("nombre_apellido", "Administrador General")
    saludo = obtener_saludo(nombre_efectivo)

    str_lit.sidebar.markdown(f"**{nombre_efectivo}** (ADMINISTRADOR)")
    
    # Control de Modo Mantenimiento en la nube (Solapa Config)
    str_lit.sidebar.markdown("---")
    str_lit.sidebar.subheader("⚙️ Control de Sistema")
    modo_mant_actual = verificar_modo_mantenimiento()
    nuevo_modo_toggle = str_lit.sidebar.toggle("🛠️ Modo Mantenimiento", value=modo_mant_actual)
    
    if nuevo_modo_toggle != modo_mant_actual:
        estado_str = "ACTIVO" if nuevo_modo_toggle else "INACTIVO"
        payload = {"action": "set_mantenimiento", "estado": estado_str}
        res = _post_apps_script(payload)
        if res.get("status") == "success":
            str_lit.cache_data.clear()
            str_lit.success(f"Modo mantenimiento cambiado a: {estado_str}")
            str_lit.rerun()
        else:
            str_lit.error("No se pudo actualizar el mantenimiento en la nube.")

    boton_refrescar()
    if str_lit.sidebar.button("🚪 Cerrar Sesión"):
        str_lit.session_state.autenticado = False
        str_lit.rerun()

    str_lit.markdown(f"### {saludo}")
    str_lit.title("📋 Control General y Despacho a Central")
    
    data = cargar_servicios_df()
    
    # Alerta visual automática
    if not data.empty and "Estado Central" in data.columns:
        pendientes_revision = data[data["Estado Central"].astype(str).str.contains("Recibido", na=False)]
        if not pendientes_revision.empty:
            str_lit.warning(f"🔔 **¡Atención Admin!** Hay **{len(pendientes_revision)}** servicio(s) con planillas de asistencia subidas por las dependencias pendientes de revisión y despacho.")

    t_expedientes, t_asistencia_qr, t_depuracion = str_lit.tabs(["📄 Expedientes y PDF", "📱 Asistencias Digitales", "🧹 Mantenimiento / Depuración Mensual"])
    
    with t_expedientes:
        str_lit.subheader("Gestión de Expedientes y Constancias")
        if not data.empty:
            col_f1, col_f2 = str_lit.columns([2, 2])
            with col_f1:
                filtro_estado = str_lit.selectbox("🔍 Filtrar expedientes por estado:", ["Todos", "🟡 Recibido - Requiere Revisión", "🟡 Pendiente", "🟢 Enviado a Central"])
            
            df_filtrado = data.copy()
            if filtro_estado != "Todos" and "Estado Central" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Estado Central"].astype(str).str.contains(filtro_estado, na=False)]

            str_lit.dataframe(df_filtrado, use_container_width=True)
            
            opciones = [f"{r['ID Servicio']} - {r['Dependencia']} ({r['Fecha Evento']}) [{r.get('Estado Central', '')}]" for _, r in df_filtrado.iterrows()]
            if opciones:
                elegido = str_lit.selectbox("Seleccionar servicio para verificar / despachar", options=opciones)
                if elegido:
                    id_sel = elegido.split(" - ")[0]
                    row_data = data[data["ID Servicio"].astype(str) == id_sel].iloc[0]

                    c_pdf, c_excel, c_estado = str_lit.columns(3)
                    with c_pdf:
                        if str_lit.button("🚀 Generar PDF con Constancias"):
                            pdf_path = generar_pdf_servicio(row_data)
                            with open(pdf_path, "rb") as f:
                                pdf_bytes = f.read()
                            os.unlink(pdf_path)
                            str_lit.download_button("📥 Descargar PDF Final", data=pdf_bytes, file_name=f"{id_sel}.pdf", mime="application/pdf")
                    
                    with c_excel:
                        df_asist = cargar_asistencia_digital_df()
                        if not df_asist.empty and "ID Servicio" in df_asist.columns:
                            df_exp = df_asist[df_asist["ID Servicio"].astype(str) == str(id_sel)]
                            if not df_exp.empty:
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    df_exp.to_excel(writer, index=False, sheet_name='Asistentes')
                                excel_data = output.getvalue()
                                str_lit.download_button(
                                    label="📊 Descargar Asistentes (Excel)",
                                    data=excel_data,
                                    file_name=f"Asistentes_{id_sel}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            else:
                                str_lit.info("Sin registros de asistencia digital.")

                    with c_estado:
                        if str_lit.button("✅ Marcar como Enviado a Central"):
                            try:
                                hoja_serv = abrir_hoja(HOJA_SERVICIOS)
                                cell = hoja_serv.find(str(id_sel))
                                if cell:
                                    hoja_serv.update_cell(cell.row, 25, "🟢 Enviado a Central")
                                    str_lit.success("¡Expediente marcado como enviado a Central!")
                                    str_lit.cache_data.clear()
                                    str_lit.rerun()
                            except Exception as e:
                                str_lit.error(f"Error al actualizar estado: {e}")
            else:
                str_lit.info("No hay expedientes en este filtro.")

    with t_asistencia_qr:
        str_lit.subheader("Listado Global de Asistencias Registradas")
        df_asist = cargar_asistencia_digital_df()
        if not df_asist.empty:
            str_lit.dataframe(df_asist, use_container_width=True)
            output_global = BytesIO()
            with pd.ExcelWriter(output_global, engine='openpyxl') as writer:
                df_asist.to_excel(writer, index=False, sheet_name='Global_Asistencias')
            str_lit.download_button("📊 Descargar Toda la Asistencia (Excel)", data=output_global.getvalue(), file_name="Asistencia_Global.xlsx")
        else:
            str_lit.info("Aún no hay registros de asistencia.")

    with t_depuracion:
        str_lit.subheader("🧹 Depuración y Archivado Mensual")
        str_lit.warning("Esta acción moverá los registros a una solapa histórica y vaciará la tabla activa para el nuevo mes.")
        periodo_actual = datetime.date.today().strftime("%Y_%m")
        periodo_inp = str_lit.text_input("Periodo del Histórico", value=periodo_actual)
        
        if str_lit.button("⚠️ Ejecutar Limpieza Mensual"):
            payload = {
                "action": "depurar_mes",
                "periodo": periodo_inp
            }
            res = _post_apps_script(payload)
            if res.get("status") == "success":
                str_lit.success(f"¡Depuración exitosa! Se migró a Historico_Asistencia_{periodo_inp}.")
                str_lit.cache_data.clear()
            else:
                str_lit.error(f"Error al depurar: {res.get('message')}")

# --- ENRUTADOR PRINCIPAL ---
modo_mantenimiento_activo = verificar_modo_mantenimiento()

if pin_servicio_qr:
    pantalla_fichaje_qr(id_servicio_defecto=pin_servicio_qr)
elif str_lit.session_state.get("cambio_clave_requerido", False):
    pantalla_cambio_clave()
elif not str_lit.session_state.autenticado:
    login()
elif str_lit.session_state.rol == "Administrador":
    vista_admin()
elif modo_mantenimiento_activo:
    str_lit.markdown("""
        <div style="text-align: center; padding: 4rem 2rem;">
            <h1>🛠️ SITIO EN MANTENIMIENTO</h1>
            <p style="font-size: 1.2rem; color: #555;">La Oficina de Adicionales de la UR V se encuentra realizando tareas de actualización en el sistema. Por favor intente nuevamente en unos momentos.</p>
        </div>
    """, unsafe_allow_html=True)
    if str_lit.button("🚪 Cerrar Sesión"):
        str_lit.session_state.autenticado = False
        str_lit.rerun()
else:
    vista_dependencia()
