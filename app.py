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
import unicodedata
from fpdf import FPDF
from io import BytesIO
from PIL import Image

# --- CONFIGURACIÓN DE PÁGINA ---
str_lit.set_page_config(page_title="UR V - Adicionales", layout="centered", page_icon="👮")

# --- CONSTANTES DE COLUMNAS ---
COL_ID_SERV = "ID Servicio"
COL_TIPO_EV = "Tipo de Evento"
COL_FECHA = "Fecha Evento"
COL_UBI = "Ubicación Predio"
COL_HORARIO = "Horario"
COL_DEP = "Dependencia"
COL_ESTADO_CENTRAL = "Estado Central"
COL_ESTADO_ASIST = "Estado Asistencia"
COL_ENLACE_ASIST = "Enlace Asistencia"

# --- ESTILOS VISUALES ---
str_lit.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background: #F5F5F5; }
    .header-container {
        background: linear-gradient(135deg, #002060, #002a80); color: white;
        padding: 2rem 1rem; border-radius: 20px; text-align: center;
        margin-bottom: 2rem; box-shadow: 0 10px 25px rgba(0,32,96,0.3);
        border-bottom: 6px solid #8B1E1E;
    }
    .badge-chamical {
        background: #8B1E1E; display: inline-block; padding: 8px 30px;
        border-radius: 50px; font-size: 1rem; font-weight: 800; color: white;
        box-shadow: 0 4px 15px rgba(139,30,30,0.5); margin-top: 10px; letter-spacing: 1px;
    }
    div.stExpander, div.stForm {
        background: #FFFFFF; border-radius: 20px; padding: 20px;
        margin-bottom: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border-left: 6px solid #002060; border: 1px solid #E0E0E0;
    }
    input, textarea, select, div[data-baseweb="input"] > div, div[data-baseweb="base-input"] {
        background-color: #FFFFFF !important; border: 1.5px solid #B0BEC5 !important;
        border-radius: 8px !important; color: #1a1a1a !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #002060, #002a80) !important; color: white !important;
    }
    .stButton > button, div.stFormSubmitButton > button {
        border-radius: 30px !important; font-weight: 700 !important;
        background: linear-gradient(135deg, #002060, #002a80) !important;
        color: white !important; border: none !important;
        height: 3.2em !important; width: 100% !important;
    }
    .monto-box {
        background: #E8F5E9; border-left: 6px solid #2E7D32; padding: 15px;
        border-radius: 10px; margin-top: 15px; text-align: center;
    }
    .monto-total { font-size: 1.5rem; font-weight: 800; color: #1B5E20; }
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
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" in str_lit.secrets:
        info = dict(str_lit.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        ruta = os.path.join(os.path.dirname(__file__), "credenciales.json")
        if not os.path.exists(ruta): ruta = "credenciales.json"
        if not os.path.exists(ruta): raise FileNotFoundError("No se encontró credenciales.json")
        credentials = Credentials.from_service_account_file(ruta, scopes=scopes)
    client = gspread.authorize(credentials)
    drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    return client, drive_service

def abrir_hoja(nombre: str):
    client, _ = conectar_google()
    ss = client.open(SHEET_NAME)
    try: return ss.worksheet(nombre)
    except gspread.WorksheetNotFound: return ss.sheet1

@str_lit.cache_data(ttl=15, show_spinner=False)
def cargar_servicios_df() -> pd.DataFrame:
    try: 
        hoja = abrir_hoja(HOJA_SERVICIOS)
        datos = hoja.get_all_values()
        if not datos: return pd.DataFrame()
        
        cabecera = datos[0]
        nuevas_cabeceras = []
        vistos = {}
        for col in cabecera:
            col_limpia = str(col).strip()
            if col_limpia in vistos:
                vistos[col_limpia] += 1
                nuevas_cabeceras.append(f"{col_limpia}_{vistos[col_limpia]}")
            else:
                vistos[col_limpia] = 1
                nuevas_cabeceras.append(col_limpia)
                
        return pd.DataFrame(datos[1:], columns=nuevas_cabeceras)
    except Exception as e: 
        logger.error(f"Error servicios: {e}")
        return pd.DataFrame()

@str_lit.cache_data(ttl=60, show_spinner=False)
def cargar_usuarios_df() -> pd.DataFrame:
    try: return pd.DataFrame(abrir_hoja(HOJA_USUARIOS).get_all_records())
    except Exception as e: logger.error(f"Error usuarios: {e}"); return pd.DataFrame()

@str_lit.cache_data(ttl=5, show_spinner=False)
def cargar_asistencia_digital_df() -> pd.DataFrame:
    try: return pd.DataFrame(abrir_hoja(HOJA_ASISTENCIA).get_all_records())
    except Exception as e: logger.error(f"Error asistencia: {e}"); return pd.DataFrame()

@str_lit.cache_data(ttl=5, show_spinner=False)
def verificar_modo_mantenimiento() -> bool:
    try:
        hoja = abrir_hoja(HOJA_CONFIG)
        for row in hoja.get_all_records():
            if str(row.get("clave", "")).strip() == "modo_mantenimiento":
                return str(row.get("valor", "")).strip().upper() == "ACTIVO"
        return False
    except Exception: return False

def boton_refrescar():
    if str_lit.sidebar.button("🔄 Refrescar datos"):
        str_lit.cache_data.clear(); str_lit.rerun()

# --- ESTADO DE SESIÓN ---
_defaults = {"autenticado": False, "usuario_dni": None, "dependencia": None, "rol": None, "nombre_apellido": None, "cambio_clave_requerido": False, "form_counter": 0}
for k, v in _defaults.items(): str_lit.session_state.setdefault(k, v)

def validar_cuil(valor: str) -> bool: return bool(re.fullmatch(r"\d{11}", str(valor).strip()))
def calcular_monto_y_modulos(hora_inicio: datetime.time, hora_fin: datetime.time):
    dt_ini = datetime.datetime.combine(datetime.date.today(), hora_inicio)
    dt_fin = datetime.datetime.combine(datetime.date.today(), hora_fin)
    if dt_fin <= dt_ini: dt_fin += datetime.timedelta(days=1)
    horas = (dt_fin - dt_ini).total_seconds() / 3600.0
    return horas, max(1, min(MAX_MODULOS_PERMITIDOS, math.ceil(horas / HORAS_POR_MODULO)))
def calcular_monto_total(jefes_por_mod: list, subs_por_mod: list) -> float:
    return sum(((j * PRECIO_JEFE) + (s * PRECIO_SUB)) * i for i, (j, s) in enumerate(zip(jefes_por_mod, subs_por_mod), start=1))
def obtener_saludo(nombre_completo: str) -> str:
    hora_actual = datetime.datetime.now().hour
    nombre_str = str(nombre_completo).strip() if nombre_completo else "Efectivo"
    if 6 <= hora_actual < 13: return f"🌅 ¡Buen día, {nombre_str}!"
    elif 13 <= hora_actual < 20: return f"☀️ ¡Buenas tardes, {nombre_str}!"
    else: return f"🌙 ¡Buenas noches, {nombre_str}!"

def _post_apps_script(payload: dict) -> dict:
    try:
        r = requests.post(URL_APPS_SCRIPT, json=payload, timeout=TIMEOUT_HTTP)
        r.raise_for_status(); return r.json()
    except Exception as e: return {"status": "error", "message": str(e)}

def limpiar_texto_archivo(texto):
    texto = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'[^a-zA-Z0-9\s]', '', texto)
    return " ".join(texto.split()).strip()

# --- LOGIN Y CAMBIO DE CLAVE ---
def _validar_credenciales(dni, clave):
    try:
        data = cargar_usuarios_df()
        if data.empty: return False, "No se pudo cargar la base de datos de usuarios."
        data.columns = [str(c).strip().lower() for c in data.columns]
        user = data[data["dni"].astype(str) == str(dni)]
        if user.empty: return False, "DNI no registrado."

        u = user.iloc[0]
        clave_pers = str(u.get("clave personal", u.get("clave_personal", ""))).strip()
        clave_temp = str(u.get("clave", u.get("clave_temporal", ""))).strip()

        requiere_cambio = (clave_pers == "" or clave_pers == str(dni) or (clave_temp != "" and clave == clave_temp))
        if requiere_cambio:
            if clave == clave_temp or clave == str(dni) or clave == clave_pers:
                str_lit.session_state.cambio_clave_requerido = True
                str_lit.session_state.dni_temp = dni
                return True, "cambio"
            return False, "Contraseña temporal incorrecta."

        if clave == clave_pers:
            str_lit.session_state.autenticado = True
            str_lit.session_state.usuario_dni = dni
            str_lit.session_state.dependencia = str(u.get("dependencia", ""))
            str_lit.session_state.rol = str(u.get("rol", "")).strip()
            col_nombre = next((c for c in u.index if "nombre" in str(c).lower()), data.columns[1])
            str_lit.session_state.nombre_apellido = str(u.get(col_nombre, f"DNI: {dni}")).strip()
            return True, "ok"
        return False, "Contraseña incorrecta."
    except Exception as e:
        return False, f"Error de conexión: {e}"

def pantalla_cambio_clave():
    str_lit.title("🔑 Definir Nueva Contraseña Personal")
    str_lit.warning("⚠️ Primer ingreso detectado. Debés cambiar tu contraseña temporal.")
    nueva = str_lit.text_input("Nueva Contraseña", type="password")
    confirm = str_lit.text_input("Confirmar Nueva Contraseña", type="password")
    if str_lit.button("Guardar Nueva Contraseña"):
        if not nueva or nueva != confirm: str_lit.error("Las contraseñas no coinciden o están vacías."); return
        try:
            hoja = abrir_hoja(HOJA_USUARIOS)
            cell = hoja.find(str(str_lit.session_state.dni_temp))
            hoja.update_cell(cell.row, 6, nueva); hoja.update_cell(cell.row, 5, "")
            str_lit.success("¡Contraseña actualizada! Iniciá sesión nuevamente.")
            str_lit.session_state.cambio_clave_requerido = False
            str_lit.session_state.dni_temp = None
            str_lit.rerun()
        except Exception as e: str_lit.error(f"Error al actualizar: {e}")

def login():
    str_lit.markdown("### 🔐 Acceso al Sistema de Adicionales")
    with str_lit.form("form_login"):
        dni = str_lit.text_input("DNI / Usuario")
        clave = str_lit.text_input("Contraseña", type="password")
        submitted = str_lit.form_submit_button("INGRESAR AL SISTEMA")
    if submitted:
        if not dni or not clave: str_lit.error("Complete DNI y contraseña."); return
        with str_lit.spinner("Verificando credenciales..."):
            exito, msg = _validar_credenciales(dni, clave)
            if exito:
                if msg == "cambio": str_lit.rerun()
                else: str_lit.rerun()
            else: str_lit.error(msg)

# --- GENERADORES DE PDF ---
def _cabecera_pdf(pdf, titulo, subtitulo):
    pdf.set_fill_color(0, 32, 96)
    pdf.rect(0, 0, 210, 20, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 10)
    pdf.set_xy(5, 6)
    pdf.cell(200, 6, "POLICÍA DE LA PROVINCIA DE LA RIOJA - UNIDAD REGIONAL QUINTA - CHAMICAL", ln=True, align="center")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(12)
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 7, titulo, ln=True, align="center")
    pdf.set_font("helvetica", "I", 9)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(0, 5, subtitulo, ln=True, align="center")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)

def descargar_imagen_drive(url):
    if not url or not str(url).startswith("http"): return None
    try:
        file_id = None
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match: file_id = match.group(1)
        else:
            match = re.search(r'id=([a-zA-Z0-9-_]+)', url)
            if match: file_id = match.group(1)
        
        if not file_id: return None
        
        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(download_url, stream=True, timeout=10)
        if response.status_code == 200:
            return BytesIO(response.content)
        return None
    except Exception as e:
        logger.error(f"Error descargando imagen: {e}")
        return None

def generar_pdf_servicio(row: pd.Series) -> str:
    pdf = FPDF(); pdf.add_page(); pdf.set_auto_page_break(auto=True, margin=15)
    _cabecera_pdf(pdf, "INFORME DE SERVICIO ADICIONAL", f"ID de Solicitud: {row.get(COL_ID_SERV, '')}")
    
    def fila_tabla(lbl1, val1, lbl2, val2, w1=35, w2=60, w3=35, w4=60):
        pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(245, 247, 250)
        pdf.cell(w1, 6, f" {lbl1}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9); pdf.cell(w2, 6, f" {val1}", 1, 0, 'L')
        pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(245, 247, 250)
        pdf.cell(w3, 6, f" {lbl2}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9); pdf.cell(w4, 6, f" {val2}", 1, 1, 'L')

    def fila_completa(lbl, val, w_lbl=45, w_val=145):
        pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(245, 247, 250)
        pdf.cell(w_lbl, 6, f" {lbl}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9); pdf.cell(w_val, 6, f" {val}", 1, 1, 'L')

    def seccion_titulo(titulo):
        pdf.ln(2); pdf.set_font("helvetica", "B", 10); pdf.set_fill_color(0, 32, 96)
        pdf.set_text_color(255, 255, 255); pdf.cell(190, 6, f"  {titulo}", 1, 1, 'L', True)
        pdf.set_text_color(0, 0, 0)

    seccion_titulo("1. DATOS GENERALES DEL EVENTO")
    fecha_ev_str = str(row.get(COL_FECHA, ''))
    fila_tabla("Dependencia", str(row.get(COL_DEP, '')), "Fecha(s)", fecha_ev_str)
    fila_tabla("Tipo de Evento", str(row.get(COL_TIPO_EV, '')), "Horario", str(row.get(COL_HORARIO, '')))
    fila_completa("Ubicación / Predio", str(row.get(COL_UBI, '')))
    fila_completa("Organizador", str(row.get('Organizador', '')))

    seccion_titulo("2. RECURSOS Y MÓDULOS ASIGNADOS")
    monto_val = row.get('Monto Total', 0)
    monto_str = f"${float(monto_val):,.2f}" if monto_val else "$0.00"
    
    js_simple = int(float(row.get('Jefes Simples', 0) or 0))
    subs_simple = int(float(row.get('Subalternos Simples', 0) or 0))
    js_doble = int(float(row.get('Jefes Dobles', 0) or 0))
    subs_doble = int(float(row.get('Subalternos Dobles', 0) or 0))
    js_triple = int(float(row.get('Jefes Triples', 0) or 0))
    subs_triple = int(float(row.get('Subalternos Triples', 0) or 0))
    
    # DISEÑO CORREGIDO: TABLA DESGLOSADA
    pdf.set_font("helvetica", "B", 9)
    pdf.set_fill_color(245, 247, 250)
    pdf.cell(45, 6, " Módulo", 1, 0, 'L', True)
    pdf.cell(35, 6, " Jefes", 1, 0, 'L', True)
    pdf.cell(35, 6, " Subalternos", 1, 1, 'L', True)
    
    pdf.set_font("helvetica", "", 9)
    if js_simple > 0 or subs_simple > 0:
        pdf.cell(45, 6, " Simple (4h)", 1, 0, 'L')
        pdf.cell(35, 6, f" {js_simple}", 1, 0, 'L')
        pdf.cell(35, 6, f" {subs_simple}", 1, 1, 'L')
    if js_doble > 0 or subs_doble > 0:
        pdf.cell(45, 6, " Doble (8h)", 1, 0, 'L')
        pdf.cell(35, 6, f" {js_doble}", 1, 0, 'L')
        pdf.cell(35, 6, f" {subs_doble}", 1, 1, 'L')
    if js_triple > 0 or subs_triple > 0:
        pdf.cell(45, 6, " Triple (12h)", 1, 0, 'L')
        pdf.cell(35, 6, f" {js_triple}", 1, 0, 'L')
        pdf.cell(35, 6, f" {subs_triple}", 1, 1, 'L')

    pdf.ln(2)
    fila_tabla("Módulos (Max)", str(row.get('Módulos', 1)), "Monto Total", monto_str)

    seccion_titulo("3. DATOS DE FACTURACIÓN Y PAGO")
    fila_completa("Pagado por", str(row.get('Pagador', '')))
    fila_completa("CUIL / CUIT Facturación", str(row.get('CUIL Pagador', '')))

    seccion_titulo("4. DOCUMENTACIÓN ADJUNTA (GOOGLE DRIVE)")
    col_hab = next((c for c in row.index if "Enlace Habilitación" in str(c)), "Enlace Habilitación")
    col_comp = next((c for c in row.index if "Enlace Comprobante" in str(c)), "Enlace Comprobante")
    col_cont = next((c for c in row.index if "Enlace Contrato" in str(c)), "Enlace Contrato")
    
    url_hab = str(row.get(col_hab, ""))
    url_comp = str(row.get(col_comp, ""))
    url_cont = str(row.get(col_cont, ""))

    def link_pdf(etiqueta, url):
        pdf.set_fill_color(245, 247, 250)
        pdf.set_font("helvetica", "B", 9)
        pdf.cell(45, 6, f" {etiqueta}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 9)
        if url.startswith("http"):
            pdf.set_text_color(0, 102, 204)
            pdf.cell(145, 6, " [ Ver Documento ]", 1, 1, 'L', link=url)
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.set_text_color(120, 120, 120)
            pdf.cell(145, 6, " No adjuntado", 1, 1, 'L')
            pdf.set_text_color(0, 0, 0)

    link_pdf("Habilitación Municipal", url_hab)
    link_pdf("Comprobante de Pago", url_comp)
    link_pdf("Contrato", url_cont)

    # INCRUSTAR IMÁGENES
    def incrustar_imagen(titulo, url):
        if not url or not str(url).startswith("http"): return
        img_data = descargar_imagen_drive(url)
        if not img_data: return
        
        try:
            img = Image.open(img_data)
            if img.mode in ("RGBA", "P"): img = img.convert("RGB")
            
            pdf.add_page()
            pdf.set_font("helvetica", "B", 12)
            pdf.set_text_color(0, 32, 96)
            pdf.cell(0, 10, f"ANEXO: {titulo.upper()}", ln=True, align="center")
            pdf.ln(5)
            
            ancho_max = 190
            ancho_img, alto_img = img.size
            proporcion = alto_img / ancho_img
            alto_render = ancho_max * proporcion
            
            if alto_render > 250:
                alto_render = 250
                ancho_max = alto_render / proporcion

            x = (210 - ancho_max) / 2
            tmp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            img.save(tmp_img.name, format="JPEG")
            
            pdf.image(tmp_img.name, x=x, y=pdf.get_y(), w=ancho_max, h=alto_render)
            os.unlink(tmp_img.name)
        except Exception as e:
            logger.error(f"Error incrustando imagen: {e}")
            pdf.set_font("helvetica", "I", 10)
            pdf.cell(0, 10, " [ No se pudo incrustar la imagen ] ", ln=True, align="center")

    if url_comp.startswith("http"): incrustar_imagen("Comprobante de Pago", url_comp)
    if url_hab.startswith("http"): incrustar_imagen("Habilitación Municipal", url_hab)
    if url_cont.startswith("http"): incrustar_imagen("Contrato", url_cont)

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name

def generar_pdf_asistencia(df_asistentes: pd.DataFrame, id_servicio: str, tipo_evento: str, fecha_evento: str, horario_evento: str = "") -> str:
    pdf = FPDF(); pdf.add_page(); pdf.set_auto_page_break(auto=True, margin=15)
    _cabecera_pdf(pdf, "PLANILLA DE ASISTENCIA - SERVICIO ADICIONAL", f"ID de Servicio: {id_servicio}")

    pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(245, 247, 250)
    pdf.cell(45, 6, " Evento:", 1, 0, 'L', True); pdf.set_font("helvetica", "", 9); pdf.cell(145, 6, f" {tipo_evento}", 1, 1, 'L')
    pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(245, 247, 250)
    pdf.cell(45, 6, " Fecha:", 1, 0, 'L', True); pdf.set_font("helvetica", "", 9); pdf.cell(145, 6, f" {fecha_evento}", 1, 1, 'L')
    pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(245, 247, 250)
    pdf.cell(45, 6, " Horario:", 1, 0, 'L', True); pdf.set_font("helvetica", "", 9); pdf.cell(145, 6, f" {horario_evento}", 1, 1, 'L')
    pdf.ln(5)

    pdf.set_font("helvetica", "B", 9); pdf.set_fill_color(0, 32, 96); pdf.set_text_color(255, 255, 255)
    pdf.cell(20, 7, " Puesto", 1, 0, 'L', True); pdf.cell(35, 7, " DNI", 1, 0, 'L', True)
    pdf.cell(45, 7, " Jerarquía", 1, 0, 'L', True); pdf.cell(90, 7, " Nombre y Apellido", 1, 1, 'L', True)
    pdf.set_text_color(0, 0, 0)

    for _, row in df_asistentes.iterrows():
        pdf.set_font("helvetica", "", 9); pdf.cell(20, 6, f" {str(row.get('Puesto Asignado', ''))}", 1, 0, 'L')
        pdf.set_font("courier", "B", 10); pdf.cell(35, 6, f" {str(row.get('DNI', ''))}", 1, 0, 'L')
        pdf.set_font("helvetica", "", 9); pdf.cell(45, 6, f" {str(row.get('Jerarquía', ''))}", 1, 0, 'L')
        pdf.cell(90, 6, f" {str(row.get('Nombre y Apellido', ''))}", 1, 1, 'L')

    pdf.ln(10); pdf.set_font("helvetica", "I", 8); pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 5, f"Documento generado el {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="center")
    pdf.cell(0, 5, f"Total de efectivos: {len(df_asistentes)}", ln=True, align="center")

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name

def _render_modulos(max_modulos: int, counter: int):
    labels_base = ["Simple (4h)", "Doble (8h)", "Triple (12h)"]
    labels_ordenados = labels_base[:max_modulos][::-1]
    jefes_dict, subs_dict = {}, {}
    for label in labels_ordenados:
        str_lit.markdown(f"**Turno {label}**")
        c1, c2 = str_lit.columns(2)
        with c1: j = str_lit.number_input(f"Jefes ({label})", 0, step=1, key=f"j_{label}_{counter}")
        with c2: s = str_lit.number_input(f"Subalternos ({label})", 0, step=1, key=f"s_{label}_{counter}")
        if "Simple" in label: jefes_dict[0] = int(j); subs_dict[0] = int(s)
        elif "Doble" in label: jefes_dict[1] = int(j); subs_dict[1] = int(s)
        elif "Triple" in label: jefes_dict[2] = int(j); subs_dict[2] = int(s)
    return [jefes_dict.get(i, 0) for i in range(3)], [subs_dict.get(i, 0) for i in range(3)]

def vista_dependencia():
    nombre_efectivo = str_lit.session_state.get("nombre_apellido", "Personal Policial")
    saludo = obtener_saludo(nombre_efectivo)
    
    str_lit.sidebar.markdown(f"**{nombre_efectivo}** ({str_lit.session_state.dependencia})")
    boton_refrescar()
    if str_lit.sidebar.button("🚪 Cerrar Sesión"): str_lit.session_state.autenticado = False; str_lit.rerun()

    str_lit.markdown(f"### {saludo}")
    tab1, tab2, tab3 = str_lit.tabs(["🟢 Cargar Nuevo Servicio", "⏳ Pendientes de Asistencia", "📜 Historial de Servicios"])
    counter = str_lit.session_state.form_counter

    with tab1:
        str_lit.header("Solicitud de Servicio Adicional")
        if f"fecha_elegida_{counter}" not in str_lit.session_state: str_lit.session_state[f"fecha_elegida_{counter}"] = False
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
        if btn_confirmar_fecha: str_lit.session_state[f"fecha_elegida_{counter}"] = True

        if str_lit.session_state.get(f"fecha_elegida_{counter}", False):
            fecha_texto = f"{fec_inicio.strftime('%d/%m/%Y')}" if (tipo_cant_dias == "1 Día único" or fec_inicio == fec_fin) else f"{fec_inicio.strftime('%d/%m/%Y')} y {fec_fin.strftime('%d/%m/%Y')}"
            cant_dias = 1 if (tipo_cant_dias == "1 Día único" or fec_inicio == fec_fin) else 2
            tipo_evento = str_lit.text_input("Tipo de evento", key=f"tipo_{counter}")
            if tipo_evento.strip():
                col_org1, col_org2 = str_lit.columns(2)
                with col_org1: nombre_organizador = str_lit.text_input("Nombre Organizador", key=f"nom_org_{counter}")
                with col_org2: dni_organizador = str_lit.text_input("DNI Organizador", key=f"dni_org_{counter}")
                
                if nombre_organizador.strip() and dni_organizador.strip():
                    es_mismo_pagador = str_lit.checkbox("El organizador es la misma persona que paga", value=True, key=f"mismo_pag_{counter}")
                    
                    if es_mismo_pagador:
                        pagado_por = f"{nombre_organizador} (DNI: {dni_organizador})"
                        cuil_pagado_por = str_lit.text_input("CUIL / CUIT del Organizador (11 dígitos)", key=f"cuil_org_{counter}")
                    else:
                        col_pag1, col_pag2, col_pag3 = str_lit.columns(3)
                        with col_pag1: nombre_pagador = str_lit.text_input("Nombre del Pagador", key=f"nom_pag_{counter}")
                        with col_pag2: dni_pagador = str_lit.text_input("DNI del Pagador", key=f"dni_pag_{counter}")
                        with col_pag3: cuil_pagado_por = str_lit.text_input("CUIL / CUIT a Facturar (11 dígitos)", key=f"cuil_pag_{counter}")
                        pagado_por = f"{nombre_pagador} (DNI: {dni_pagador})"

                    if pagado_por.strip() and cuil_pagado_por.strip():
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
                            
                            str_lit.markdown(f"""
                                <div class="monto-box">
                                    <div style="font-size: 0.9rem; color: #555; font-weight: bold; margin-bottom: 5px;">TOTAL ESTIMADO A PAGAR (aprox)</div>
                                    <div class="monto-total">${monto_total:,.2f}</div>
                                    <div style="font-size: 0.8rem; color: #666; margin-top: 5px;">Jefes: {total_jefes} | Subalternos: {total_subalternos} (Módulos: {max_modulos})</div>
                                </div>
                            """, unsafe_allow_html=True)
                            
                            with str_lit.form(f"form_envio_{counter}"):
                                file_comp = str_lit.file_uploader("Comprobante de Pago (OBLIGATORIO)", type=["png", "jpg", "jpeg", "pdf"], key=f"comp_{counter}")
                                file_hab = str_lit.file_uploader("Habilitación Municipal (Opcional)", type=["png", "jpg", "jpeg", "pdf"], key=f"hab_{counter}")
                                file_cont = str_lit.file_uploader("Contrato (Opcional)", type=["png", "jpg", "jpeg", "pdf"], key=f"cont_{counter}")
                                submitted = str_lit.form_submit_button("🚀 Enviar Solicitud Definitiva")

                            if submitted:
                                if total_jefes == 0 and total_subalternos == 0: 
                                    str_lit.error("Asigne al menos un efectivo.")
                                elif not file_comp: 
                                    str_lit.error("Debe subir el Comprobante de Pago obligatoriamente.")
                                elif not validar_cuil(cuil_pagado_por): 
                                    str_lit.error("El CUIL / CUIT ingresado es inválido (deben ser 11 números).")
                                else:
                                    id_servicio = f"SERV-{int(datetime.datetime.now().timestamp())}"
                                    payload = {
                                        "action": "crear_servicio",
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
                                        "file_comp_base64": base64.b64encode(file_comp.getvalue()).decode("utf-8"),
                                        "name_comp": f"{id_servicio}_comprobante.{file_comp.name.split('.')[-1]}", 
                                        "type_comp": file_comp.type,
                                        "file_hab_base64": base64.b64encode(file_hab.getvalue()).decode("utf-8") if file_hab else "",
                                        "name_hab": f"{id_servicio}_habilitacion.{file_hab.name.split('.')[-1]}" if file_hab else "", 
                                        "type_hab": file_hab.type if file_hab else "",
                                        "file_cont_base64": base64.b64encode(file_cont.getvalue()).decode("utf-8") if file_cont else "",
                                        "name_cont": f"{id_servicio}_contrato.{file_cont.name.split('.')[-1]}" if file_cont else "", 
                                        "type_cont": file_cont.type if file_cont else "",
                                    }
                                    with str_lit.spinner("Enviando solicitud a la nube..."):
                                        res = _post_apps_script(payload)
                                        if res.get("status") == "success":
                                            str_lit.success("¡Solicitud registrada con éxito!")
                                            str_lit.session_state[f"fecha_elegida_{counter}"] = False
                                            str_lit.session_state.form_counter += 1
                                            str_lit.cache_data.clear(); str_lit.rerun()
                                        else: 
                                            str_lit.error(f"Error al enviar: {res.get('message')}")

    with tab2:
        str_lit.header("Servicios Pendientes de Planilla de Asistencia")
        data = cargar_servicios_df()
        if not data.empty and COL_DEP in data.columns:
            df_dep = data[data[COL_DEP].astype(str) == str(str_lit.session_state.dependencia)]
            col_a1 = next((c for c in df_dep.columns if "asistencia día 1" in str(c).lower() or "asistencia 1" in str(c).lower()), None)
            col_a2 = next((c for c in df_dep.columns if "asistencia día 2" in str(c).lower() or "asistencia 2" in str(c).lower()), None)
            col_gen = next((c for c in df_dep.columns if "enlace asistencia" in str(c).lower()), None)

            def esta_pendiente(r):
                if not col_a1 and not col_gen: return False
                fec = str(r.get(COL_FECHA, ""))
                a1_val = str(r.get(col_a1, r.get(col_gen, ""))).strip()
                a2_val = str(r.get(col_a2, "")).strip()
                return (a1_val == "" or a2_val == "") if " y " in fec else (a1_val == "")

            df_pend = df_dep[df_dep.apply(esta_pendiente, axis=1)]
            if df_pend.empty: str_lit.success("🎉 ¡Sin pendientes!")
            else:
                for _, row in df_pend.iterrows():
                    id_serv = row.get(COL_ID_SERV) or row.get('id_servicio') or row.iloc[0]
                    fec_ev = str(row.get(COL_FECHA) or '')
                    with str_lit.expander(f"📁 EXPEDIENTE: {id_serv} | 📅 {fec_ev}"):
                        with str_lit.form(f"form_asist_1_{id_serv}"):
                            file_a1 = str_lit.file_uploader("Planilla Asistencia (Escaneada / Foto)", type=["png", "jpg", "jpeg", "pdf"], key=f"f1_{id_serv}")
                            sub_a1 = str_lit.form_submit_button("Subir Archivo al Expediente")
                        if sub_a1 and file_a1:
                            payload = {"action": "subir_asistencia", "id_servicio": str(id_serv), "dia_num": 1,
                                "file_asist_base64": base64.b64encode(file_a1.getvalue()).decode("utf-8"),
                                "name_asist": f"{id_serv}_asistencia.{file_a1.name.split('.')[-1]}", "type_asist": file_a1.type}
                            with str_lit.spinner("Subiendo archivo..."):
                                res = _post_apps_script(payload)
                                if res.get("status") == "success":
                                    str_lit.success("¡Planilla adjuntada!")
                                    str_lit.cache_data.clear(); str_lit.rerun()
                                else: str_lit.error(f"Error: {res.get('message')}")

    with tab3:
        str_lit.header("Historial General")
        data = cargar_servicios_df()
        if not data.empty and COL_DEP in data.columns:
            str_lit.dataframe(data[data[COL_DEP].astype(str) == str(str_lit.session_state.dependencia)], use_container_width=True)

# --- PANEL DE ADMINISTRACIÓN ---
def vista_admin():
    nombre_efectivo = str_lit.session_state.get("nombre_apellido", "Administrador General")
    saludo = obtener_saludo(nombre_efectivo)

    str_lit.sidebar.markdown(f"**{nombre_efectivo}** (ADMINISTRADOR)")
    str_lit.sidebar.markdown("---")
    str_lit.sidebar.subheader("⚙️ Control de Sistema")
    modo_mant_actual = verificar_modo_mantenimiento()
    nuevo_modo_toggle = str_lit.sidebar.toggle("🛠️ Modo Mantenimiento", value=modo_mant_actual)
    
    if nuevo_modo_toggle != modo_mant_actual:
        estado_str = "ACTIVO" if nuevo_modo_toggle else "INACTIVO"
        with str_lit.spinner("Actualizando..."):
            res = _post_apps_script({"action": "set_mantenimiento", "estado": estado_str})
            if res.get("status") == "success":
                str_lit.cache_data.clear(); str_lit.success(f"Modo: {estado_str}"); str_lit.rerun()
            else: str_lit.error("No se pudo actualizar el mantenimiento.")

    boton_refrescar()
    if str_lit.sidebar.button("🚪 Cerrar Sesión"): str_lit.session_state.autenticado = False; str_lit.rerun()

    str_lit.markdown(f"### {saludo}")
    str_lit.title("📋 Control General y Despacho a Central")
    data = cargar_servicios_df()
    
    if not data.empty and COL_ESTADO_CENTRAL in data.columns:
        pendientes = data[data[COL_ESTADO_CENTRAL].astype(str).str.contains("Recibido", na=False)]
        if not pendientes.empty: str_lit.warning(f"🔔 Hay **{len(pendientes)}** servicio(s) pendientes de revisión y despacho.")

    tab_exp, tab_enlaces, tab_asist, tab_depuracion = str_lit.tabs([
        "📄 Expedientes", "📲 Enlaces (Asistencia)", "📊 Asistencias Digitales", "🧹 Mantenimiento"
    ])

    with tab_exp:
        str_lit.subheader("Gestión de Expedientes (Central)")
        if not data.empty:
            filtro_estado = str_lit.radio("Filtrar expedientes:", ["Todas", "Pendientes", "Enviadas"], horizontal=True, key="filtro_exp")
            df_filtrado = data.copy()
            if filtro_estado == "Pendientes" and COL_ESTADO_CENTRAL in df_filtrado.columns:
                df_filtrado = df_filtrado[~df_filtrado[COL_ESTADO_CENTRAL].astype(str).str.contains("Enviado", na=False)]
            elif filtro_estado == "Enviadas" and COL_ESTADO_CENTRAL in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado[COL_ESTADO_CENTRAL].astype(str).str.contains("Enviado", na=False)]

            str_lit.dataframe(df_filtrado, use_container_width=True)
            opciones = [f"{r.get(COL_ID_SERV, '')} - {r.get(COL_DEP, '')} ({r.get(COL_FECHA, '')})" for _, r in df_filtrado.iterrows()]
            if opciones:
                elegido = str_lit.selectbox("Seleccionar servicio para despachar", options=opciones, key="sel_exp")
                if elegido:
                    id_sel = elegido.split(" - ")[0]
                    row_data = data[data[COL_ID_SERV].astype(str) == id_sel].iloc[0]
                    str_lit.markdown("---")
                    c_pdf, c_estado = str_lit.columns(2)
                    with c_pdf:
                        if str_lit.button("🚀 Generar PDF con Constancias"):
                            with str_lit.spinner("Generando PDF..."):
                                pdf_path = generar_pdf_servicio(row_data)
                                with open(pdf_path, "rb") as f: pdf_bytes = f.read()
                                os.unlink(pdf_path)
                                
                                nombre_ev = limpiar_texto_archivo(row_data.get(COL_TIPO_EV, ""))
                                fec_sel = str(row_data.get(COL_FECHA, ""))
                                try:
                                    partes_fecha = fec_sel.split("/")
                                    fecha_formateada = f"{partes_fecha[0].zfill(2)}{partes_fecha[1].zfill(2)}{partes_fecha[2][-2:]}" if len(partes_fecha) == 3 else fec_sel.replace("/", "")
                                except: fecha_formateada = fec_sel.replace("/", "")
                                
                                nombre_archivo = f"Constancia_{nombre_ev}_{fecha_formateada}.pdf".replace(" ", "_")
                                
                                str_lit.download_button(label=f"📥 Descargar {nombre_archivo}", data=pdf_bytes, file_name=nombre_archivo, mime="application/pdf")
                    with c_estado:
                        if str_lit.button("✅ Marcar como Enviado a Central"):
                            try:
                                hoja_serv = abrir_hoja(HOJA_SERVICIOS)
                                cell = hoja_serv.find(str(id_sel))
                                if cell:
                                    hoja_serv.update_cell(cell.row, 25, "🟢 Enviado a Central")
                                    str_lit.success("¡Marcado como enviado a Central!")
                                    str_lit.cache_data.clear(); str_lit.rerun()
                            except Exception as e: str_lit.error(f"Error: {e}")

    with tab_enlaces:
        str_lit.subheader("Enviar Enlace de Asistencia al Encargado")
        if not data.empty:
            filtro_estado_env = str_lit.radio("Filtrar eventos:", ["Todas", "Pendientes", "Enviadas"], horizontal=True, key="filtro_env")
            df_filtrado_env = data.copy()
            if filtro_estado_env == "Pendientes" and COL_ENLACE_ASIST in df_filtrado_env.columns:
                df_filtrado_env = df_filtrado_env[df_filtrado_env[COL_ENLACE_ASIST].astype(str).str.strip() == ""]
            elif filtro_estado_env == "Enviadas" and COL_ENLACE_ASIST in df_filtrado_env.columns:
                df_filtrado_env = df_filtrado_env[df_filtrado_env[COL_ENLACE_ASIST].astype(str).str.strip() != ""]

            str_lit.dataframe(df_filtrado_env[[COL_ID_SERV, COL_TIPO_EV, COL_FECHA, COL_ENLACE_ASIST]], use_container_width=True)
            opciones_env = [f"{r.get(COL_ID_SERV, '')} - {r.get(COL_DEP, '')} ({r.get(COL_FECHA, '')})" for _, r in df_filtrado_env.iterrows()]
            if opciones_env:
                elegido_env = str_lit.selectbox("Seleccionar servicio para enviar enlace", options=opciones_env, key="sel_env")
                if elegido_env:
                    id_sel_env = elegido_env.split(" - ")[0]
                    row_data_env = data[data[COL_ID_SERV].astype(str) == id_sel_env].iloc[0]
                    link_formulario_limpio = f"{URL_APPS_SCRIPT}?id_servicio={id_sel_env}"
                    str_lit.markdown("---")
                    str_lit.markdown("##### 📤 Enviar Planilla al Encargado del Evento")
                    mensaje_wa = f"Hola, te envío el enlace para cargar la planilla de asistencia del servicio *{id_sel_env}* ({row_data_env.get(COL_TIPO_EV, '')}):\n\n{link_formulario_limpio}"
                    url_wa_directo = f"https://api.whatsapp.com/send?text={requests.utils.quote(mensaje_wa)}"
                    str_lit.markdown(f'<a href="{url_wa_directo}" target="_blank" style="text-decoration:none;"><button style="background-color: #25D366; color: white; border: none; padding: 12px 20px; font-size: 1rem; font-weight: bold; border-radius: 8px; cursor: pointer; width: 100%; margin-bottom: 10px;">📲 Enviar Formulario por WhatsApp</button></a>', unsafe_allow_html=True)
                    with str_lit.expander("Ver enlace en texto"): str_lit.code(link_formulario_limpio, language="text")
                    if str_lit.button("✅ Marcar Enlace como Enviado"):
                        try:
                            hoja_serv = abrir_hoja(HOJA_SERVICIOS)
                            cell = hoja_serv.find(str(id_sel_env))
                            if cell:
                                hoja_serv.update_cell(cell.row, 27, "🟢 Enviado")
                                str_lit.success("¡Enlace marcado como enviado!")
                                str_lit.cache_data.clear(); str_lit.rerun()
                        except Exception as e: str_lit.error(f"Error: {e}")

    with tab_asist:
        str_lit.subheader("Listado y Descarga de Asistencias")
        df_asist = cargar_asistencia_digital_df()
        if not df_asist.empty and COL_ID_SERV in df_asist.columns:
            df_agrupado = df_asist.groupby(COL_ID_SERV).agg(Efectivos_Cargados=("DNI", "count")).reset_index()
            if COL_ID_SERV in data.columns:
                df_resumen = pd.merge(df_agrupado, data, on=COL_ID_SERV, how="left")
            else:
                df_resumen = df_agrupado
                df_resumen[COL_TIPO_EV] = "Evento Desconocido"
                df_resumen[COL_FECHA] = ""
                df_resumen[COL_UBI] = ""

            if COL_ESTADO_ASIST not in df_resumen.columns:
                df_resumen[COL_ESTADO_ASIST] = ""
                
            filtro_asist = str_lit.radio("Filtrar por estado de envío:", ["Todas", "Pendientes", "Enviadas a Central"], horizontal=True, key="filtro_asist")
            if filtro_asist == "Pendientes":
                df_mostrar = df_resumen[~df_resumen[COL_ESTADO_ASIST].astype(str).str.contains("Enviada", na=False)]
            elif filtro_asist == "Enviadas a Central":
                df_mostrar = df_resumen[df_resumen[COL_ESTADO_ASIST].astype(str).str.contains("Enviada", na=False)]
            else:
                df_mostrar = df_resumen

            cols_a_mostrar = [COL_ID_SERV, COL_TIPO_EV, COL_FECHA, "Efectivos_Cargados"]
            if COL_ESTADO_ASIST in df_mostrar.columns: cols_a_mostrar.append(COL_ESTADO_ASIST)
            str_lit.dataframe(df_mostrar[cols_a_mostrar], use_container_width=True)
            str_lit.markdown("---")
            
            opciones_eventos = [f"{r.get(COL_ID_SERV, '')} | {r.get(COL_TIPO_EV, 'Evento')} - {r.get(COL_FECHA, '')} ({r.get(COL_UBI, '')})" for _, r in df_mostrar.iterrows()]
            evento_elegido = str_lit.selectbox("Seleccioná el evento:", options=opciones_eventos, key="sel_asist")
            
            if evento_elegido:
                id_sel_asist = evento_elegido.split(" | ")[0]
                df_filtrado_evento = df_asist[df_asist[COL_ID_SERV].astype(str) == id_sel_asist]
                row_serv_asist = df_resumen[df_resumen[COL_ID_SERV].astype(str) == id_sel_asist].iloc[0]
                
                c_btn_pdf, c_btn_marcar = str_lit.columns(2)
                with c_btn_pdf:
                    if str_lit.button("📄 Generar PDF de Asistencia"):
                        with str_lit.spinner("Generando PDF..."):
                            pdf_path_asist = generar_pdf_asistencia(df_filtrado_evento, id_sel_asist, str(row_serv_asist.get(COL_TIPO_EV, "")), str(row_serv_asist.get(COL_FECHA, "")), str(row_serv_asist.get(COL_HORARIO, "")))
                            with open(pdf_path_asist, "rb") as f: pdf_bytes_asist = f.read()
                            os.unlink(pdf_path_asist)
                            
                            nombre_ev = limpiar_texto_archivo(row_serv_asist.get(COL_TIPO_EV, ""))
                            ubi_ev = limpiar_texto_archivo(row_serv_asist.get(COL_UBI, ""))
                            fec_sel = str(row_serv_asist.get(COL_FECHA, ""))
                            try:
                                partes_fecha = fec_sel.split("/")
                                fecha_formateada = f"{partes_fecha[0].zfill(2)}{partes_fecha[1].zfill(2)}{partes_fecha[2][-2:]}" if len(partes_fecha) == 3 else fec_sel.replace("/", "")
                            except: fecha_formateada = fec_sel.replace("/", "")
                            
                            nombre_final_pdf = f"asistencia {nombre_ev} predio {ubi_ev} {fecha_formateada}.pdf".replace(" ", "_")
                            str_lit.download_button(label=f"📥 Descargar {nombre_final_pdf}", data=pdf_bytes_asist, file_name=nombre_final_pdf, mime="application/pdf")

                with c_btn_marcar:
                    if not str(row_serv_asist.get(COL_ESTADO_ASIST, "")).startswith("🟢"):
                        if str_lit.button("✅ Marcar Asistencia como Enviada"):
                            with str_lit.spinner("Actualizando..."):
                                res = _post_apps_script({"action": "marcar_asistencia_enviada", "id_servicio": id_sel_asist})
                                if res.get("status") == "success":
                                    str_lit.success("¡Planilla de asistencia marcada como enviada!")
                                    str_lit.cache_data.clear()
                                    str_lit.rerun()
                                else:
                                    str_lit.error(f"Error: {res.get('message')}")
                    else:
                        str_lit.success("✅ Esta planilla ya fue marcada como enviada a Central.")
        else:
            str_lit.info("Aún no hay registros de asistencia o falta la columna 'ID Servicio'.")

    with tab_depuracion:
        str_lit.subheader("🧹 Depuración y Archivado Mensual")
        str_lit.warning("Esta acción moverá los registros a una solapa histórica.")
        periodo_inp = str_lit.text_input("Periodo del Histórico", value=datetime.date.today().strftime("%Y_%m"))
        if str_lit.button("⚠️ Ejecutar Limpieza Mensual"):
            with str_lit.spinner("Depurando..."):
                res = _post_apps_script({"action": "depurar_mes", "periodo": periodo_inp})
                if res.get("status") == "success":
                    str_lit.success(f"¡Migrado a Historico_Asistencia_{periodo_inp}!")
                    str_lit.cache_data.clear()
                else: str_lit.error(f"Error: {res.get('message')}")

# --- ENRUTADOR PRINCIPAL ---
modo_mantenimiento_activo = verificar_modo_mantenimiento()

if str_lit.session_state.get("cambio_clave_requerido", False):
    pantalla_cambio_clave()
elif not str_lit.session_state.autenticado:
    login()
elif str_lit.session_state.rol == "Administrador":
    vista_admin()
elif modo_mantenimiento_activo:
    str_lit.markdown("""<div style="text-align: center; padding: 4rem 2rem;"><h1>🛠️ SITIO EN MANTENIMIENTO</h1><p style="font-size: 1.2rem; color: #555;">Por favor intente nuevamente en unos momentos.</p></div>""", unsafe_allow_html=True)
    if str_lit.button("🚪 Cerrar Sesión"): str_lit.session_state.autenticado = False; str_lit.rerun()
else:
    vista_dependencia()
