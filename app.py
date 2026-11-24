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
import streamlit.components.v1 as components
import time

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
    .suspension-box {
        background: #FFEBEE; border: 2px solid #C62828; border-radius: 8px;
        padding: 15px; text-align: center; margin-bottom: 20px;
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
    URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbwh6t07cnZ5n2w1cBFRQo6ZIdIab3nYJ-iASS6VSia5J5ca18-SD11OJSzYnHm-NFTZ/exec"

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

@str_lit.cache_data(ttl=5, show_spinner=False)
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

# --- FUNCIONES DE MEMORIA DE ORGANIZADORES Y PAGADORES ---
@str_lit.cache_data(ttl=5, show_spinner=False)
def cargar_organizadores_df() -> pd.DataFrame:
    try: 
        df = pd.DataFrame(abrir_hoja("Organizadores").get_all_records())
        if not df.empty: df.columns = [str(c).strip().lower() for c in df.columns]
        return df
    except Exception as e: 
        logger.error(f"Error organizadores: {e}")
        return pd.DataFrame()

@str_lit.cache_data(ttl=5, show_spinner=False)
def cargar_pagadores_df() -> pd.DataFrame:
    try: 
        df = pd.DataFrame(abrir_hoja("Pagadores").get_all_records())
        if not df.empty: df.columns = [str(c).strip().lower() for c in df.columns]
        return df
    except Exception as e: 
        logger.error(f"Error pagadores: {e}")
        return pd.DataFrame()
# ---------------------------------------------------------

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

# --- NUEVA FUNCIÓN PARA GENERAR IDs CORTOS ---
def generar_id_corto(dependencia, fecha_str, df_existente):
    # Tomamos las primeras 4 letras de la dependencia
    dep_code = limpiar_texto_archivo(dependencia)[:4].upper()
    # Formateamos la fecha a DDMMAA
    try:
        partes = fecha_str.split("/")
        fecha_code = f"{partes[0]}{partes[1]}{partes[2][-2:]}"
    except:
        fecha_code = datetime.date.today().strftime("%d%m%y")
    
    base_id = f"{dep_code}-{fecha_code}"
    
    # Buscamos el último número usado para esa combinación
    correlativo = 1
    if not df_existente.empty and COL_ID_SERV in df_existente.columns:
        ids_existentes = df_existente[COL_ID_SERV].astype(str).tolist()
        # Filtramos los que empiezan con nuestra base_id
        coincidencias = [i for i in ids_existentes if i.startswith(base_id)]
        if coincidencias:
            # Extraemos el número más alto
            numeros = []
            for c in coincidencias:
                try:
                    num = int(c.split("-")[-1].split("_")[0])
                    numeros.append(num)
                except:
                    pass
            if numeros:
                correlativo = max(numeros) + 1
                
    return f"{base_id}-{correlativo:02d}"

# --- REPORTE DE SUSPENSIÓN ---
def generar_pdf_suspension(row: pd.Series, motivo: str) -> str:
    pdf = FPDF(); pdf.add_page()
    _cabecera_pdf(pdf, "INFORME DE SUSPENSIÓN DE SERVICIO", f"ID de Servicio: {row.get(COL_ID_SERV, '')}")
    
    # Cartel Rojo Gigante
    pdf.set_fill_color(198, 40, 40) # Rojo
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 24)
    pdf.cell(0, 20, "EVENTO SUSPENDIDO", ln=True, align="center", fill=True)
    
    pdf.ln(5)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 10, f"Motivo: {motivo}", ln=True, align="center")
    
    pdf.ln(10)
    
    def fila_completa(lbl, val, w_lbl=45, w_val=145):
        pdf.set_font("helvetica", "B", 10); pdf.set_fill_color(245, 247, 250)
        pdf.cell(w_lbl, 8, f" {lbl}", 1, 0, 'L', True)
        pdf.set_font("helvetica", "", 10); pdf.cell(w_val, 8, f" {val}", 1, 1, 'L')

    def seccion_titulo(titulo):
        pdf.ln(4); pdf.set_font("helvetica", "B", 11); pdf.set_fill_color(0, 32, 96)
        pdf.set_text_color(255, 255, 255); pdf.cell(190, 8, f"  {titulo}", 1, 1, 'L', True)
        pdf.set_text_color(0, 0, 0)

    seccion_titulo("DATOS DEL EVENTO ORIGINAL")
    fila_completa("Dependencia", str(row.get(COL_DEP, '')))
    fila_completa("Tipo de Evento", str(row.get(COL_TIPO_EV, '')))
    fila_completa("Fecha Original", str(row.get(COL_FECHA, '')))
    fila_completa("Horario", str(row.get(COL_HORARIO, '')))
    fila_completa("Ubicación / Predio", str(row.get(COL_UBI, '')))
    fila_completa("Organizador", str(row.get('Organizador', '')))
    
    monto_val = row.get('Monto Total', 0)
    monto_str = f"${float(monto_val):,.2f}" if monto_val else "$0.00"
    fila_completa("Monto Total", monto_str)

    pdf.ln(15)
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 10, "_________________________________________________", ln=True, align="center")
    pdf.cell(0, 5, "Firma y Aclaración del Administrador", ln=True, align="center")
    
    fecha_emision = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    pdf.ln(5)
    pdf.cell(0, 5, f"Reporte generado el: {fecha_emision}", ln=True, align="center")

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name

# --- GENERADORES DE PDF Y HTML ---
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

def descargar_imagen_drive(url_or_id):
    if not url_or_id: return None
    try:
        file_id = str(url_or_id).strip()
        
        if file_id.startswith("data:image"):
            try:
                header, encoded = file_id.split(",", 1)
                img = Image.open(BytesIO(base64.b64decode(encoded)))
                out = BytesIO()
                img.convert("RGB").save(out, format="JPEG", quality=50)
                return BytesIO(out.getvalue())
            except Exception as e:
                return None
                
        if file_id.startswith("http"):
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', file_id)
            if match: file_id = match.group(1)
            else:
                match = re.search(r'id=([a-zA-Z0-9-_]+)', file_id)
                if match: file_id = match.group(1)
                else: return None

        download_url = f"https://drive.google.com/thumbnail?id={file_id}&sz=w1000"
        response = requests.get(download_url, stream=True, timeout=10)
        
        if response.status_code == 200:
            try:
                img = Image.open(BytesIO(response.content))
                img.verify() 
                return BytesIO(response.content)
            except Exception as e:
                return None
        else:
            return None
    except Exception as e:
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
    fila_tabla("Tipo de Evento", str(row.get(COL_TIPO_EV, '')), "Horario", str(row.get(COL_HORARIO, '')), 35, 70, 25, 60)
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

def generar_reporte_html(df_asistentes, id_servicio, tipo_evento, fecha_evento, horario_evento, ubicacion_evento):
    html = "<html><head><meta charset='utf-8'>"
    html += "<style>"
    html += "body { font-family: Arial, sans-serif; padding: 20px; background: #F5F5F5; }"
    html += ".card { background: white; padding: 20px; border-radius: 10px; max-width: 800px; margin: auto; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }"
    html += ".header { text-align: center; border-bottom: 2px solid #002060; padding-bottom: 10px; margin-bottom: 20px; }"
    html += ".header h1 { color: #002060; font-size: 18px; margin: 0; }"
    html += ".header h2 { color: #8B1E1E; font-size: 14px; margin: 5px 0; }"
    html += ".info { background: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #002060; font-size: 14px; }"
    
    html += ".seccion-box { border: 2px solid #ddd; border-radius: 8px; padding: 15px; margin-bottom: 20px; }"
    html += ".seccion-triple { border-color: #5e35b1; }"
    html += ".seccion-doble { border-color: #1565c0; }"
    html += ".seccion-simple { border-color: #2e7d32; }"
    html += ".seccion-titulo { color: white; padding: 8px; border-radius: 5px; font-weight: bold; margin-bottom: 10px; font-size: 14px; }"
    html += ".bg-triple { background: #5e35b1; }"
    html += ".bg-doble { background: #1565c0; }"
    html += ".bg-simple { background: #2e7d32; }"

    html += "table { width: 100%; border-collapse: collapse; }"
    html += "th { background: #002060; color: white; padding: 10px; text-align: left; font-size: 13px; }"
    html += "td { padding: 10px; border-bottom: 1px solid #ddd; font-size: 13px; vertical-align: middle; }"
    html += ".dni { font-family: monospace; font-weight: bold; font-size: 15px; background: #e9ecef; padding: 4px 8px; border-radius: 4px; user-select: all; color: #8B1E1E; }"
    html += ".firma { max-width: 150px; max-height: 70px; }"
    html += "@media print { body { background: white; padding: 0; } .card { box-shadow: none; } button { display: none
