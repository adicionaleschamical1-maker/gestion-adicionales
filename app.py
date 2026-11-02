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

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestión de Servicios Adicionales", page_icon="🛡️", layout="wide")

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

# URL de Apps Script: primero secrets, fallback a constante
try:
    URL_APPS_SCRIPT = st.secrets["URL_APPS_SCRIPT"]
except (KeyError, FileNotFoundError):
    URL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxqVp5T0fVGtPncuwkocbXbU5TSHLgPqgmD0A5pQ6CmuEU3C-UrGx2eZHxg2HTLUjU1/exec"


# --- CONEXIÓN GOOGLE ---
@st.cache_resource
def conectar_google():
    """Autoriza y devuelve (cliente gspread, servicio Drive)."""
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
            raise FileNotFoundError(
                "No se encontró credenciales.json ni st.secrets['gcp_service_account']"
            )
        credentials = Credentials.from_service_account_file(ruta, scopes=scopes)

    client = gspread.authorize(credentials)
    drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    return client, drive_service


def abrir_hoja(nombre: str):
    """Abre una hoja del spreadsheet principal."""
    client, _ = conectar_google()
    ss = client.open(SHEET_NAME)
    try:
        return ss.worksheet(nombre)
    except gspread.WorksheetNotFound:
        return ss.sheet1


@st.cache_data(ttl=60, show_spinner=False)
def cargar_servicios_df() -> pd.DataFrame:
    """Carga hoja Servicios como DataFrame (cache 60 s)."""
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
    """Valida CUIL/CUIT argentino (11 dígitos)."""
    return bool(re.fullmatch(r"\d{11}", str(valor).strip()))


def calcular_monto_y_modulos(hora_inicio: datetime.time, hora_fin: datetime.time):
    """Retorna (diferencia_horas, max_modulos)."""
    dt_ini = datetime.datetime.combine(datetime.date.today(), hora_inicio)
    dt_fin = datetime.datetime.combine(datetime.date.today(), hora_fin)
    if dt_fin <= dt_ini:
        dt_fin += datetime.timedelta(days=1)
    horas = (dt_fin - dt_ini).total_seconds() / 3600.0
    max_mod = max(1, min(4, math.ceil(horas / HORAS_POR_MODULO)))
    return horas, max_mod


def calcular_monto_total(jefes_por_mod: list, subs_por_mod: list) -> float:
    """Suma ponderada por módulo."""
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
    if st.button("Guardar Nueva Contraseña", type="primary"):
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
    st.title("🛡️ Acceso a Servicios Adicionales")
    dni = st.text_input("DNI / Usuario")
    clave = st.text_input("Contraseña", type="password")
    if not st.button("Ingresar", type="primary"):
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
    """Descarga imagen de Drive vía API oficial. Retorna ruta temporal JPG o None."""
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
    """Genera PDF con datos + constancias. Retorna ruta."""
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
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Colaborador")
    boton_refrescar()
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    tab1, tab2, tab3 = st.tabs(
        ["🟢 Cargar Nuevo Servicio", "🔵 Pendientes de Asistencia", "📜 Historial de Servicios"]
    )

    counter = st.session_state.form_counter

    # ---------- TAB 1 ----------
    with tab1:
        st.header(f"Solicitud de Servicio Adicional - {st.session_state.dependencia}")

        fecha = st.date_input("Fecha del Evento", min_value=datetime.date.today(), key=f"fecha_{counter}")
        st.subheader("📋 Datos del Evento")
        tipo_evento = st.text_input(
            "Tipo de Evento (Ej: Partido de fútbol, Baile, Recital, etc.)", key=f"tipo_{counter}"
        )
        
        # Campos separados para el nombre y el DNI del organizador
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

    # ---------- TAB 3 ----------
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
    st.sidebar.write("👑 **Administrador General**")
    st.sidebar.write(f"🏢 **Dependencia:** {st.session_state.dependencia}")
    st.sidebar.write("👤 **Rol:** Administrador")
    boton_refrescar()
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

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

        if st.button("🚀 Generar PDF con Constancias", type="primary"):
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
                type="primary",
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
