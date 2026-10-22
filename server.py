import os
import io
import re
import sys
import json
import time
import shutil
import socket
import smtplib
import tempfile
import subprocess
import threading
import unicodedata
from datetime import datetime

# Asegurar codificación UTF-8 en consola de Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

from flask import Flask, render_template, request, jsonify, send_from_directory
import pandas as pd

try:
    import win32file
    import pythoncom
    import win32com.client as win32
except ImportError:
    win32 = None
    win32file = None
    pythoncom = None

app = Flask(__name__, template_folder="templates", static_folder="static")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
os.makedirs(ASSETS_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(BASE_DIR, "clientes.xlsx")
CSV_PATH = os.path.join(BASE_DIR, "clientes.csv")
ASESORES_PATH = os.path.join(BASE_DIR, "asesores.json")
PLANTILLA_PATH = os.path.join(BASE_DIR, "plantilla_correo.html")
HISTORIAL_PATH = os.path.join(BASE_DIR, "historial_envios.csv")
CONFIG_CORREO_PATH = os.path.join(BASE_DIR, "config_correo.json")

# Constantes MAPI para incrustación de imagen en Outlook
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# Estado global de envío
ESTADO_ENVIO = {
    "activo": False,
    "pausado": False,
    "cancelar": False,
    "total": 0,
    "actual": 0,
    "enviados": 0,
    "errores": 0,
    "modo": "CORREO_TEST",
    "mensaje_actual": "",
    "logs": []
}

DF_CACHE = None
DF_FUENTE = ""
DF_MAPEO = {}

# ----------------- UTILIDADES DE LIMPIEZA Y VALIDACIÓN -----------------

def limpiar_texto_simple(texto) -> str:
    """Elimina acentos y caracteres especiales dejando solo alfanuméricos en minúsculas."""
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    nfkd = unicodedata.normalize('NFKD', texto)
    limpio = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^a-z0-9]', '', limpio)

def normalizar_texto(texto: str) -> str:
    """Normaliza texto preservando espacios para comparaciones geográficas."""
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    nfkd = unicodedata.normalize('NFKD', texto)
    limpio = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'\s+', ' ', limpio)

def limpiar_valor(val, default="") -> str:
    """Sanitiza valores de Excel evitando que aparezca 'nan', 'none', etc."""
    if pd.isna(val) or val is None:
        return default
    s = str(val).strip()
    if s.lower() in ["nan", "none", "null", "undefined", "n/a"]:
        return default
    return s

def validar_correo(correo: str) -> bool:
    """Valida formato básico de correo electrónico."""
    if not correo or not isinstance(correo, str):
        return False
    correo = correo.strip().lower()
    patron = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(patron, correo))

def agregar_log(mensaje: str, tipo: str = "info"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    entry = {"hora": timestamp, "texto": mensaje, "tipo": tipo}
    ESTADO_ENVIO["logs"].append(entry)
    if len(ESTADO_ENVIO["logs"]) > 300:
        ESTADO_ENVIO["logs"] = ESTADO_ENVIO["logs"][-300:]

# ----------------- REGISTRO EN HISTORIAL CSV -----------------

def cargar_historial_enviados() -> set:
    """Carga los correos que ya han sido enviados previamente."""
    if not os.path.exists(HISTORIAL_PATH):
        return set()
    try:
        df_log = pd.read_csv(HISTORIAL_PATH, encoding='utf-8')
        if 'Correo' in df_log.columns and 'Estado' in df_log.columns:
            enviados = set(df_log[df_log['Estado'] == 'ENVIADO']['Correo'].dropna().astype(str).str.strip().str.lower())
            return enviados
    except Exception:
        pass
    return set()

def registrar_en_historial(correo: str, nombre: str, asesor: str, estado: str, detalle: str = "", modo: str = ""):
    """Registra cada envío en historial_envios.csv para trazabilidad."""
    archivo_existe = os.path.exists(HISTORIAL_PATH)
    registro = {
        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Correo": correo,
        "Nombre": nombre,
        "Asesor_Asignado": asesor,
        "Modo": modo,
        "Estado": estado,
        "Detalle": detalle
    }
    try:
        df_nuevo = pd.DataFrame([registro])
        df_nuevo.to_csv(
            HISTORIAL_PATH,
            mode='a',
            index=False,
            header=not archivo_existe,
            encoding='utf-8'
        )
    except Exception as e:
        print(f"[AVISO] No se pudo escribir en historial: {e}")

# ----------------- LECTURA SEGURA DE BASE DE DATOS -----------------

def leer_excel_seguro(ruta):
    """Lee archivo Excel incluso si está abierto en otra ventana de Microsoft Excel."""
    if win32file is not None:
        try:
            handle = win32file.CreateFile(
                ruta,
                win32file.GENERIC_READ,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                None,
                win32file.OPEN_EXISTING,
                win32file.FILE_ATTRIBUTE_NORMAL,
                None
            )
            buffer = bytearray()
            while True:
                hr, data = win32file.ReadFile(handle, 128 * 1024)
                if not data:
                    break
                buffer.extend(data)
            handle.Close()
            return pd.read_excel(io.BytesIO(buffer))
        except Exception:
            pass
    try:
        return pd.read_excel(ruta)
    except PermissionError:
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, f"temp_excel_{int(time.time())}.xlsx")
        shutil.copy2(ruta, temp_path)
        try:
            df = pd.read_excel(temp_path)
            return df
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

def obtener_dataframe_clientes(forzar=False):
    """Detecta y carga clientes.xlsx o clientes.csv mapeando columnas inteligentemente."""
    global DF_CACHE, DF_FUENTE, DF_MAPEO
    if not forzar and DF_CACHE is not None:
        return DF_CACHE, DF_FUENTE, DF_MAPEO

    df = None
    fuente = ""
    if os.path.exists(EXCEL_PATH):
        try:
            df = leer_excel_seguro(EXCEL_PATH)
            fuente = "clientes.xlsx"
        except Exception as e:
            agregar_log(f"Error leyendo Excel: {e}", "error")

    if df is None and os.path.exists(CSV_PATH):
        try:
            df = pd.read_csv(CSV_PATH, encoding='utf-8')
            fuente = "clientes.csv"
        except Exception:
            try:
                df = pd.read_csv(CSV_PATH, encoding='latin-1')
                fuente = "clientes.csv (latin-1)"
            except Exception as e:
                agregar_log(f"Error leyendo CSV: {e}", "error")

    if df is None:
        DF_CACHE = None
        DF_FUENTE = ""
        DF_MAPEO = {}
        return None, "", {}

    df.columns = [str(col).strip() for col in df.columns]

    mapeo = {"nombre": None, "correo": None, "ubicacion": None, "empresa": None, "celular": None, "nit": None, "depto": None}
    
    for col in df.columns:
        norm = limpiar_texto_simple(col)
        if not mapeo["correo"] and any(k in norm for k in ["email", "correo"]):
            mapeo["correo"] = col
        elif not mapeo["nombre"] and any(k in norm for k in ["razon", "social", "nombre", "cliente"]):
            mapeo["nombre"] = col
        elif not mapeo["ubicacion"] and any(k in norm for k in ["ciudad", "municipio", "ubicacion"]):
            mapeo["ubicacion"] = col
        elif not mapeo["depto"] and any(k in norm for k in ["depto", "departamento", "estado"]):
            mapeo["depto"] = col
        elif not mapeo["celular"] and any(k in norm for k in ["celular", "telefono", "movil"]):
            mapeo["celular"] = col
        elif not mapeo["nit"] and any(k in norm for k in ["identificaci", "nit", "cedula", "documento"]):
            mapeo["nit"] = col
        elif not mapeo["empresa"] and any(k in norm for k in ["empresa", "compania"]):
            mapeo["empresa"] = col

    # Si no hay ubicación primaria, usar departamento
    if not mapeo["ubicacion"] and mapeo["depto"]:
        mapeo["ubicacion"] = mapeo["depto"]

    DF_CACHE = df
    DF_FUENTE = fuente
    DF_MAPEO = mapeo
    return DF_CACHE, DF_FUENTE, DF_MAPEO

# ----------------- CONFIGURACIÓN Y ASESORES -----------------

def cargar_config_correo():
    default_config = {
        "metodo": "OUTLOOK",
        "smtp_host": "smtp.office365.com",
        "smtp_port": 587,
        "smtp_usuario": "ximena_garciamo@soy.sena.edu.co",
        "smtp_password": "",
        "remitente_nombre": "Dirección Comercial Corporativa",
        "asunto": "Información exclusiva para {nombre} - Actualización Corporativa",
        "saludo": "Estimado(a) {nombre},",
        "cuerpo": "Nos complace saludarte y ponernos en contacto contigo para compartirte información relevante sobre nuestros servicios y las novedades diseñadas especialmente para atender tus requerimientos en <strong>{ubicacion}</strong>.",
        "cuerpo_destacado": "Adjunto a este correo encontrarás el documento detallado en formato PDF con el catálogo completo y las especificaciones que hemos preparado para ti.",
        "cuerpo_cierre": "Con el propósito de brindarte un acompañamiento oportuno y personalizado en tu zona, hemos asignado a un asesor comercial exclusivo para apoyarte:",
        "correo_test": "tu_correo@empresa.com"
    }
    if not os.path.exists(CONFIG_CORREO_PATH):
        with open(CONFIG_CORREO_PATH, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        return default_config

    try:
        with open(CONFIG_CORREO_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            for k, v in default_config.items():
                if k not in cfg:
                    cfg[k] = v
            return cfg
    except Exception:
        return default_config

def guardar_config_correo(cfg):
    with open(CONFIG_CORREO_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

def cargar_asesores():
    if not os.path.exists(ASESORES_PATH):
        return {}
    with open(ASESORES_PATH, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except Exception:
            return {}

def obtener_asesor_para_ubicacion(ubicacion: str, asesores: dict) -> dict:
    ubicacion_norm = normalizar_texto(ubicacion)
    if ubicacion_norm:
        for clave, datos in asesores.items():
            clave_norm = normalizar_texto(clave)
            if clave != "default" and clave_norm and (clave_norm == ubicacion_norm or clave_norm in ubicacion_norm or ubicacion_norm in clave_norm):
                return datos

    return asesores.get("default", {
        "nombre": "Equipo de Asesoría Corporativa",
        "cargo": "Atención y Servicio al Cliente Nacional",
        "correo": "contacto@empresa.com",
        "telefono": "+57 (601) 800 0000",
        "ciudad_display": "Nacional"
    })

# ----------------- RENDERIZADOR ROBUSTO DE PLANTILLA -----------------

def renderizar_correo(datos_cliente: dict, asesor: dict, config_textos: dict, imagen_nombre: str = None) -> tuple:
    nombre = limpiar_valor(datos_cliente.get("nombre"), "Estimado Cliente")
    correo = limpiar_valor(datos_cliente.get("correo"), "")
    ubicacion = limpiar_valor(datos_cliente.get("ubicacion"), "su ciudad")
    empresa = limpiar_valor(datos_cliente.get("empresa"), nombre)
    celular = limpiar_valor(datos_cliente.get("celular"), "")

    asunto_raw = config_textos.get("asunto") or "Información exclusiva para {nombre} - Actualización Corporativa"
    saludo_raw = config_textos.get("saludo") or "Estimado(a) {nombre},"
    cuerpo_raw = config_textos.get("cuerpo") or "Nos complace saludarte y ponernos en contacto contigo para compartirte información relevante sobre nuestros servicios en <strong>{ubicacion}</strong>."
    cuerpo_destacado_raw = config_textos.get("cuerpo_destacado") or "Adjunto a este correo encontrarás el documento detallado en formato PDF."
    cuerpo_cierre_raw = config_textos.get("cuerpo_cierre") or "Con el propósito de brindarte un acompañamiento oportuno, hemos asignado a un asesor exclusivo:"

    # Plantilla HTML Base
    plantilla = ""
    if os.path.exists(PLANTILLA_PATH):
        with open(PLANTILLA_PATH, "r", encoding="utf-8") as f:
            plantilla = f.read()

    # Reemplazo de banner
    if imagen_nombre:
        tag_banner = f"""
        <tr>
          <td align="center" style="background-color: #0f2d59;">
            <img src="/assets/{imagen_nombre}" width="620" style="width: 100%; max-width: 620px; height: auto; display: block;" alt="Banner Corporativo">
          </td>
        </tr>
        """
    else:
        tag_banner = ""

    html = plantilla.replace("{TAG_IMAGEN_BANNER}", tag_banner)

    # Reemplazo de bloques de texto principales
    html = html.replace("{saludo}", saludo_raw)
    html = html.replace("{cuerpo}", cuerpo_raw)
    html = html.replace("{cuerpo_destacado}", cuerpo_destacado_raw)
    html = html.replace("{cuerpo_cierre}", cuerpo_cierre_raw)

    # Reemplazo de comodines tanto en HTML como en Asunto
    def aplicar_comodines(texto):
        if not texto: return ""
        return (texto.replace("{nombre}", nombre)
                     .replace("{correo}", correo)
                     .replace("{ubicacion}", ubicacion)
                     .replace("{empresa}", empresa)
                     .replace("{celular}", celular)
                     .replace("{asesor_nombre}", asesor.get("nombre", "Equipo Comercial"))
                     .replace("{asesor_cargo}", asesor.get("cargo", "Asesor Comercial"))
                     .replace("{asesor_correo}", asesor.get("correo", ""))
                     .replace("{asesor_telefono}", asesor.get("telefono", ""))
                     .replace("{asesor_zona}", asesor.get("ciudad_display", ubicacion)))

    asunto = aplicar_comodines(asunto_raw)
    html = aplicar_comodines(html)

    return asunto, html

# ----------------- RUTAS DE LA API WEB -----------------

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/assets/<path:filename>")
def serve_asset(filename):
    return send_from_directory(ASSETS_DIR, filename)

@app.route("/api/estado-general")
def estado_general():
    df, fuente, mapeo = obtener_dataframe_clientes()
    total_clientes = len(df) if df is not None else 0
    total_con_correo = 0

    ciudades = []
    if df is not None:
        col_ubi = mapeo.get("ubicacion")
        if col_ubi and col_ubi in df.columns:
            ciudades = sorted([
                c for c in df[col_ubi].dropna().astype(str).str.strip().unique()
                if c.lower() not in ["nan", "none", "null", ""]
            ])
        col_cor = mapeo.get("correo")
        if col_cor and col_cor in df.columns:
            total_con_correo = int(df[col_cor].dropna().astype(str).str.strip().apply(validar_correo).sum())

    banner_existe = os.path.exists(os.path.join(ASSETS_DIR, "banner_correo.png"))
    pdf_existe = os.path.exists(os.path.join(ASSETS_DIR, "informacion_corporativa.pdf"))
    cfg_correo = cargar_config_correo()

    # Comprobación de Outlook activo
    outlook_info = verificar_outlook_instalado()

    return jsonify({
        "archivo_conectado": fuente,
        "total_registros": total_clientes,
        "total_con_correo": total_con_correo,
        "columnas_detectadas": list(df.columns) if df is not None else [],
        "mapeo_columnas": mapeo,
        "ciudades_disponibles": ciudades,
        "banner_activo": "banner_correo.png" if banner_existe else None,
        "pdf_activo": "informacion_corporativa.pdf" if pdf_existe else None,
        "outlook_status": outlook_info,
        "asesores": cargar_asesores(),
        "historial_enviados_count": len(cargar_historial_enviados()),
        "config_correo": {
            "metodo": cfg_correo.get("metodo", "OUTLOOK"),
            "smtp_host": cfg_correo.get("smtp_host", "smtp.office365.com"),
            "smtp_port": cfg_correo.get("smtp_port", 587),
            "smtp_usuario": cfg_correo.get("smtp_usuario", ""),
            "tiene_password": bool(cfg_correo.get("smtp_password")),
            "remitente_nombre": cfg_correo.get("remitente_nombre", ""),
            "asunto": cfg_correo.get("asunto", ""),
            "saludo": cfg_correo.get("saludo", ""),
            "cuerpo": cfg_correo.get("cuerpo", ""),
            "cuerpo_destacado": cfg_correo.get("cuerpo_destacado", ""),
            "cuerpo_cierre": cfg_correo.get("cuerpo_cierre", ""),
            "correo_test": cfg_correo.get("correo_test", "tu_correo@empresa.com")
        }
    })

def verificar_outlook_instalado():
    if win32 is None:
        return {"ok": False, "activo": False, "mensaje": "pywin32 no disponible"}
    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        ol = win32.GetActiveObject("Outlook.Application")
        return {"ok": True, "activo": True, "mensaje": "Outlook Desktop activo y conectado en tu equipo"}
    except Exception:
        try:
            ol = win32.Dispatch("Outlook.Application")
            return {"ok": True, "activo": True, "mensaje": "Outlook Desktop conectado con éxito"}
        except Exception:
            return {"ok": True, "activo": False, "mensaje": "Outlook instalado (Ábrelo en Windows para enviar sin contraseñas)"}

@app.route("/api/recargar-datos", methods=["POST"])
def recargar_datos():
    df, fuente, mapeo = obtener_dataframe_clientes(forzar=True)
    total = len(df) if df is not None else 0
    total_con_correo = 0
    if df is not None and mapeo.get("correo"):
        col_c = mapeo["correo"]
        total_con_correo = int(df[col_c].dropna().astype(str).str.strip().apply(validar_correo).sum())
    return jsonify({"ok": True, "total": total, "total_con_correo": total_con_correo, "fuente": fuente})

@app.route("/api/clientes")
def listar_clientes():
    df, fuente, mapeo = obtener_dataframe_clientes()
    if df is None:
        return jsonify({"clientes": [], "total": 0, "total_con_correo": 0})

    ciudad_filtro = request.args.get("ciudad", "").strip()
    busqueda = request.args.get("q", "").strip().lower()
    pagina = int(request.args.get("pagina", 1))
    limite = int(request.args.get("limite", 50))
    solo_correo = request.args.get("solo_correo", "1") in ["1", "true", "True"]

    col_nombre = mapeo.get("nombre") or df.columns[0]
    col_correo = mapeo.get("correo") or (df.columns[1] if len(df.columns) > 1 else df.columns[0])
    col_ubicacion = mapeo.get("ubicacion")
    col_depto = mapeo.get("depto")

    enviados_set = cargar_historial_enviados()
    asesores = cargar_asesores()

    filtrado = df

    # Filtro opcional: solo registros con correo válido
    if solo_correo and col_correo in filtrado.columns:
        mascara_correo = filtrado[col_correo].fillna("").astype(str).str.strip().apply(validar_correo)
        filtrado = filtrado[mascara_correo]

    if ciudad_filtro and col_ubicacion in filtrado.columns:
        mascara_ciudad = filtrado[col_ubicacion].fillna("").astype(str).str.strip().str.lower() == ciudad_filtro.lower()
        filtrado = filtrado[mascara_ciudad]

    if busqueda:
        mascara = filtrado[col_nombre].fillna("").astype(str).str.lower().str.contains(busqueda, na=False)
        if col_correo in filtrado.columns:
            mascara = mascara | filtrado[col_correo].fillna("").astype(str).str.lower().str.contains(busqueda, na=False)
        if col_ubicacion in filtrado.columns:
            mascara = mascara | filtrado[col_ubicacion].fillna("").astype(str).str.lower().str.contains(busqueda, na=False)
        filtrado = filtrado[mascara]

    total = len(filtrado)
    inicio = (pagina - 1) * limite
    fin = inicio + limite
    sub_df = filtrado.iloc[inicio:fin]

    lista = []
    for idx, row in sub_df.iterrows():
        ubic = limpiar_valor(row[col_ubicacion]) if col_ubicacion and col_ubicacion in sub_df.columns else ""
        if not ubic and col_depto and col_depto in sub_df.columns:
            ubic = limpiar_valor(row[col_depto])
        if not ubic:
            ubic = "Nacional"

        correo_val = limpiar_valor(row[col_correo]) if col_correo in sub_df.columns else ""
        es_valido = validar_correo(correo_val)
        ya_enviado = correo_val.lower() in enviados_set

        asesor = obtener_asesor_para_ubicacion(ubic, asesores)

        lista.append({
            "indice_original": int(idx),
            "nombre": limpiar_valor(row[col_nombre], "Cliente"),
            "correo": correo_val,
            "correo_valido": es_valido,
            "ya_enviado": ya_enviado,
            "ubicacion": ubic,
            "asesor": asesor.get("nombre", "General"),
            "asesor_zona": asesor.get("ciudad_display", ubic)
        })

    return jsonify({"clientes": lista, "total": total, "pagina": pagina, "limite": limite})

@app.route("/api/seleccionar-filtrados", methods=["POST"])
def seleccionar_filtrados():
    df, fuente, mapeo = obtener_dataframe_clientes()
    if df is None:
        return jsonify({"indices": [], "total": 0})

    payload = request.json or {}
    ciudad_filtro = payload.get("ciudad", "").strip()
    busqueda = payload.get("q", "").strip().lower()
    solo_correo = payload.get("solo_correo", True)

    col_nombre = mapeo.get("nombre") or df.columns[0]
    col_correo = mapeo.get("correo") or (df.columns[1] if len(df.columns) > 1 else df.columns[0])
    col_ubicacion = mapeo.get("ubicacion")

    filtrado = df

    if solo_correo and col_correo in filtrado.columns:
        mascara_correo = filtrado[col_correo].fillna("").astype(str).str.strip().apply(validar_correo)
        filtrado = filtrado[mascara_correo]

    if ciudad_filtro and col_ubicacion in filtrado.columns:
        mascara_ciudad = filtrado[col_ubicacion].fillna("").astype(str).str.strip().str.lower() == ciudad_filtro.lower()
        filtrado = filtrado[mascara_ciudad]

    if busqueda:
        mascara = filtrado[col_nombre].fillna("").astype(str).str.lower().str.contains(busqueda, na=False)
        if col_correo in filtrado.columns:
            mascara = mascara | filtrado[col_correo].fillna("").astype(str).str.lower().str.contains(busqueda, na=False)
        if col_ubicacion in filtrado.columns:
            mascara = mascara | filtrado[col_ubicacion].fillna("").astype(str).str.lower().str.contains(busqueda, na=False)
        filtrado = filtrado[mascara]

    return jsonify({"indices": [int(i) for i in filtrado.index], "total": len(filtrado)})

@app.route("/api/guardar-textos-correo", methods=["POST"])
def guardar_textos_correo():
    nuevos = request.json or {}
    cfg = cargar_config_correo()
    for campo in ["asunto", "saludo", "cuerpo", "cuerpo_destacado", "cuerpo_cierre"]:
        if campo in nuevos:
            cfg[campo] = nuevos[campo]
    guardar_config_correo(cfg)
    return jsonify({"ok": True, "mensaje": "Textos guardados exitosamente"})

@app.route("/api/guardar-config-smtp", methods=["POST"])
def guardar_config_smtp():
    nuevos = request.json or {}
    cfg = cargar_config_correo()
    if "metodo" in nuevos: cfg["metodo"] = nuevos["metodo"]
    if "smtp_usuario" in nuevos: cfg["smtp_usuario"] = nuevos["smtp_usuario"].strip()
    if "smtp_host" in nuevos:
        host_val = nuevos["smtp_host"].strip()
        # Auto-ajuste de host si se introduce un correo Gmail
        if cfg.get("smtp_usuario", "").lower().endswith("@gmail.com") and (host_val == "smtp.office365.com" or not host_val):
            host_val = "smtp.gmail.com"
        cfg["smtp_host"] = host_val
    if "smtp_port" in nuevos: cfg["smtp_port"] = int(nuevos["smtp_port"])
    if "remitente_nombre" in nuevos: cfg["remitente_nombre"] = nuevos["remitente_nombre"].strip()
    if "correo_test" in nuevos: cfg["correo_test"] = nuevos["correo_test"].strip()
    if nuevos.get("smtp_password"):
        cfg["smtp_password"] = nuevos["smtp_password"]

    guardar_config_correo(cfg)
    return jsonify({"ok": True, "mensaje": "Configuración del emisor guardada correctamente"})

@app.route("/api/probar-smtp", methods=["POST"])
def probar_smtp():
    payload = request.json or {}
    cfg = cargar_config_correo()
    host = payload.get("smtp_host") or cfg.get("smtp_host", "smtp.office365.com")
    port = int(payload.get("smtp_port") or cfg.get("smtp_port", 587))
    user = payload.get("smtp_usuario") or cfg.get("smtp_usuario", "")
    password = payload.get("smtp_password") or cfg.get("smtp_password", "")

    if not user:
        return jsonify({"ok": False, "error": "Por favor escribe tu correo corporativo emisor"}), 400
    if not password:
        return jsonify({"ok": False, "error": "Ingresa la contraseña del correo para probar la autenticación SMTP"}), 400

    if user.lower().endswith("@gmail.com") and host == "smtp.office365.com":
        return jsonify({
            "ok": False,
            "error": "Conflicto detectado: Tu correo es de Gmail (@gmail.com) pero el servidor configurado es Microsoft Office 365 (smtp.office365.com). Cambia el servidor a 'smtp.gmail.com' (puerto 587) y usa una Contraseña de Aplicación de Google de 16 caracteres."
        }), 400

    try:
        if port == 465:
            server = smtplib.SMTP_SSL(host, port, timeout=12)
        else:
            server = smtplib.SMTP(host, port, timeout=12)
            server.starttls()

        server.login(user, password)
        server.quit()
        return jsonify({"ok": True, "mensaje": f"¡Conexión y autenticación exitosa con {host}:{port} para {user}!"})
    except smtplib.SMTPAuthenticationError:
        return jsonify({"ok": False, "error": "Fallo de autenticación: Usuario o contraseña incorrectos (si tienes doble factor MFA en Office 365 / SENA, genera una 'Contraseña de aplicación')."}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error conectando con el servidor SMTP: {str(e)}"}), 400

def obtener_outlook_app():
    if win32 is None:
        raise RuntimeError("pywin32 no está instalado en este sistema.")
    if pythoncom is not None:
        pythoncom.CoInitialize()
    try:
        return win32.GetActiveObject("Outlook.Application")
    except Exception:
        try:
            return win32.Dispatch("Outlook.Application")
        except Exception:
            raise RuntimeError("Microsoft Outlook no está abierto ni disponible en tu computador. Por favor abre la aplicación Outlook en Windows antes de continuar.")

@app.route("/api/probar-outlook", methods=["POST"])
def probar_outlook():
    try:
        ol = obtener_outlook_app()
        ns = ol.GetNamespace("MAPI")
        cuenta = ns.Accounts.Item(1).DisplayName if ns.Accounts.Count > 0 else "Cuenta predeterminada"
        return jsonify({
            "ok": True,
            "mensaje": f"¡Microsoft Outlook conectado con éxito! Cuenta activa: {cuenta}",
            "cuenta": cuenta
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 400

@app.route("/api/abrir-outlook", methods=["POST"])
def abrir_outlook():
    exito = False

    # 1. Intentar abrir mediante Shell protocol (Nuevo Outlook o registrado)
    try:
        subprocess.Popen(["cmd", "/c", "start", "ms-outlook:"], shell=True)
        exito = True
    except Exception:
        pass

    # 2. Intentar abrir ejecutable clásico
    rutas_outlook = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
    ]
    for r in rutas_outlook:
        if os.path.exists(r):
            try:
                subprocess.Popen([r, "/recycle"])
                exito = True
                break
            except Exception:
                pass

    return jsonify({
        "ok": True,
        "mensaje": "Se envió la orden de apertura a Windows. Si no se visualiza al frente, haz clic en el icono de Outlook en tu barra de tareas o menú Inicio."
    })

@app.route("/api/preview", methods=["POST"])
def preview_email():
    payload = request.json or {}
    indice = payload.get("indice", 0)
    imagen_nombre = payload.get("imagen", "banner_correo.png")

    cfg_textos = {
        "asunto": payload.get("asunto"),
        "saludo": payload.get("saludo"),
        "cuerpo": payload.get("cuerpo"),
        "cuerpo_destacado": payload.get("cuerpo_destacado"),
        "cuerpo_cierre": payload.get("cuerpo_cierre")
    }
    cfg_base = cargar_config_correo()
    for k, v in cfg_base.items():
        if not cfg_textos.get(k):
            cfg_textos[k] = v

    df, fuente, mapeo = obtener_dataframe_clientes()
    datos_cliente = {
        "nombre": "Empresa Cliente Ejemplo S.A.S",
        "correo": "cliente.ejemplo@empresa.com",
        "ubicacion": "Bogotá D.C.",
        "empresa": "Empresa Cliente Ejemplo S.A.S",
        "celular": "+57 300 000 0000"
    }

    if df is not None and 0 <= indice < len(df):
        row = df.iloc[indice]
        col_nom = mapeo.get("nombre")
        col_cor = mapeo.get("correo")
        col_ubi = mapeo.get("ubicacion")
        col_cel = mapeo.get("celular")
        col_emp = mapeo.get("empresa")

        if col_nom: datos_cliente["nombre"] = limpiar_valor(row[col_nom], "Cliente")
        if col_cor: datos_cliente["correo"] = limpiar_valor(row[col_cor], "")
        if col_ubi: datos_cliente["ubicacion"] = limpiar_valor(row[col_ubi], "su ciudad")
        if col_cel: datos_cliente["celular"] = limpiar_valor(row[col_cel], "")
        datos_cliente["empresa"] = limpiar_valor(row[col_emp], datos_cliente["nombre"]) if col_emp else datos_cliente["nombre"]

    asesores = cargar_asesores()
    asesor = obtener_asesor_para_ubicacion(datos_cliente["ubicacion"], asesores)

    asunto, html = renderizar_correo(datos_cliente, asesor, cfg_textos, imagen_nombre)
    return jsonify({"asunto": asunto, "html": html, "asesor": asesor, "cliente": datos_cliente})

@app.route("/api/subir-banner", methods=["POST"])
def subir_banner():
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400
    archivo = request.files["archivo"]
    if archivo.filename == "":
        return jsonify({"error": "Archivo vacío"}), 400

    destino = os.path.join(ASSETS_DIR, "banner_correo.png")
    archivo.save(destino)
    return jsonify({"ok": True, "mensaje": "Banner actualizado exitosamente", "archivo": "banner_correo.png"})

@app.route("/api/subir-pdf", methods=["POST"])
def subir_pdf():
    if "archivo" not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400
    archivo = request.files["archivo"]
    if archivo.filename == "":
        return jsonify({"error": "Archivo vacío"}), 400

    destino = os.path.join(ASSETS_DIR, "informacion_corporativa.pdf")
    archivo.save(destino)
    return jsonify({"ok": True, "mensaje": "Archivo PDF corporativo actualizado", "archivo": "informacion_corporativa.pdf"})

@app.route("/api/guardar-asesores", methods=["POST"])
def guardar_asesores():
    nuevos = request.json or {}
    with open(ASESORES_PATH, "w", encoding="utf-8") as f:
        json.dump(nuevos, f, indent=2, ensure_ascii=False)
    return jsonify({"ok": True, "mensaje": "Tabla de asesores actualizada correctamente"})

# ----------------- MOTOR PRINCIPAL DE ENVÍO DE CORREOS -----------------

def enviar_correo_smtp(smtp_server, remitente, destinatario, asunto, html_cuerpo, ruta_banner, ruta_pdf):
    """Envía correo vía SMTP con banner incrustado (CID) y PDF adjunto."""
    msg_raiz = MIMEMultipart("related")
    msg_raiz["Subject"] = asunto
    msg_raiz["From"] = remitente
    msg_raiz["To"] = destinatario

    msg_alt = MIMEMultipart("alternative")
    msg_raiz.attach(msg_alt)

    # Reemplazar cualquier ruta de banner por cid:banner_principal (incluso con cache-buster ?t=...)
    html_procesado = re.sub(r'src=["\']/assets/[^"\'\s]+["\']', 'src="cid:banner_principal"', html_cuerpo)
    msg_html = MIMEText(html_procesado, "html", "utf-8")
    msg_alt.attach(msg_html)

    # Incrustar banner vía CID
    if ruta_banner and os.path.exists(ruta_banner):
        with open(ruta_banner, "rb") as f:
            msg_img = MIMEImage(f.read())
            msg_img.add_header("Content-ID", "<banner_principal>")
            msg_img.add_header("Content-Disposition", "inline", filename="banner_correo.png")
            msg_raiz.attach(msg_img)

    # Adjuntar PDF
    if ruta_pdf and os.path.exists(ruta_pdf):
        with open(ruta_pdf, "rb") as f:
            msg_pdf = MIMEApplication(f.read(), _subtype="pdf")
            msg_pdf.add_header("Content-Disposition", "attachment", filename=os.path.basename(ruta_pdf))
            msg_raiz.attach(msg_pdf)

    smtp_server.send_message(msg_raiz)

def worker_envio(indices, modo, config_textos, correo_test):
    """Hilo de ejecución que procesa envíos en segundo plano."""
    global ESTADO_ENVIO
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

    ESTADO_ENVIO["activo"] = True
    ESTADO_ENVIO["cancelar"] = False
    ESTADO_ENVIO["total"] = len(indices)
    ESTADO_ENVIO["actual"] = 0
    ESTADO_ENVIO["enviados"] = 0
    ESTADO_ENVIO["errores"] = 0
    ESTADO_ENVIO["modo"] = modo
    ESTADO_ENVIO["logs"] = []

    cfg_correo = cargar_config_correo()
    metodo = cfg_correo.get("metodo", "OUTLOOK")

    # En modo BORRADORES se requiere forzosamente Outlook
    if modo == "BORRADORES":
        metodo = "OUTLOOK"

    agregar_log(f"Iniciando proceso: modo '{modo}' con motor '{metodo}' para {len(indices)} destinatarios", "info")

    smtp_conn = None
    outlook_app = None

    try:
        # Inicialización del motor seleccionado
        if metodo == "OUTLOOK":
            try:
                outlook_app = obtener_outlook_app()
                ns = outlook_app.GetNamespace("MAPI")
                cuenta = ns.Accounts.Item(1).DisplayName if ns.Accounts.Count > 0 else "Outlook Principal"
                agregar_log(f"Conectado a Microsoft Outlook con la cuenta: {cuenta}", "success")
            except Exception as e:
                agregar_log(f"{str(e)}", "error")
                return

        elif metodo == "SMTP":
            host = cfg_correo.get("smtp_host", "smtp.office365.com")
            port = int(cfg_correo.get("smtp_port", 587))
            user = cfg_correo.get("smtp_usuario", "").strip()
            password = cfg_correo.get("smtp_password", "")

            if not user or not password:
                agregar_log("Falta correo o contraseña SMTP. Configúralos en el botón superior de Configuración.", "error")
                return

            try:
                agregar_log(f"Conectando al servidor SMTP {host}:{port}...", "info")
                if port == 465:
                    smtp_conn = smtplib.SMTP_SSL(host, port, timeout=15)
                else:
                    smtp_conn = smtplib.SMTP(host, port, timeout=15)
                    smtp_conn.starttls()
                smtp_conn.login(user, password)
                agregar_log(f"Autenticación SMTP exitosa con: {user}", "success")
            except Exception as e:
                agregar_log(f"Error de conexión SMTP: {e}", "error")
                return

        # Cargar base de datos
        df, fuente, mapeo = obtener_dataframe_clientes()
        if df is None:
            agregar_log("Error: No se pudo cargar la base de datos de clientes.", "error")
            return

        asesores = cargar_asesores()
        col_nom = mapeo.get("nombre")
        col_cor = mapeo.get("correo")
        col_ubi = mapeo.get("ubicacion")
        col_depto = mapeo.get("depto")

        ruta_banner = os.path.abspath(os.path.join(ASSETS_DIR, "banner_correo.png")) if os.path.exists(os.path.join(ASSETS_DIR, "banner_correo.png")) else None
        ruta_pdf = os.path.abspath(os.path.join(ASSETS_DIR, "informacion_corporativa.pdf")) if os.path.exists(os.path.join(ASSETS_DIR, "informacion_corporativa.pdf")) else None

        rem_nom = cfg_correo.get('remitente_nombre', '').strip()
        smtp_u = cfg_correo.get('smtp_usuario', '').strip()
        remitente_display = f"{rem_nom} <{smtp_u}>" if rem_nom else smtp_u

        for pos, idx in enumerate(indices, 1):
            if ESTADO_ENVIO["cancelar"]:
                agregar_log("Envío detenido a petición del usuario.", "warning")
                break

            ESTADO_ENVIO["actual"] = pos
            row = df.iloc[idx]
            nombre_cli = limpiar_valor(row[col_nom], "Cliente") if col_nom else "Cliente"
            correo_cli = limpiar_valor(row[col_cor], "") if col_cor else ""
            ubic_cli = limpiar_valor(row[col_ubi]) if col_ubi else ""
            if not ubic_cli and col_depto:
                ubic_cli = limpiar_valor(row[col_depto])
            if not ubic_cli:
                ubic_cli = "Nacional"

            # Validación de correo
            if not validar_correo(correo_cli) and modo != "CORREO_TEST":
                agregar_log(f"[{pos}/{len(indices)}] Correo inválido '{correo_cli}' ({nombre_cli}). Omitido.", "warning")
                ESTADO_ENVIO["errores"] += 1
                continue

            asesor = obtener_asesor_para_ubicacion(ubic_cli, asesores)

            # Destinatario según modo
            if modo == "CORREO_TEST":
                destinatario = correo_test
            else:
                destinatario = correo_cli

            asunto, html = renderizar_correo(
                {"nombre": nombre_cli, "correo": correo_cli, "ubicacion": ubic_cli},
                asesor, config_textos, "banner_correo.png" if ruta_banner else None
            )

            if modo == "CORREO_TEST":
                asunto = f"[PRUEBA] {asunto}"
            elif modo == "BORRADORES":
                asunto = f"[BORRADOR] {asunto}"

            try:
                if metodo == "OUTLOOK":
                    mail = outlook_app.CreateItem(0)
                    mail.To = destinatario
                    mail.Subject = asunto

                    html_outlook = re.sub(r'src=["\']/assets/[^"\'\s]+["\']', 'src="cid:banner_principal"', html)

                    if ruta_banner:
                        att_b = mail.Attachments.Add(ruta_banner)
                        att_b.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, "banner_principal")
                        att_b.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)

                    if ruta_pdf:
                        mail.Attachments.Add(ruta_pdf)

                    mail.HTMLBody = html_outlook

                    if modo == "BORRADORES":
                        mail.Save()
                        agregar_log(f"[{pos}/{len(indices)}] [BORRADOR OUTLOOK] Guardado para: {nombre_cli} ({destinatario})", "success")
                        registrar_en_historial(destinatario, nombre_cli, asesor.get('nombre'), 'BORRADOR_GUARDADO', 'Borrador en Outlook', modo)
                    else:
                        mail.Send()
                        agregar_log(f"[{pos}/{len(indices)}] [ENVIADO OUTLOOK] {nombre_cli} ({destinatario}) -> Asesor: {asesor.get('nombre')}", "success")
                        registrar_en_historial(destinatario, nombre_cli, asesor.get('nombre'), 'ENVIADO', 'Outlook Desktop', modo)

                elif metodo == "SMTP":
                    enviar_correo_smtp(smtp_conn, remitente_display, destinatario, asunto, html, ruta_banner, ruta_pdf)
                    agregar_log(f"[{pos}/{len(indices)}] [ENVIADO SMTP] {nombre_cli} ({destinatario}) -> Asesor: {asesor.get('nombre')}", "success")
                    registrar_en_historial(destinatario, nombre_cli, asesor.get('nombre'), 'ENVIADO', 'SMTP Directo', modo)

                ESTADO_ENVIO["enviados"] += 1

            except Exception as e:
                agregar_log(f"[{pos}/{len(indices)}] Error enviando a {destinatario}: {e}", "error")
                registrar_en_historial(destinatario, nombre_cli, asesor.get('nombre'), 'ERROR', str(e), modo)
                ESTADO_ENVIO["errores"] += 1

            # Pausa de seguridad entre envíos
            if modo == "REAL":
                time.sleep(2.5)
                # Pausa por lote cada 50 correos
                if pos % 50 == 0 and pos < len(indices):
                    agregar_log("Pausa de descanso de 30 segundos (Protección anti-spam Microsoft 365)...", "info")
                    time.sleep(30)
            else:
                time.sleep(0.8)

        agregar_log("Proceso finalizado.", "info")

    except Exception as e:
        agregar_log(f"Error general en el proceso: {e}", "error")
    finally:
        if smtp_conn:
            try:
                smtp_conn.quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        ESTADO_ENVIO["activo"] = False

@app.route("/api/iniciar-envio", methods=["POST"])
def iniciar_envio():
    global ESTADO_ENVIO
    if ESTADO_ENVIO["activo"]:
        return jsonify({"error": "Ya hay un envío en curso. Espera a que termine o pulsa Detener."}), 400

    data = request.json or {}
    indices = data.get("indices", [])
    modo = data.get("modo", "CORREO_TEST")
    correo_test = data.get("correo_test", "tu_correo@empresa.com").strip()

    if modo == "CORREO_TEST" and not validar_correo(correo_test):
        return jsonify({"error": "Por favor escribe un correo de prueba válido (ej. tu_correo@empresa.com)"}), 400

    cfg_textos = {
        "asunto": data.get("asunto"),
        "saludo": data.get("saludo"),
        "cuerpo": data.get("cuerpo"),
        "cuerpo_destacado": data.get("cuerpo_destacado"),
        "cuerpo_cierre": data.get("cuerpo_cierre")
    }

    if not indices:
        return jsonify({"error": "No seleccionaste ningún cliente para enviar"}), 400

    hilo = threading.Thread(
        target=worker_envio,
        args=(indices, modo, cfg_textos, correo_test),
        daemon=True
    )
    hilo.start()

    return jsonify({"ok": True, "mensaje": f"Envío iniciado en modo {modo} para {len(indices)} destinatarios"})

@app.route("/api/estado-envio")
def estado_envio():
    return jsonify(ESTADO_ENVIO)

@app.route("/api/cancelar-envio", methods=["POST"])
def cancelar_envio():
    global ESTADO_ENVIO
    ESTADO_ENVIO["cancelar"] = True
    return jsonify({"ok": True, "mensaje": "Cancelación solicitada"})

def obtener_ip_local():
    """Detecta la IP local de la máquina en la red local/LAN para acceso desde otros dispositivos."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def encontrar_puerto_disponible(puerto_inicial=5000, max_intentos=10):
    """Encuentra un puerto libre para evitar conflictos si el puerto 5000 ya está ocupado."""
    for offset in range(max_intentos):
        p = puerto_inicial + offset
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("0.0.0.0", p))
                return p
            except OSError:
                continue
    return puerto_inicial

if __name__ == "__main__":
    env_host = os.environ.get("HOST", "0.0.0.0")
    env_port = os.environ.get("PORT")
    if env_port:
        port = int(env_port)
    else:
        port = encontrar_puerto_disponible(5000)

    ip_lan = obtener_ip_local()

    print("\n" + "=" * 72)
    print(" [SISTEMA] PANEL DE ENVIOS MASIVOS CORPORATIVOS - SERVIDOR ACTIVO")
    print("=" * 72)
    print(f" [LOCAL]   Acceso en esta maquina:         http://localhost:{port}")
    if ip_lan != "127.0.0.1":
        print(f" [RED LAN] Acceso desde la red (WiFi/LAN):  http://{ip_lan}:{port}")
    print(f" [HOST]    Escuchando en:                  http://{env_host}:{port}")
    print("=" * 72 + "\n")

    app.run(host=env_host, port=port, debug=False)
