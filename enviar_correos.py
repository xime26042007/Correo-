"""
================================================================================
SISTEMA DE ENVÍO MASIVO DE CORREOS CORPORATIVOS CON PYTHON & OUTLOOK
================================================================================
Características:
- Integración nativa con Microsoft Outlook Desktop (MAPI).
- Asignación dinámica de asesores comerciales según ubicación geográfica.
- Incrustación de imagen/banner directamente en el cuerpo (usando Content-ID / CID).
- Adjunto de archivo PDF corporativo opcional.
- Control de ritmo (Throttling) para evitar bloqueos de cuenta en Microsoft 365.
- Control de reanudación y prevención de duplicados (historial_envios.csv).
- Modo de prueba seguro (Borradores o correo de test).
"""

import os
import sys
import time
import json
import re
import unicodedata
from datetime import datetime
import pandas as pd

# Importar configuración
import config

try:
    import win32com.client as win32
except ImportError:
    print("[ERROR] Falta la librería 'pywin32'.")
    print("Por favor instálala ejecutando: pip install pywin32")
    sys.exit(1)

# Constantes MAPI para incrustación de imágenes (Content-ID y ocultar de adjuntos regulares)
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

def normalizar_texto(texto: str) -> str:
    """Normaliza un texto eliminando tildes, espacios extra y convirtiendo a minúsculas."""
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    # Eliminar acentos y tildes
    nfkd = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

def validar_correo(correo: str) -> bool:
    """Verifica si una cadena tiene formato básico de correo electrónico."""
    if not isinstance(correo, str):
        return False
    patron = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return bool(re.match(patron, correo.strip()))

def leer_excel_seguro(ruta):
    """Lee un archivo Excel incluso si está abierto en Microsoft Excel."""
    try:
        import win32file, io
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
        import tempfile, shutil
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, f"temp_excel_{int(time.time())}.xlsx")
        shutil.copy2(ruta, temp_path)
        try:
            return pd.read_excel(temp_path)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

def cargar_base_clientes():
    """Carga el archivo de clientes (.xlsx o .csv) con soporte para archivos abiertos y mapeo flexible."""
    ruta_xlsx = config.ARCHIVO_CLIENTES
    ruta_csv = os.path.splitext(config.ARCHIVO_CLIENTES)[0] + ".csv"

    if os.path.exists(ruta_xlsx):
        print(f"[OK] Leyendo base de datos Excel: {ruta_xlsx}")
        df = leer_excel_seguro(ruta_xlsx)
    elif os.path.exists(ruta_csv):
        print(f"[OK] Leyendo base de datos CSV: {ruta_csv}")
        try:
            df = pd.read_csv(ruta_csv, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(ruta_csv, encoding='latin-1')
    else:
        raise FileNotFoundError(
            f"No se encontró el archivo de clientes ni en '{ruta_xlsx}' ni en '{ruta_csv}'."
        )

    # Limpiar nombres de columnas
    df.columns = [str(col).strip() for col in df.columns]

    # Mapeo inteligente de columnas si no coinciden exactamente con config.py
    col_nombre = None
    col_correo = None
    col_ubicacion = None

    for col in df.columns:
        norm = normalizar_texto(col)
        if not col_correo and any(k in norm for k in ["email", "correo"]):
            col_correo = col
        elif not col_nombre and any(k in norm for k in ["razon social", "nombre", "cliente", "empresa"]):
            col_nombre = col
        elif not col_ubicacion and any(k in norm for k in ["ciudad", "ubicacion", "depto", "estado", "municipio"]):
            col_ubicacion = col

    if col_nombre:
        config.COLUMNA_NOMBRE = col_nombre
    if col_correo:
        config.COLUMNA_CORREO = col_correo
    if col_ubicacion:
        config.COLUMNA_UBICACION = col_ubicacion

    print(f"[COLUMNAS MAPEADAS] Nombre: '{config.COLUMNA_NOMBRE}' | Correo: '{config.COLUMNA_CORREO}' | Ubicación: '{config.COLUMNA_UBICACION}'")
    return df


def cargar_tabla_asesores():
    """Carga el diccionario o archivo JSON de asesores por ubicación."""
    if not os.path.exists(config.ARCHIVO_ASESORES):
        print(f"[AVISO] No se encontró '{config.ARCHIVO_ASESORES}'. Se usará asesor genérico.")
        return {}

    with open(config.ARCHIVO_ASESORES, "r", encoding="utf-8") as f:
        asesores = json.load(f)

    # Normalizar las claves para búsqueda insensible a mayúsculas y tildes
    asesores_normalizados = {}
    for clave, datos in asesores.items():
        asesores_normalizados[normalizar_texto(clave)] = datos

    return asesores_normalizados

def obtener_asesor_para_ubicacion(ubicacion: str, asesores: dict) -> dict:
    """Retorna los datos del asesor según la ubicación del cliente o el asesor por defecto."""
    ubicacion_norm = normalizar_texto(ubicacion)
    
    # Búsqueda exacta
    if ubicacion_norm in asesores:
        return asesores[ubicacion_norm]
    
    # Búsqueda parcial (ej. si dice "Bogotá D.C." y la clave es "bogota")
    for clave, datos in asesores.items():
        if clave != "default" and (clave in ubicacion_norm or ubicacion_norm in clave):
            return datos

    # Si no coincide, retornar el asesor por defecto
    return asesores.get("default", {
        "nombre": "Atención al Cliente Corporativo",
        "cargo": "Área de Asesoría Comercial",
        "correo": "contacto@empresa.com",
        "telefono": "+57 (601) 800 0000",
        "ciudad_display": "Nacional"
    })

def cargar_historial_enviados() -> set:
    """Carga los correos que ya han sido enviados previamente para no duplicar."""
    if not os.path.exists(config.ARCHIVO_HISTORIAL_LOG):
        return set()

    try:
        df_log = pd.read_csv(config.ARCHIVO_HISTORIAL_LOG, encoding='utf-8')
        if 'Correo' in df_log.columns and 'Estado' in df_log.columns:
            enviados = set(df_log[df_log['Estado'] == 'ENVIADO']['Correo'].dropna().str.strip().str.lower())
            return enviados
    except Exception as e:
        print(f"[AVISO] No se pudo leer el historial previo: {e}")
    return set()

def registrar_en_historial(correo: str, nombre: str, asesor: str, estado: str, detalle: str = ""):
    """Registra en el archivo CSV el resultado de cada intento de envío."""
    archivo_existe = os.path.exists(config.ARCHIVO_HISTORIAL_LOG)
    
    registro = {
        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Correo": correo,
        "Nombre": nombre,
        "Asesor_Asignado": asesor,
        "Estado": estado,
        "Detalle": detalle
    }
    
    df_nuevo = pd.DataFrame([registro])
    df_nuevo.to_csv(
        config.ARCHIVO_HISTORIAL_LOG,
        mode='a',
        index=False,
        header=not archivo_existe,
        encoding='utf-8'
    )

def inicializar_outlook():
    """Inicia la conexión COM con la aplicación de Outlook de escritorio."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        cuentas = namespace.Accounts
        if cuentas.Count > 0:
            cuenta_principal = cuentas.Item(1).DisplayName
            print(f"[OK] Conectado a Microsoft Outlook con la cuenta: {cuenta_principal}")
        else:
            print("[AVISO] Conectado a Outlook pero no se detectaron cuentas activas.")
        return outlook
    except Exception as e:
        print(f"[ERROR] No se pudo conectar a Microsoft Outlook: {e}")
        print("Asegúrate de que la aplicación Outlook de escritorio esté abierta e iniciada.")
        sys.exit(1)

def preparar_cuerpo_html(plantilla_html: str, datos_cliente: dict, datos_asesor: dict, tiene_banner: bool) -> str:
    """Reemplaza los comodines de la plantilla HTML con los datos del cliente y su asesor."""
    html = plantilla_html

    # Reemplazo de la imagen del banner
    if tiene_banner:
        tag_banner = """
        <tr>
          <td align="center" style="background-color: #0f2d59;">
            <img src="cid:banner_principal" width="620" style="width: 100%; max-width: 620px; height: auto; display: block;" alt="Banner Corporativo">
          </td>
        </tr>
        """
    else:
        tag_banner = ""

    html = html.replace("{TAG_IMAGEN_BANNER}", tag_banner)

    # Reemplazo de bloques de texto dinámicos desde config_correo.json si existe
    saludo = "Estimado(a) {nombre},"
    cuerpo = "Nos complace saludarte y ponernos en contacto contigo para compartirte información relevante sobre nuestros servicios y las novedades diseñadas especialmente para atender tus requerimientos en <strong>{ubicacion}</strong>."
    cuerpo_destacado = "Adjunto a este correo encontrarás el documento detallado en formato PDF con el catálogo completo y las especificaciones que hemos preparado para ti."
    cuerpo_cierre = "Con el propósito de brindarte un acompañamiento oportuno y personalizado en tu zona, hemos asignado a un asesor comercial exclusivo para apoyarte:"

    ruta_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_correo.json")
    if os.path.exists(ruta_cfg):
        try:
            with open(ruta_cfg, "r", encoding="utf-8") as f_cfg:
                c_json = json.load(f_cfg)
                saludo = c_json.get("saludo") or saludo
                cuerpo = c_json.get("cuerpo") or cuerpo
                cuerpo_destacado = c_json.get("cuerpo_destacado") or cuerpo_destacado
                cuerpo_cierre = c_json.get("cuerpo_cierre") or cuerpo_cierre
        except Exception:
            pass

    html = html.replace("{saludo}", saludo)
    html = html.replace("{cuerpo}", cuerpo)
    html = html.replace("{cuerpo_destacado}", cuerpo_destacado)
    html = html.replace("{cuerpo_cierre}", cuerpo_cierre)

    # Reemplazo de datos del cliente
    nombre_val = str(datos_cliente.get(config.COLUMNA_NOMBRE, 'Estimado Cliente'))
    if str(nombre_val).strip().lower() in ['nan', 'none', '']:
        nombre_val = 'Estimado Cliente'
    ubic_val = str(datos_cliente.get(config.COLUMNA_UBICACION, 'su ciudad'))
    if str(ubic_val).strip().lower() in ['nan', 'none', '']:
        ubic_val = 'su ciudad'
    correo_val = str(datos_cliente.get(config.COLUMNA_CORREO, ''))

    html = html.replace("{nombre}", nombre_val)
    html = html.replace("{correo}", correo_val)
    html = html.replace("{ubicacion}", ubic_val)

    # Reemplazar cualquier otra columna que exista en el cliente (ej. {Empresa})
    for col, val in datos_cliente.items():
        placeholder = f"{{{col}}}"
        if placeholder in html:
            v_str = str(val) if not pd.isna(val) else ""
            html = html.replace(placeholder, v_str)

    # Reemplazo de datos del asesor
    html = html.replace("{asesor_nombre}", datos_asesor.get("nombre", "Equipo Comercial"))
    html = html.replace("{asesor_cargo}", datos_asesor.get("cargo", "Asesor Comercial"))
    html = html.replace("{asesor_correo}", datos_asesor.get("correo", ""))
    html = html.replace("{asesor_telefono}", datos_asesor.get("telefono", ""))
    html = html.replace("{asesor_zona}", datos_asesor.get("ciudad_display", ubic_val))

    return html

def procesar_envios():
    """Función principal de orquestación y envío de correos."""
    print("=" * 70)
    print("INICIANDO SISTEMA DE ENVÍO MASIVO DE CORREOS CORPORATIVOS")
    print("=" * 70)

    # 1. Validar existencia de plantilla HTML
    if not os.path.exists(config.ARCHIVO_PLANTILLA_HTML):
        print(f"[ERROR] No existe la plantilla HTML: {config.ARCHIVO_PLANTILLA_HTML}")
        return

    with open(config.ARCHIVO_PLANTILLA_HTML, "r", encoding="utf-8") as f:
        plantilla_html = f.read()

    # 2. Validar archivos de adjuntos e imagen
    tiene_banner = False
    if config.ARCHIVO_IMAGEN_BANNER and os.path.exists(config.ARCHIVO_IMAGEN_BANNER):
        tiene_banner = True
        print(f"[OK] Banner de imagen detectado: {config.ARCHIVO_IMAGEN_BANNER}")
    else:
        print("[INFO] No se incrustará banner o no se encontró la imagen.")

    tiene_pdf = False
    if config.ARCHIVO_PDF_ADJUNTO and os.path.exists(config.ARCHIVO_PDF_ADJUNTO):
        tiene_pdf = True
        print(f"[OK] Archivo PDF adjunto detectado: {config.ARCHIVO_PDF_ADJUNTO}")
    else:
        print("[INFO] No se adjuntará PDF o no se encontró el archivo.")

    # 3. Cargar asesores y base de clientes
    asesores = cargar_tabla_asesores()
    df_clientes = cargar_base_clientes()
    total_registros = len(df_clientes)
    print(f"[INFO] Registros cargados en la base de datos: {total_registros}")

    # 4. Cargar historial para evitar reenvíos
    enviados_previamente = cargar_historial_enviados()
    if enviados_previamente:
        print(f"[INFO] Correos ya enviados anteriormente (se omitirán): {len(enviados_previamente)}")

    # 5. Filtrar clientes pendientes
    pendientes = []
    for idx, row in df_clientes.iterrows():
        correo = str(row[config.COLUMNA_CORREO]).strip()
        if correo.lower() not in enviados_previamente:
            pendientes.append(row)

    print(f"[INFO] Destinatarios pendientes por procesar: {len(pendientes)}")
    if not pendientes:
        print("[FINALIZADO] Todos los correos de la base de datos ya fueron enviados exitosamente.")
        return

    # 6. Verificación de Modo Prueba / Producción
    if config.MODO_PRUEBA:
        print("\n" + "!" * 70)
        print("MODO DE PRUEBA ACTIVADO:")
        print(f"- Acción configurada: {config.ACCION_MODO_PRUEBA}")
        if config.ACCION_MODO_PRUEBA == "BORRADORES":
            print(f"- Los correos se crearán y guardarán en la carpeta 'Borradores' de Outlook.")
            print(f"- Límite de prueba: {config.CANTIDAD_PRUEBAS} correos.")
        else:
            print(f"- Se enviarán un máximo de {config.CANTIDAD_PRUEBAS} correos únicamente a: {config.CORREO_TEST}")
        print("!" * 70 + "\n")
        pendientes = pendientes[:config.CANTIDAD_PRUEBAS]
    else:
        print("\n" + "#" * 70)
        print("ADVERTENCIA: ESTÁS EN MODO DE ENVÍO REAL A CLIENTES")
        print(f"Total a enviar: {len(pendientes)} correos.")
        print(f"Pausa entre correos: {config.PAUSA_ENTRE_CORREOS_SEG} segundos.")
        print(f"Lotes de {config.TAMANO_LOTE} correos con pausa de {config.PAUSA_ENTRE_LOTES_SEG} seg.")
        tiempo_estimado_min = (len(pendientes) * config.PAUSA_ENTRE_CORREOS_SEG + (len(pendientes) // config.TAMANO_LOTE) * config.PAUSA_ENTRE_LOTES_SEG) / 60
        print(f"Tiempo total estimado: {tiempo_estimado_min:.1f} minutos.")
        print("#" * 70)
        confirmacion = input("\n¿Confirmas el inicio del envío masivo? Escribe 'SI' para continuar: ").strip().upper()
        if confirmacion != "SI":
            print("[CANCELADO] Operación cancelada por el usuario.")
            return

    # 7. Conectar a Outlook
    outlook = inicializar_outlook()

    # 8. Bucle de procesamiento
    enviados_en_esta_sesion = 0
    errores_en_esta_sesion = 0

    for i, cliente in enumerate(pendientes, 1):
        nombre_cliente = str(cliente.get(config.COLUMNA_NOMBRE, '')).strip()
        correo_cliente = str(cliente.get(config.COLUMNA_CORREO, '')).strip()
        ubicacion_cliente = str(cliente.get(config.COLUMNA_UBICACION, '')).strip()

        # Validar sintaxis del correo
        if not validar_correo(correo_cliente):
            print(f"[{i}/{len(pendientes)}] [ERROR SINTAXIS] Correo inválido '{correo_cliente}' para {nombre_cliente}. Saltando.")
            registrar_en_historial(correo_cliente, nombre_cliente, "N/A", "ERROR", "Formato de correo inválido")
            errores_en_esta_sesion += 1
            continue

        # Obtener asesor asignado dinámicamente según ubicación
        asesor = obtener_asesor_para_ubicacion(ubicacion_cliente, asesores)

        # Destinatario según modo
        destinatario_final = config.CORREO_TEST if (config.MODO_PRUEBA and config.ACCION_MODO_PRUEBA == "ENVIAR_TEST") else correo_cliente

        # Asunto personalizado
        asunto = config.ASUNTO_CORREO.format(
            nombre=nombre_cliente,
            ubicacion=ubicacion_cliente,
            asesor_nombre=asesor.get("nombre", "")
        )
        if config.MODO_PRUEBA:
            asunto = f"[PRUEBA ({config.ACCION_MODO_PRUEBA})] " + asunto

        try:
            # Crear ítem de correo en Outlook (0 = olMailItem)
            mail = outlook.CreateItem(0)
            mail.To = destinatario_final
            mail.Subject = asunto

            # Incrustar banner como imagen CID en el cuerpo HTML
            if tiene_banner:
                ruta_abs_banner = os.path.abspath(config.ARCHIVO_IMAGEN_BANNER)
                att_banner = mail.Attachments.Add(ruta_abs_banner)
                # Configurar Content-ID para referenciarlo como cid:banner_principal
                att_banner.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, "banner_principal")
                att_banner.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, True)

            # Adjuntar PDF si existe
            if tiene_pdf:
                ruta_abs_pdf = os.path.abspath(config.ARCHIVO_PDF_ADJUNTO)
                mail.Attachments.Add(ruta_abs_pdf)

            # Preparar y asignar el cuerpo HTML
            cuerpo_html = preparar_cuerpo_html(plantilla_html, cliente.to_dict(), asesor, tiene_banner)
            mail.HTMLBody = cuerpo_html

            # Enviar o guardar en borradores
            if config.MODO_PRUEBA and config.ACCION_MODO_PRUEBA == "BORRADORES":
                mail.Save()
                estado = "BORRADOR_GUARDADO"
                print(f"[{i}/{len(pendientes)}] [BORRADOR] {nombre_cliente} ({correo_cliente}) -> Asesor: {asesor.get('nombre')} [{asesor.get('ciudad_display', ubicacion_cliente)}]")
            else:
                mail.Send()
                estado = "ENVIADO"
                print(f"[{i}/{len(pendientes)}] [ENVIADO] {nombre_cliente} ({destinatario_final}) -> Asesor: {asesor.get('nombre')} [{asesor.get('ciudad_display', ubicacion_cliente)}]")

            # Registrar en log
            registrar_en_historial(correo_cliente, nombre_cliente, asesor.get("nombre", ""), estado)
            enviados_en_esta_sesion += 1

        except Exception as e:
            print(f"[{i}/{len(pendientes)}] [FALLO] Error enviando a {correo_cliente}: {e}")
            registrar_en_historial(correo_cliente, nombre_cliente, asesor.get("nombre", ""), "ERROR", str(e))
            errores_en_esta_sesion += 1

        # Control de ritmo (Throttling) entre correos
        time.sleep(config.PAUSA_ENTRE_CORREOS_SEG)

        # Control de pausas por lotes para proteger la reputación de la cuenta
        if i % config.TAMANO_LOTE == 0 and i < len(pendientes):
            print(f"\n>>> Lote de {config.TAMANO_LOTE} completado. Pausa de seguridad de {config.PAUSA_ENTRE_LOTES_SEG} seg para Exchange... <<<\n")
            time.sleep(config.PAUSA_ENTRE_LOTES_SEG)

    # Resumen final
    print("\n" + "=" * 70)
    print("PROCESO FINALIZADO")
    print(f"- Exitosos / Procesados: {enviados_en_esta_sesion}")
    print(f"- Fallidos: {errores_en_esta_sesion}")
    print(f"- Historial registrado en: {config.ARCHIVO_HISTORIAL_LOG}")
    if config.MODO_PRUEBA and config.ACCION_MODO_PRUEBA == "BORRADORES":
        print("-> Abre tu Outlook y revisa la carpeta 'Borradores' para inspeccionar cómo quedaron los correos.")
    print("=" * 70)

if __name__ == "__main__":
    procesar_envios()
