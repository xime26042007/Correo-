import os

# ==============================================================================
# CONFIGURACIÓN GENERAL DE ENVÍO MASIVO POR OUTLOOK
# ==============================================================================

# Directorio raíz del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 1. ARCHIVOS DE ENTRADA Y RECURSOS
# Archivo de base de datos con los clientes (puede ser .xlsx o .csv)
ARCHIVO_CLIENTES = os.path.join(BASE_DIR, "clientes.xlsx")

# Archivo con el mapeo de asesores según la ubicación/ciudad
ARCHIVO_ASESORES = os.path.join(BASE_DIR, "asesores.json")

# Plantilla HTML del correo
ARCHIVO_PLANTILLA_HTML = os.path.join(BASE_DIR, "plantilla_correo.html")

# Imagen para incrustar en el cuerpo del correo (dejar None si no se usará)
# Nota: La imagen se incrusta usando CID para que se vea directamente en el cuerpo
ARCHIVO_IMAGEN_BANNER = os.path.join(BASE_DIR, "assets", "banner_correo.png")

# Archivo PDF a adjuntar (dejar None o comentar si no se desea adjuntar)
ARCHIVO_PDF_ADJUNTO = os.path.join(BASE_DIR, "assets", "informacion_corporativa.pdf")

# Archivo donde se registra el progreso y estado de cada correo
ARCHIVO_HISTORIAL_LOG = os.path.join(BASE_DIR, "historial_envios.csv")

# 2. CONTENIDO DEL MENSAJE
# Asunto del correo (admite comodines como {nombre})
ASUNTO_CORREO = "Información exclusiva para {nombre} - Actualización Corporativa"

# 3. CONTROL DE SEGURIDAD Y MODO DE PRUEBA
# IMPORTANTE:
# Deja MODO_PRUEBA = True para hacer pruebas antes de enviar a los 4000 clientes.
# En modo prueba NO se enviará a los clientes reales.
MODO_PRUEBA = True

# Opciones para MODO_PRUEBA:
# 'BORRADORES': Crea el correo y lo guarda en la carpeta Borradores de Outlook (no envía).
# 'ENVIAR_TEST': Envía los correos de prueba únicamente al correo especificado en CORREO_TEST.
ACCION_MODO_PRUEBA = "BORRADORES"  # Opciones: "BORRADORES" o "ENVIAR_TEST"
CORREO_TEST = "tu_correo_personal_o_pruebas@empresa.com"
CANTIDAD_PRUEBAS = 3  # Cantidad de correos a procesar en modo prueba

# 4. GESTIÓN DE VELOCIDAD Y LÍMITES DE MICROSOFT 365 (EXCHANGE)
# Microsoft 365 impone un límite de ~30 mensajes por minuto.
# Si se envían más rápido, la cuenta corporativa puede ser bloqueada por spam.
PAUSA_ENTRE_CORREOS_SEG = 2.5  # Pausa entre cada correo individual (2.5 a 3.5 seg)
TAMANO_LOTE = 50                # Cantidad de correos por cada bloque/lote
PAUSA_ENTRE_LOTES_SEG = 45      # Pausa de descanso tras completar cada lote (en segundos)

# 5. NOMBRES DE COLUMNAS ESPERADOS EN EL EXCEL / CSV
COLUMNA_NOMBRE = "Nombre"
COLUMNA_CORREO = "Correo"
COLUMNA_UBICACION = "Ubicacion"
# Puedes añadir más columnas según tu base de datos (ej. "Empresa", "Saldo", etc.)
