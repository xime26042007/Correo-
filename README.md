# Guía Completa de Envío Masivo por Outlook Corporativo

Este proyecto automatiza el envío masivo de correos corporativos utilizando **Python** y **Microsoft Outlook de escritorio (MAPI)**, con asignación dinámica de asesores por ubicación geográfica, incrustación de imagen/banner en el cuerpo, archivo PDF adjunto y control estricto de límites para proteger tu cuenta de correo contra bloqueos por spam.

---

## 📁 Estructura del Proyecto

```text
CORREOS_MASIVOS/
├── config.py                 # Configuración general (tiempos, modo prueba, rutas)
├── asesores.json             # Tabla de mapeo: Ubicación -> Asesor comercial asignado
├── plantilla_correo.html     # Plantilla HTML profesional compatible con Outlook
├── enviar_correos.py         # Script principal de envío y orquestación
├── clientes.csv              # Base de datos de ejemplo (o clientes.xlsx)
├── historial_envios.csv      # Log automático para evitar duplicados y permitir reanudación
├── assets/
│   ├── banner_correo.png     # Banner corporativo (incrustado vía CID en el cuerpo)
│   └── informacion_corporativa.pdf # Documento PDF a adjuntar
└── requirements.txt          # Dependencias necesarias (pandas, openpyxl, pywin32)
```

---

## ⚠️ Aspectos Críticos para Enviar 4,000 Correos por Outlook

Microsoft 365 / Exchange Online aplica límites automáticos de seguridad:
1. **Límite de velocidad (Submission Rate)**: Máximo **30 correos por minuto**. Si envías más rápido, el servidor bloqueará temporalmente la cuenta con el error `550 5.7.501 Service unavailable`.
2. **Límite diario de destinatarios**: Generalmente **10,000 destinatarios/día** en cuentas estándar de Microsoft 365. Con 4,000 correos estás dentro del margen diario, pero **debe enviarse con pausas controladas**.
3. **Tiempo estimado para 4,000 correos**:
   - Con pausa de **2.5 segundos por correo** y descansos de **45 segundos cada 50 correos**, el envío total tomará aproximadamente **2.5 a 3 horas**.
4. **Sistema Anti-Duplicados**:
   - Cada correo enviado se registra en `historial_envios.csv`.
   - Si la computadora se apaga, se desconecta internet o detienes el script en el correo 1,500, al volver a ejecutarlo **continuará automáticamente desde el 1,501** sin reenviar a los primeros 1,500.

---

## 🚀 Paso a Paso para la Puesta en Marcha

### Paso 1: Instalar Python (si no lo tienes instalado)
En una terminal de PowerShell:
```powershell
winget install --id Python.Python.3.12 --source winget --accept-package-agreements --accept-source-agreements
```
*(Reinicia la terminal tras instalar para que cargue la variable PATH)*.

### Paso 2: Instalar las dependencias
```powershell
pip install -r requirements.txt
```

### Paso 3: Configurar Asesores por Ubicación
Edita el archivo `asesores.json` para agregar las ciudades o departamentos de tus clientes y los datos del asesor correspondiente:
```json
{
  "bogota": {
    "nombre": "Carlos Andrés Mendoza",
    "cargo": "Asesor Senior Regional Centro",
    "correo": "carlos.mendoza@empresa.com",
    "telefono": "+57 310 123 4567",
    "ciudad_display": "Bogotá D.C."
  },
  "default": {
    "nombre": "Equipo de Asesoría Corporativa",
    "cargo": "Atención al Cliente Nacional",
    "correo": "contacto@empresa.com",
    "telefono": "+57 (601) 800 0000",
    "ciudad_display": "Nacional"
  }
}
```
*Si un cliente tiene una ubicación que no coincide con ninguna ciudad, automáticamente se le asignará el asesor `default`.*

### Paso 4: Cargar la Base de Clientes
Puedes colocar tu archivo como `clientes.xlsx` o `clientes.csv` en la carpeta raíz.
Asegúrate de que tenga al menos las siguientes columnas:
- **`Nombre`**: Nombre completo o tratamiento del cliente.
- **`Correo`**: Correo electrónico del destinatario.
- **`Ubicacion`**: Ciudad o zona del cliente (ej. Bogotá, Medellín, Cali, etc.).
*(Opcional: puedes agregar columnas extra como `Empresa`, `Identificacion`, etc., y llamarlas en el correo usando `{Empresa}`).*

### Paso 5: Ejecutar la Prueba Segura (Modo Borradores)
Por defecto, `config.py` viene en **`MODO_PRUEBA = True`** y **`ACCION_MODO_PRUEBA = "BORRADORES"`**.

Ejecuta el script:
```powershell
python enviar_correos.py
```
- No enviará correos reales a nadie.
- Creará los primeros 3 correos con la imagen, el PDF y el asesor asignado según la ubicación, y los guardará en tu carpeta **Borradores** de Outlook.
- Abre Outlook, inspecciona los borradores y comprueba que el diseño, los datos y la firma se vean perfectos.

### Paso 6: Lanzar el Envío Masivo Real
Cuando hayas verificado que los borradores son correctos:
1. Abre `config.py`.
2. Cambia `MODO_PRUEBA = False`.
3. Ejecuta de nuevo:
   ```powershell
   python enviar_correos.py
   ```
4. El script te mostrará un resumen del tiempo estimado y te pedirá confirmar escribiendo `SI`.
5. Deja la computadora encendida y Outlook abierto mientras procesa los envíos.
