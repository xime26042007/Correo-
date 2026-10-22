// ==============================================================================
// LÓGICA FRONTEND COMPLETA CON CONTROL TOTAL DE BOTONES Y ENVÍO ROBUSTO
// ==============================================================================

let estadoApp = {
  paginaActual: 1,
  limitePorPagina: 50,
  totalClientes: 0,
  totalConCorreo: 0,
  clientesVisibles: [],
  indicesSeleccionados: new Set(),
  asesoresData: {},
  intervaloProgreso: null,
  clientePreviewIndex: 0,
  ultimoCampoId: "inputCuerpo",
  debounceTimer: null,
  previewDebounceTimer: null
};

document.addEventListener("DOMContentLoaded", () => {
  cargarEstadoInicial();

  const usuarioInput = document.getElementById("inputSmtpUsuario");
  if (usuarioInput) {
    usuarioInput.addEventListener("input", (e) => {
      const email = e.target.value.trim().toLowerCase();
      const hostInput = document.getElementById("inputSmtpHost");
      const portInput = document.getElementById("inputSmtpPort");
      if (hostInput && portInput) {
        if (email.endsWith("@gmail.com")) {
          if (hostInput.value === "smtp.office365.com" || !hostInput.value) {
            hostInput.value = "smtp.gmail.com";
            portInput.value = "587";
          }
        } else if (email.endsWith("@outlook.com") || email.endsWith("@hotmail.com") || email.endsWith(".edu.co")) {
          if (hostInput.value === "smtp.gmail.com" || !hostInput.value) {
            hostInput.value = "smtp.office365.com";
            portInput.value = "587";
          }
        }
      }
    });
  }
});

// Toast flotante elegante
function mostrarToast(mensaje, tipo = "info") {
  const container = document.getElementById("toastContainer");
  if (!container) return;

  const toast = document.createElement("div");
  toast.className = `toast toast-${tipo}`;
  
  const iconos = {
    success: "✅",
    info: "ℹ️",
    warning: "⚠️",
    error: "❌"
  };
  
  toast.innerHTML = `<span>${iconos[tipo] || "ℹ️"}</span> <span>${escapeHtml(mensaje)}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateY(-10px)";
    toast.style.transition = "all 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// 1. CARGA INICIAL DEL ESTADO
async function cargarEstadoInicial() {
  try {
    const res = await fetch("/api/estado-general");
    const data = await res.json();

    // Estado Excel
    const badgeExcel = document.getElementById("badgeExcel");
    const txtExcel = document.getElementById("excelStatusText");
    if (data.total_registros > 0) {
      badgeExcel.classList.add("connected");
      const conCorreo = data.total_con_correo ? data.total_con_correo.toLocaleString() : "0";
      txtExcel.textContent = `Excel Conectado: ${data.total_registros.toLocaleString()} clientes (${conCorreo} con correo)`;
    } else {
      txtExcel.textContent = "Sin base de datos detectada";
    }

    // Configuración de Textos y Emisor
    if (data.config_correo) {
      const cfg = data.config_correo;
      if (cfg.asunto) document.getElementById("inputAsunto").value = cfg.asunto;
      if (cfg.saludo) document.getElementById("inputSaludo").value = cfg.saludo;
      if (cfg.cuerpo) document.getElementById("inputCuerpo").value = cfg.cuerpo;
      if (cfg.cuerpo_destacado) document.getElementById("inputCuerpoDestacado").value = cfg.cuerpo_destacado;
      if (cfg.cuerpo_cierre) document.getElementById("inputCuerpoCierre").value = cfg.cuerpo_cierre;
      if (cfg.correo_test) document.getElementById("inputCorreoTest").value = cfg.correo_test;
      
      // Modal Configuración
      const metodo = cfg.metodo || "OUTLOOK";
      document.getElementById("selectMetodoEnvio").value = metodo;
      document.getElementById("inputSmtpHost").value = cfg.smtp_host || "smtp.office365.com";
      document.getElementById("inputSmtpPort").value = cfg.smtp_port || 587;
      document.getElementById("inputSmtpUsuario").value = cfg.smtp_usuario || "";
      document.getElementById("inputRemitenteNombre").value = cfg.remitente_nombre || "Dirección Comercial Corporativa";
      
      if (cfg.tiene_password) {
        document.getElementById("txtPasswordGuardada").textContent = "● Contraseña guardada en el servidor (déjala en blanco para conservarla)";
      } else {
        document.getElementById("txtPasswordGuardada").textContent = "";
      }

      actualizarBadgeMetodo(metodo, cfg.smtp_usuario);
      toggleCamposMetodo(metodo);
    }

    // Contador total de clientes
    document.getElementById("countTotalClientes").textContent = data.total_con_correo ? data.total_con_correo.toLocaleString() : data.total_registros.toLocaleString();
    estadoApp.totalClientes = data.total_registros;
    estadoApp.totalConCorreo = data.total_con_correo;

    // Selector de ciudades
    const selectCiudad = document.getElementById("selectCiudad");
    selectCiudad.innerHTML = '<option value="">Todas las ubicaciones</option>';
    if (data.ciudades_disponibles) {
      data.ciudades_disponibles.forEach(ciudad => {
        const opt = document.createElement("option");
        opt.value = ciudad;
        opt.textContent = ciudad;
        selectCiudad.appendChild(opt);
      });
    }

    // Asesores
    estadoApp.asesoresData = data.asesores || {};
    renderizarResumenAsesores();

    // Cargar primera página de clientes
    await cargarClientes(1);

    // Cargar vista previa inicial
    actualizarPreview();

  } catch (err) {
    console.error("Error cargando estado inicial:", err);
    mostrarToast("Error conectando con el servidor: " + err.message, "error");
  }
}

function actualizarBadgeMetodo(metodo, usuario) {
  const txtBadge = document.getElementById("txtMetodoBadge");
  if (!txtBadge) return;
  if (metodo === "OUTLOOK") {
    txtBadge.textContent = "Outlook Desktop";
  } else {
    txtBadge.textContent = usuario ? `SMTP (${usuario})` : "SMTP Corporativo";
  }
}

// 2. RECARGAR BASE DE DATOS
async function recargarBaseDatos() {
  mostrarToast("Recargando base de datos desde Excel...", "info");
  try {
    const res = await fetch("/api/recargar-datos", { method: "POST" });
    const data = await res.json();
    if (data.ok) {
      mostrarToast(`¡Base de datos actualizada! ${data.total.toLocaleString()} clientes cargados (${data.total_con_correo.toLocaleString()} con correo).`, "success");
      await cargarEstadoInicial();
    } else {
      mostrarToast("No se pudo recargar la base de datos", "error");
    }
  } catch (err) {
    mostrarToast("Error recargando: " + err.message, "error");
  }
}

// 3. CARGA Y FILTRADO DE CLIENTES
async function cargarClientes(pagina = 1) {
  estadoApp.paginaActual = pagina;
  const busqueda = document.getElementById("inputBusqueda").value.trim();
  const ciudad = document.getElementById("selectCiudad").value.trim();
  const soloValidos = document.getElementById("chkSoloValidos").checked ? "1" : "0";

  const tbody = document.getElementById("tbodyClientes");
  tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4" style="color: #94a3b8; text-align: center; padding: 24px;">Buscando registros en la base de datos...</td></tr>';

  try {
    const url = `/api/clientes?pagina=${pagina}&limite=${estadoApp.limitePorPagina}&q=${encodeURIComponent(busqueda)}&ciudad=${encodeURIComponent(ciudad)}&solo_correo=${soloValidos}`;
    const res = await fetch(url);
    const data = await res.json();

    estadoApp.clientesVisibles = data.clientes;
    renderizarTablaClientes(data.clientes);

    const inicio = (pagina - 1) * estadoApp.limitePorPagina + (data.total > 0 ? 1 : 0);
    const fin = Math.min(pagina * estadoApp.limitePorPagina, data.total);
    document.getElementById("txtInfoPaginacion").textContent = `Mostrando ${inicio} a ${fin} de ${data.total.toLocaleString()} destinatarios`;
    document.getElementById("txtNumPagina").textContent = `Página ${pagina} de ${Math.ceil(data.total / estadoApp.limitePorPagina) || 1}`;
    
    document.getElementById("btnPagePrev").disabled = pagina <= 1;
    document.getElementById("btnPageNext").disabled = fin >= data.total;

  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-4" style="color: #ef4444; text-align: center;">Error cargando clientes: ${err.message}</td></tr>`;
  }
}

function filtrarClientesDebounced() {
  clearTimeout(estadoApp.debounceTimer);
  estadoApp.debounceTimer = setTimeout(() => {
    cargarClientes(1);
  }, 180);
}

function filtrarClientes() {
  cargarClientes(1);
}

function renderizarTablaClientes(clientes) {
  const tbody = document.getElementById("tbodyClientes");
  if (!clientes || clientes.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4" style="color: #94a3b8; text-align: center; padding: 24px;">No se encontraron registros con los filtros actuales.</td></tr>';
    return;
  }

  tbody.innerHTML = "";
  clientes.forEach(cli => {
    const isChecked = estadoApp.indicesSeleccionados.has(cli.indice_original);
    const tr = document.createElement("tr");

    let estadoBadge = "";
    if (cli.ya_enviado) {
      estadoBadge = '<span class="badge-status-mini badge-enviado">✓ Enviado</span>';
    } else if (!cli.correo_valido) {
      estadoBadge = '<span class="badge-status-mini badge-invalido">Sin correo</span>';
    }

    tr.innerHTML = `
      <td>
        <input type="checkbox" class="chk-cliente" data-idx="${cli.indice_original}" ${isChecked ? 'checked' : ''} onchange="toggleSeleccionCliente(${cli.indice_original}, this.checked)">
      </td>
      <td>
        <span class="client-name">${escapeHtml(cli.nombre)}</span>
      </td>
      <td>
        <span class="client-email">${cli.correo ? escapeHtml(cli.correo) : '<em style="color:#ef4444;">Sin correo</em>'}</span>
        ${estadoBadge}
      </td>
      <td>
        <span class="client-badge-city">${escapeHtml(cli.ubicacion)}</span>
      </td>
      <td>
        <strong style="color: #60a5fa;">${escapeHtml(cli.asesor)}</strong>
        <div style="font-size: 11px; color: #94a3b8;">${escapeHtml(cli.asesor_zona)}</div>
      </td>
      <td>
        <button type="button" class="btn-preview-mini" onclick="verPreviewCliente(${cli.indice_original})">Ver Correo</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  actualizarContadorSeleccion();
}

// 4. SELECCIÓN DE CLIENTES
function toggleSeleccionCliente(indice, checked) {
  if (checked) {
    estadoApp.indicesSeleccionados.add(indice);
  } else {
    estadoApp.indicesSeleccionados.delete(indice);
  }
  actualizarContadorSeleccion();
}

function toggleMaster(masterCheckbox) {
  seleccionarTodosVisibles(masterCheckbox.checked);
}

function seleccionarTodosVisibles(select) {
  estadoApp.clientesVisibles.forEach(cli => {
    // Si se seleccionan, priorizar los que tengan correo válido
    if (select) {
      if (cli.correo_valido) {
        estadoApp.indicesSeleccionados.add(cli.indice_original);
      }
    } else {
      estadoApp.indicesSeleccionados.delete(cli.indice_original);
    }
  });

  document.querySelectorAll(".chk-cliente").forEach(chk => {
    const idx = parseInt(chk.dataset.idx);
    chk.checked = estadoApp.indicesSeleccionados.has(idx);
  });
  document.getElementById("chkMaster").checked = select;

  actualizarContadorSeleccion();
  if (select) {
    mostrarToast(`${estadoApp.indicesSeleccionados.size} clientes seleccionados`, "info");
  } else {
    mostrarToast("Selección limpiada", "info");
  }
}

async function seleccionarTodosFiltrados() {
  const busqueda = document.getElementById("inputBusqueda").value.trim();
  const ciudad = document.getElementById("selectCiudad").value.trim();
  const soloValidos = document.getElementById("chkSoloValidos").checked;

  mostrarToast("Seleccionando todos los destinatarios filtrados...", "info");
  try {
    const res = await fetch("/api/seleccionar-filtrados", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ q: busqueda, ciudad: ciudad, solo_correo: soloValidos })
    });
    const data = await res.json();
    if (data.indices) {
      data.indices.forEach(idx => estadoApp.indicesSeleccionados.add(idx));
      document.querySelectorAll(".chk-cliente").forEach(chk => {
        const idx = parseInt(chk.dataset.idx);
        if (estadoApp.indicesSeleccionados.has(idx)) {
          chk.checked = true;
        }
      });
      document.getElementById("chkMaster").checked = true;
      actualizarContadorSeleccion();
      mostrarToast(`¡${data.total.toLocaleString()} clientes válidos seleccionados!`, "success");
    }
  } catch (err) {
    mostrarToast("Error en selección masiva: " + err.message, "error");
  }
}

function actualizarContadorSeleccion() {
  const total = estadoApp.indicesSeleccionados.size;
  document.getElementById("countSeleccionados").textContent = total.toLocaleString();
}

function cambiarPagina(delta) {
  cargarClientes(estadoApp.paginaActual + delta);
}

// 5. EDITOR DE TEXTOS Y PREVISUALIZADOR EN VIVO
function marcarCampoEnfocado(campoId) {
  estadoApp.ultimoCampoId = campoId;
}

function insertarEnCampoActivo(tag) {
  const campoId = estadoApp.ultimoCampoId || "inputCuerpo";
  const el = document.getElementById(campoId);
  if (!el) return;

  const start = el.selectionStart || el.value.length;
  const end = el.selectionEnd || el.value.length;
  const val = el.value;

  el.value = val.substring(0, start) + tag + val.substring(end);
  el.selectionStart = el.selectionEnd = start + tag.length;
  el.focus();

  mostrarToast(`Comodín ${tag} insertado`, "info");
  actualizarPreview();
}

function actualizarPreviewDebounced() {
  clearTimeout(estadoApp.previewDebounceTimer);
  estadoApp.previewDebounceTimer = setTimeout(() => {
    actualizarPreview();
  }, 220);
}

async function actualizarPreview() {
  const asunto = document.getElementById("inputAsunto").value;
  const saludo = document.getElementById("inputSaludo").value;
  const cuerpo = document.getElementById("inputCuerpo").value;
  const cuerpoDestacado = document.getElementById("inputCuerpoDestacado").value;
  const cuerpoCierre = document.getElementById("inputCuerpoCierre").value;

  try {
    const res = await fetch("/api/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        indice: estadoApp.clientePreviewIndex,
        asunto: asunto,
        saludo: saludo,
        cuerpo: cuerpo,
        cuerpo_destacado: cuerpoDestacado,
        cuerpo_cierre: cuerpoCierre,
        imagen: "banner_correo.png"
      })
    });
    const data = await res.json();
    const iframe = document.getElementById("iframePreview");
    if (iframe) {
      iframe.srcdoc = data.html;
    }
    if (data.cliente && data.cliente.nombre) {
      document.getElementById("previewClienteNombre").textContent = data.cliente.nombre;
    }
  } catch (err) {
    console.error("Error actualizando preview:", err);
  }
}

function verPreviewCliente(indiceOriginal) {
  estadoApp.clientePreviewIndex = indiceOriginal;
  const cli = estadoApp.clientesVisibles.find(c => c.indice_original === indiceOriginal);
  const nombre = cli ? cli.nombre : "Cliente";
  document.getElementById("previewClienteNombre").textContent = nombre;
  cambiarPestana('preview');
  actualizarPreview();
}

function cambiarPestana(tabName) {
  document.querySelectorAll(".tab-btn").forEach(btn => btn.classList.remove("active"));
  document.querySelectorAll(".tab-view").forEach(view => view.classList.remove("active"));

  if (tabName === 'clientes') {
    document.getElementById("tabBtnClientes").classList.add("active");
    document.getElementById("viewClientes").classList.add("active");
  } else {
    document.getElementById("tabBtnPreview").classList.add("active");
    document.getElementById("viewPreview").classList.add("active");
    actualizarPreview();
  }
}

async function guardarTextosCorreo() {
  const datos = {
    asunto: document.getElementById("inputAsunto").value,
    saludo: document.getElementById("inputSaludo").value,
    cuerpo: document.getElementById("inputCuerpo").value,
    cuerpo_destacado: document.getElementById("inputCuerpoDestacado").value,
    cuerpo_cierre: document.getElementById("inputCuerpoCierre").value
  };

  try {
    const res = await fetch("/api/guardar-textos-correo", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(datos)
    });
    const data = await res.json();
    if (data.ok) {
      mostrarToast("¡Textos de la campaña guardados exitosamente!", "success");
      actualizarPreview();
    }
  } catch (err) {
    mostrarToast("Error guardando textos: " + err.message, "error");
  }
}

function restablecerTextos() {
  document.getElementById("inputAsunto").value = "Información exclusiva para {nombre} - Actualización Corporativa";
  document.getElementById("inputSaludo").value = "Estimado(a) {nombre},";
  document.getElementById("inputCuerpo").value = "Nos complace saludarte y ponernos en contacto contigo para compartirte información relevante sobre nuestros servicios y las novedades diseñadas especialmente para atender tus requerimientos en <strong>{ubicacion}</strong>.";
  document.getElementById("inputCuerpoDestacado").value = "Adjunto a este correo encontrarás el documento detallado en formato PDF con el catálogo completo y las especificaciones que hemos preparado para ti.";
  document.getElementById("inputCuerpoCierre").value = "Con el propósito de brindarte un acompañamiento oportuno y personalizado en tu zona, hemos asignado a un asesor comercial exclusivo para apoyarte:";
  mostrarToast("Textos restablecidos a los valores corporativos", "info");
  actualizarPreview();
}

// 6. SUBIDA DE ARCHIVOS (BANNER Y PDF)
async function subirNuevaImagen(e) {
  const file = e.target.files[0];
  if (!file) return;

  mostrarToast("Subiendo nueva imagen del banner...", "info");
  const formData = new FormData();
  formData.append("archivo", file);

  try {
    const res = await fetch("/api/subir-banner", {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (data.ok) {
      document.getElementById("imgBannerMini").src = "/assets/banner_correo.png?t=" + new Date().getTime();
      document.getElementById("txtBannerName").textContent = file.name;
      mostrarToast("¡Imagen actualizada exitosamente!", "success");
      actualizarPreview();
    } else {
      mostrarToast(data.error || "Error al subir la imagen", "error");
    }
  } catch (err) {
    mostrarToast("Error subiendo imagen: " + err.message, "error");
  }
}

async function subirNuevoPdf(e) {
  const file = e.target.files[0];
  if (!file) return;

  mostrarToast("Subiendo documento PDF corporativo...", "info");
  const formData = new FormData();
  formData.append("archivo", file);

  try {
    const res = await fetch("/api/subir-pdf", {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (data.ok) {
      document.getElementById("txtPdfName").textContent = file.name;
      mostrarToast("¡Documento PDF actualizado exitosamente!", "success");
    } else {
      mostrarToast(data.error || "Error al subir PDF", "error");
    }
  } catch (err) {
    mostrarToast("Error subiendo PDF: " + err.message, "error");
  }
}

// 7. GESTIÓN DE ASESORES POR UBICACIÓN
function renderizarResumenAsesores() {
  const contenedor = document.getElementById("listaAsesoresResumen");
  contenedor.innerHTML = "";
  for (const [clave, datos] of Object.entries(estadoApp.asesoresData)) {
    const div = document.createElement("div");
    div.className = "advisor-row-item";
    div.innerHTML = `
      <span class="advisor-badge-city">${escapeHtml(datos.ciudad_display || clave.toUpperCase())}</span>
      <span>${escapeHtml(datos.nombre)}</span>
      <span style="color: #94a3b8; font-size: 11px;">${escapeHtml(datos.telefono || '')}</span>
    `;
    contenedor.appendChild(div);
  }
}

function abrirModalAsesores() {
  const modal = document.getElementById("modalAsesores");
  const contenedor = document.getElementById("contenedorAsesoresForm");
  contenedor.innerHTML = "";

  for (const [clave, datos] of Object.entries(estadoApp.asesoresData)) {
    const card = document.createElement("div");
    card.style.display = "grid";
    card.style.gridTemplateColumns = "1.2fr 1.5fr 1.2fr 1.5fr 1.2fr auto";
    card.style.gap = "8px";
    card.style.alignItems = "center";
    card.className = "advisor-card-edit";
    card.innerHTML = `
      <input type="text" class="input-city input-control" placeholder="Ciudad / Zona" value="${escapeHtml(datos.ciudad_display || clave)}">
      <input type="text" class="input-name input-control" placeholder="Nombre Asesor" value="${escapeHtml(datos.nombre || '')}">
      <input type="text" class="input-role input-control" placeholder="Cargo" value="${escapeHtml(datos.cargo || 'Asesor Comercial')}">
      <input type="text" class="input-email input-control" placeholder="Correo Asesor" value="${escapeHtml(datos.correo || '')}">
      <input type="text" class="input-phone input-control" placeholder="Teléfono" value="${escapeHtml(datos.telefono || '')}">
      <button type="button" class="btn-text" style="color: #ef4444; font-size: 20px;" onclick="this.parentElement.remove()" title="Eliminar">&times;</button>
    `;
    contenedor.appendChild(card);
  }

  modal.classList.add("open");
}

function cerrarModalAsesores() {
  document.getElementById("modalAsesores").classList.remove("open");
}

function agregarNuevoAsesor() {
  const contenedor = document.getElementById("contenedorAsesoresForm");
  const card = document.createElement("div");
  card.style.display = "grid";
  card.style.gridTemplateColumns = "1.2fr 1.5fr 1.2fr 1.5fr 1.2fr auto";
  card.style.gap = "8px";
  card.style.alignItems = "center";
  card.className = "advisor-card-edit";
  card.innerHTML = `
    <input type="text" class="input-city input-control" placeholder="Ciudad / Región">
    <input type="text" class="input-name input-control" placeholder="Nombre Asesor">
    <input type="text" class="input-role input-control" placeholder="Cargo" value="Asesor Comercial Regional">
    <input type="text" class="input-email input-control" placeholder="Correo Asesor">
    <input type="text" class="input-phone input-control" placeholder="Teléfono">
    <button type="button" class="btn-text" style="color: #ef4444; font-size: 20px;" onclick="this.parentElement.remove()" title="Eliminar">&times;</button>
  `;
  contenedor.appendChild(card);
}

async function guardarAsesores() {
  const cards = document.querySelectorAll(".advisor-card-edit");
  const nuevosAsesores = {};

  cards.forEach(c => {
    const ciudad = c.querySelector(".input-city").value.trim();
    const nombre = c.querySelector(".input-name").value.trim();
    const cargo = c.querySelector(".input-role").value.trim();
    const email = c.querySelector(".input-email").value.trim();
    const tel = c.querySelector(".input-phone").value.trim();

    if (ciudad) {
      const claveNorm = ciudad.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
      nuevosAsesores[claveNorm] = {
        nombre: nombre || "Asesor Comercial",
        cargo: cargo || "Asesor Comercial Regional",
        correo: email,
        telefono: tel,
        ciudad_display: ciudad
      };
    }
  });

  if (!nuevosAsesores["default"]) {
    nuevosAsesores["default"] = {
      nombre: "Equipo de Asesoría Corporativa",
      cargo: "Atención y Servicio al Cliente Nacional",
      correo: "contacto@empresa.com",
      telefono: "+57 (601) 800 0000",
      ciudad_display: "Nacional"
    };
  }

  try {
    const res = await fetch("/api/guardar-asesores", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(nuevosAsesores)
    });
    const data = await res.json();
    if (data.ok) {
      estadoApp.asesoresData = nuevosAsesores;
      renderizarResumenAsesores();
      cerrarModalAsesores();
      mostrarToast("¡Tabla de asesores guardada exitosamente!", "success");
      cargarClientes(estadoApp.paginaActual);
    }
  } catch (err) {
    mostrarToast("Error guardando asesores: " + err.message, "error");
  }
}

// 8. CONFIGURACIÓN DEL EMISOR (SMTP / OUTLOOK)
function abrirModalConfigCorreo() {
  document.getElementById("modalConfigCorreo").classList.add("open");
}

function cerrarModalConfigCorreo() {
  document.getElementById("modalConfigCorreo").classList.remove("open");
}

function toggleCamposMetodo(metodo) {
  const camposSmtp = document.getElementById("camposSmtp");
  const camposOutlook = document.getElementById("camposOutlook");

  if (metodo === "OUTLOOK") {
    camposOutlook.style.display = "block";
    camposSmtp.style.display = "none";
  } else {
    camposOutlook.style.display = "none";
    camposSmtp.style.display = "block";
  }
}

function togglePasswordVisibilidad() {
  const el = document.getElementById("inputSmtpPassword");
  el.type = el.type === "password" ? "text" : "password";
}

async function probarConexionOutlook() {
  const msgBox = document.getElementById("statusOutlookMsg");
  msgBox.innerHTML = '<span style="color: #60a5fa;">Verificando conexión con Microsoft Outlook...</span>';

  try {
    const res = await fetch("/api/probar-outlook", { method: "POST" });
    const data = await res.json();
    if (data.ok) {
      msgBox.innerHTML = `<span style="color: #10b981;">✅ ${escapeHtml(data.mensaje)}</span>`;
      mostrarToast("Outlook conectado exitosamente", "success");
    } else {
      msgBox.innerHTML = `<span style="color: #f43f5e;">❌ ${escapeHtml(data.error)}</span>`;
      mostrarToast(data.error, "error");
    }
  } catch (err) {
    msgBox.innerHTML = `<span style="color: #f43f5e;">❌ Error: ${escapeHtml(err.message)}</span>`;
  }
}

async function abrirOutlookDesktop() {
  mostrarToast("Iniciando Microsoft Outlook en tu equipo...", "info");
  try {
    const res = await fetch("/api/abrir-outlook", { method: "POST" });
    const data = await res.json();
    if (data.ok) {
      mostrarToast(data.mensaje, "info");
      setTimeout(probarConexionOutlook, 2500);
    } else {
      mostrarToast(data.error, "warning");
    }
  } catch (err) {
    mostrarToast("Error: " + err.message, "error");
  }
}

async function probarConexionSmtp() {
  const usuario = document.getElementById("inputSmtpUsuario").value.trim();
  const password = document.getElementById("inputSmtpPassword").value;
  const host = document.getElementById("inputSmtpHost").value.trim();
  const port = document.getElementById("inputSmtpPort").value;

  if (!usuario) {
    mostrarToast("Por favor escribe tu correo corporativo emisor.", "warning");
    return;
  }

  mostrarToast("Probando conexión con el servidor SMTP...", "info");
  try {
    const res = await fetch("/api/probar-smtp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        smtp_usuario: usuario,
        smtp_password: password,
        smtp_host: host,
        smtp_port: port
      })
    });
    const data = await res.json();
    if (res.ok) {
      alert("✅ " + data.mensaje);
      mostrarToast("Conexión SMTP exitosa", "success");
    } else {
      alert("❌ " + (data.error || "Fallo de conexión"));
      mostrarToast(data.error || "Fallo de conexión", "error");
    }
  } catch (err) {
    alert("❌ Error: " + err.message);
  }
}

async function guardarConfigSmtp() {
  const metodo = document.getElementById("selectMetodoEnvio").value;
  const usuario = document.getElementById("inputSmtpUsuario").value.trim();
  const password = document.getElementById("inputSmtpPassword").value;
  const remitente = document.getElementById("inputRemitenteNombre").value.trim();
  const host = document.getElementById("inputSmtpHost").value.trim();
  const port = document.getElementById("inputSmtpPort").value;

  const datos = {
    metodo: metodo,
    smtp_usuario: usuario,
    smtp_password: password,
    remitente_nombre: remitente,
    smtp_host: host,
    smtp_port: port
  };

  try {
    const res = await fetch("/api/guardar-config-smtp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(datos)
    });
    const data = await res.json();
    if (data.ok) {
      actualizarBadgeMetodo(metodo, usuario);
      mostrarToast("¡Configuración guardada exitosamente!", "success");
      cerrarModalConfigCorreo();
    }
  } catch (err) {
    mostrarToast("Error guardando: " + err.message, "error");
  }
}

// 9. EJECUCIÓN Y MONITOREO DEL ENVÍO
function confirmarEnvioReal() {
  const seleccionados = Array.from(estadoApp.indicesSeleccionados);
  if (seleccionados.length === 0) {
    mostrarToast("Por favor selecciona al menos un cliente de la tabla antes de enviar.", "warning");
    return;
  }

  const confirmacion = confirm(`¿Estás completamente seguro de iniciar el envío masivo real a los ${seleccionados.length} destinatarios seleccionados?`);
  if (confirmacion) {
    ejecutarEnvio('REAL');
  }
}

async function ejecutarEnvio(modo) {
  let seleccionados = Array.from(estadoApp.indicesSeleccionados);

  // Si no hay seleccionados y es prueba o borradores, tomar los primeros 3 visibles o el actual de preview
  if (seleccionados.length === 0) {
    if (modo === "CORREO_TEST") {
      seleccionados = [estadoApp.clientePreviewIndex];
    } else if (modo === "BORRADORES") {
      const visiblesConCorreo = estadoApp.clientesVisibles.filter(c => c.correo_valido).map(c => c.indice_original);
      seleccionados = visiblesConCorreo.slice(0, 3);
      if (seleccionados.length === 0) {
        mostrarToast("No hay destinatarios con correo en la página actual. Selecciona clientes en la tabla.", "warning");
        return;
      }
    } else {
      mostrarToast("Por favor selecciona al menos un cliente de la tabla.", "warning");
      return;
    }
  }

  const asunto = document.getElementById("inputAsunto").value;
  const saludo = document.getElementById("inputSaludo").value;
  const cuerpo = document.getElementById("inputCuerpo").value;
  const cuerpoDestacado = document.getElementById("inputCuerpoDestacado").value;
  const cuerpoCierre = document.getElementById("inputCuerpoCierre").value;
  const correoTest = document.getElementById("inputCorreoTest").value.trim();

  if (modo === "CORREO_TEST" && (!correoTest || !correoTest.includes("@"))) {
    mostrarToast("Ingresa un correo de prueba válido en la casilla 'Enviar prueba a:'", "warning");
    document.getElementById("inputCorreoTest").focus();
    return;
  }

  // Abrir modal de progreso
  document.getElementById("modalProgreso").classList.add("open");
  document.getElementById("badgeModoEnvio").textContent = modo;
  document.getElementById("btnCancelarEnvio").style.display = "inline-block";
  document.getElementById("btnCerrarModalProgreso").style.display = "none";
  document.getElementById("progressBarFill").style.width = "0%";
  document.getElementById("txtStatProgreso").textContent = `0 / ${seleccionados.length}`;
  document.getElementById("txtStatExitosos").textContent = "0";
  document.getElementById("txtStatErrores").textContent = "0";
  document.getElementById("consoleLogs").innerHTML = "";

  try {
    const res = await fetch("/api/iniciar-envio", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        indices: seleccionados,
        modo: modo,
        asunto: asunto,
        saludo: saludo,
        cuerpo: cuerpo,
        cuerpo_destacado: cuerpoDestacado,
        cuerpo_cierre: cuerpoCierre,
        correo_test: correoTest
      })
    });
    const data = await res.json();
    if (!res.ok) {
      alert(data.error || "Error al iniciar envío");
      cerrarModalProgreso();
      return;
    }

    iniciarSondeoProgreso();

  } catch (err) {
    mostrarToast("Error de conexión: " + err.message, "error");
    cerrarModalProgreso();
  }
}

function iniciarSondeoProgreso() {
  if (estadoApp.intervaloProgreso) clearInterval(estadoApp.intervaloProgreso);

  estadoApp.intervaloProgreso = setInterval(async () => {
    try {
      const res = await fetch("/api/estado-envio");
      const data = await res.json();

      document.getElementById("txtStatProgreso").textContent = `${data.actual} / ${data.total}`;
      document.getElementById("txtStatExitosos").textContent = data.enviados;
      document.getElementById("txtStatErrores").textContent = data.errores;

      const pct = data.total > 0 ? (data.actual / data.total) * 100 : 0;
      document.getElementById("progressBarFill").style.width = pct + "%";

      const consoleBox = document.getElementById("consoleLogs");
      consoleBox.innerHTML = "";
      data.logs.forEach(log => {
        const div = document.createElement("div");
        div.className = `log-line log-${log.tipo}`;
        div.innerHTML = `<span class="log-time">[${log.hora}]</span> <span>${escapeHtml(log.texto)}</span>`;
        consoleBox.appendChild(div);
      });
      consoleBox.scrollTop = consoleBox.scrollHeight;

      if (!data.activo) {
        clearInterval(estadoApp.intervaloProgreso);
        document.getElementById("btnCancelarEnvio").style.display = "none";
        document.getElementById("btnCerrarModalProgreso").style.display = "inline-block";
        mostrarToast("¡Proceso finalizado!", "success");
        // Refrescar tabla para mostrar estados
        cargarClientes(estadoApp.paginaActual);
      }

    } catch (err) {
      console.error("Error sondeando progreso:", err);
    }
  }, 750);
}

async function cancelarEnvio() {
  await fetch("/api/cancelar-envio", { method: "POST" });
  document.getElementById("btnCancelarEnvio").textContent = "Deteniendo...";
  mostrarToast("Cancelación solicitada al servidor", "warning");
}

function cerrarModalProgreso() {
  document.getElementById("modalProgreso").classList.remove("open");
  if (estadoApp.intervaloProgreso) clearInterval(estadoApp.intervaloProgreso);
}

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
