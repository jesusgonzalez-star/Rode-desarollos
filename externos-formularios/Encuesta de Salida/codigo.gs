// Retorna la URL correcta del webapp (/exec) desde las propiedades del script
function getScriptUrl() {
  var url = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  if (!url) {
    url = ScriptApp.getService().getUrl();
  }
  return url;
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'login';

  const pages = {
    'landing':       'Landing',
    'login':         'Index',
    'encuestaSalida': 'FormEncuestaSalida'
  };

  const titulos = {
    'landing':       'Portal RR.HH. — Relix Water',
    'login':         'Acceso al Portal — Relix Water',
    'encuestaSalida': 'Encuesta de Salida'
  };

  const templateName = pages[page] || 'Index';
  const titulo       = titulos[page] || 'Portal RR.HH.';

  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    template.scriptUrl = getScriptUrl();

    return template.evaluate()
        .setTitle(titulo + ' — Relix Water')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    const errHtml = '<html><body style="font-family:monospace;padding:40px;background:#1e293b;color:#f8fafc;">' +
      '<h2 style="color:#f87171;">❌ Error al cargar: ' + templateName + '</h2>' +
      '<pre style="background:#0f172a;padding:20px;border-radius:8px;color:#fca5a5;white-space:pre-wrap;">' + err.toString() + '</pre>' +
      '<a href="?" style="color:#60a5fa;">← Volver al Portal</a>' +
      '</body></html>';
    return HtmlService.createHtmlOutput(errHtml).setTitle('Error — ' + templateName);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function obtenerDatosDesdeDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return { error: "No se encontró la hoja CONFIG" };

  let userEmail = "";
  try { userEmail = Session.getActiveUser().getEmail().toLowerCase(); } catch(e) { userEmail = "desconocido"; }

  const fullData = sheet.getDataRange().getValues();
  const headers = fullData[0].map(h => h.toString().trim().toUpperCase());
  const extraerColumna = (n) => {
    let idx = headers.indexOf(n.toUpperCase());
    if (idx === -1) idx = headers.findIndex(h => h.includes(n.toUpperCase()));
    return idx === -1 ? [] : [...new Set(fullData.slice(1).map(row => row[idx]).filter(String))].sort();
  };

  let jefaturas = [];
  try {
    let jIdx = headers.indexOf('JEFE');
    if (jIdx === -1) jIdx = headers.indexOf('REPORTA');
    if (jIdx === -1) jIdx = headers.indexOf('JEFATURAS');
    let cIdx = headers.indexOf('CARGO JEFE');
    if (cIdx === -1) cIdx = headers.indexOf('CARGO JEFATURA');
    if (jIdx !== -1) {
      jefaturas = [...new Set(fullData.slice(1).map(row => {
        const nombre = row[jIdx] ? row[jIdx].toString().trim().toUpperCase() : "";
        const cargo  = (cIdx !== -1 && row[cIdx]) ? row[cIdx].toString().trim().toUpperCase() : "";
        return (nombre && cargo) ? cargo + " - " + nombre : (nombre || cargo || "");
      }).filter(String))].sort();
    }
  } catch(e) { jefaturas = []; }

  let gerencias = extraerColumna('GERENCIA');
  let areas = extraerColumna('AREA');

  return {
    usuarioActual:    userEmail,
    nombreSolicitante: userEmail.split('@')[0].toUpperCase(),
    ccs:              extraerColumna('CENTRO').length > 0 ? extraerColumna('CENTRO') : extraerColumna('CC'),
    jefaturas:        jefaturas,
    cargosExistentes: extraerColumna('CARGO'),
    gerencias:        gerencias,
    areas:            areas
  };
}


function guardarSolicitud(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let nombreHoja = 'RESPUESTAS';

  if (d.tipoFormulario === 'ENCUESTA_SALIDA') {
    nombreHoja = 'RESPUESTAS SALIDA';
  }

  let s = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);

  let prefijo = "SAL-26";
  if (d.tipoFormulario === 'ENCUESTA_SALIDA') prefijo = "SAL-26";

  const lastRow = s.getLastRow();
  let numeroCorrelativo = 1;
  if (lastRow > 1) {
    const ultimoFolio = s.getRange(lastRow, 1).getValue().toString();
    const soloNumero = ultimoFolio.replace(prefijo, "");
    if (soloNumero) numeroCorrelativo = parseInt(soloNumero, 10) + 1;
  }
  const folioFinal = prefijo + numeroCorrelativo.toString().padStart(3, '0');

  let headers = [];
  let fila = [];
  const t = Utilities.getUuid().substring(0, 8).toUpperCase();

  if (d.tipoFormulario === 'ENCUESTA_SALIDA') {
    headers = [
      "FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "NOMBRE EMPLEADO", "RUT", "EDAD", "GÉNERO",
      "GERENCIA", "ÁREA", "JEFATURA DIRECTA", "TIEMPO PERMANENCIA", "FECHA TÉRMINO",
      "MOTIVO SALIDA", "OPORTUNIDADES CRECIMIENTO", "RECONOCIMIENTO",
      "RETROALIMENTACIÓN", "MEJORAS JEFATURA", "AMBIENTE LABORAL", "EQUILIBRIO VIDA-TRABAJO",
      "RECOMENDACIÓN", "VOLVERÍA TRABAJAR"
    ];

    fila = [
      folioFinal,
      t,
      "COMPLETADA",
      new Date(),
      d.nombreEmpleado || "N/A",
      d.rut || "N/A",
      d.edad || "N/A",
      d.genero || "N/A",
      d.gerencia || "N/A",
      d.area || "N/A",
      d.jefatura || "N/A",
      d.tiempoPermancia || "N/A",
      d.fechaTermino || "N/A",
      d.motivoSalida || "N/A",
      d.oportunidadesCrecimiento || "N/A",
      d.reconocimiento || "N/A",
      d.retroalimentacion || "N/A",
      d.mejorasJefatura || "N/A",
      d.ambienteLaboral || "N/A",
      d.equilibrioVidaTrabajo || "N/A",
      d.recomendacion || "N/A",
      d.volveriaTrabajar || "N/A"
    ];
  }

  if (s.getLastRow() === 0) {
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0D7DAF").setFontColor("white");
  }

  try {
    s.appendRow(fila);

    console.log("✅ Encuesta de salida guardada: " + folioFinal);

    let nombreParaArchivo = (d.nombreEmpleado || "N/A").toUpperCase();
    const mappingNombres = {
      'ENCUESTA_SALIDA': 'Encuesta Salida'
    };
    const tipoLabel = mappingNombres[d.tipoFormulario] || 'Solicitud';
    const nombreArchivoPDF = `RELIX ${tipoLabel} #${folioFinal} ${nombreParaArchivo.trim()}`;

    return {
      status: "OK",
      folio: folioFinal,
      token: t,
      fechaHoy: new Date().toLocaleDateString('es-CL'),
      fullData: d,
      nombreArchivoPDF: nombreArchivoPDF
    };
  } catch(e) {
    return {status: "ERROR", msg: e.toString()};
  }
}

// Plantilla HTML corporativa
function generarPlantillaCorreoRRHH(titulo, nombre, mensajePrincipal, folio, solicitante, detalle1, detalle2, piePagina) {
  const LOGO_URL = 'https://relixwater.com/wp-content/uploads/2025/10/Logo-Relix-water-150px.png';
  return `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 620px; margin: auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background-color: #0D7DAF; padding: 28px; text-align: center;">
        <img src="${LOGO_URL}" alt="ReliX Water" style="width: 140px; height: auto;">
      </div>
      <div style="padding: 32px;">
        <h2 style="color: #0D7DAF; margin-top: 0; font-size: 20px; border-bottom: 2px solid #f0f4f8; padding-bottom: 12px;">${titulo}</h2>
        <p style="font-size: 15px;">Estimado/a <strong>${nombre}</strong>,</p>
        <p style="font-size: 15px; color: #555;">${mensajePrincipal}</p>
        <div style="background-color: #f8fafc; padding: 20px; border-radius: 8px; border-left: 5px solid #f89d31; margin: 24px 0;">
          <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
            <tr><td style="padding: 6px 0; color: #777;"><strong>Folio:</strong></td><td style="font-weight: bold; color: #0D7DAF;">${folio}</td></tr>
            <tr><td style="padding: 6px 0; color: #777;"><strong>Solicitante:</strong></td><td>${solicitante}</td></tr>
            ${detalle1 ? `<tr><td style="padding: 6px 0; color: #777;"><strong>Detalle:</strong></td><td>${detalle1}</td></tr>` : ''}
            ${detalle2 ? `<tr><td style="padding: 6px 0; color: #777;"><strong>Área / C.C.:</strong></td><td>${detalle2}</td></tr>` : ''}
          </table>
        </div>
        <p style="font-size: 13px; color: #444; background: #fffdf0; padding: 12px; border: 1px solid #fde68a; border-radius: 6px;">
          <strong>NOTA:</strong> ${piePagina}
        </p>
      </div>
      <div style="background-color: #f1f5f9; padding: 14px; text-align: center; font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;">
        © ${new Date().getFullYear()} RELIX WATER — DEPTO. RR.HH. Y RECLUTAMIENTO
      </div>
    </div>
  `;
}

function validarUsuario(u, p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PERMISOS');
  if (!sheet) return { success: false, msg: "Configuración de seguridad no encontrada" };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, msg: "No hay usuarios registrados" };

  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  const uIdx = headers.indexOf('USUARIO');
  const cIdx = headers.indexOf('CLAVE');
  const rIdx = headers.indexOf('ROL');

  if (uIdx === -1 || cIdx === -1) return { success: false, msg: "Estructura de la hoja PERMISOS incorrecta" };

  const uInput = u.toString().trim().toLowerCase();
  const pInput = p.toString().trim();

  for (let i = 1; i < data.length; i++) {
    const userDB = data[i][uIdx] ? data[i][uIdx].toString().trim().toLowerCase() : "";
    const passDB = data[i][cIdx] ? data[i][cIdx].toString().trim() : "";
    const rolDB  = (rIdx !== -1 && data[i][rIdx]) ? data[i][rIdx].toString().trim().toUpperCase() : "EMPLEADO";

    if (userDB === uInput && passDB === pInput) {
      return {
        success: true,
        token: "session_" + Utilities.getUuid().substring(0,8),
        email: userDB,
        rol: rolDB
      };
    }
  }

  return { success: false, msg: "Credenciales inválidas." };
}
