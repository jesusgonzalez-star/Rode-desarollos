// ═══════════════════════════════════════════════════════════════════════
// CONFIGURACIÓN CENTRAL — editar solo este bloque al instalar en nueva cuenta
// ═══════════════════════════════════════════════════════════════════════
const CONFIG = {
  DOMINIO:            '@relixwater.com',
  EMAIL_RRHH:         'rode.palma@relixwater.com',
  EMAIL_NOTIF_FIJO:   'rodepalma@gmail.com',
  CC_REC_CONT:        'nugget.drawings89@gmail.com',
  CC_RESTO:           'reejvv.draws@gmail.com',
  CLAVE_FIJA:         'Relix.9470',
  ENCUESTA_URL:       'https://REEMPLAZAR_CON_URL_DE_TU_ENCUESTA',
  FOLDER_DESCRIPTORS: 'DESCRIPTORS_RECLUTAMIENTO',
  FOLDER_CONT_FILES:  'CONTRATACION_DOCS'
};
// ═══════════════════════════════════════════════════════════════════════

// Retorna la URL correcta del webapp (/exec) desde las propiedades del script
function getScriptUrl() {
  var url = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  if (!url) {
    url = ScriptApp.getService().getUrl();
  }
  return url;
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'landing';

  // ── AUTENTICACIÓN — cuenta Google activa ─────────────────────────
  let userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(err) {}

  if (!userEmail || !userEmail.endsWith(CONFIG.DOMINIO)) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:monospace;padding:60px;background:#0D7DAF;color:white;text-align:center;">' +
      '<h2 style="font-size:18px;letter-spacing:2px;">ACCESO RESTRINGIDO</h2>' +
      '<p style="margin-top:12px;opacity:0.8;">Portal exclusivo para cuentas <b>' + CONFIG.DOMINIO + '</b>.</p>' +
      '<p style="margin-top:8px;font-size:12px;opacity:0.6;">Cuenta detectada: ' + (userEmail || 'ninguna') + '</p>' +
      '</body></html>'
    ).setTitle('Acceso Restringido — Relix Water');
  }

  // ── ROL — leer desde hoja PERMISOS ───────────────────────────────
  let userRol = 'EMPLEADO';
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PERMISOS');
    if (sheet) {
      const data    = sheet.getDataRange().getValues();
      const headers = data[0].map(h => h.toString().trim().toUpperCase());
      const uIdx    = headers.indexOf('USUARIO');
      const rIdx    = headers.indexOf('ROL');
      if (uIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const u = data[i][uIdx] ? data[i][uIdx].toString().trim().toLowerCase() : '';
          if (u === userEmail && rIdx !== -1 && data[i][rIdx]) {
            userRol = data[i][rIdx].toString().trim().toUpperCase();
            break;
          }
        }
      }
    }
  } catch(err) {}

  // ── CONTROL DE ACCESO POR PÁGINA ─────────────────────────────────
  const ACCESO = {
    'landing':       ['RRHH', 'ADMIN', 'RRLL', 'REC Y SEL', 'EMPLEADO'],
    'reclutamiento': ['RRHH', 'ADMIN', 'RRLL', 'REC Y SEL'],
    'contratacion':  ['RRHH', 'ADMIN', 'RRLL', 'REC Y SEL'],
    'modificacion':  ['RRHH', 'ADMIN', 'RRLL'],
    'permisos':      ['RRHH', 'ADMIN', 'RRLL', 'REC Y SEL', 'EMPLEADO'],
    'revocacion':    ['RRHH', 'ADMIN', 'RRLL'],
    'evaluacion':    ['RRHH', 'ADMIN', 'RRLL'],
    'horasextra':    ['RRHH', 'ADMIN', 'RRLL', 'REC Y SEL', 'EMPLEADO'],
    'finiquito':     ['RRHH', 'ADMIN', 'RRLL'],
    'configuracion': ['RRHH', 'ADMIN']
  };

  const rolesPermitidos = ACCESO[page] || ['RRHH', 'ADMIN'];
  if (!rolesPermitidos.includes(userRol)) {
    const url = getScriptUrl() + '?page=landing';
    return HtmlService.createHtmlOutput(
      '<html><head><meta http-equiv="refresh" content="0;url=' + url + '"></head><body></body></html>'
    );
  }

  // ── SERVIR TEMPLATE ───────────────────────────────────────────────
  const pages = {
    'landing':       'Landing',
    'reclutamiento': 'FormReclutamiento',
    'contratacion':  'FormContratacion',
    'modificacion':  'FormModificacion',
    'permisos':      'FormPermisos',
    'revocacion':    'FormRevocacion',
    'evaluacion':    'FormEvaluacion',
    'horasextra':    'FormHorasExtra',
    'finiquito':     'FormFiniquito',
    'configuracion': 'FormConfig'
  };

  const titulos = {
    'landing':       'Portal RR.HH. — Relix Water',
    'reclutamiento': 'Solicitud de Reclutamiento',
    'contratacion':  'Contratación de Personal',
    'modificacion':  'Modificación de Condiciones',
    'permisos':      'Solicitud de Permiso',
    'revocacion':    'Revocación Carta de Aviso',
    'evaluacion':    'Evaluación de Personal',
    'horasextra':    'Pacto de Horas Extraordinarias',
    'finiquito':     'Cálculo de Finiquito',
    'configuracion': 'Configuración del Sistema'
  };

  const templateName = pages[page] || 'Landing';
  const titulo       = titulos[page] || 'Portal RR.HH.';

  try {
    const template     = HtmlService.createTemplateFromFile(templateName);
    template.scriptUrl = getScriptUrl();
    template.userEmail = userEmail;
    template.userRol   = userRol;
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

function FORZAR_AUTORIZACION_CORREOS() {
  // Esta función es solo para forzar la ventana de permisos de Google
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Prueba de Permisos", "Tus permisos de correo ya están configurados correctamente.");
}

function AUTORIZAR_URLFETCH() {
  // Ejecutar UNA VEZ desde el editor para autorizar UrlFetchApp
  var r = UrlFetchApp.fetch('https://api.boostr.cl/holidays.json', { muteHttpExceptions: true });
  Logger.log('Status: ' + r.getResponseCode() + ' | Bytes: ' + r.getContentText().length);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function calcularSLA48Horas(fechaInicio) {
  let fecha = new Date(fechaInicio);
  const feriados = ['2026-01-01', '2026-04-03', '2026-04-04', '2026-05-01', '2026-05-21', '2026-06-29', '2026-07-16', '2026-08-15', '2026-09-18', '2026-09-19', '2026-10-12', '2026-10-31', '2026-11-01', '2026-12-08', '2026-12-25'];
  function esHabil(d) {
    const day = d.getDay();
    if (day === 0 || day === 6) return false;
    const isoDate = d.toISOString().split('T')[0];
    return !feriados.includes(isoDate);
  }
  if (fecha.getHours() >= 18) { fecha.setDate(fecha.getDate() + 1); fecha.setHours(9, 0, 0, 0); }
  let diasSumados = 0;
  while (diasSumados < 2) { fecha.setDate(fecha.getDate() + 1); if (esHabil(fecha)) diasSumados++; }
  return fecha.toLocaleDateString('es-CL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}


function obtenerDatosDesdeDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return { error: "No se encontró la hoja CONFIG" };

  // Obtener email sin depender de People API (que puede no estar habilitada)
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

  return {
    usuarioActual:    userEmail,
    nombreSolicitante: userEmail.split('@')[0].toUpperCase(),
    ccs:              extraerColumna('CENTRO').length > 0 ? extraerColumna('CENTRO') : extraerColumna('CC'),
    tiposIngreso:     extraerColumna('TIPO DE INGRESO'),
    jefaturas:        jefaturas,
    ubicaciones:      extraerColumna('Ubicación Física'),
    contratos:        extraerColumna('Tipo de Contrato'),
    cargosExistentes: extraerColumna('CARGO'),
    estado:           extraerColumna('Estado Civil'),
    motivos:          extraerColumna('Motivo'),
    motivosMod:       extraerColumna('Motivos_Cont'),
    motivosDesv:      extraerColumna('Motiv_Desv'),
    causalesIAS:      extraerColumna('IAS_Desv'),
    horarios:         extraerColumna('HORARIO'),
    afps:             extraerColumna('AFP').length > 0 ? extraerColumna('AFP') : ["CAPITAL","CUPRUM","HABITAT","MODELO","PLANVITAL","PROVIDA","UNO"]
  };
}

// ─── CONFIGURACIÓN — lectura y escritura de listas en hoja CONFIG ────────────

function obtenerConfigListas() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return { error: 'Hoja CONFIG no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());

  let ccIdx = headers.indexOf('CENTRO');
  if (ccIdx === -1) ccIdx = headers.findIndex(h => h.includes('CENTRO'));
  if (ccIdx === -1) ccIdx = headers.indexOf('CC');

  let jIdx = headers.indexOf('JEFE');
  if (jIdx === -1) jIdx = headers.indexOf('REPORTA');
  if (jIdx === -1) jIdx = headers.indexOf('JEFATURAS');
  let cIdx = headers.indexOf('CARGO JEFE');
  if (cIdx === -1) cIdx = headers.indexOf('CARGO JEFATURA');

  const ccs = ccIdx !== -1
    ? [...new Set(data.slice(1).map(r => r[ccIdx] ? r[ccIdx].toString().trim().toUpperCase() : '').filter(Boolean))].sort()
    : [];

  const jefsMap = {};
  if (jIdx !== -1) {
    data.slice(1).forEach(r => {
      const nombre = r[jIdx] ? r[jIdx].toString().trim().toUpperCase() : '';
      const cargo  = (cIdx !== -1 && r[cIdx]) ? r[cIdx].toString().trim().toUpperCase() : '';
      if (nombre) jefsMap[nombre + '||' + cargo] = { nombre, cargo };
    });
  }
  const jefaturas = Object.values(jefsMap).sort((a, b) => a.nombre.localeCompare(b.nombre));

  return { ccs, jefaturas, ccIdx, jIdx, cIdx };
}

function actualizarCC(viejo, nuevo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return { error: 'Hoja CONFIG no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  let ccIdx = headers.indexOf('CENTRO');
  if (ccIdx === -1) ccIdx = headers.findIndex(h => h.includes('CENTRO'));
  if (ccIdx === -1) ccIdx = headers.indexOf('CC');
  if (ccIdx === -1) return { error: 'Columna CC no encontrada' };

  if (!viejo) {
    // Agregar nueva fila
    const fila = new Array(headers.length).fill('');
    fila[ccIdx] = nuevo.toString().trim().toUpperCase();
    sheet.appendRow(fila);
  } else {
    const viejoVal = viejo.toString().trim().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      const val = data[i][ccIdx] ? data[i][ccIdx].toString().trim().toUpperCase() : '';
      if (val === viejoVal) {
        sheet.getRange(i + 1, ccIdx + 1).setValue(nuevo ? nuevo.toString().trim().toUpperCase() : '');
      }
    }
  }
  return { success: true };
}

function actualizarJefatura(vNombre, vCargo, nNombre, nCargo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return { error: 'Hoja CONFIG no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  let jIdx = headers.indexOf('JEFE');
  if (jIdx === -1) jIdx = headers.indexOf('REPORTA');
  if (jIdx === -1) jIdx = headers.indexOf('JEFATURAS');
  let cIdx = headers.indexOf('CARGO JEFE');
  if (cIdx === -1) cIdx = headers.indexOf('CARGO JEFATURA');
  if (jIdx === -1) return { error: 'Columna Jefatura no encontrada' };

  if (!vNombre) {
    // Agregar nueva fila
    const fila = new Array(headers.length).fill('');
    fila[jIdx] = nNombre.toString().trim().toUpperCase();
    if (cIdx !== -1) fila[cIdx] = nCargo ? nCargo.toString().trim().toUpperCase() : '';
    sheet.appendRow(fila);
  } else {
    const vN = vNombre.toString().trim().toUpperCase();
    const vC = vCargo ? vCargo.toString().trim().toUpperCase() : '';
    for (let i = 1; i < data.length; i++) {
      const rN = data[i][jIdx] ? data[i][jIdx].toString().trim().toUpperCase() : '';
      const rC = (cIdx !== -1 && data[i][cIdx]) ? data[i][cIdx].toString().trim().toUpperCase() : '';
      if (rN === vN && rC === vC) {
        sheet.getRange(i + 1, jIdx + 1).setValue(nNombre ? nNombre.toString().trim().toUpperCase() : '');
        if (cIdx !== -1) sheet.getRange(i + 1, cIdx + 1).setValue(nCargo ? nCargo.toString().trim().toUpperCase() : '');
      }
    }
  }
  return { success: true };
}

function obtenerUsuariosRoles() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PERMISOS');
  if (!sheet) return { error: 'Hoja PERMISOS no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  const uIdx    = headers.indexOf('USUARIO');
  const rIdx    = headers.indexOf('ROL');
  if (uIdx === -1) return { error: 'Columna USUARIO no encontrada' };

  const usuarios = data.slice(1)
    .map((row, i) => ({
      fila:  i + 2,
      email: row[uIdx] ? row[uIdx].toString().trim().toLowerCase() : '',
      rol:   rIdx !== -1 && row[rIdx] ? row[rIdx].toString().trim().toUpperCase() : 'EMPLEADO'
    }))
    .filter(u => u.email);

  return { usuarios, uIdx, rIdx };
}

function actualizarUsuarioRol(vEmail, nEmail, nRol) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PERMISOS');
  if (!sheet) return { error: 'Hoja PERMISOS no encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toUpperCase());
  const uIdx    = headers.indexOf('USUARIO');
  const rIdx    = headers.indexOf('ROL');
  if (uIdx === -1) return { error: 'Columna USUARIO no encontrada' };

  if (!vEmail) {
    // Agregar nuevo usuario
    const fila = new Array(headers.length).fill('');
    fila[uIdx] = nEmail.toString().trim().toLowerCase();
    if (rIdx !== -1) fila[rIdx] = nRol ? nRol.toString().trim().toUpperCase() : 'EMPLEADO';
    sheet.appendRow(fila);
  } else {
    const vE = vEmail.toString().trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const val = data[i][uIdx] ? data[i][uIdx].toString().trim().toLowerCase() : '';
      if (val === vE) {
        if (!nEmail) {
          // Eliminar — borrar fila completa
          sheet.deleteRow(i + 1);
        } else {
          sheet.getRange(i + 1, uIdx + 1).setValue(nEmail.toString().trim().toLowerCase());
          if (rIdx !== -1) sheet.getRange(i + 1, rIdx + 1).setValue(nRol ? nRol.toString().trim().toUpperCase() : 'EMPLEADO');
        }
        break;
      }
    }
  }
  return { success: true };
}

// Obtiene UF del último día del mes anterior desde mindicador.cl (open-source, sin auth)
function obtenerUFMesAnterior() {
  try {
    var hoy = new Date();
    var ultimoDiaMesAnterior = new Date(hoy.getFullYear(), hoy.getMonth(), 0);
    var yyyy = ultimoDiaMesAnterior.getFullYear();
    var mm = ('0' + (ultimoDiaMesAnterior.getMonth() + 1)).slice(-2);
    var dd = ('0' + ultimoDiaMesAnterior.getDate()).slice(-2);
    var url = 'https://mindicador.cl/api/uf/' + dd + '-' + mm + '-' + yyyy;
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return { error: 'API UF no disponible' };
    var j = JSON.parse(r.getContentText());
    var valor = (j && j.serie && j.serie[0]) ? j.serie[0].valor : null;
    if (!valor) return { error: 'Sin datos UF para ' + dd + '-' + mm + '-' + yyyy };
    return { uf: valor, fecha: dd + '-' + mm + '-' + yyyy };
  } catch (e) {
    return { error: e.toString() };
  }
}

// Obtiene feriados Chile desde date.nager.at (open-source, sin auth)
function obtenerFeriadosAPI(años) {
  var opts = { muteHttpExceptions: true };
  var fechas = [];
  if (Array.isArray(años)) {
    años.forEach(function(a) {
      try {
        var r = UrlFetchApp.fetch('https://date.nager.at/api/v3/PublicHolidays/' + a + '/CL', opts);
        if (r.getResponseCode() !== 200) return;
        var j = JSON.parse(r.getContentText());
        if (!Array.isArray(j)) return;
        j.forEach(function(item) { fechas.push(item.date.substring(0, 10)); });
      } catch(e) {}
    });
  }
  return fechas.length > 0 ? [...new Set(fechas)] : [];
}

function guardarSolicitud(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let nombreHoja = 'RESPUESTAS';
  if (d.tipoFormulario === 'RECLUTAMIENTO') nombreHoja = 'RESPUESTAS REC';
  else if (d.tipoFormulario === 'CONTRATACIÓN') nombreHoja = 'RESPUESTAS CONT';
  else if (d.tipoFormulario === 'MODIFICACIÓN DE CONDICIONES') nombreHoja = 'RESPUESTAS MOD';
  else if (d.tipoFormulario === 'SOLICITUD DE PERMISO') nombreHoja = 'RESPUESTAS SOL';
  else if (d.tipoFormulario === 'REVOCACIÓN CARTA DE AVISO') nombreHoja = 'RESPUESTAS REV';
  else if (d.tipoFormulario === 'EVALUACIÓN DE PERSONAL A DESVINCULAR') nombreHoja = 'RESPUESTAS EVAL';
  else if (d.tipoFormulario === 'PACTO HORAS EXTRAORDINARIAS') nombreHoja = 'RESPUESTAS HEX';
  else if (d.tipoFormulario === 'FINIQUITO') nombreHoja = 'RESPUESTAS FIN';

  let s = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);

  let prefijo = "REC-26";
  if (d.tipoFormulario === 'CONTRATACIÓN') prefijo = "CON-26";
  else if (d.tipoFormulario === 'MODIFICACIÓN DE CONDICIONES') prefijo = "MOD-26";
  else if (d.tipoFormulario === 'SOLICITUD DE PERMISO') prefijo = "SOL-26";
  else if (d.tipoFormulario === 'REVOCACIÓN CARTA DE AVISO') prefijo = "REV-26";
  else if (d.tipoFormulario === 'EVALUACIÓN DE PERSONAL A DESVINCULAR') prefijo = "EVA-26";
  else if (d.tipoFormulario === 'PACTO HORAS EXTRAORDINARIAS') prefijo = "HEX-26";
  else if (d.tipoFormulario === 'FINIQUITO') prefijo = "FIN-26";

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

  if (d.tipoFormulario === 'RECLUTAMIENTO') {
    if (d.fileB64) {
      const folderName = CONFIG.FOLDER_DESCRIPTORS;
      const fileUrl = (function(base64, nombre, folderName) {
        try {
          const rootFolder = DriveApp.getRootFolder();
          const folders = DriveApp.getFoldersByName(folderName);
          const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
          const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', nombre);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return file.getUrl();
        } catch (e) { return "ERROR DRIVE: " + e.toString(); }
      })(d.fileB64, d.fileName, folderName);
      d.descriptorCargo = fileUrl;
    }

    // Actualizar hoja CONFIG con nuevo cargo si corresponde
    if (d.tipoIngreso === 'NUEVO CARGO') {
      try {
        const sheetConfig = ss.getSheetByName('CONFIG');
        const configData = sheetConfig.getDataRange().getValues();
        const configHeaders = configData[0].map(h => h.toString().trim().toUpperCase());
        const cargoColIdx = configHeaders.indexOf('CARGO');
        if (cargoColIdx !== -1) {
          const listCargos = configData.slice(1).map(r => r[cargoColIdx] ? r[cargoColIdx].toString().trim().toUpperCase() : "");
          const valCargo = d.cargo ? d.cargo.toString().trim().toUpperCase() : "";
          if (valCargo && !listCargos.includes(valCargo)) {
            const newFilaRow = new Array(configHeaders.length).fill("");
            newFilaRow[cargoColIdx] = valCargo;
            sheetConfig.appendRow(newFilaRow);
          }
        }
      } catch (e) { console.error("Error en actualizar CONFIG:", e); }
    }

    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "CARGO SOLICITANTE", "C.C. SOLICITANTE", "MOTIVO CONTRATACIÓN", "CANTIDAD VACANTES", "ESTUDIOS/PROFESIÓN", "CARGO A CUBRIR", "DESCRIPTOR CARGO", "SUELDO LÍQUIDO", "COSTO EMPRESA", "CENTRO COSTO", "CARTA OFERTA", "PRESUPUESTO", "JEFE APROBADOR", "REQUIERE CORREO", "TIPO CONTRATO", "FECHA INICIO", "FECHA TÉRMINO", "RESPONSABLE", "OBSERVACIONES"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.cargoSolicitante || "N/A", d.ccSolicitante || "N/A", d.tipoIngreso || "N/A", d.cantidadVacantes || "1", d.estudiosVacante || "N/A", d.cargo || "N/A", d.descriptorCargo || "N/A", d.sueldoLiquido || "N/A", d.costoEmpresa || "N/A", d.cc || "N/A", d.cartaOferta || "N/A", d.presupuesto || "N/A", d.jefeAprobador || "N/A", d.necesitaCorreo || "N/A", d.tipoContrato || "N/A", d.fechaInicio || "N/A", d.fechaTermino || "N/A", d.responsableRec || "N/A", d.observacionesRec || "N/A"];
  } else if (d.tipoFormulario === 'CONTRATACIÓN') {
    const contFolder = CONFIG.FOLDER_CONT_FILES;
    const subirArchivoCont = function(base64, nombre, folder) {
      try {
        const folders = DriveApp.getFoldersByName(folder);
        const f = folders.hasNext() ? folders.next() : DriveApp.createFolder(folder);
        const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/octet-stream', nombre);
        const file = f.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      } catch (e) { return "ERROR DRIVE: " + e.toString(); }
    };
    if (d.fileSaludB64) d.certSaludUrl = subirArchivoCont(d.fileSaludB64, d.fileSaludName || 'cert_salud', contFolder);
    if (d.fileDiscB64)  d.certDiscUrl  = subirArchivoCont(d.fileDiscB64,  d.fileDiscName  || 'cert_discapacidad', contFolder);

    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "NOMBRE CANDIDATO", "RUT", "NACIONALIDAD", "ESTADO CIVIL", "DIRECCIÓN", "TELÉFONO", "EMAIL", "HORARIO", "FECHA INGRESO", "AFP", "SALUD", "PLAN SALUD", "TIPO CONTRATO", "FECHA TÉRMINO", "CENTRO DE COSTO", "JEFATURA DIRECTA", "SUELDO LÍQUIDO", "SOLICITUD EQUIPOS", "FAMILIAR EMPRESA", "V°B° LABORAL", "CERT. DISCAPACIDAD", "EMERGENCIA NOMBRE", "EMERGENCIA TELÉFONO", "EMERGENCIA VÍNCULO", "ENFERMEDADES CRÓNICAS", "ALERGIAS", "TRATAMIENTO MÉDICO", "CV", "CÉDULA", "CERT. AFP", "CERT. SALUD", "ANTECEDENTES", "CUENTA BANCARIA", "RESIDENCIA", "ESTUDIOS", "FINIQUITO", "FICHA PERSONAL", "OBSERVACIONES DOCUMENTOS", "ARCHIVO CERT. SALUD", "ARCHIVO CERT. DISC.", "EQUIPOS SOLICITADOS", "OBS. EQUIPOS"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.nombre1 || "N/A", d.rut || "N/A", d.nacionalidad || "N/A", d.estadoCivil || "N/A", d.direccion || "N/A", d.telefono || "N/A", d.emailPersonal || "N/A", d.horario || "N/A", d.fechaIngreso || "N/A", d.afp || "N/A", d.salud || "N/A", d.planSalud || "N/A", d.tipoContrato || "N/A", d.fechaTermino || "N/A", d.cont_cc || "N/A", d.cont_jefeDirecto || "N/A", d.contSueldoLiquido || "N/A", d.contSolEquipos || "N/A", d.familiarEmpresa || "N/A", d.vbLaboral || "N/A", d.certDiscapacidad || "N/A", d.emgNombre || "N/A", d.emgTelefono || "N/A", d.emgVinculo || "N/A", d.emgEnfermedades || "N/A", d.emgAlergias || "N/A", d.emgTratamiento || "N/A", d.docCv || "NO", d.docCedula || "NO", d.docAfp || "NO", d.docSalud || "NO", d.docAntecedentes || "NO", d.docBanco || "NO", d.docResidencia || "NO", d.docEstudios || "NO", d.docFiniquito || "NO", d.docFicha || "NO", d.obsDocs || "N/A", d.certSaludUrl || "N/A", d.certDiscUrl || "N/A", d.contEquipos || "N/A", d.contEquiposObs || "N/A"];
  } else if (d.tipoFormulario === 'MODIFICACIÓN DE CONDICIONES') {
    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "NOMBRE TRABAJADOR", "RUT", "CENTRO COSTO ACTUAL", "MOTIVO CAMBIO CARGO", "VIGENCIA DESDE", "CARGO ANTERIOR", "SUELDO L. ANTERIOR", "C.C. ANTERIOR", "JORNADA ANTERIOR", "NUEVO CARGO", "SUELDO L. NUEVO", "C.C. NUEVO", "NUEVA JORNADA", "JEFE DIRECTO", "FECHA FIRMA JEFE", "SUPERINTENDENTE", "FECHA FIRMA SUPER", "GERENTE ÁREA", "FECHA FIRMA GERENTE", "OBSERVACIONES"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.nombre1 || "N/A", d.rut || "N/A", d.modCcActual || "N/A", d.modMotivo || "N/A", d.modVigencia || "N/A", d.modCargoAnt || "N/A", d.modSueldoAnt || "N/A", d.modCcAnt || "N/A", d.modJornadaAnt || "N/A", d.modNuevoCargo || "N/A", d.modSueldoNuev || "N/A", d.modCcNuev || "N/A", d.modJornadaNuevo || "N/A", d.modJefe || "N/A", d.modFJefe || "N/A", d.modSuper || "N/A", d.modFSuper || "N/A", d.modGerente || "N/A", d.modFGerente || "N/A", d.modObservaciones || "N/A"];
  } else if (d.tipoFormulario === 'SOLICITUD DE PERMISO') {
    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "NOMBRE TRABAJADOR", "RUT", "FECHA INGRESO", "CENTRO DE COSTO", "CARGO", "JEFE DIRECTO", "DESDE", "HASTA", "MOTIVO", "OBSERVACIONES"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.nombre1 || "N/A", d.rut || "N/A", d.fechaIngresoTrabajador || "N/A", d.permCc || "N/A", d.permCargo || "N/A", d.permJefe || "N/A", d.permFechaInicio || "N/A", d.permFechaFin || "N/A", d.permMotivo || "N/A", d.permObservacion || "N/A"];
  } else if (d.tipoFormulario === 'REVOCACIÓN CARTA DE AVISO') {
    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "CARGO SOLICITANTE", "C.C. SOLICITANTE", "NOMBRE TRABAJADOR", "RUT", "FECHA CARTA ORIGINAL", "MOTIVO REVOCACIÓN"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.revCargoSol || "N/A", d.revCCSol || "N/A", d.nombre1 || "N/A", d.rut || "N/A", d.revocFechaOriginal || "N/A", d.revMotivo || "N/A"];
  } else if (d.tipoFormulario === 'PACTO HORAS EXTRAORDINARIAS') {
    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "NOMBRE TRABAJADOR", "RUT", "FECHA CONTRATO", "ÁREA", "JEFE DIRECTO", "FECHA PACTO", "HORA INICIO", "HORA FIN", "TOTAL HORAS", "DESCRIPCIÓN ACTIVIDAD"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.nombre1 || "N/A", d.rut || "N/A", d.hexFechaContrato || "N/A", d.hexArea || "N/A", d.hexJefe || "N/A", d.hexFecha || "N/A", d.hexHoraInicio || "N/A", d.hexHoraFin || "N/A", d.hexHorasExtra || "0.0", d.hexDescripcion || "N/A"];
  } else if (d.tipoFormulario === 'FINIQUITO') {
    headers = ["FOLIO","TOKEN","ESTADO","FECHA REGISTRO","SOLICITANTE","EMAIL SOLICITANTE","NOMBRE TRABAJADOR","RUT","CARGO","CENTRO COSTO","FECHA INICIO","FECHA TÉRMINO","ANTIGÜEDAD","CAUSAL","CON AVISO PREVIO","SUELDO BASE","BONOS","GRATIFICACIÓN","MOVILIZACIÓN","TOPE IMPONIBLE","BASE TOTAL","CÁLCULO DIARIO","BASE IAS C/TOPE","MESES PERÍODO","DÍAS HÁB. PROP.","VAC. USADAS","SALDO DÍAS","DÍAS INHÁBILES","TOTAL VAC. DÍAS","HAB. VACACIONES","HAB. INDEMNIZACIÓN","HAB. ACUERDO MARCO","HAB. MES AVISO","SUB-TOTAL HABERES","DESC. AFC","DESC. PRÉSTAMO","SUB-TOTAL DESCUENTOS","TOTAL FINIQUITO","SUELDO MES","TOTAL PAGAR","UF","UTM","TOPE IAS LEGAL","BASE TRIBUTABLE","FACTOR","REBAJA","IUT","ACUERDO MARCO","EMPRESA ACUERDO"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante||"N/A", d.emailSolicitante||"N/A", d.nombre1||"N/A", d.rut||"N/A", d.finCargo||"N/A", d.finCentroCosto||"N/A", d.finFechaInicio||"N/A", d.finFechaTermino||"N/A", d.finAntiguedad||"N/A", d.finCausal||"N/A", d.finConAvisoPrevio||"NO", d.finSueldoBase||"N/A", d.finBonos||"N/A", d.finGratificacion||"N/A", d.finMovilizacion||"N/A", d.finTopeImponible||"N/A", d.finBaseTotal||"N/A", d.finCalculoDiario||"N/A", d.finBaseIASConTope||"N/A", d.finMesesPeriodo||"N/A", d.finDiasHabilesProp||"N/A", d.finVacUsadas||"N/A", d.finSaldoDias||"N/A", d.finDiasInhabiles||"N/A", d.finTotalVacDias||"N/A", d.finHabVacaciones||"N/A", d.finHabIndemnizacion||"N/A", d.finHabAcuerdo||"N/A", d.finHabMesAviso||"N/A", d.finSubTotalHaberes||"N/A", d.finDescAFC||"N/A", d.finDescPrestamo||"N/A", d.finSubTotalDescuentos||"N/A", d.finTotalFiniquito||"N/A", d.finSueldoMes||"N/A", d.finTotalPagar||"N/A", d.finUF||"N/A", d.finUTM||"N/A", d.finTopeIndemnizacion||"N/A", d.finBaseTributable||"N/A", d.finFactor||"N/A", d.finRebaja||"N/A", d.finIUT||"N/A", d.finAplicaAcuerdo||"NO", d.finTipoAcuerdo||"N/A"];
  } else if (d.tipoFormulario === 'EVALUACIÓN DE PERSONAL A DESVINCULAR') {
    headers = ["FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "SOLICITANTE", "EMAIL SOLICITANTE", "NOMBRE TRABAJADOR", "RUT", "CORREO PERSONAL", "CARGO", "ÁREA/PROYECTO", "F. TÉRMINO", "MOTIVO DESVINCULACIÓN", "FUNDAMENTO", "FORMULARIO SALIDA", "PUNTUACIÓN FINAL", "PROM. PERSONAL", "PROM. PROFESIONAL", "PROM. SEGURIDAD", "PROM. CALIDAD", "NOMBRE EVALUADOR", "C.C. EVALUADOR", "CARGO EVALUADOR", "RECOMIENDA", "OBS 1", "OBS 2", "IMPLEMENTOS PENDIENTES", "DOC CONTRATO", "DOC DESCRIPTOR", "DOC ASISTENCIA", "DOC REGLAMENTO", "DOC IRL", "DOC ORIGINALES", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "C1", "C2", "C3", "C4", "C5", "C6", "D1", "D2", "D3", "D4", "D5", "D6", "D7"];
    fila = [folioFinal, t, "PENDIENTE", new Date(), d.nombreSolicitante || "N/A", d.emailSolicitante || "N/A", d.nombre1 || "N/A", d.rut || "N/A", d.correoPersonal || "N/A", d.evaCargo || "N/A", d.evaArea || "N/A", d.evaFTermino || "N/A", d.evaMotivoLista || "N/A", d.evaMotivo || "N/A", d.evaFormularioSalida || "N/A", d.evaPuntuacion || "0.0", d.promPersonal || "0.0", d.promProf || "0.0", d.promSeg || "0.0", d.promCal || "0.0", d.evaEvaluador || "N/A", d.evaCCEvaluador || "N/A", d.evaCargoEvaluador || "N/A", d.evaRecomienda || "N/A", d.evaObs1 || "N/A", d.evaObs2 || "N/A", d.evaImplementos || "N/A", d.evaDoc1 || "N/A", d.evaDoc2 || "N/A", d.evaDoc3 || "N/A", d.evaDoc4 || "N/A", d.evaDoc5 || "N/A", d.evaDoc6 || "N/A", d.asp1 || "0", d.asp2 || "0", d.asp3 || "0", d.asp4 || "0", d.asp5 || "0", d.asp6 || "0", d.asp7 || "0", d.asp8 || "0", d.asp9 || "0", d.asp10 || "0", d.asp11 || "0", d.asp12 || "0", d.asp13 || "0", d.asp14 || "0", d.prof1 || "0", d.prof2 || "0", d.prof3 || "0", d.prof4 || "0", d.prof5 || "0", d.prof6 || "0", d.prof7 || "0", d.seg1 || "0", d.seg2 || "0", d.seg3 || "0", d.seg4 || "0", d.seg5 || "0", d.seg6 || "0", d.cal1 || "0", d.cal2 || "0", d.cal3 || "0", d.cal4 || "0", d.cal5 || "0", d.cal6 || "0", d.cal7 || "0"];
  }

  if (s.getLastRow() === 0) {
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0D7DAF").setFontColor("white");
  }

  try {
    s.appendRow(fila);

    // ── ENCUESTA DE SALIDA — solo se envía cuando el motivo es RENUNCIA ──────
    const motivoEva = (d.evaMotivoLista || '').toString().toUpperCase();
    const esRenuncia = motivoEva.includes('RENUNCIA');
    if (d.tipoFormulario === 'EVALUACIÓN DE PERSONAL A DESVINCULAR' && esRenuncia) {
      const ENCUESTA_URL = CONFIG.ENCUESTA_URL;
      const emailEvaluado = (d.correoPersonal || '').toString().trim();
      if (emailEvaluado && emailEvaluado.includes('@')) {
        try {
          MailApp.sendEmail({
            to: emailEvaluado,
            subject: 'Encuesta de Salida — Relix Water',
            htmlBody: generarPlantillaCorreoRRHH(
              'Encuesta de Salida',
              d.nombre1 || 'Colaborador/a',
              'Como parte de nuestro proceso de mejora continua, te invitamos a completar la siguiente encuesta de salida. Tu opinión es muy importante para nosotros.',
              '—', d.nombre1 || 'N/A',
              '<a href="' + ENCUESTA_URL + '" style="color:#0D7DAF;">Ir a la encuesta</a>', '',
              'Esta encuesta es completamente anónima y voluntaria.'
            )
          });
        } catch(eEncuesta) { console.error('Error enviando encuesta salida: ' + eEncuesta.toString()); }
      }
    }

    const EMAIL_RRHH  = CONFIG.EMAIL_RRHH;
    const CC_REC_CONT = CONFIG.CC_REC_CONT;
    const CC_RESTO    = CONFIG.CC_RESTO;

    try {
      const tipoLabel = d.tipoFormulario || "SOLICITUD";
      const nombreDoc = d.nombreSolicitante || "N/A";
      const detalle1 = d.cargo || d.nombre1 || d.evaCargo || "N/A";
      const detalle2 = d.ccSolicitante || d.modCcActual || d.permCc || "N/A";

      const esRecOCont = (tipoLabel === "RECLUTAMIENTO" || tipoLabel === "CONTRATACIÓN");
      const ccDestinatario = esRecOCont ? CC_REC_CONT : CC_RESTO;

      const htmlNotif = generarPlantillaCorreoRRHH(
        "Nueva Solicitud Recibida",
        "Equipo RR.HH.",
        `Se ha registrado una nueva solicitud de tipo <strong>${tipoLabel}</strong> en el sistema.`,
        folioFinal, nombreDoc, detalle1, detalle2,
        "Ingresa a la planilla para revisar y actualizar el estado."
      );
      MailApp.sendEmail({
        to: EMAIL_RRHH,
        subject: `🔔 [${tipoLabel}] Nueva Solicitud: ${folioFinal}`,
        htmlBody: htmlNotif
      });

      if (d.emailSolicitante && d.emailSolicitante.includes("@") && d.emailSolicitante !== "N/A") {
        const htmlConfirm = generarPlantillaCorreoRRHH(
          "Confirmación de Solicitud",
          d.nombreSolicitante || "Solicitante",
          `Tu solicitud de tipo <strong>${tipoLabel}</strong> fue recibida con éxito bajo el folio <strong>${folioFinal}</strong>.`,
          folioFinal, nombreDoc, detalle1, detalle2,
          "Te notificaremos por este medio cuando haya actualizaciones en el estado de tu solicitud."
        );
        MailApp.sendEmail({
          to: d.emailSolicitante,
          cc: ccDestinatario,
          subject: `✅ Confirmación Solicitud RR.HH.: ${folioFinal}`,
          htmlBody: htmlConfirm
        });
      }
    } catch (err) {
      console.error("❌ Error al enviar notificación: " + err.toString());
    }

    let nombreParaArchivo = "SOLICITUD";
    if (d.tipoFormulario === 'RECLUTAMIENTO') nombreParaArchivo = (d.cargo || "N/A").toUpperCase();
    else nombreParaArchivo = (d.nombre1 || "N/A").toUpperCase();

    const mappingNombres = {
      'RECLUTAMIENTO': 'Reclutamiento',
      'CONTRATACIÓN': 'Contratación',
      'SOLICITUD DE PERMISO': 'Permiso',
      'REVOCACIÓN CARTA DE AVISO': 'Revocación',
      'EVALUACIÓN DE PERSONAL A DESVINCULAR': 'Evaluación',
      'MODIFICACIÓN DE CONDICIONES': 'Modificación',
      'PACTO HORAS EXTRAORDINARIAS': 'HorasExtra',
      'FINIQUITO': 'Finiquito'
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
  } catch(e) { return {status: "ERROR", msg: e.toString()}; }
}

//-------------
// FUNCIÓN PARA CREAR EL TRIGGER INSTALABLE AUTOMÁTICAMENTE
// Ejecuta esta función UNA VEZ desde el editor de Apps Script
//-------------
function SETUP_TRIGGER() {
  // Eliminar triggers duplicados de esta función
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarNotificaciones') {
      ScriptApp.deleteTrigger(t);
      console.log('🗑️ Trigger anterior eliminado.');
    }
  });

  // Crear trigger instalable onEdit
  ScriptApp.newTrigger('enviarNotificaciones')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  console.log('✅ Trigger instalable creado exitosamente para enviarNotificaciones.');
}


function enviarNotificaciones(e) {
  if (!e || !e.source) {
    console.error("Ejecución manual detectada. Para probar, modifique una celda real en la hoja.");
    return;
  }

  const sheet  = e.source.getActiveSheet();
  const nombre = sheet.getName();

  const HOJAS_ACTIVAS = ["RESPUESTAS REC", "RESPUESTAS CONT", "RESPUESTAS SOL", "RESPUESTAS HEX"];
  if (!HOJAS_ACTIVAS.includes(nombre)) return;

  const range   = e.range;
  const fila    = range.getRow();
  const columna = range.getColumn();
  if (fila <= 1) return;

  const valorNuevo   = range.getValue()  ? String(range.getValue()).trim().toUpperCase()  : "";
  const valorAnterior = e.oldValue       ? String(e.oldValue).trim().toUpperCase()        : "";

  const EMAIL_NOTIFICACION_FIJO = CONFIG.EMAIL_NOTIF_FIJO;
  const COLUMNA_FOLIO  = 1;
  const COLUMNA_ESTADO = 3;

  // Detectar columna email desde headers de la hoja
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toUpperCase());
  let emailColIdx = headers.indexOf('EMAIL SOLICITANTE');
  if (emailColIdx === -1) emailColIdx = headers.indexOf('EMAIL');
  const COLUMNA_EMAIL = emailColIdx !== -1 ? emailColIdx + 1 : -1;

  if (columna === COLUMNA_FOLIO && valorNuevo && !valorAnterior) {
    try {
      MailApp.sendEmail({
        to: EMAIL_NOTIFICACION_FIJO,
        subject: "🔔 Nueva Solicitud Manual — " + nombre,
        body: "Se ha ingresado un nuevo folio manualmente.\n\nHoja: " + nombre + "\nFolio: " + valorNuevo
      });
    } catch (err) {
      console.error("❌ Error enviando correo de folio manual: " + err.toString());
    }
  }

  if (columna === COLUMNA_ESTADO) {
    const estadosValidos = ["EN PROCESO", "PROCESO", "TERMINADO", "RECHAZADO"];
    if (!estadosValidos.includes(valorNuevo)) return;
    if (COLUMNA_EMAIL === -1) { console.error("⚠️ Hoja sin columna email: " + nombre); return; }

    const emailSolicitante = sheet.getRange(fila, COLUMNA_EMAIL).getValue().toString().trim();
    const folio            = sheet.getRange(fila, COLUMNA_FOLIO).getValue();
    const nombreSol        = sheet.getRange(fila, 5).getValue();

    if (!emailSolicitante || emailSolicitante === "N/A" || !emailSolicitante.includes("@")) {
      console.error("⚠️ Email inválido → '" + emailSolicitante + "'");
      return;
    }

    try {
      const colorEstado = valorNuevo === "RECHAZADO" ? "#e53e3e" : valorNuevo === "TERMINADO" ? "#10B981" : "#0D7DAF";
      const htmlEstado  = generarPlantillaCorreoRRHH(
        "Actualización de tu Solicitud", nombreSol,
        `El estado de tu solicitud ha sido actualizado a: <strong style="color:${colorEstado};">${valorNuevo}</strong>.`,
        folio, nombreSol, "Ver detalles en la planilla", "",
        "Este es un aviso automático del Departamento RR.HH. - Relix Water."
      );
      MailApp.sendEmail({ to: emailSolicitante, subject: `📋 Actualización Solicitud RR.HH.: ${folio}`, htmlBody: htmlEstado });
      console.log("✅ Correo enviado a " + emailSolicitante);
    } catch (err) {
      console.error("❌ Error al enviar correo: " + err.toString());
    }
  }
}

// ─────────────────────────────────────────────────
// PLANTILLA HTML CORPORATIVA PARA TODOS LOS CORREOS
// ─────────────────────────────────────────────────
function generarPlantillaCorreoRRHH(titulo, nombre, mensajePrincipal, folio, solicitante, detalle1, detalle2, piePagina) {
  const LOGO_URL = 'https://relixwater.com/wp-content/uploads/2025/10/Logo-Relix-water-150px.png';
  return `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 620px; margin: auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background-color: #ffffff; padding: 28px; text-align: center; border-bottom: 4px solid #0D7DAF;">
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

// --- FUNCIONES DE SESIÓN Y SEGURIDAD ---

/**
 * Valida las credenciales del usuario (Versión simplificada)
 * @param {string} u Usuario/Email
 * @param {string} p Contraseña
 * @return {object} Resultado de la validación
 */
function validarUsuario(u, p) {
  const DOMINIO    = CONFIG.DOMINIO;
  const CLAVE_FIJA = CONFIG.CLAVE_FIJA;

  const email = u.toString().trim().toLowerCase();
  const clave = p.toString().trim();

  if (!email.endsWith(DOMINIO))
    return { success: false, msg: "Acceso solo para cuentas @relixwater.com" };

  if (clave !== CLAVE_FIJA)
    return { success: false, msg: "Clave incorrecta." };

  // Buscar rol en hoja PERMISOS — si no existe, EMPLEADO
  let rol = 'EMPLEADO';
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PERMISOS');
    if (sheet) {
      const data    = sheet.getDataRange().getValues();
      const headers = data[0].map(h => h.toString().trim().toUpperCase());
      const uIdx    = headers.indexOf('USUARIO');
      const rIdx    = headers.indexOf('ROL');
      if (uIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const userDB = data[i][uIdx] ? data[i][uIdx].toString().trim().toLowerCase() : "";
          if (userDB === email && rIdx !== -1 && data[i][rIdx]) {
            rol = data[i][rIdx].toString().trim().toUpperCase();
            break;
          }
        }
      }
    }
  } catch(e) {}

  return {
    success: true,
    token: "session_" + Utilities.getUuid().substring(0, 8),
    email: email,
    rol: rol
  };
}
function DIAGNOSTICO_FERIADOS() {
  var opts = {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'application/json, text/plain, */*',
      'Accept-Language': 'es-CL,es;q=0.9'
    }
  };
  var resp = UrlFetchApp.fetch('https://api.boostr.cl/holidays.json', opts);
  Logger.log('Código: ' + resp.getResponseCode());
  Logger.log('Contenido: ' + resp.getContentText().substring(0, 500));
}
