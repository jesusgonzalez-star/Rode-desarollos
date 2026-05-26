
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Ficha Personal del Trabajador')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

  const fullData = sheet.getDataRange().getValues();
  const headers = fullData[0].map(h => h.toString().trim().toUpperCase());

  const extraerColumna = (n) => {
    let idx = headers.indexOf(n.toUpperCase());
    if (idx === -1) {
      idx = headers.findIndex(h => h.includes(n.toUpperCase()));
    }
    return idx === -1 ? [] : [...new Set(fullData.slice(1).map(row => row[idx]).filter(String))].sort();
  };

  return {
    ccs: extraerColumna('CC'),
    tiposIngreso: extraerColumna('TIPO DE INGRESO'),
    jefaturas: extraerColumna('JEFATURAS'),
    ubicaciones: extraerColumna('Ubicación Física'),
    contratos: extraerColumna('Tipo de Contrato'),
    dominios: extraerColumna('Dominio de Correo Electrónico'),
    grupos: extraerColumna('Grupos de Distribución de Correo'),
    sistemas: extraerColumna('Acceso a Sistemas'),
    software: extraerColumna('Licencias de Software Específicas'),
    directorios: extraerColumna('Directorios Softland'),
    cargosExistentes: extraerColumna('CARGO'),
    perfilesSoftland: extraerColumna('Perfil Softland'),
    monitores: extraerColumna('Monitores'),
    equipamientoLista: extraerColumna('Equipamiento'), // NUEVA COLUMNA DINÁMICA
    estado: extraerColumna('Estado Civil'),
    afp: extraerColumna('AFP'),
    bancos: extraerColumna('Bancos'),
    salud: extraerColumna('Sistema de Salud'),
    motivos: extraerColumna('Motivo')
  };
}

function guardarSolicitud(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = 'RESPUESTAS FICHA';
  let s = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);

  const lastRow = s.getLastRow();
  let numeroCorrelativo = 1;
  if (lastRow > 1) {
    const ultimoFolio = s.getRange(lastRow, 1).getValue().toString();
    const match = ultimoFolio.match(/(\d+)$/);
    if (match) {
      const strNumero = match[0];
      const soloNumero = strNumero.length > 3 ? strNumero.slice(-3) : strNumero;
      numeroCorrelativo = parseInt(soloNumero, 10) + 1;
    }
  }
  
  const prefijo = "FIC-26";
  const folioFinal = prefijo + numeroCorrelativo.toString().padStart(3, '0');

  let headers = [];
  let fila = [];
  const t = Utilities.base64Encode(Math.random().toString()).substring(0, 12);

  // SIEMPRE ES FICHA PERSONAL (se eliminó evaluación por solicitud)
  headers = [
    "FOLIO", "TOKEN", "ESTADO", "FECHA REGISTRO", "NOMBRE COMPLETO", "RUT", "ESTADO CIVIL", "FECHA NACIMIENTO", "SEXO", 
    "DIRECCIÓN", "COMUNA", "CIUDAD", "EMAIL", "CELULAR", "NACIONALIDAD", 
    "TALLA POLERA", "TALLA PANTALÓN", "CALZADO", 
    "ESTUDIOS MEDIOS", "ESTUDIOS SUPERIORES", "OTROS ESTUDIOS", "CERTIFICACIONES",
    "BANCO", "TIPO CUENTA", "N° CUENTA",
    "SISTEMA SALUD", "PLAN SALUD", "AFP",
    "F_NOMBRE_1", "F_RUT_1", "F_PARENTESCO_1", "F_FECHA_1", "F_SEXO_1",
    "F_NOMBRE_2", "F_RUT_2", "F_PARENTESCO_2", "F_FECHA_2", "F_SEXO_2",
    "F_NOMBRE_3", "F_RUT_3", "F_PARENTESCO_3", "F_FECHA_3", "F_SEXO_3",
    "FAM_OBSERVACIONES",
    "E_NOMBRE_1", "E_PARENTESCO_1", "E_TELÉFONO_1",
    "E_NOMBRE_2", "E_PARENTESCO_2", "E_TELÉFONO_2",
    "ENFERMEDADES CRÓNICAS", "ALERGIAS", "TRATAMIENTO MÉDICO",
    "VÍNCULO EMPRESA", "NOMBRE VÍNCULO"
  ];
  
  fila = [
    folioFinal, t, "PENDIENTE", new Date(), 
    d.nombre1 || "N/A", d.rut || "N/A", d.estadoCivil || "N/A", d.fechaNac || "N/A", d.sexo || "N/A",
    d.direccion || "N/A", d.comuna || "N/A", d.ciudad || "N/A", d.emailPersonal || "N/A", d.telefono || "N/A", d.nacionalidad || "N/A",
    d.tallaPolera || "N/A", d.tallaPantalon || "N/A", d.calzado || "N/A",
    d.estudiosMedios || "N/A", d.estudiosSuperiores || "N/A", d.estudiosOtros || "N/A", d.certificaciones || "N/A",
    d.banco || "N/A", d.tipoCuenta || "N/A", d.numCuenta || "N/A",
    d.salud || "N/A", d.planSalud || "N/A", d.afp || "N/A",
    d.fam_nombre_1 || "N/A", d.fam_rut_1 || "N/A", d.fam_parentesco_1 || "N/A", d.fam_fecha_1 || "N/A", d.fam_sexo_1 || "N/A",
    d.fam_nombre_2 || "N/A", d.fam_rut_2 || "N/A", d.fam_parentesco_2 || "N/A", d.fam_fecha_2 || "N/A", d.fam_sexo_2 || "N/A",
    d.fam_nombre_3 || "N/A", d.fam_rut_3 || "N/A", d.fam_parentesco_3 || "N/A", d.fam_fecha_3 || "N/A", d.fam_sexo_3 || "N/A",
    d.fam_observaciones || "N/A",
    d.emg_nombre_1 || "N/A", d.emg_parentesco_1 || "N/A", d.emg_telefono_1 || "N/A",
    d.emg_nombre_2 || "N/A", d.emg_parentesco_2 || "N/A", d.emg_telefono_2 || "N/A",
    d.emgEnfermedades || "N/A", d.emgAlergias || "N/A", d.emgTratamiento || "N/A",
    d.familiarEmpresa || "N/A", d.familiarNombre || "N/A"
  ];

  if (s.getLastRow() === 0) {
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0D7DAF").setFontColor("white");
  }

  try {
    s.appendRow(fila);
    const nombreParaArchivo = (d.nombre1 || "N/A").toUpperCase().trim();
    const fechaHoy = new Date().toLocaleDateString('es-CL');
    const nombreArchivoPDF = "RELIX Ficha Personal #" + folioFinal + " " + nombreParaArchivo;

    const htmlFicha = generarHtmlFicha(d, folioFinal, t, fechaHoy);
    try {
      const folder = DriveApp.getFolderById(CARPETA_FICHAS_ID);
      const blob = Utilities.newBlob(htmlFicha, 'text/html', nombreArchivoPDF + '.html');
      folder.createFile(blob);
    } catch(eDrive) {
      Logger.log('Drive error: ' + eDrive.toString());
    }

    return { status: "OK", folio: folioFinal, token: t, fechaHoy: fechaHoy, fullData: d, nombreArchivoPDF: nombreArchivoPDF };
  } catch(e) { return {status: "ERROR", msg: e.toString()}; }
}

function generarHtmlFicha(d, folio, token, fecha) {
  const v = (val) => (val && val !== 'N/A') ? val : '';
  const fam = (n, r, p) => n ? '<div class="box" style="grid-column:span 2"><span class="lbl">NOMBRE</span><span class="val">' + n + '</span></div><div class="box"><span class="lbl">RUT</span><span class="val">' + r + '</span></div><div class="box"><span class="lbl">VÍNCULO</span><span class="val">' + p + '</span></div>' : '';
  const content = `
    <div class="st">01. IDENTIFICACIÓN TRABAJADOR</div><div class="grid">
      <div class="box" style="grid-column:span 2"><span class="lbl">NOMBRE</span><span class="val">${d.nombre1}</span></div>
      <div class="box"><span class="lbl">RUT</span><span class="val">${d.rut}</span></div>
      <div class="box"><span class="lbl">ESTADO CIVIL</span><span class="val">${d.estadoCivil}</span></div>
      <div class="box"><span class="lbl">FECHA NACIMIENTO</span><span class="val">${d.fechaNac}</span></div>
      <div class="box"><span class="lbl">SEXO</span><span class="val">${d.sexo}</span></div>
      <div class="box"><span class="lbl">NACIONALIDAD</span><span class="val">${d.nacionalidad}</span></div>
      <div class="box" style="grid-column:span 3"><span class="lbl">DIRECCIÓN</span><span class="val">${d.direccion}, ${d.comuna}, ${d.ciudad}</span></div>
      <div class="box"><span class="lbl">TELÉFONO</span><span class="val">${d.telefono}</span></div>
      <div class="box" style="grid-column:span 3"><span class="lbl">EMAIL</span><span class="val">${d.emailPersonal}</span></div>
    </div>
    <div class="st">02. TALLAS Y CALZADO</div><div class="grid">
      <div class="box"><span class="lbl">POLERA</span><span class="val">${d.tallaPolera}</span></div>
      <div class="box"><span class="lbl">PANTALÓN</span><span class="val">${d.tallaPantalon}</span></div>
      <div class="box"><span class="lbl">CALZADO</span><span class="val">${d.calzado}</span></div>
    </div>
    <div class="st">03. FORMACIÓN EDUCACIONAL</div><div class="grid">
      <div class="box" style="grid-column:span 2"><span class="lbl">ESTUDIOS MEDIOS</span><span class="val">${d.estudiosMedios}</span></div>
      <div class="box" style="grid-column:span 2"><span class="lbl">ESTUDIOS SUPERIORES</span><span class="val">${d.estudiosSuperiores}</span></div>
      <div class="box" style="grid-column:span 2"><span class="lbl">OTROS ESTUDIOS</span><span class="val">${d.estudiosOtros}</span></div>
      <div class="box" style="grid-column:span 2"><span class="lbl">CERTIFICACIONES</span><span class="val">${d.certificaciones}</span></div>
    </div>
    <div class="st">04. DATOS BANCARIOS</div><div class="grid">
      <div class="box"><span class="lbl">BANCO</span><span class="val">${d.banco}</span></div>
      <div class="box"><span class="lbl">TIPO CUENTA</span><span class="val">${d.tipoCuenta}</span></div>
      <div class="box"><span class="lbl">N° CUENTA</span><span class="val">${d.numCuenta}</span></div>
    </div>
    <div class="st">05. SALUD Y PREVISIÓN</div><div class="grid">
      <div class="box"><span class="lbl">SISTEMA SALUD</span><span class="val">${d.salud}</span></div>
      <div class="box"><span class="lbl">PLAN SALUD</span><span class="val">${d.planSalud}</span></div>
      <div class="box"><span class="lbl">AFP</span><span class="val">${d.afp}</span></div>
    </div>
    <div class="st">06. CARGAS FAMILIARES</div><div class="grid">
      ${fam(v(d.fam_nombre_1), v(d.fam_rut_1), v(d.fam_parentesco_1))}
      ${fam(v(d.fam_nombre_2), v(d.fam_rut_2), v(d.fam_parentesco_2))}
      ${fam(v(d.fam_nombre_3), v(d.fam_rut_3), v(d.fam_parentesco_3))}
      <div class="box" style="grid-column:span 4"><span class="lbl">OBSERVACIONES</span><span class="val">${d.fam_observaciones || 'SIN OBSERVACIONES'}</span></div>
    </div>
    <div class="st">07. CONTACTO DE EMERGENCIA</div><div class="grid">
      <div class="box" style="grid-column:span 2"><span class="lbl">CONTACTO 1</span><span class="val">${d.emg_nombre_1} (${d.emg_parentesco_1})</span></div>
      <div class="box"><span class="lbl">TELÉFONO</span><span class="val">${d.emg_telefono_1}</span></div>
      ${v(d.emg_nombre_2) ? '<div class="box" style="grid-column:span 2"><span class="lbl">CONTACTO 2</span><span class="val">' + d.emg_nombre_2 + ' (' + d.emg_parentesco_2 + ')</span></div><div class="box"><span class="lbl">TELÉFONO</span><span class="val">' + d.emg_telefono_2 + '</span></div>' : ''}
      <div class="box" style="grid-column:span 4"><span class="lbl">DATOS MÉDICOS</span><span class="val">${d.emgEnfermedades} / ${d.emgAlergias} / ${d.emgTratamiento}</span></div>
    </div>
    <div class="st">08. DECLARACIÓN</div>
    <div class="box"><span class="lbl">VÍNCULO CON TRABAJADORES VIGENTES</span><span class="val">${d.familiarEmpresa} ${d.familiarNombre && d.familiarNombre !== 'N/A' ? '(' + d.familiarNombre + ')' : ''}</span></div>
    <div class="cert-text">Al firmar a continuación, certifico que los datos proporcionados han sido completados correctamente y de acuerdo a la realidad, sin hacer omisión a lo solicitado en el presente formulario.</div>
    <div class="firma-box"><div style="width:250px"><div class="linea"></div><span class="val">FIRMA DEL TRABAJADOR</span></div><div class="val">FECHA: ${fecha}</div></div>`;

  return '<html><head><style>@import url("https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@400;700&display=swap");@page{margin:0;size:letter}body{font-family:"Roboto Mono",monospace;padding:30px 45px;font-size:8px;color:#1e293b;margin:0;text-transform:uppercase}.h{border-bottom:2px solid #0D7DAF;margin-bottom:20px;padding-bottom:10px;display:flex;justify-content:space-between;align-items:center}.st{background:#0D7DAF;color:white;padding:4px 10px;font-weight:bold;margin-top:10px;margin-bottom:4px;font-size:9px}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}.box{border:1px solid #cbd5e1;padding:5px;background:#f8fafc}.lbl{color:#64748b;font-size:7px;font-weight:bold;display:block}.val{font-weight:bold;font-size:9px}.firma-box{margin-top:120px;display:flex;justify-content:space-between;align-items:flex-end}.linea{border-top:1px solid black;width:180px;margin-bottom:5px}.cert-text{margin-top:40px;text-align:justify;font-size:8px;line-height:1.4;text-transform:uppercase}</style></head><body><div class="h"><div><b style="font-size:12px;color:#0D7DAF">RELIX WATER S.A.</b><br>Ficha Personal v1.0</div><div style="text-align:right"><b>FOLIO N°: ' + folio + '</b><br>TOKEN: <b>' + token + '</b><br>REGISTRO: ' + fecha + '</div></div>' + content + '</body></html>';
}

const CARPETA_FICHAS_ID = '1MYtUotz5x7ZGwpJajrZIZsXpSgAUBAJz';

function guardarEnDrive(htmlContent, filename) {
  try {
    const folder = DriveApp.getFolderById(CARPETA_FICHAS_ID);
    const blob = Utilities.newBlob(htmlContent, 'text/html', filename + '.html');
    const file = folder.createFile(blob);
    return { status: 'OK', url: file.getUrl() };
  } catch(e) {
    return { status: 'ERROR', msg: e.toString() };
  }
}
