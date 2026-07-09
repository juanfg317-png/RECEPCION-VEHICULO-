/**
 * BACKEND para "Recepción de Camionetas"
 * -----------------------------------------------------------
 * Guarda las recepciones, la flotilla y los conductores en
 * 3 hojas de un Google Sheet: "Recepciones", "Flotilla", "Conductores".
 *
 * INSTALACIÓN:
 * 1. Crea un Google Sheet nuevo (sheets.new).
 * 2. Menú Extensiones > Apps Script.
 * 3. Borra el contenido de Code.gs y pega TODO este archivo.
 * 4. Guarda (icono de disquete).
 * 5. Implementar > Nueva implementación > tipo "Aplicación web".
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 * 6. Autoriza los permisos que pida Google.
 * 7. Copia la URL que termina en /exec y pégala en el HTML
 *    en la constante GAS_URL.
 * 8. Cada vez que edites este script, vuelve a hacer
 *    "Implementar > Administrar implementaciones > Editar > Nueva versión".
 */

const SHEET_RECEPCIONES = 'Recepciones';
const SHEET_FLOTILLA = 'Flotilla';
const SHEET_CONDUCTORES = 'Conductores';

function getSheet_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if(!sheet){
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  try{
    const action = e.parameter.action;
    if(action === 'getRegistry') return jsonOut_({ registry: getRegistry_() });
    if(action === 'getDrivers') return jsonOut_({ drivers: getDrivers_() });
    return jsonOut_({ error: 'Acción no reconocida: ' + action });
  }catch(err){
    return jsonOut_({ error: String(err) });
  }
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if(action === 'saveReception') return jsonOut_(saveReception_(body));
    if(action === 'saveVehicle') return jsonOut_(saveVehicle_(body));
    if(action === 'deleteVehicle') return jsonOut_(deleteVehicle_(body));
    if(action === 'saveDriver') return jsonOut_(saveDriver_(body));
    if(action === 'deleteDriver') return jsonOut_(deleteDriver_(body));
    return jsonOut_({ error: 'Acción no reconocida: ' + action });
  }catch(err){
    return jsonOut_({ error: String(err) });
  }
}

// ---------- Recepciones ----------
function saveReception_(body){
  const sheet = getSheet_(SHEET_RECEPCIONES, [
    'Timestamp','Patente','Marca','Modelo','Año','Fecha','Hora','Km','Conductor',
    'Resultado','Observaciones','Puntos declarados (JSON)'
  ]);
  sheet.appendRow([
    new Date(), body.placas || '', body.marca || '', body.modelo || '', body.anio || '',
    body.fecha || '', body.hora || '', body.km || '', body.conductor || '',
    body.resultado || '', body.observaciones || '',
    body.puntos || '[]'
  ]);
  return { ok: true };
}

// ---------- Flotilla ----------
function getRegistry_(){
  const sheet = getSheet_(SHEET_FLOTILLA, ['Patente','Marca','Modelo','Año']);
  const rows = sheet.getDataRange().getValues();
  const registry = {};
  for(let i=1;i<rows.length;i++){
    const [placa, marca, modelo, anio] = rows[i];
    if(!placa) continue;
    registry[placa] = { marca: marca||'', modelo: modelo||'', anio: anio||'' };
  }
  return registry;
}

function saveVehicle_(body){
  const sheet = getSheet_(SHEET_FLOTILLA, ['Patente','Marca','Modelo','Año']);
  const rows = sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0] === body.placa){
      sheet.getRange(i+1,1,1,4).setValues([[body.placa, body.marca||'', body.modelo||'', body.anio||'']]);
      return { ok: true };
    }
  }
  sheet.appendRow([body.placa, body.marca||'', body.modelo||'', body.anio||'']);
  return { ok: true };
}

function deleteVehicle_(body){
  const sheet = getSheet_(SHEET_FLOTILLA, ['Patente','Marca','Modelo','Año']);
  const rows = sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0] === body.placa){ sheet.deleteRow(i+1); break; }
  }
  return { ok: true };
}

// ---------- Conductores ----------
function getDrivers_(){
  const sheet = getSheet_(SHEET_CONDUCTORES, ['Nombre']);
  const rows = sheet.getDataRange().getValues();
  const drivers = [];
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]) drivers.push(rows[i][0]);
  }
  return drivers;
}

function saveDriver_(body){
  const sheet = getSheet_(SHEET_CONDUCTORES, ['Nombre']);
  const rows = sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0] === body.nombre) return { ok: true };
  }
  sheet.appendRow([body.nombre]);
  return { ok: true };
}

function deleteDriver_(body){
  const sheet = getSheet_(SHEET_CONDUCTORES, ['Nombre']);
  const rows = sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0] === body.nombre){ sheet.deleteRow(i+1); break; }
  }
  return { ok: true };
}
