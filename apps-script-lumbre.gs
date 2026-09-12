const LUMBRE_SHEET_ID = '1nepC2ds5t6aUc3J-_QJNTqKOCtRlnZDOXfLmu2Tep88';
const FOTO_FOLDER_NAME = 'Lumbre — fotos catálogo';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'productos';
  try {
    if (action === 'productos') return jsonResponse({ ok: true, productos: getProductos_() });
    if (action === 'pedidos') return jsonResponse({ ok: true, pedidos: getPedidos_() });
    if (action === 'solicitudes') return jsonResponse({ ok: true, solicitudes: getSolicitudes_() });
    return jsonResponse({ ok: false, error: 'Acción no válida' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = data.action;
    if (action === 'guardarProducto') return jsonResponse(guardarProducto_(data));
    if (action === 'crearPedido') return jsonResponse(crearPedido_(data));
    if (action === 'actualizarEstado') return jsonResponse(actualizarEstado_(data));
    if (action === 'crearSolicitud') return jsonResponse(crearSolicitud_(data));
    return jsonResponse({ ok: false, error: 'Acción no válida' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function ss_() { return SpreadsheetApp.openById(LUMBRE_SHEET_ID); }

function getProductos_() {
  const sh = ss_().getSheetByName('Productos');
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).filter(r => r[0] && r[1]).map(r => ({
    id: Number(r[0]), name: r[1], desc: r[2] || '', price: Number(r[3]) || 0,
    stock: Number(r[4]) || 0, weight: r[5] || '', img: r[6] || '', active: r[7] !== false
  }));
}

function guardarProducto_(d) {
  if (!d.name) throw new Error('Falta el nombre del producto.');
  const sh = ss_().getSheetByName('Productos');
  const values = sh.getDataRange().getValues();
  let photoUrl = d.photoUrl || '';
  if (d.photoData) photoUrl = guardarFoto_(d.photoData, d.name);

  let id = Number(d.id || 0);
  let row = -1;
  if (id) {
    for (let i = 1; i < values.length; i++) if (Number(values[i][0]) === id) { row = i + 1; break; }
  }
  if (!id) {
    id = values.slice(1).reduce((m, r) => Math.max(m, Number(r[0]) || 0), 0) + 1;
  }
  const newRow = [id, d.name, d.desc || '', Number(d.price) || 0, Number(d.stock) || 0, d.weight || '', photoUrl, d.active !== false];
  if (row > 0) sh.getRange(row, 1, 1, 8).setValues([newRow]); else sh.appendRow(newRow);
  return { ok: true, producto: { id, name: d.name, img: photoUrl } };
}

function guardarFoto_(dataUrl, name) {
  const m = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('Formato de foto no válido.');
  const mime = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  const ext = mime.indexOf('png') > -1 ? 'png' : mime.indexOf('webp') > -1 ? 'webp' : 'jpg';
  const folders = DriveApp.getFoldersByName(FOTO_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOTO_FOLDER_NAME);
  const file = folder.createFile(Utilities.newBlob(bytes, mime, name.replace(/[^a-z0-9áéíóúñ _-]/gi, '') + '-' + Date.now() + '.' + ext));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w1400';
}

function crearPedido_(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (!d.cliente || !d.whatsapp || !Array.isArray(d.items) || !d.items.length) throw new Error('Pedido incompleto.');
    const catalogo = getProductos_();
    const items = d.items.map(i => {
      const p = catalogo.find(x => x.id === Number(i.id));
      if (!p || p.active === false) throw new Error('Uno de los productos ya no está disponible.');
      const qty = Math.max(1, Number(i.qty) || 1);
      if (qty > p.stock) throw new Error('Stock insuficiente para ' + p.name + '. Quedan ' + p.stock + '.');
      return { id: p.id, name: p.name, qty, price: p.price };
    });
    const total = items.reduce((s, i) => s + i.price * i.qty, 0);
    const id = 'LUM-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Santiago', 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random()*90+10);
    const sh = ss_().getSheetByName('Pedidos');
    sh.appendRow([new Date(), d.cliente, d.whatsapp, JSON.stringify(items), total, 'Nuevo', d.notas || '', id, false]);
    return { ok: true, id, total, estado: 'Nuevo' };
  } finally { lock.releaseLock(); }
}

function getPedidos_() {
  const sh = ss_().getSheetByName('Pedidos');
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).filter(r => r[0] || r[7]).map(r => {
    let items = [];
    try { items = JSON.parse(r[3] || '[]'); } catch (_) {}
    const resumen = Array.isArray(items) && items.length ? items.map(i => i.qty + '× ' + i.name).join(', ') : String(r[3] || '');
    return { fecha: formatDate_(r[0]), cliente: r[1] || '', whatsapp: r[2] || '', productos: resumen, items, total: Number(r[4]) || 0, estado: r[5] || 'Nuevo', notas: r[6] || '', id: r[7] || '', stockDescontado: r[8] === true };
  }).reverse();
}

function actualizarEstado_(d) {
  if (!d.id || !d.status) throw new Error('Falta pedido o estado.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = ss_().getSheetByName('Pedidos');
    const rows = sh.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) if (String(rows[i][7]) === String(d.id)) { rowIndex = i + 1; break; }
    if (rowIndex < 0) throw new Error('Pedido no encontrado.');
    const row = sh.getRange(rowIndex, 1, 1, 9).getValues()[0];
    let items = []; try { items = JSON.parse(row[3] || '[]'); } catch (_) {}
    const already = row[8] === true;

    if (d.status === 'Pagado' && !already) {
      descontarStock_(items);
      sh.getRange(rowIndex, 9).setValue(true);
    }
    if (d.status === 'Cancelado' && already) {
      devolverStock_(items);
      sh.getRange(rowIndex, 9).setValue(false);
    }
    sh.getRange(rowIndex, 6).setValue(d.status);
    return { ok: true, id: d.id, estado: d.status };
  } finally { lock.releaseLock(); }
}

function descontarStock_(items) {
  const sh = ss_().getSheetByName('Productos');
  const rows = sh.getDataRange().getValues();
  const changes = [];
  items.forEach(i => {
    let found = false;
    for (let r = 1; r < rows.length; r++) {
      if (Number(rows[r][0]) === Number(i.id)) {
        found = true;
        const actual = Number(rows[r][4]) || 0;
        if (actual < Number(i.qty)) throw new Error('No alcanza el stock de ' + rows[r][1] + '. Quedan ' + actual + '.');
        changes.push({ row: r + 1, stock: actual - Number(i.qty) });
        break;
      }
    }
    if (!found) throw new Error('Producto del pedido no encontrado.');
  });
  changes.forEach(c => sh.getRange(c.row, 5).setValue(c.stock));
}

function devolverStock_(items) {
  const sh = ss_().getSheetByName('Productos');
  const rows = sh.getDataRange().getValues();
  items.forEach(i => {
    for (let r = 1; r < rows.length; r++) if (Number(rows[r][0]) === Number(i.id)) {
      sh.getRange(r + 1, 5).setValue((Number(rows[r][4]) || 0) + Number(i.qty)); break;
    }
  });
}

function crearSolicitud_(d) {
  if (!d.cliente || !d.whatsapp || !d.idea) throw new Error('Solicitud incompleta.');
  const id = 'SOL-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Santiago', 'yyMMdd-HHmmss');
  ss_().getSheetByName('Solicitudes personalizadas').appendRow([new Date(), d.cliente, d.whatsapp, d.tipo || '', d.idea, d.fechaRequerida || '', 'Nueva', id]);
  return { ok: true, id };
}

function getSolicitudes_() {
  const sh = ss_().getSheetByName('Solicitudes personalizadas');
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).filter(r => r[0] || r[7]).map(r => ({ fecha: formatDate_(r[0]), cliente:r[1]||'', whatsapp:r[2]||'', tipo:r[3]||'', idea:r[4]||'', fechaRequerida:r[5]||'', estado:r[6]||'Nueva', id:r[7]||'' })).reverse();
}

function formatDate_(v) {
  if (!v) return '';
  try { return Utilities.formatDate(new Date(v), Session.getScriptTimeZone() || 'America/Santiago', 'dd/MM/yyyy HH:mm'); } catch (_) { return String(v); }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
