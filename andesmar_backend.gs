/**
 * ══════════════════════════════════════════════════════════════
 * ANDESMAR CARGAS — Backend Google Apps Script
 * Base de datos compartida + Emails via Gmail
 * ══════════════════════════════════════════════════════════════
 *
 * INSTALACIÓN:
 * 1. sheets.google.com → Nuevo Sheet → copiar el ID de la URL
 * 2. script.google.com → Nuevo proyecto → pegar este código
 * 3. Reemplazar SHEET_ID con el ID del Sheet
 * 4. Ejecutar setup() UNA sola vez
 * 5. Implementar → Web App → Ejecutar como: Yo · Acceso: Cualquiera
 * 6. Copiar la URL → pegarla en el tablero (Admin → GAS URL)
 * ══════════════════════════════════════════════════════════════
 */

const SHEET_ID  = 'REEMPLAZAR_CON_ID_DEL_SHEET';
const FROM_NAME = 'Andesmar Cargas · Gestión de Proyectos';
const ADMIN_EMAIL = 'cynthiacastro@andesmar.com.ar';

// ── Nombres de hojas
const SH = {
  DB:        'DB',          // Base de datos JSON completa
  USERS:     'Usuarios',    // Lista de usuarios (para consulta rápida)
  INVITES:   'Invitaciones',
  LOG:       'Log',
  MEDICIONES:'Mediciones',
  GESTIONES: 'Gestiones',
};

// ══════════════════════════════════════════════════════════════
// ENTRY POINTS
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e?.parameter?.action || 'ping';
  try {
    if (action === 'ping')   return json({ ok:true, status:'active' });
    if (action === 'getDB')  return json(getDB());
    return json({ ok:false, error:'Unknown GET action' });
  } catch(err) {
    return json({ ok:false, error:err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || '';
    let result;

    switch(action) {
      case 'saveDB':       result = saveDB(data.db);          break;
      case 'saveUsers':    result = saveUsers(data.users);    break;
      case 'saveInvite':   result = saveInvite(data.invite);  break;
      case 'getInvite':    result = getInvite(data.token);    break;
      case 'activateInvite': result = activateInvite(data.token, data.user); break;
      case 'sendEmail':    result = sendEmail(data);          break;
      case 'sendInviteEmail': result = sendInviteEmail(data); break;
      case 'logMedicion':  result = logMedicion(data);        break;
      case 'logGestion':   result = logGestion(data);         break;
      default: result = { ok:false, error:'Unknown action: ' + action };
    }

    return json(result);
  } catch(err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return json({ ok:false, error:err.message });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
// BASE DE DATOS — Lee y escribe el JSON completo en el Sheet
// ══════════════════════════════════════════════════════════════
function getDB() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.DB);
  if (!sheet || sheet.getLastRow() < 2) return { ok:true, db:null };
  const raw = sheet.getRange(2, 1).getValue();
  try {
    return { ok:true, db: JSON.parse(raw) };
  } catch(e) {
    return { ok:true, db:null };
  }
}

function saveDB(db) {
  if (!db) return { ok:false, error:'No DB provided' };
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.DB);
  const json  = JSON.stringify(db);
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1).setValue(json);
  } else {
    sheet.getRange(2, 1).setValue(json);
  }
  // Backup timestamp
  sheet.getRange(2, 2).setValue(new Date());
  return { ok:true };
}

// ══════════════════════════════════════════════════════════════
// USUARIOS
// ══════════════════════════════════════════════════════════════
function saveUsers(users) {
  if (!users) return { ok:false };
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  // Clear and rewrite
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow()-1, 6).clearContent();
  users.forEach((u, i) => {
    sheet.getRange(i+2, 1, 1, 6).setValues([[
      u.id, u.username || u.email, u.name, u.area, u.role, u.email
    ]]);
  });
  return { ok:true };
}

// ══════════════════════════════════════════════════════════════
// INVITACIONES
// ══════════════════════════════════════════════════════════════
function saveInvite(invite) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.INVITES);
  sheet.appendRow([
    invite.token,
    invite.email,
    invite.fullName,
    invite.sector,
    invite.rol,
    new Date(invite.expiry),
    'pending',
    invite.createdBy,
    new Date()
  ]);
  return { ok:true };
}

function getInvite(token) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.INVITES);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      return {
        ok:true,
        invite: {
          token:    data[i][0],
          email:    data[i][1],
          fullName: data[i][2],
          sector:   data[i][3],
          rol:      data[i][4],
          expiry:   new Date(data[i][5]).getTime(),
          status:   data[i][6],
        }
      };
    }
  }
  return { ok:false, error:'Token not found' };
}

function activateInvite(token, user) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.INVITES);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.getRange(i+1, 7).setValue('active');
      sheet.getRange(i+1, 9).setValue(new Date());
      return { ok:true };
    }
  }
  return { ok:false, error:'Token not found' };
}

// ══════════════════════════════════════════════════════════════
// EMAILS
// ══════════════════════════════════════════════════════════════
function sendEmail(data) {
  const to      = data.to_email;
  const toName  = data.to_name  || '';
  const subject = data.subject  || 'Notificación — Andesmar Gestión de Proyectos';
  const message = data.message  || '';

  if (!to || !to.includes('@')) return { ok:false, error:'Email inválido' };

  const html = buildNotifHtml(toName, subject, message);

  GmailApp.sendEmail(to, subject, message, {
    htmlBody: html,
    name: FROM_NAME,
    replyTo: ADMIN_EMAIL
  });

  logRow(SH.LOG, [new Date(), 'EMAIL', to, subject, 'ok', '']);
  return { ok:true };
}

function sendInviteEmail(data) {
  const to      = data.to_email  || '';
  const toName  = data.to_name   || '';
  const sector  = data.sector    || '';
  const rol     = data.rol       || '';
  const link    = data.link      || '';

  if (!to || !link) return { ok:false, error:'Faltan datos' };

  const subject  = '🎉 Invitación — Andesmar Gestión de Proyectos';
  const bodyText = `Hola ${toName},\n\nFuiste invitado/a al Gestor de Proyectos de Andesmar Cargas.\n\nEmail: ${to}\nSector: ${sector}\nRol: ${rol}\n\nActivá tu cuenta:\n${link}\n\n(Link válido 7 días)\n\n— Andesmar Cargas`;

  const html = `
<div style="font-family:Arial,sans-serif;max-width:540px;margin:0 auto;background:#f0f3f8;padding:20px">
  <div style="background:#003481;border-radius:12px 12px 0 0;padding:22px 28px;text-align:center">
    <div style="color:#01feff;font-size:15px;font-weight:700;letter-spacing:.05em;text-transform:uppercase">Gestión de Proyectos</div>
    <div style="color:rgba(255,255,255,.5);font-size:11px;margin-top:3px">Andesmar Cargas</div>
  </div>
  <div style="background:white;border-radius:0 0 12px 12px;padding:28px;border:1px solid #cdd5e4;border-top:none">
    <h2 style="color:#003481;margin:0 0 16px;font-size:18px">🎉 Fuiste invitado/a</h2>
    <p style="color:#3d4a60;font-size:14px;line-height:1.6">Hola <strong>${toName}</strong>,</p>
    <p style="color:#3d4a60;font-size:14px;line-height:1.6">Te invitamos a unirte al <strong>Gestor de Proyectos de Andesmar Cargas</strong>.</p>
    <div style="background:#eef2f8;border-radius:8px;padding:14px 16px;margin:16px 0;font-size:13px;color:#1e2733">
      <div><strong>Email:</strong> ${to}</div>
      <div><strong>Sector:</strong> ${sector}</div>
      <div><strong>Rol:</strong> ${rol}</div>
    </div>
    <div style="text-align:center;margin:24px 0">
      <a href="${link}" style="background:#003481;color:white;padding:13px 30px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">Activar mi cuenta →</a>
    </div>
    <p style="color:#8a96ab;font-size:11px;text-align:center">Link válido por 7 días. Si no esperabas esta invitación, ignorá este email.</p>
  </div>
</div>`;

  GmailApp.sendEmail(to, subject, bodyText, {
    htmlBody: html,
    name: FROM_NAME,
    replyTo: ADMIN_EMAIL
  });

  logRow(SH.LOG, [new Date(), 'INVITE', to, toName, 'ok', sector+'/'+rol]);
  return { ok:true };
}

function buildNotifHtml(toName, subject, message) {
  const isVenc = subject.includes('vencida') || subject.includes('VENCIDA');
  const isProx = subject.includes('vencer')  || subject.includes('próxima');
  const color  = isVenc ? '#E05252' : isProx ? '#F47B20' : '#003481';
  const icon   = isVenc ? '⚠️' : isProx ? '🔔' : '📋';
  const msgHtml = message.replace(/\n/g, '<br>');

  return `
<div style="font-family:Arial,sans-serif;max-width:540px;margin:0 auto;background:#f0f3f8;padding:20px">
  <div style="background:#003481;border-radius:12px 12px 0 0;padding:18px 24px;display:flex;align-items:center;justify-content:space-between">
    <div style="color:#01feff;font-size:13px;font-weight:700;letter-spacing:.04em;text-transform:uppercase">Gestión de Proyectos · Andesmar</div>
    <span style="font-size:24px">${icon}</span>
  </div>
  <div style="background:white;border-radius:0 0 12px 12px;padding:24px;border:1px solid #cdd5e4;border-top:none">
    <div style="background:${color}18;border-left:4px solid ${color};border-radius:0 8px 8px 0;padding:12px 16px;margin-bottom:18px">
      <div style="font-size:14px;font-weight:700;color:${color}">${subject}</div>
    </div>
    <p style="color:#3d4a60;font-size:14px">Hola <strong>${toName}</strong>,</p>
    <div style="background:#f5f6f8;border-radius:8px;padding:14px 16px;font-size:13px;color:#1e2733;line-height:1.7">${msgHtml}</div>
    <p style="color:#8a96ab;font-size:11px;margin-top:20px;text-align:center">
      Mensaje automático del Gestor de Proyectos de Andesmar Cargas.<br>
      Consultas: <a href="mailto:${ADMIN_EMAIL}" style="color:#003481">${ADMIN_EMAIL}</a>
    </p>
  </div>
</div>`;
}

// ══════════════════════════════════════════════════════════════
// KPI LOGGING
// ══════════════════════════════════════════════════════════════
function logMedicion(data) {
  logRow(SH.MEDICIONES, [
    new Date(), data.sector||'', data.kpi||'', data.periodo||'',
    data.modo||'', data.valor||'', data.objetivo||'', data.nota||'',
    data.usuarioSector||'', data.usuarioActivo||''
  ]);
  return { ok:true };
}

function logGestion(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SH.GESTIONES);
  const id    = 'GES-' + String(sheet.getLastRow()).padStart(3,'0');
  sheet.appendRow([
    id, new Date(), data.sector||'', data.kpi||'', data.obs||'',
    data.mejora||'', data.owner||'', 'Abierta',
    data.usuarioSector||'', data.usuarioActivo||''
  ]);
  return { ok:true };
}

function logRow(sheetName, row) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.appendRow(row);
}

// ══════════════════════════════════════════════════════════════
// SETUP — ejecutar UNA sola vez
// ══════════════════════════════════════════════════════════════
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const sheets = {
    [SH.DB]:         ['JSON_DB', 'Última actualización'],
    [SH.USERS]:      ['ID','Email','Nombre','Sector','Rol','Email_2'],
    [SH.INVITES]:    ['Token','Email','Nombre','Sector','Rol','Expiry','Status','CreadoPor','FechaCreacion','FechaActivacion'],
    [SH.LOG]:        ['Fecha','Tipo','Destinatario','Asunto','Estado','Detalle'],
    [SH.MEDICIONES]: ['Fecha','Sector','KPI','Período','Modo','Valor','Objetivo','Nota','UsuarioSector','UsuarioActivo'],
    [SH.GESTIONES]:  ['ID','Fecha','Sector','KPI','Observación','Mejora','Owner','Estado','UsuarioSector','UsuarioActivo'],
  };

  Object.entries(sheets).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#003481')
        .setFontColor('white');
      sh.setFrozenRows(1);
    }
  });

  // Dropdown Estado en Gestiones col H
  const ges = ss.getSheetByName(SH.GESTIONES);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Abierta','En proceso','Cerrada'], true).build();
  ges.getRange('H2:H1000').setDataValidation(rule);

  Logger.log('✅ Setup OK — hojas creadas: ' + Object.keys(sheets).join(', '));
}

// ══════════════════════════════════════════════════════════════
// TEST
// ══════════════════════════════════════════════════════════════
function testConexion() {
  const db = getDB();
  Logger.log('DB actual: ' + JSON.stringify(db).slice(0,200));
}

function testEmail() {
  const r = sendEmail({
    to_email: ADMIN_EMAIL,
    to_name:  'Cynthia Castro',
    subject:  '🔔 Test — Gestor Andesmar OK',
    message:  'El sistema de notificaciones está funcionando correctamente.\n\nProyecto: LION 360 × Andesmar\nTarea: Verificación de conectividad\nEstado: ✅ OK'
  });
  Logger.log('Test email: ' + JSON.stringify(r));
}
