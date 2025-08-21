/**************** CONFIG ****************/
const SHEET_ID = '1XRntaEWa93ilPKQwnYAB4xzosiWBBgdTaC5wobCE0yI';

const CFG = {
  TZ: 'Europe/Madrid',
  SEND_AT_HOUR: 9,                 // helper antiguo (no afecta al exacto)
  START_HOUR: 9,                   // timeline desde 09:00
  END_HOUR: 22,
  INCLUDE_ALLDAY: true,

  SHEET_EMPLEADOS: 'EMPLEADOS',
  SHEET_PRIORIDADES: 'PRIORIDADES',
  SHEET_NOTAS: 'NOTAS',

  BRAND: 'MyCityHome',
  SENDER_NAME: 'City · Agenda diaria',
  SUBJECT: 'Agenda de hoy — {{fecha}}',

  // Visual
  HERO_BG: 'https://mycityhome.es/wp-content/uploads/2023/04/cropped-Copia-de-calle-argumosa-14-scaled-3.jpg',
  LOGO_URL: 'https://mycityhome.es/wp-content/uploads/2018/02/cropped-logo-mch.png',

  C: {
    azul: '#0ea5e9',
    naranja: '#fbbf24',
    gris: '#6b7280',
    texto: '#111827',
    bg: '#f6f8fb',
    borde: '#f1e7c2',
    sombra: '0 10px 24px rgba(0,0,0,.10)'
  },

  // Dominio por defecto para resolver user -> correo (si no hay columna "usuario")
  DOMAIN_FALLBACK: 'mycityhome.es'
};

/*************** MENÚ EN LA HOJA ***************/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Agenda diaria')
    .addItem('Enviar a todos (hoy)', 'sendAllToday')
    .addItem('Enviar fila seleccionada', 'sendSelectedRow')
    .addSeparator()
    .addItem('Crear trigger 09:00 exacto', 'setupDailyTriggerExact')
    .addItem('Crear respaldo 09:05', 'setupDailyTriggerBackup')
    .addItem('Borrar todos los triggers', 'deleteAllTriggers')
    .addItem('Listar triggers (Log)', 'listTriggers')
    .addSeparator()
    .addItem('Helper antiguo 08:30', 'setupDailyTrigger') // opcional
    .addToUi();
}

function sendSelectedRow(){
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_EMPLEADOS);
  if (!sh) return SpreadsheetApp.getUi().alert('No existe la hoja "'+CFG.SHEET_EMPLEADOS+'".');
  if (SpreadsheetApp.getActiveSheet().getName() !== CFG.SHEET_EMPLEADOS){
    return SpreadsheetApp.getUi().alert('Ve a la hoja "'+CFG.SHEET_EMPLEADOS+'", selecciona una fila y vuelve a ejecutar.');
  }
  const r = sh.getActiveRange().getRow();
  if (r <= 1) return SpreadsheetApp.getUi().alert('Selecciona una fila con datos.');
  const corp = (sh.getRange(r,1).getDisplayValue() || '').trim();
  if (!corp) return SpreadsheetApp.getUi().alert('Falta email corporativo en la fila.');
  sendOne(corp);
  SpreadsheetApp.getUi().alert('Enviado a: ' + corp);
}

/*************** TRIGGERS ***************/
// EXACTO 09:00 (zona horaria del proyecto)
function setupDailyTriggerExact() {
  const fn = 'sendAllToday';
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fn)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(fn)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(0)
    .create();

  Logger.log('Trigger creado: sendAllToday @ 09:00 exacto (diario)');
}

// Respaldo 09:05
function setupDailyTriggerBackup() {
  ScriptApp.newTrigger('sendAllToday')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(5)
    .create();
  Logger.log('Trigger de respaldo: sendAllToday @ 09:05 (diario)');
}

// Helper antiguo (aprox 08:30) — opcional
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendAllToday')
    .forEach(t => ScriptApp.deleteTrigger(t));
  let t = ScriptApp.newTrigger('sendAllToday').timeBased().everyDays(1).atHour(CFG.SEND_AT_HOUR);
  if (t.nearMinute) t = t.nearMinute(30);
  t.create();
  Logger.log('Trigger diario creado (aprox ' + CFG.SEND_AT_HOUR + ':30).');
}

// Utilidades
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.length) { Logger.log('No hay triggers.'); return; }
  triggers.forEach(t => {
    Logger.log(`• fn=${t.getHandlerFunction()} — id=${t.getUniqueId()} — source=${t.getTriggerSource()}`);
  });
}
function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(ScriptApp.deleteTrigger);
  Logger.log('Todos los triggers eliminados.');
}
// Envío inmediato manual (para probar)
function sendNow() { sendAllToday(); Logger.log('sendAllToday ejecutado (manual).'); }

/*************** ENVIOS ***************/
function sendAllToday() {
  const empleados = readEmpleados_().filter(e => e.activo);
  if (!empleados.length) throw new Error('No hay filas activas en "'+CFG.SHEET_EMPLEADOS+'".');

  const now = new Date();
  const start = startOfDay_(now, CFG.TZ);
  const end   = endOfDay_(now, CFG.TZ);
  const fechaBonita = Utilities.formatDate(now, CFG.TZ, "EEEE d 'de' MMMM yyyy").replace(/^\w/, c => c.toUpperCase());
  const subject = CFG.SUBJECT.replace('{{fecha}}', Utilities.formatDate(now, CFG.TZ, 'dd/MM'));

  empleados.forEach(emp => {
    try {
      const events = getAllEventsFor_(emp, start, end); // corporativo + link (si hay)
      const html = renderPlannerHtml_({
        name: emp.nombre || guessNameFromEmail_(emp.corp_email),
        dept: emp.depto || '',
        fecha: fechaBonita,
        events,
        startHour: CFG.START_HOUR,
        endHour: CFG.END_HOUR
      });
      const plain = stripHtml_(html);
      const to = emp.personal_email || emp.corp_email;
      GmailApp.sendEmail(to, subject, plain, { htmlBody: html, name: CFG.SENDER_NAME, cc: emp.corp_email });
      Utilities.sleep(150);
    } catch (e) {
      Logger.log('Error con ' + emp.corp_email + ': ' + e);
    }
  });
}

function sendOne(corpEmail) {
  const emp = readEmpleados_().find(x => x.corp_email.toLowerCase() === String(corpEmail||'').toLowerCase());
  if (!emp) throw new Error('No encontrado en "'+CFG.SHEET_EMPLEADOS+'": ' + corpEmail);

  const now = new Date();
  const start = startOfDay_(now, CFG.TZ);
  const end   = endOfDay_(now, CFG.TZ);
  const fechaBonita = Utilities.formatDate(now, CFG.TZ, "EEEE d 'de' MMMM yyyy").replace(/^\w/, c => c.toUpperCase());
  const subject = CFG.SUBJECT.replace('{{fecha}}', Utilities.formatDate(now, CFG.TZ, 'dd/MM'));

  const events = getAllEventsFor_(emp, start, end);

  const html = renderPlannerHtml_({
    name: emp.nombre || guessNameFromEmail_(emp.corp_email),
    dept: emp.depto || '',
    fecha: fechaBonita,
    events,
    startHour: CFG.START_HOUR,
    endHour: CFG.END_HOUR
  });
  const plain = stripHtml_(html);
  const to = emp.personal_email || emp.corp_email;
  GmailApp.sendEmail(to, subject, plain, { htmlBody: html, name: CFG.SENDER_NAME, cc: emp.corp_email });
}

/*************** SHEET I/O ***************/
function readEmpleados_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(CFG.SHEET_EMPLEADOS);
  if (!sh) throw new Error('Falta la pestaña "'+CFG.SHEET_EMPLEADOS+'".');

  const vals = sh.getDataRange().getDisplayValues();
  if (vals.length <= 1) return [];

  // A: corporativo | B: LINK CALENDARIO | C: personal | D: nombre | E: departamento | F: activo | G: usuario (opcional)
  return vals.slice(1).map(r => ({
    corp_email: (r[0]||'').trim(),
    link_calendar: (r[1]||'').trim(),
    personal_email: (r[2]||'').trim(),
    nombre: (r[3]||'').trim(),
    depto: (r[4]||'').trim(),
    activo: String(r[5]||'TRUE').toUpperCase() !== 'FALSE',
    usuario: (r[6]||'').trim()
  })).filter(x => x.corp_email && x.corp_email.includes('@'));
}

/*************** CALENDARIOS (mezcla corp + link) ***************/
function getAllEventsFor_(emp, start, end){
  let all = [];

  // 1) Corporativo (siempre)
  try {
    const calCorp = CalendarApp.getCalendarById(emp.corp_email);
    if (calCorp) {
      let evs = calCorp.getEvents(start, end);
      if (!CFG.INCLUDE_ALLDAY) evs = evs.filter(e => !e.isAllDayEvent());
      all = all.concat(evs);
    } else {
      Logger.log('Sin acceso al cal corporativo: '+emp.corp_email);
    }
  } catch(e){ Logger.log('Corp cal error '+emp.corp_email+': '+e); }

  // 2) Enlace adicional (si hay)
  try {
    const id = resolveCalId_(emp.link_calendar);
    if (id) {
      const calLink = CalendarApp.getCalendarById(id);
      if (calLink) {
        let evs2 = calLink.getEvents(start, end);
        if (!CFG.INCLUDE_ALLDAY) evs2 = evs2.filter(e => !e.isAllDayEvent());
        all = all.concat(evs2);
      } else {
        Logger.log('No se encuentra calendario LINK (ID): '+id);
      }
    }
  } catch(e){ Logger.log('Link cal error '+emp.link_calendar+': '+e); }

  // dedupe por título + inicio + fin
  const seen = new Set();
  const out = [];
  for (const e of all){
    const key = [e.getTitle(), e.getStartTime().getTime(), e.getEndTime().getTime()].join('|');
    if (!seen.has(key)){ seen.add(key); out.push(e); }
  }
  return out;
}

// Convierte link con ?cid=... o ID directo en ID usable por CalendarApp
function resolveCalId_(val){
  if (!val) return '';
  const s = String(val).trim();
  if (!/^https?:/i.test(s) && /@/.test(s)) return s;      // ya es ID
  const m = s.match(/[?&]cid=([^&]+)/);
  if (m) return decodeURIComponent(m[1]);
  return ''; // si es .ics puro, suscribir antes para obtener @import.calendar...
}

/*************** ENDPOINT WEB PARA EL ERP ***************/
/*
  /exec?dept=atic            -> agenda de todo el departamento (empleados activos)
  /exec?user=jaquidn         -> agenda de un usuario (por username)
  /exec?email=xxx@...        -> agenda de un email concreto
  Opcional: &date=YYYY-MM-DD
*/
function doGet(e){
  try{
    const p = e.parameter || {};
    const dateStr = (p.date || '').trim();
    const targetDate = dateStr ? new Date(dateStr + 'T00:00:00') : new Date();
    const start = startOfDay_(targetDate, CFG.TZ);
    const end   = endOfDay_(targetDate, CFG.TZ);

    if (p.dept) {
      const deptSlug = slug_(p.dept);
      const empleadosDept = readEmpleados_().filter(x => x.activo && slug_(x.depto) === deptSlug);

      const employees = empleadosDept.map(emp => {
        const events = getAllEventsFor_(emp, start, end);
        return {
          name: emp.nombre || guessNameFromEmail_(emp.corp_email),
          email: emp.corp_email,
          dept: emp.depto || '',
          items: events.map(ev => packEventForApi_(ev))
        };
      });

      return _json({
        ok: true,
        scope: 'dept',
        dept: p.dept,
        date: Utilities.formatDate(start, CFG.TZ, 'yyyy-MM-dd'),
        tz: CFG.TZ,
        employees
      });
    }

    const raw = (p.email || p.user || '').trim();
    if (!raw) return _json({ ok:false, error:'Falta dept, user o email' }, 400);

    const corpEmail = resolveCorpEmail_(raw);
    if (!corpEmail || !/@mycityhome\.es$/i.test(corpEmail)) {
      return _json({ ok:false, error:'Usuario/email inválido' }, 400);
    }
    const emp = readEmpleados_().find(x => x.corp_email.toLowerCase() === corpEmail.toLowerCase());
    if (!emp) return _json({ ok:false, error:'Empleado no encontrado' }, 404);

    const events = getAllEventsFor_(emp, start, end);
    return _json({
      ok: true,
      scope: 'user',
      email: corpEmail,
      name: emp.nombre || guessNameFromEmail_(corpEmail),
      dept: emp.depto || '',
      date: Utilities.formatDate(start, CFG.TZ, 'yyyy-MM-dd'),
      tz: CFG.TZ,
      items: events.map(ev => packEventForApi_(ev))
    });

  } catch(err){
    return _json({ ok:false, error:String(err) }, 500);
  }
}

/*************** RENDER HTML ***************/
function renderPlannerHtml_({ name, dept, fecha, events, startHour, endHour }) {
  const grid = buildHourGrid_(events, startHour, endHour);
  const C = CFG.C;
  const font = "font-family:'Segoe UI',Arial,Helvetica,sans-serif;";
  const logo = CFG.LOGO_URL
    ? `<img src="${CFG.LOGO_URL}" alt="${CFG.BRAND}" style="height:54px;max-width:240px;border:0;display:block;margin:0 auto;">`
    : `<div style="font-weight:800;color:${C.texto};text-align:center;font-size:22px">${CFG.BRAND}</div>`;

  // pintar 09:00 -> 22:00 en orden fijo
  let leftCol = '';
  for (let h = startHour; h <= endHour; h++) {
    const hh = String(h).padStart(2,'0');
    const items = grid[hh] || [];
    const lines = items.length
      ? items.map(ev => {
          const horas = ev.allDay ? 'Todo el día' : `${ev.hi}–${ev.hf}`;
          const where = ev.where ? `<div style="color:${C.gris};font-size:12px;margin-top:2px">${escape_(ev.where)}</div>` : '';
          const descRaw = ev.desc ? cleanDesc_(ev.desc) : '';
          const desc = descRaw ? `<div style="color:${C.gris};font-size:12px;margin-top:6px">${descRaw}</div>` : '';
          const btn = addToCalendarBtn_(ev);
          return `<div style="margin:10px 0 14px 0">
                    <div style="font-weight:900;color:${C.texto};font-size:17px">${escape_(ev.title||'(sin título)')}</div>
                    <div style="color:${C.gris};font-size:13px;margin:2px 0 6px">${horas}</div>
                    ${where}${desc}${btn}
                  </div>`;
        }).join('')
      : `<div style="color:${C.gris};font-size:12px;opacity:.7">—</div>`;

    leftCol += `
      <tr>
        <td style="width:56px;padding:10px 6px 10px 12px;color:${C.azul};font-weight:800;white-space:nowrap">${hh}:00</td>
        <td style="padding:10px 14px;border-left:1px solid #eee">${lines}</td>
      </tr>`;
  }

  return `
  <!-- Fondo general -->
  <div style="
    ${font}
    background-image:linear-gradient(rgba(255,255,255,.86), rgba(255,255,255,.86)), url('${CFG.HERO_BG}');
    background-size:cover;background-position:center;background-repeat:no-repeat;
    padding:0;margin:0;">
    <!-- Cabecera/hero -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:980px;margin:0 auto;">
      <tr><td style="height:34px"></td></tr>
      <tr><td>${logo}</td></tr>
      <tr><td style="height:8px"></td></tr>
      <tr>
        <td align="left" style="padding:0 16px">
          <div style="font-size:40px;line-height:1.1;font-weight:900;color:#0b3b60;letter-spacing:-.4px">
            Hola <span style="color:${C.naranja}">${escape_(name)}</span>
          </div>
          <div style="font-size:13px;color:#0b3b60;opacity:.85;margin-top:4px">
            Esta es tu agenda de hoy para el departamento <b>${escape_(dept || '—')}</b>.
            También puedes revisarla desde tu cuenta corporativa.
          </div>
          <div style="height:14px"></div>
        </td>
      </tr>
    </table>

    <!-- Tarjeta de agenda -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" align="center" style="max-width:980px;margin:0 auto;">
      <tr><td style="height:10px"></td></tr>
      <tr>
        <td style="padding:0 16px">
          <div style="background:#fff;border:1px solid ${C.borde};border-radius:18px;box-shadow:${C.sombra};overflow:hidden;">
            <div style="background:linear-gradient(180deg,#fde68a 0%, ${C.naranja} 80%);padding:14px 16px;display:flex;align-items:center;justify-content:space-between;">
              <div style="font-weight:900;color:#0b3b60">Agenda — ${escape_(fecha)}</div>
            </div>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
              ${leftCol}
            </table>
          </div>
        </td>
      </tr>

      <!-- Mensaje motivacional -->
      <tr><td style="height:14px"></td></tr>
      <tr>
        <td style="padding:0 16px">
          <div style="background:#ffffff;border:1px solid ${C.borde};border-radius:14px;box-shadow:${C.sombra};padding:12px 16px;text-align:center;">
            <div style="font-weight:900;color:#0b3b60;margin-bottom:4px">¡Que tengas un gran día!</div>
            <div style="color:${C.gris};font-size:13px">
              Respira, prioriza y avanza paso a paso. Si algo no sale perfecto hoy, mañana saldrá mejor. 
            </div>
          </div>
        </td>
      </tr>
      <tr><td style="height:16px"></td></tr>
      <tr><td style="text-align:center;color:#6b7280;font-size:11px">Enviado automáticamente por ${CFG.SENDER_NAME}</td></tr>
      <tr><td style="height:12px"></td></tr>
    </table>
  </div>`;
}

/*************** TIMELINE ***************/
function buildHourGrid_(events, startHour, endHour){
  const tz = CFG.TZ;
  const grid = {};
  for (let h=startHour; h<=endHour; h++) grid[String(h).padStart(2,'0')] = [];

  events.forEach(e => {
    const allDay = e.isAllDayEvent();
    const s = e.getStartTime(), f = e.getEndTime();
    const hi = allDay ? '00:00' : Utilities.formatDate(s, tz, 'HH:mm');
    const hf = allDay ? '23:59' : Utilities.formatDate(f, tz, 'HH:mm');
    const hourKey = allDay ? String(startHour).padStart(2,'0') : Utilities.formatDate(s, tz, 'HH');
    const where = e.getLocation ? (e.getLocation() || '') : '';
    const desc = e.getDescription ? (e.getDescription() || '') : '';
    if (grid[hourKey]) grid[hourKey].push({
      title: e.getTitle(),
      hi, hf, where, desc,
      allDay
    });
  });

  return grid;
}

/*************** BOTÓN “AÑADIR A CALENDARIO” ***************/
function addToCalendarBtn_(ev){
  try{
    const title = encodeURIComponent(ev.title || 'Evento');
    const startISO = toGCalDateTime_(parseHH_(ev.hi));
    const endISO   = toGCalDateTime_(parseHH_(ev.hf));
    const details  = encodeURIComponent('');
    const location = encodeURIComponent(ev.where || '');
    const url = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${title}&dates=${startISO}/${endISO}&details=${details}&location=${location}`;

    return `<a href="${url}" target="_blank"
              style="display:inline-block;padding:10px 14px;border-radius:8px;background:#1d9bf0;color:#fff;text-decoration:none;font-weight:800;font-size:13px"
              onmouseover="this.style.background='${CFG.C.naranja}'"
              onmouseout="this.style.background='#1d9bf0'">
              Añadir a tu calendario
            </a>`;
  } catch(e){ return ''; }
}
function parseHH_(hhmm){
  const tz = CFG.TZ;
  const now = new Date();
  const [h,m] = hhmm.split(':').map(Number);
  const y = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const mo= Number(Utilities.formatDate(now, tz, 'MM'))-1;
  const d = Number(Utilities.formatDate(now, tz, 'dd'));
  return new Date(y,mo,d,h,m,0);
}
function toGCalDateTime_(d){
  const pad=n=>String(n).padStart(2,'0');
  return d.getUTCFullYear()+pad(d.getUTCMonth()+1)+pad(d.getUTCDate())+'T'+pad(d.getUTCHours())+pad(d.getUTCMinutes())+pad(d.getUTCSeconds())+'Z';
}

/*************** HELPERS ***************/
function cleanDesc_(s){
  if (!s) return '';
  let t = String(s);
  t = t.replace(/\r\n/g, '\n');
  t = t.replace(/<(br|br\/)>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<p[^>]*>/gi, '\n');
  t = t.replace(/<[^>]+>/g, '');
  t = t.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/_([^_]+)_/g, '<i>$1</i>');
  t = linkify_(t);
  t = shorten_(t, 360);
  t = escapeExceptTags_(t, ['b','i','a','br']);
  t = t.replace(/\n{2,}/g, '\n').replace(/\n/g, '<br>');
  return t;
}
function linkify_(txt){ return txt.replace(/https?:\/\/[^\s]+/g, m => `<a href="${m}" target="_blank">${m}</a>`); }
function escapeExceptTags_(html, allow){
  const allowRe = new RegExp(`</?(${allow.join('|')})(\\s+[^>]*)?>`,'gi');
  const placeholders = [];
  html = html.replace(allowRe, m => { placeholders.push(m); return `@@PH${placeholders.length-1}@@`; });
  html = escape_(html);
  placeholders.forEach((m, i) => { html = html.replace(`@@PH${i}@@`, m); });
  return html;
}
function startOfDay_(d, tz){ return new Date(Utilities.formatDate(d, tz, 'yyyy-MM-dd')+'T00:00:00'+tzOffset_(tz)); }
function endOfDay_(d, tz){   return new Date(Utilities.formatDate(d, tz, 'yyyy-MM-dd')+'T23:59:59'+tzOffset_(tz)); }
function tzOffset_(tz){ return Utilities.formatDate(new Date(), tz, 'XXX'); }
function guessNameFromEmail_(email){ return email.split('@')[0].replace(/[._-]+/g,' ').replace(/\b\w/g, m=>m.toUpperCase()); }
function escape_(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function stripHtml_(html){
  return String(html||'')
    .replace(/<style[\s\S]*?<\/style>/gi,'')
    .replace(/<script[\s\S]*?<\/script>/gi,'')
    .replace(/<!--[\s\S]*?-->/g,'')
    .replace(/<\/?[^>]+>/g,'')
    .replace(/&nbsp;/g,' ')
    .trim();
}
function shorten_(s, max) {
  s = String(s || '').trim();
  if (s.length <= max) return s;
  const cut = s.slice(0, max);
  const last = Math.max(cut.lastIndexOf('. '), cut.lastIndexOf(' '));
  return cut.slice(0, last > 40 ? last : max) + '…';
}
function slug_(s){ return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/g,'').trim(); }

/*************** API HELPERS ***************/
function resolveCorpEmail_(idOrEmail){
  const v = String(idOrEmail||'').trim();
  if (!v) return '';
  if (/@/.test(v)) return v.toLowerCase(); // ya es email

  const empleados = readEmpleados_();
  const byUser = empleados.find(e => (e.usuario||'').toLowerCase() === v.toLowerCase());
  if (byUser) return byUser.corp_email.toLowerCase();

  const byLocal = empleados.find(e => e.corp_email.split('@')[0].toLowerCase() === v.toLowerCase());
  if (byLocal) return byLocal.corp_email.toLowerCase();

  if (CFG.DOMAIN_FALLBACK) return `${v.toLowerCase()}@${CFG.DOMAIN_FALLBACK}`;
  return '';
}

function packEventForApi_(ev){
  const s = ev.getStartTime(), f = ev.getEndTime();
  return {
    title: ev.getTitle() || '(sin título)',
    start: Utilities.formatDate(s, CFG.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    end:   Utilities.formatDate(f, CFG.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    allDay: ev.isAllDayEvent(),
    where:  (ev.getLocation && ev.getLocation()) || '',
    desc:   (ev.getDescription && ev.getDescription()) || '',
    addUrl: _gcalAddUrl(ev.getTitle(), s, f, (ev.getLocation && ev.getLocation()) || '')
  };
}

function _gcalAddUrl(title, start, end, loc){
  const t = encodeURIComponent(title||'Evento');
  const s = toGCalDateTime_(start);
  const e = toGCalDateTime_(end);
  const l = encodeURIComponent(loc||'');
  return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${t}&dates=${s}/${e}&location=${l}`;
}

function _json(obj, status){
  const out = ContentService.createTextOutput(JSON.stringify(obj, null, 2));
  out.setMimeType(ContentService.MimeType.JSON);
  if (status) {
    // Apps Script no permite setear el status en TextOutput, pero client-side se maneja por contenido
  }
  return out;
}

/*************** API JSON PARA EL WIDGET (no rompe lo existente) ***************/
const API = {
  TOKEN: '', // p.ej. con passwordsgenerator.net
  DEFAULT_TZ: CFG.TZ || 'Europe/Madrid'
};

// GET -> https://script.google.com/macros/s/DEPLOYMENT_ID/exec?token=XXX&mode=dept&dept=Atic
// o     https://.../exec?token=XXX&mode=user&user=proyectostic@mycityhome.es
function doGet(e) {
  try {
    const q = e && e.parameter ? e.parameter : {};
    if (API.TOKEN && (!q.token || q.token !== API.TOKEN)) {
      return json_({ ok: false, error: 'invalid_token' }, 403);
    }

    const tz = API.DEFAULT_TZ;
    const today = new Date();
    const dayStart = startOfDay_(today, tz);
    const dayEnd   = endOfDay_(today, tz);
    const empleados = readEmpleados_();

    let payload;

    if ((q.mode || '').toLowerCase() === 'dept') {
      const deptKey = String(q.dept || '').trim().toLowerCase();
      if (!deptKey) return json_({ ok:false, error:'missing_dept' }, 400);

      const delDepto = empleados.filter(e => (e.depto||'').trim().toLowerCase() === deptKey);
      const users = [];

      for (const emp of delDepto) {
        const evs = getAllEventsFor_(emp, dayStart, dayEnd);
        users.push({
          name: emp.nombre || guessNameFromEmail_(emp.corp_email),
          email: emp.corp_email,
          dept: emp.depto || '',
          events: evs.map(serializeEvent_)
        });
      }

      payload = {
        ok: true,
        mode: 'dept',
        dept: deptKey,
        date: Utilities.formatDate(today, tz, 'yyyy-MM-dd'),
        tz,
        users
      };

    } else { // mode=user (por defecto)
      const key = String(q.user || '').trim();
      if (!key) return json_({ ok:false, error:'missing_user' }, 400);

      const emp = resolveUser_(empleados, key);
      if (!emp) return json_({ ok:false, error:'user_not_found', key }, 404);

      const evs = getAllEventsFor_(emp, dayStart, dayEnd);
      payload = {
        ok: true,
        mode: 'user',
        user: {
          name: emp.nombre || guessNameFromEmail_(emp.corp_email),
          email: emp.corp_email,
          dept: emp.depto || ''
        },
        date: Utilities.formatDate(today, tz, 'yyyy-MM-dd'),
        tz,
        events: evs.map(serializeEvent_)
      };
    }

    return json_(payload);
  } catch (err) {
    return json_({ ok:false, error:String(err) }, 500);
  }
}

// -------- helpers de la API --------
function json_(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  if (code) {
    // Apps Script no deja setear el status code directamente.
    // Lo devolvemos embebido en el JSON y el cliente lo valida.
    // (Para GET simples, los navegadores permiten CORS sin preflight.)
  }
  return out;
}

// Serializa un evento de CalendarApp a JSON "amigable"
function serializeEvent_(e) {
  const tz = API.DEFAULT_TZ;
  const allDay = e.isAllDayEvent();
  const start = e.getStartTime();
  const end   = e.getEndTime();
  return {
    title: e.getTitle() || '',
    startISO: start.toISOString(),
    endISO: end.toISOString(),
    startLocal: Utilities.formatDate(start, tz, "HH:mm"),
    endLocal: Utilities.formatDate(end, tz, "HH:mm"),
    where: e.getLocation ? (e.getLocation() || '') : '',
    desc: e.getDescription ? (e.getDescription() || '') : '',
    allDay
  };
}

// Busca por email completo, local-part (antes del @) o por nombre aproximado
function resolveUser_(empleados, key) {
  const k = key.toLowerCase();

  // 1) email exacto
  let emp = empleados.find(e => (e.corp_email||'').toLowerCase() === k || (e.personal_email||'').toLowerCase() === k);
  if (emp) return emp;

  // 2) local-part (antes del @)
  emp = empleados.find(e => (e.corp_email||'').split('@')[0].toLowerCase() === k);
  if (emp) return emp;

  // 3) nombre que contenga
  emp = empleados.find(e => (e.nombre||'').toLowerCase().includes(k));
  return emp || null;
}

