// ===== CONFIG =====
var MAIL_TO = 'technikteam@lg-n.de';
var MAIL_CC = '';
var MAIL_SUBJECT_PREFIX = 'Neue Buchung - Technik-Team';
var MIN_LEAD_DAYS = 14;
// ==================

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Buchungen') || ss.insertSheet('Buchungen');

    // --- Payload lesen: JSON ODER form-urlencoded ---
    var obj = {};
    if (e && e.postData && e.postData.type === 'application/json') {
      try { obj = JSON.parse(e.postData.contents || '{}'); } catch (_) { obj = {}; }
    } else if (e && e.parameter) {
      obj = e.parameter; // kommt bei x-www-form-urlencoded rein
    } else {
      obj = {};
    }

    // Validieren
    var val = validatePayload(obj);
    if (!val.ok) return jsonResp({ ok:false, error: val.error });

    // Kopfzeile, falls neu
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Eingang (Timestamp)',
        'Tag der Veranstaltung',
        'von',
        'bis',
        'Name der Veranstaltung',
        'Veranstaltungsort',
        'Beschreibung',
        'Personal benötigt',
        'Benötigte Technik',
        'Zusätzliche Informationen',
        'Absender Name',
        'Absender E-Mail'
      ]);
    }

    // Datensatz
    var now = new Date();
    sheet.appendRow([
      now,
      obj.eventDate || '',
      obj.timeFrom || '',
      obj.timeTo || '',
      obj.eventName || '',
      obj.location || '',
      obj.description || '',
      normYesNo(obj.staffRequired),
      obj.tech || '',
      obj.extra || '',
      obj.senderName || '',
      obj.senderEmail || ''
    ]);

    // Mail
    try {
      var html = formatMail(obj, ss.getUrl());
      var subject = MAIL_SUBJECT_PREFIX + ' - ' + (obj.eventDate || '') + ' - ' + (obj.eventName || '');
      var opts = { htmlBody: html };
      if (MAIL_CC && MAIL_CC.trim()) opts.cc = MAIL_CC.trim();
      MailApp.sendEmail(MAIL_TO, subject, 'HTML erforderlich', opts);
    } catch (mailErr) { Logger.log('Mail error: ' + mailErr); }

    return jsonResp({ ok:true });

  } catch (err) {
    return jsonResp({ ok:false, error: String(err) });
  }
}

// ---- Helpers ----
function jsonResp(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function normYesNo(v) {
  if (!v) return 'nein';
  v = String(v).toLowerCase();
  return (v === 'ja' || v === 'yes' || v === 'true' || v === '1') ? 'ja' : 'nein';
}

function validatePayload(o) {
  function str(x){ return x!==undefined && x!==null && String(x).trim()!==''; }
  var miss=[];
  if(!str(o.eventDate))  miss.push('Tag der Veranstaltung');
  if(!str(o.timeFrom))   miss.push('von');
  if(!str(o.timeTo))     miss.push('bis');
  if(!str(o.eventName))  miss.push('Name der Veranstaltung');
  if(!str(o.location))   miss.push('Veranstaltungsort');
  if(!str(o.senderName)) miss.push('Absender Name');
  if(!str(o.senderEmail)) miss.push('Absender E-Mail');
  if(miss.length) return {ok:false,error:'Pflichtfelder fehlen: '+miss.join(', ')};

  var d = new Date(o.eventDate); if (isNaN(d.getTime())) return {ok:false,error:'Ungueltiges Datum (YYYY-MM-DD).'};
  var today=new Date(); var minDate=new Date(today.getFullYear(),today.getMonth(),today.getDate()); minDate.setDate(minDate.getDate()+MIN_LEAD_DAYS);
  var eventDateOnly=new Date(d.getFullYear(),d.getMonth(),d.getDate());
  if(eventDateOnly < minDate) return {ok:false,error:'Vorlauf zu kurz. Mindestens '+MIN_LEAD_DAYS+' Tage.'};

  if(!/^\d{2}:\d{2}$/.test(String(o.timeFrom)) || !/^\d{2}:\d{2}$/.test(String(o.timeTo))) return {ok:false,error:'Zeitformat ungueltig (HH:MM).'};
  if(String(o.timeFrom) >= String(o.timeTo)) return {ok:false,error:'Die Endzeit muss nach der Startzeit liegen.'};

  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(o.senderEmail))) return {ok:false,error:'Absender E-Mail ungueltig.'};

  if (o.description && String(o.description).length > 5000) return {ok:false,error:'Beschreibung zu lang.'};
  if (o.tech && String(o.tech).length > 2000) return {ok:false,error:'Benoetigte Technik zu lang.'};
  if (o.extra && String(o.extra).length > 2000) return {ok:false,error:'Zusaetzliche Informationen zu lang.'};

  return { ok:true };
}

function formatMail(o, sheetUrl) {
  function esc(s){ s=(s||'').toString(); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  var html='';
  html+='<h2>Neue Buchung - Technik-Team</h2><ul>';
  html+='<li><b>Tag:</b> '+esc(o.eventDate)+'</li>';
  html+='<li><b>Zeit:</b> '+esc(o.timeFrom)+' - '+esc(o.timeTo)+'</li>';
  html+='<li><b>Name:</b> '+esc(o.eventName)+'</li>';
  html+='<li><b>Ort:</b> '+esc(o.location)+'</li>';
  html+='<li><b>Personal benötigt:</b> '+esc(normYesNo(o.staffRequired))+'</li>';
  html+='<li><b>Technik:</b> '+esc(o.tech)+'</li>';
  html+='<li><b>Beschreibung:</b> '+esc(o.description)+'</li>';
  html+='<li><b>Zusatzinfos:</b> '+esc(o.extra)+'</li>';
  html+='<li><b>Absender:</b> '+esc(o.senderName)+' ('+esc(o.senderEmail)+')</li>';
  html+='</ul>';
  if (sheetUrl) html+='<p><a href="'+esc(sheetUrl)+'">Zur Tabelle</a></p>';
  return html;
}
