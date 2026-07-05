// ============================================================
//  Quinta da Aldeia — Google Apps Script (Full Backend)
//  Handles: Bookings, Payments, Finance, Checklist Responses,
//           Templates, Message History, Anniversary & Instalment automation
//  Deploy as Web App — Execute as Me — Anyone can access
// ============================================================

// ---- CONFIGURATION ----
var SENDER_EMAIL  = 'quintadaaldeia@gmail.com';
var DISCOUNT      = '15%';
var BOOKING_URL   = 'https://yoursite.netlify.app';
var DAYS_BEFORE   = 7;
// -----------------------

var SHEETS = {
  bookings:          'Bookings',
  payments:          'Payments',
  finance:           'Finance',
  messageHistory:    'MessageHistory',
  checklistResponses:'ChecklistResponses',
  templates:         'Templates',
  instalments:       'Instalments',
};

// ============================================================
//  ROUTER
// ============================================================

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var result;
  try {
    var action = e.parameter.action;
    var sheet  = e.parameter.sheet || 'bookings';
    switch(action) {
      case 'getAll':   result = getAll(sheet);               break;
      case 'add':      result = addRow(sheet, e.parameter);  break;
      case 'update':   result = updateRow(sheet, e.parameter); break;
      case 'delete':       result = deleteRow(sheet, e.parameter); break;
      case 'addBooking':   result = addBookingWithConfirmation(e.parameter); break;
      case 'submitQuote':  result = submitQuote(e.parameter); break;
      case 'getById':  result = getById(sheet, e.parameter.id); break;
      default:         result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }
  var output = ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
//  SHEET DEFINITIONS — columns per sheet
// ============================================================

var SCHEMA = {
  bookings:          ['id','name','type','date','guests','contact','status','notes'],
  payments:          ['id','booking_id','name','type','event_date','total','email','mode','dep_status','dep_due','fin_status','fin_due'],
  instalments:       ['id','booking_id','customer_name','email','instalment_num','due_date','amount','paid'],
  finance:           ['id','type','cat','desc','amount','date','notes'],
  messageHistory:    ['date_sent','couple_name','email','message_type','channel','status'],
  checklistResponses:['id','submitted_at','noiva_nome','noiva_tel','noivo_nome','noivo_tel','email','num_pessoas',
                      'data_chegada','hora_chegada','data_saida','hora_saida','hora_inicio','hora_termino',
                      'local_altar','local_buffet','local_toldos','musica_info',
                      'noiva_casa','hora_banheira','rec_noiva','noivo_casa','rec_noivo',
                      'pessoas_autorizadas','pessoa_inspecao','rita','manoel','rental_items','obs','booking_id'],
  templates:         ['id','channel','label','subject','body','updated_at'],
};

// ============================================================
//  GENERIC CRUD
// ============================================================

function getOrCreateSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var cols = SCHEMA[Object.keys(SHEETS).find(k => SHEETS[k] === name)] || [];
    if (cols.length) {
      sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
      sheet.getRange(1, 1, 1, cols.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function sheetKey(sheetName) {
  return Object.keys(SHEETS).find(k => SHEETS[k] === sheetName) || sheetName;
}

function getAll(sheetName) {
  var sheet = getOrCreateSheet(SHEETS[sheetName] || sheetName);
  var data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { rows: [], sheet: sheetName };
  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[h] = val;
    });
    return obj;
  }).filter(function(r) { return r[headers[0]]; }); // filter empty rows
  return { rows: rows, sheet: sheetName };
}

function addRow(sheetName, params) {
  var sheet   = getOrCreateSheet(SHEETS[sheetName] || sheetName);
  var key     = sheetKey(sheetName);
  var cols    = SCHEMA[key] || [];
  var id      = params.id || (sheetName + '_' + new Date().getTime());
  var row     = cols.map(function(col) {
    if (col === 'id') return id;
    if (col === 'submitted_at' || col === 'updated_at' || col === 'date_sent') {
      return params[col] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }
    return params[col] !== undefined ? params[col] : '';
  });
  sheet.appendRow(row);
  var result = {};
  cols.forEach(function(col, i) { result[col] = row[i]; });
  return { success: true, row: result, sheet: sheetName };
}

function updateRow(sheetName, params) {
  var sheet   = getOrCreateSheet(SHEETS[sheetName] || sheetName);
  var key     = sheetKey(sheetName);
  var cols    = SCHEMA[key] || [];
  var data    = sheet.getDataRange().getValues();
  var id      = String(params.id);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      var row = cols.map(function(col) {
        if (col === 'id') return id;
        if (col === 'updated_at') return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        return params[col] !== undefined ? params[col] : data[i][cols.indexOf(col)];
      });
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true, sheet: sheetName };
    }
  }
  return { error: 'Row not found: ' + id };
}

function deleteRow(sheetName, params) {
  var sheet = getOrCreateSheet(SHEETS[sheetName] || sheetName);
  var data  = sheet.getDataRange().getValues();
  var id    = String(params.id);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { success: true, sheet: sheetName };
    }
  }
  return { error: 'Row not found: ' + id };
}

function getById(sheetName, id) {
  var result = getAll(sheetName);
  var row    = (result.rows || []).find(function(r) { return String(r.id) === String(id); });
  return row ? { row: row } : { error: 'Not found' };
}

// ============================================================
//  ANNIVERSARY AUTOMATION  (daily trigger — 8am)
// ============================================================

function checkAnniversaries() {
  var result   = getAll('bookings');
  var weddings = (result.rows || []).filter(function(b) {
    return String(b.type).toLowerCase() === 'wedding' &&
           String(b.status).toLowerCase() === 'confirmed' &&
           b.contact && String(b.contact).includes('@') && b.date;
  });

  var today = new Date(); today.setHours(0,0,0,0);

  weddings.forEach(function(b) {
    var wDate = new Date(b.date + 'T12:00:00');
    if (isNaN(wDate)) return;
    var name  = b.name || 'Caros noivos';
    var email = b.contact;

    // 6-month
    var sixM = new Date(wDate); sixM.setMonth(sixM.getMonth() + 6);
    if (daysDiff(today, sixM) === 0) {
      send6MonthEmail(name, email, wDate);
      logMsg(name, email, '6 meses — parabéns', 'Email', 'Sent');
    }

    // Yearly
    for (var y = 1; y <= 20; y++) {
      var ann = new Date(wDate); ann.setFullYear(wDate.getFullYear() + y);
      var diff = daysDiff(today, ann);
      if (diff === DAYS_BEFORE) { sendReminderEmail(name, email, wDate, y); logMsg(name, email, y+'º aniv — aviso', 'Email', 'Sent'); }
      if (diff === 0)           { sendAnniversaryEmail(name, email, wDate, y); logMsg(name, email, y+'º aniv — dia', 'Email', 'Sent'); }
    }
  });
  Logger.log('checkAnniversaries done');
}

// ============================================================
//  INSTALMENT REMINDERS  (daily trigger — 8am)
// ============================================================

function checkInstalments() {
  var result = getAll('instalments');
  var rows   = (result.rows || []).filter(function(r) {
    return String(r.paid).toLowerCase() !== 'true' && String(r.paid) !== 'overdue';
  });

  var today    = new Date(); today.setHours(0,0,0,0);
  var tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);
  var tz       = Session.getScriptTimeZone();
  var todayStr    = Utilities.formatDate(today,    tz, 'yyyy-MM-dd');
  var tomorrowStr = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');

  rows.forEach(function(r) {
    var email   = r.email || '';
    var name    = r.customer_name || 'Cliente';
    var due     = String(r.due_date).split('T')[0];
    var amount  = r.amount || 0;
    var instNum = r.instalment_num || '';

    if (!email.includes('@')) return;

    if (due === tomorrowStr) {
      sendInstalmentReminder(name, email, instNum, amount, due, false);
      logMsg(name, email, 'Parcela '+instNum+' — aviso 1 dia', 'Email', 'Sent');
    }
    if (due < todayStr) {
      sendInstalmentReminder(name, email, instNum, amount, due, true);
      updateRow('instalments', { id: r.id, paid: 'overdue' });
      logMsg(name, email, 'Parcela '+instNum+' — overdue', 'Email', 'Sent');
    }
  });
  Logger.log('checkInstalments done');
}

// ============================================================
//  CHECKLIST REMINDERS  (daily trigger — 8am)
// ============================================================

function checklistReminders() {
  var result   = getAll('bookings');
  var weddings = (result.rows || []).filter(function(b) {
    return String(b.type).toLowerCase() === 'wedding' &&
           String(b.status).toLowerCase() === 'confirmed' &&
           b.contact && String(b.contact).includes('@') && b.date;
  });

  var today  = new Date(); today.setHours(0,0,0,0);
  var target = new Date(today); target.setDate(target.getDate() + 21);
  var tz     = Session.getScriptTimeZone();
  var targetStr = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

  weddings.forEach(function(b) {
    var eventDate = String(b.date).split('T')[0];
    if (eventDate !== targetStr) return;
    var name    = b.name || 'Caros noivos';
    var email   = b.contact;
    var formUrl = BOOKING_URL + '/wedding-form.html?id=' + b.id;
    sendChecklistLink(name, email, formUrl, eventDate);
    logMsg(name, email, 'Wedding checklist link', 'Email', 'Sent');
    Logger.log('Checklist sent to: ' + name);
  });
}

// ============================================================
//  MASTER DAILY TRIGGER — run this one function on a daily timer
// ============================================================

function dailyTrigger() {
  checkAnniversaries();
  checkInstalments();
  checklistReminders();
  Logger.log('Daily trigger completed: ' + new Date());
}

// ============================================================
//  EMAIL HELPERS
// ============================================================

function daysDiff(d1, d2) { return Math.round((d2 - d1) / 86400000); }

function sendMail(to, subject, body) {
  GmailApp.sendEmail(to, subject, body, { name: 'Quinta da Aldeia', replyTo: SENDER_EMAIL });
}

function logMsg(name, email, type, channel, status) {
  var tz = Session.getScriptTimeZone();
  addRow('messageHistory', {
    date_sent: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    couple_name: name, email: email,
    message_type: type, channel: channel, status: status
  });
}

function send6MonthEmail(name, email, wDate) {
  sendMail(email, 'Já passaram 6 meses! Parabéns, ' + name + ' 💍',
    'Caros ' + name + ',\n\nJá passaram 6 meses desde o vosso casamento na Quinta da Aldeia!\n\n' +
    'Para celebrar este marco, oferecemos ' + DISCOUNT + ' de desconto para uma nova reserva:\n' + BOOKING_URL + '/quote.html\n\nCom carinho,\nEquipa Quinta da Aldeia');
}

function sendReminderEmail(name, email, wDate, years) {
  sendMail(email, 'Falta 1 semana para o vosso ' + years + 'º aniversário! 🌸',
    'Caros ' + name + ',\n\nFalta apenas uma semana para o vosso ' + years + 'º aniversário de casamento!\n\n' +
    'Celebrem em grande na Quinta da Aldeia com ' + DISCOUNT + ' de desconto:\n' + BOOKING_URL + '/quote.html\n\nCom os melhores cumprimentos,\nEquipa Quinta da Aldeia');
}

function sendAnniversaryEmail(name, email, wDate, years) {
  var tz   = Session.getScriptTimeZone();
  var wLbl = Utilities.formatDate(wDate, tz, 'd MMMM yyyy');
  sendMail(email, 'Feliz ' + years + 'º Aniversário de Casamento, ' + name + '! 💍',
    'Caros ' + name + ',\n\nFeliz ' + years + 'º Aniversário de Casamento!\n\n' +
    'Foi uma honra receber o vosso casamento na Quinta da Aldeia no dia ' + wLbl + '.\n\n' +
    'Voltem para celebrar — ' + DISCOUNT + ' de desconto exclusivo:\n' + BOOKING_URL + '/quote.html\n\n' +
    'Com muito carinho,\nEquipa Quinta da Aldeia');
}

function sendInstalmentReminder(name, email, instNum, amount, dueDate, isOverdue) {
  var tz     = Session.getScriptTimeZone();
  var dueFmt = Utilities.formatDate(new Date(dueDate + 'T12:00:00'), tz, 'd MMMM yyyy');
  var amtFmt = 'R$ ' + Number(amount).toLocaleString('pt-BR', {minimumFractionDigits:2});
  if (isOverdue) {
    sendMail(email, 'Parcela em atraso — Quinta da Aldeia',
      'Caro(a) ' + name + ',\n\nA parcela nº ' + instNum + ' de ' + amtFmt + ' (vencimento: ' + dueFmt + ') ainda não foi liquidada.\n\n' +
      'Por favor efectue o pagamento ou contacte-nos.\n\nEquipa Quinta da Aldeia');
  } else {
    sendMail(email, 'Lembrete: parcela vence amanhã — Quinta da Aldeia',
      'Caro(a) ' + name + ',\n\nA parcela nº ' + instNum + ' de ' + amtFmt + ' vence amanhã, ' + dueFmt + '.\n\n' +
      'Por favor certifique-se de que o pagamento é efectuado atempadamente.\n\nEquipa Quinta da Aldeia');
  }
}

function sendChecklistLink(name, email, formUrl, eventDate) {
  var tz      = Session.getScriptTimeZone();
  var dateFmt = Utilities.formatDate(new Date(eventDate + 'T12:00:00'), tz, 'd MMMM yyyy');
  sendMail(email, 'Check-list do vosso casamento — Quinta da Aldeia 💍',
    'Caros ' + name + ',\n\nFaltam 3 semanas para o vosso casamento em ' + dateFmt + '!\n\n' +
    'Por favor preencham o check-list:\n' + formUrl + '\n\n' +
    'Com os melhores cumprimentos,\nEquipa Quinta da Aldeia');
}

// ============================================================
//  BOOKING CONFIRMATION EMAIL
//  Called automatically when a confirmed booking is added
//  via the web app — the app passes action=confirmEmail
// ============================================================

function sendBookingConfirmation(p) {
  // p = { name, contact, date, guests, total, type, id }
  if (!p.contact || !String(p.contact).includes('@')) return;
  if (String(p.status).toLowerCase() !== 'confirmed') return;

  var tz       = Session.getScriptTimeZone();
  var eventDt  = new Date(p.date + 'T12:00:00');
  var dateFmt  = Utilities.formatDate(eventDt, tz, 'd MMMM yyyy');
  var total    = Number(p.total) || 0;
  var deposit  = total * 0.3;
  var balance  = total * 0.7;
  var depFmt   = 'R$ ' + deposit.toLocaleString('pt-BR', {minimumFractionDigits:2});
  var balFmt   = 'R$ ' + balance.toLocaleString('pt-BR', {minimumFractionDigits:2});
  var totFmt   = 'R$ ' + total.toLocaleString('pt-BR', {minimumFractionDigits:2});
  var checklistUrl = BOOKING_URL + '/wedding-form.html?id=' + p.id;

  var subject = 'Confirmação de reserva — Quinta da Aldeia 💍';
  var body = [
    'Caros ' + p.name + ',',
    '',
    'É com grande alegria que confirmamos a vossa reserva na Quinta da Aldeia!',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'DETALHES DA RESERVA',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'Evento:          ' + p.type,
    'Data:            ' + dateFmt,
    'Nº de convidados: ' + (p.guests || '—'),
    'Valor total:     ' + totFmt,
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'PAGAMENTOS',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'Sinal (30%):     ' + depFmt,
    'Restante (70%):  ' + balFmt,
    '',
    'Por favor liquidem o sinal para confirmar definitivamente a vossa reserva.',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'PRÓXIMOS PASSOS',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    (p.type === 'Wedding' ? 'Cerca de 3 semanas antes do vosso dia, receberão o link para preencher o check-list do casamento com todos os detalhes do evento.' : ''),
    '',
    'Para qualquer questão, não hesitem em contactar-nos.',
    '',
    'Estamos muito entusiasmados por partilhar este dia especial convosco! 🌿',
    '',
    'Com os melhores cumprimentos,',
    'Equipa Quinta da Aldeia',
    BOOKING_URL
  ].filter(function(l, i, arr) {
    // Remove consecutive empty lines
    return !(l === '' && arr[i-1] === '');
  }).join('\n');

  sendMail(p.contact, subject, body);
  logMsg(p.name, p.contact, 'Booking confirmation', 'Email', 'Sent');
  Logger.log('Confirmation email sent to: ' + p.name);
}

// Override addRow to trigger confirmation email for confirmed bookings
// The web app calls action=addBooking for new bookings to trigger this
function addBookingWithConfirmation(params) {
  var result = addRow('bookings', params);
  if (params.status === 'confirmed' && params.contact && params.contact.includes('@')) {
    try {
      sendBookingConfirmation({
        id:      result.row ? result.row.id : params.id,
        name:    params.name,
        contact: params.contact,
        date:    params.date,
        guests:  params.guests,
        total:   params.total || 0,
        type:    params.type,
        status:  params.status
      });
    } catch(e) {
      Logger.log('Confirmation email error: ' + e);
    }
  }
  return result;
}

// ============================================================
//  QUOTE REQUEST FORM — submitQuote action
//  Saves to QuoteRequests sheet + sends email notification
// ============================================================

// NOTIFICATION_EMAIL is set inside submitQuote() to avoid global variable timing issues

// Add 'submitQuote' to the router — handles form submissions
// Router already uses action param — add this case:
// case 'submitQuote': result = submitQuote(e.parameter); break;
// (Add this line inside the switch in handleRequest above)

function submitQuote(p) {
  var NOTIFICATION_EMAIL = SENDER_EMAIL; // Uses the SENDER_EMAIL set at the top
  // 1. Save to QuoteRequests sheet
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('QuoteRequests');
  if (!sheet) {
    sheet = ss.insertSheet('QuoteRequests');
    var headers = ['id','submitted_at','tipo','data_pref','data_alt','convidados',
                   'servicos','hospedes','nome','telefone','email','obs','status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var id  = p.id || ('qr_' + new Date().getTime());

  sheet.appendRow([
    id,
    now,
    p.tipo         || '',
    p.data_pref    || '',
    p.data_alt     || '',
    p.convidados   || '',
    p.servicos     || '',
    p.hospedes     || '',
    p.nome         || '',
    p.telefone     || '',
    p.email        || '',
    p.obs          || '',
    'Novo'
  ]);

  // 2. Send notification email to the team
  var fmtDate = function(d) {
    if (!d) return '—';
    var parts = d.split('-');
    return parts.length === 3 ? parts[2]+'/'+parts[1]+'/'+parts[0] : d;
  };

  var subject = '🌿 Novo pedido de orçamento — ' + (p.nome || 'Cliente') + ' | Quinta da Aldeia';

  var body = [
    'Olá!',
    '',
    'Recebemos um novo pedido de orçamento através do formulário online.',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'DETALHES DO PEDIDO',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'Nome:             ' + (p.nome       || '—'),
    'Telefone:         ' + (p.telefone   || '—'),
    'E-mail:           ' + (p.email      || '—'),
    '',
    'Tipo de evento:   ' + (p.tipo       || '—'),
    'Data preferida:   ' + fmtDate(p.data_pref),
    'Data alternativa: ' + fmtDate(p.data_alt),
    'Convidados:       ' + (p.convidados || '—'),
    '',
    'Serviços:         ' + (p.servicos   || '—'),
    (p.hospedes ? 'Nº de hóspedes:   ' + p.hospedes : ''),
    '',
    'Observações:',
    (p.obs ? p.obs : '(Sem observações)'),
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '',
    'Acesse o painel para ver todos os pedidos:',
    BOOKING_URL,
    '',
    '— Sistema Quinta da Aldeia'
  ].filter(function(l, i, arr) {
    return !(l === '' && arr[i-1] === '');
  }).join('\n');

  GmailApp.sendEmail(NOTIFICATION_EMAIL, subject, body, {
    name: 'Quinta da Aldeia — Sistema',
    replyTo: p.email || SENDER_EMAIL
  });

  // 3. Send confirmation to the client
  if (p.email && p.email.includes('@')) {
    var clientSubject = 'Recebemos o seu pedido! — Quinta da Aldeia 🌿';
    var clientBody = [
      'Olá, ' + (p.nome || '') + '!',
      '',
      'Obrigado pelo seu interesse na Quinta da Aldeia.',
      'Recebemos o seu pedido de orçamento e entraremos em contacto em até 48 horas com uma proposta personalizada.',
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
      'RESUMO DO SEU PEDIDO',
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
      'Tipo de evento:   ' + (p.tipo       || '—'),
      'Data preferida:   ' + fmtDate(p.data_pref),
      'Convidados:       ' + (p.convidados || '—'),
      'Serviços:         ' + (p.servicos   || '—'),
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
      '',
      'Conheça mais sobre a nossa quinta:',
      '📷 Instagram: https://www.instagram.com/ceremonialquintadaaldeia',
      '🌿 Site: https://www.quintadaaldeia.com.br',
      '',
      'Qualquer dúvida, estamos à disposição.',
      '',
      'Com os melhores cumprimentos,',
      'Equipa Quinta da Aldeia'
    ].join('\n');

    GmailApp.sendEmail(p.email, clientSubject, clientBody, {
      name: 'Quinta da Aldeia',
      replyTo: SENDER_EMAIL
    });
  }

  return { success: true };
}

// ============================================================
//  TEST FUNCTION — run this manually in Apps Script to verify
//  emails are working before going live
// ============================================================
function testSubmitQuote() {
  var testData = {
    id:         'test_' + new Date().getTime(),
    tipo:       'Casamento',
    data_pref:  '2026-08-15',
    data_alt:   '2026-08-22',
    convidados: '101 – 150',
    servicos:   'Aluguel do espaço, Iluminação',
    hospedes:   '',
    nome:       'Teste Silva',
    telefone:   '(11) 9 9999-9999',
    email:      SENDER_EMAIL,
    obs:        'Este é um pedido de teste para verificar o sistema de e-mail.'
  };
  
  var result = submitQuote(testData);
  Logger.log('Test result: ' + JSON.stringify(result));
  Logger.log('Check your inbox at: ' + SENDER_EMAIL);
}
