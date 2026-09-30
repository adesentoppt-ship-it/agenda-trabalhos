/**
 * AGENDA DE TRABALHOS — Google Apps Script
 * Guarda os trabalhos numa folha do Google Sheets e cria eventos no
 * Google Calendar com alarmes (notificação no telemóvel).
 *
 * >>> MUDE O PIN ABAIXO antes de publicar <<<
 */
const PIN_ADMIN = '1234';   // PIN do dono (vê tudo, apaga trabalhos)
const PIN_TECNICO = '5678'; // PIN do técnico (vê agenda, edita e coloca valores)
const ALARMES_MIN = [60, 15]; // avisos antes do trabalho (minutos)

const FOLHA = 'Trabalhos';
const COLS = ['id', 'data', 'hora', 'duracao', 'nome', 'nif', 'morada', 'telefone',
  'servico', 'tecnico', 'estado', 'valor', 'pagamento', 'notas',
  'eventoId', 'criado', 'atualizado', 'fatura', 'lembrado'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Agenda de Trabalhos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/* Chamado pela app alojada no GitHub Pages (fetch POST, corpo JSON em text/plain) */
function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    out = { ok: true, r: api(b.pin, b.acao, b.dados || {}) };
  } catch (err) {
    out = { ok: false, erro: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ponto de entrada único chamado pela app ---------- */
function api(pin, acao, dados) {
  const papel = papel_(pin);
  if (!papel) throw new Error('PIN errado');
  switch (acao) {
    case 'entrar': return { papel: papel, jobs: listar_() }; // 1 só pedido ao entrar
    case 'listar': return listar_();
    case 'guardar': return guardar_(dados, papel);
    case 'subscrever': return subscrever_(dados, papel);   // ativar notificações neste telemóvel
    case 'aviso': return ultimoAviso_(papel);              // texto da última notificação
    case 'apagar':
      if (papel !== 'admin') throw new Error('Só o dono pode apagar trabalhos');
      return apagar_(dados.id);
    default: throw new Error('Ação desconhecida');
  }
}

function papel_(pin) {
  pin = String(pin || '').trim();
  if (pin === PIN_ADMIN) return 'admin';
  if (pin === PIN_TECNICO) return 'tecnico';
  return null;
}

/* ---------- folha ---------- */
function folha_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FOLHA);
  if (!sh) {
    sh = ss.insertSheet(FOLHA);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // texto, para não estragar datas/NIF
    sh.getRange(2, COLS.indexOf('valor') + 1, sh.getMaxRows() - 1, 1).setNumberFormat('0.00');
  }
  if (sh.getLastColumn() < COLS.length) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  return sh;
}

function listar_() {
  const sh = folha_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, COLS.length).getDisplayValues();
  return vals.filter(r => r[0]).map(r => {
    const o = {};
    COLS.forEach((c, i) => o[c] = r[i]);
    o.valor = o.valor === '' ? '' : Number(String(o.valor).replace(',', '.'));
    return o;
  });
}

function linhaDe_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function guardar_(d, papel) {
  let aviso = null;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = folha_();
    const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    let linha = d.id ? linhaDe_(sh, d.id) : -1;
    let atual = {};
    if (linha > 0) {
      const r = sh.getRange(linha, 1, 1, COLS.length).getDisplayValues()[0];
      COLS.forEach((c, i) => atual[c] = r[i]);
    } else {
      atual = { id: Utilities.getUuid().slice(0, 8), criado: agora, eventoId: '' };
    }
    const t = Object.assign({}, atual);
    // O dono marca o trabalho (cliente, NIF, morada, data). O técnico só fecha: estado, valor, pagamento, notas.
    if (papel === 'tecnico' && linha < 0) throw new Error('Só o dono pode criar trabalhos');
    // Contribuinte: o técnico só o coloca no fim, e só se o cliente quiser fatura.
    let campos = papel === 'tecnico'
      ? ['estado', 'valor', 'pagamento', 'notas', 'fatura']
      : ['data', 'hora', 'duracao', 'nome', 'nif', 'morada', 'telefone', 'servico',
        'tecnico', 'estado', 'valor', 'pagamento', 'notas', 'fatura'];
    if (papel === 'tecnico' && d.fatura === 'Sim') campos = campos.concat(['nif', 'nome']);
    campos.forEach(k => {
      if (d[k] !== undefined) t[k] = String(d[k]).trim();
    });
    if (!t.data || !t.hora) throw new Error('Falta a data ou a hora');
    if (!t.nome) throw new Error('Falta o nome do cliente');
    t.nif = String(t.nif || '').replace(/\D/g, '');
    if (t.fatura === 'Sim' && t.estado === 'Concluído' && t.nif.length !== 9) throw new Error('Com fatura é preciso o contribuinte (9 dígitos)');
    t.valor = (t.valor === undefined || t.valor === null || String(t.valor).trim() === '') ? '' : String(Number(String(t.valor).replace(',', '.')) || 0);
    t.estado = t.estado || 'Agendado';
    t.atualizado = agora;

    if (linha > 0 && (t.data !== atual.data || t.hora !== atual.hora)) t.lembrado = '';
    try { t.eventoId = sincronizarEvento_(t); } catch (e) { /* calendário falhou: guarda na mesma */ }

    const row = [COLS.map(c => t[c] === undefined ? '' : t[c])];
    if (linha > 0) sh.getRange(linha, 1, 1, COLS.length).setValues(row);
    else sh.appendRow(row[0]);
    aviso = avisoPara_(papel, linha < 0 ? null : atual, t);
    return listar_();
  } finally {
    lock.releaseLock();
    if (aviso) { try { notificar_(aviso[0], aviso[1], aviso[2]); } catch (e) {} }
  }
}

/* Decide quem recebe notificação depois de guardar */
function avisoPara_(papel, antes, t) {
  const dm = t.data ? t.data.slice(8, 10) + '/' + t.data.slice(5, 7) : '';
  const onde = t.nome + (t.morada ? '\n' + t.morada : '');
  if (papel === 'admin' && !antes && t.estado !== 'Cancelado')
    return ['tecnico', '🔧 Novo trabalho: ' + (t.servico || 'Trabalho'), dm + ' às ' + t.hora + ' · ' + onde];
  if (papel === 'admin' && antes && t.estado === 'Cancelado' && antes.estado !== 'Cancelado')
    return ['tecnico', '❌ Trabalho cancelado', dm + ' às ' + t.hora + ' · ' + t.nome];
  if (papel === 'admin' && antes && (antes.data !== t.data || antes.hora !== t.hora || antes.morada !== t.morada))
    return ['tecnico', '📅 Trabalho alterado: ' + (t.servico || 'Trabalho'), dm + ' às ' + t.hora + ' · ' + onde];
  if (papel === 'tecnico' && antes && t.estado === 'Concluído' && antes.estado !== 'Concluído')
    return ['admin', '✅ Concluído: ' + t.nome, (t.valor ? Number(t.valor).toFixed(2).replace('.', ',') + ' €' : 'sem valor') +
      (t.pagamento ? ' · ' + t.pagamento : '') + (t.fatura === 'Sim' ? ' · 🧾 com fatura' : '')];
  return null;
}

function apagar_(id) {
  const sh = folha_();
  const linha = linhaDe_(sh, id);
  if (linha > 0) {
    const evId = sh.getRange(linha, COLS.indexOf('eventoId') + 1).getDisplayValue();
    if (evId) { try { CalendarApp.getEventById(evId).deleteEvent(); } catch (e) {} }
    sh.deleteRow(linha);
  }
  return listar_();
}

/* ---------- Google Calendar (notificações) ---------- */
function sincronizarEvento_(t) {
  const cal = CalendarApp.getDefaultCalendar();
  let ev = null;
  if (t.eventoId) { try { ev = CalendarApp.getEventById(t.eventoId); } catch (e) { ev = null; } }

  if (t.estado === 'Cancelado') {
    if (ev) ev.deleteEvent();
    return '';
  }
  const tz = Session.getScriptTimeZone();
  const inicio = Utilities.parseDate(t.data + ' ' + t.hora, tz, 'yyyy-MM-dd HH:mm');
  const fim = new Date(inicio.getTime() + (Number(t.duracao) || 60) * 60000);
  const titulo = (t.estado === 'Concluído' ? '✅ ' : '🔧 ') + (t.servico || 'Trabalho') + ' – ' + t.nome;
  const desc = [
    'Cliente: ' + t.nome,
    t.nif ? 'NIF: ' + t.nif : '',
    t.telefone ? 'Telefone: ' + t.telefone : '',
    t.tecnico ? 'Técnico: ' + t.tecnico : '',
    t.notas ? 'Notas: ' + t.notas : ''
  ].filter(Boolean).join('\n');

  if (!ev) ev = cal.createEvent(titulo, inicio, fim);
  else { ev.setTitle(titulo); ev.setTime(inicio, fim); }
  ev.setLocation(t.morada || '');
  ev.setDescription(desc);
  ev.removeAllReminders();
  if (t.estado !== 'Concluído') ALARMES_MIN.forEach(m => ev.addPopupReminder(m));
  return ev.getId();
}

/* Execute uma vez à mão no editor para autorizar e criar a folha */
function configurar() {
  folha_();
  folhaAvisos_();
  CalendarApp.getDefaultCalendar();
  UrlFetchApp.getRequest('https://www.google.com'); // pede autorização para enviar notificações
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'lembretes'))
    ScriptApp.newTrigger('lembretes').timeBased().everyMinutes(10).create();
  Logger.log('Pronto! Agora faça Implementar > Nova implementação.');
}
