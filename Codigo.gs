/**
* AGENDA DE TRABALHOS — Google Apps Script
* Guarda os trabalhos numa folha do Google Sheets e cria eventos no
* Google Calendar com alarmes (notificação no telemóvel).
*
* Os utilizadores e PINs estão no ficheiro Utilizadores.gs (só no Apps Script).
* papel 'admin' -> marca trabalhos, vê faturação, apaga
* papel 'tecnico'-> vê os seus trabalhos e fecha-os (valor, pagamento, fatura)
* faz: true -> um admin que também faz trabalhos (recebe os avisos de técnico)
*/
const ALARMES_MIN = [60, 15]; // avisos antes do trabalho (minutos)

const FOLHA = 'Trabalhos';
const COLS = ['id', 'data', 'hora', 'duracao', 'nome', 'nif', 'morada', 'telefone',
'servico', 'tecnico', 'estado', 'valor', 'pagamento', 'notas',
'eventoId', 'criado', 'atualizado', 'fatura', 'lembrado', 'iva', 'total', 'orcamento','fotos', 'pagoEm', 'origem', 'material', 'concluidoEm', 'faturaFeita', 'faturaAviso'];
const IVA_TAXA = 0.23;

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
const u = utilizador_(pin);
if (!u) throw new Error('PIN errado');
const papel = u.papel;
const podeFinancas = papel === 'admin' || !!u.financas;
switch (acao) {
case 'entrar': return { papel: papel, nome: u.nome, financas: podeFinancas, trabalhadores: trabalhadores_(), jobs: visiveis_(u, listar_()) };
case 'listar': return visiveis_(u, listar_());
case 'guardar': return visiveis_(u, guardar_(dados, u));
    case 'foto': return visiveis_(u, addFoto_(dados, u));      // juntar foto (fica no Google Drive)
    case 'verfoto': return verFoto_(dados, u);
    case 'apagarfoto': return visiveis_(u, apagarFoto_(dados, u));
    case 'comissoes': if (papel !== 'admin') throw new Error('Sem acesso'); return comissoes_();
    case 'pagarcomissao': if (papel !== 'admin') throw new Error('Sem acesso'); return pagarComissao_(dados, u);
    case 'subscrever': return subscrever_(dados, u.nome); // ativar notificações neste telemóvel
case 'aviso': return ultimoAviso_(u.nome); // texto da última notificação
case 'apagar':
if (papel !== 'admin') throw new Error('Só o dono pode apagar trabalhos');
return visiveis_(u, apagar_(dados.id));
/* ---------- Finanças (serviço + casa) — dono, e quem tiver 'financas:true' no Utilizadores.gs ---------- */
case 'fin_tudo': checarFinancas_(podeFinancas); return finTudo_();
case 'fin_conta_saldo': checarFinancas_(podeFinancas); return finContaSaldo_(dados);
case 'fin_mov_novo': checarFinancas_(podeFinancas); return finMovNovo_(dados);
case 'fin_mov_apagar': checarFinancas_(podeFinancas); return finMovApagar_(dados.id);
case 'fin_bem_novo': checarFinancas_(podeFinancas); return finBemNovo_(dados);
case 'fin_bem_editar': checarFinancas_(podeFinancas); return finBemEditar_(dados);
case 'fin_bem_apagar': checarFinancas_(podeFinancas); return finBemApagar_(dados.id);
case 'fin_cambio': checarFinancas_(podeFinancas); return finCambio_(dados);
case 'fin_cdi': checarFinancas_(podeFinancas); return finCdi_(dados);
default: throw new Error('Ação desconhecida');
}
}

function utilizador_(pin) {
pin = String(pin || '').trim();
const u = UTILIZADORES.find(x => x.pin === pin);
/* 'financas: true' no Utilizadores.gs dá acesso às Finanças a quem não é admin (ex.: a esposa/sócia). */
return u ? { nome: u.nome, papel: u.papel, faz: u.papel === 'tecnico' || !!u.faz, financas: !!u.financas } : null;
}
/* Vários técnicos no mesmo trabalho: 'Alex + Lucas' */
function tecs_(s) { return String(s || '').split(/\s*[+,]\s*/).filter(Boolean); }
function trabalhadores_() { return UTILIZADORES.filter(x => x.papel === 'tecnico' || x.faz).map(x => x.nome); }
function admins_() { return UTILIZADORES.filter(x => x.papel === 'admin').map(x => x.nome); }
/* Um técnico só vê os trabalhos dele e os que ainda não têm técnico */
function visiveis_(u, jobs) {
// o técnico não vê de onde veio o trabalho (origem)
  return u.papel === 'admin' ? jobs : jobs.filter(j => !j.tecnico || tecs_(j.tecnico).includes(u.nome)).map(j => { const o = Object.assign({}, j); delete o.origem; return o; });
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
if (sh.getLastColumn() < COLS.length) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    if (COLS.length > 26) sh.getRange(1, 27, sh.getMaxRows(), COLS.length - 26).setNumberFormat('@'); // texto (datas não mudam)
  }
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
    o.orcamentos = lerOrcs_(o.orcamento);
    try { o.fotos = o.fotos ? JSON.parse(o.fotos) : []; } catch (e) { o.fotos = []; }
return o;
});
}

/* Orçamentos: lista [{d: descrição, v: valor, a: aceite}] guardada em JSON na coluna 'orcamento' */
function lerOrcs_(txt) {
txt = String(txt || '').trim();
if (!txt) return [];
if (txt.charAt(0) === '[') { try { return JSON.parse(txt); } catch (e) { return []; } }
const n = Number(txt.replace(',', '.')); return n ? [{ d: '', v: n, a: false }] : []; // formato antigo (1 valor)
}
function limparOrcs_(lista) {
if (!Array.isArray(lista)) return [];
return lista.slice(0, 20).map(o => ({ d: String(o.d || '').trim().slice(0, 120), v: Math.round((Number(String(o.v).replace(',', '.')) || 0) * 100) / 100, a: !!o.a }))
.filter(o => o.d || o.v);
}

function linhaDe_(sh, id) {
const n = sh.getLastRow() - 1;
if (n < 1) return -1;
const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
return -1;
}

function guardar_(d, u) {
const papel = u.papel;
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
? ['estado', 'valor', 'pagamento', 'notas', 'fatura', 'material']
: ['data', 'hora', 'duracao', 'nome', 'nif', 'morada', 'telefone', 'servico',
'tecnico', 'estado', 'valor', 'pagamento', 'notas', 'fatura', 'pagoEm', 'origem', 'material', 'faturaFeita'];
    if (papel === 'tecnico' && d.fatura === 'Sim') campos = campos.concat(['nif', 'nome']);
    // "Por pagar": só a Mariana ou o Alex (admin) marcam como pago
    if (papel === 'tecnico' && atual.pagamento === 'Por pagar' && d.pagamento !== undefined && d.pagamento !== 'Por pagar')
      throw new Error('Só a Mariana ou o Alex podem marcar como pago');
    if (atual.pagamento === 'Por pagar' && d.pagamento && d.pagamento !== 'Por pagar' && !d.pagoEm)
      d.pagoEm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
if (papel === 'tecnico' && atual.tecnico && !tecs_(atual.tecnico).includes(u.nome)) throw new Error('Este trabalho é de ' + atual.tecnico);
    // o técnico pode dizer que fez o trabalho com outro (ex.: 'Lucas + Alex'), mas tem de se incluir a si próprio
    if (papel === 'tecnico' && d.tecnico !== undefined) {
      const lista = tecs_(d.tecnico).filter(n => trabalhadores_().includes(n));
      if (!lista.includes(u.nome)) throw new Error('O técnico tem de estar incluído');
      t.tecnico = lista.join(' + ');
    }
if (papel === 'tecnico' && !atual.tecnico) t.tecnico = u.nome; // fica atribuído a quem o fez
campos.forEach(k => {
if (d[k] !== undefined) t[k] = String(d[k]).trim();
});
if (!t.data || !t.hora) throw new Error('Falta a data ou a hora');
if (!t.nome) throw new Error('Falta o nome do cliente');
t.nif = String(t.nif || '').replace(/\D/g, '');
if (t.fatura === 'Sim' && t.estado === 'Concluído' && t.nif.length !== 9) throw new Error('Com fatura é preciso o contribuinte (9 dígitos)');
t.valor = (t.valor === undefined || t.valor === null || String(t.valor).trim() === '') ? '' : String(Number(String(t.valor).replace(',', '.')) || 0);
if (d.orcamentos !== undefined) { const l = limparOrcs_(d.orcamentos); t.orcamento = l.length ? JSON.stringify(l) : ''; }
// IVA 23% só com fatura: o valor do técnico é SEM IVA; o cliente paga valor + IVA
const base = Number(t.valor) || 0;
t.iva = (t.fatura === 'Sim' && t.valor !== '') ? (Math.round(base * IVA_TAXA * 100) / 100).toFixed(2) : (t.valor !== '' ? '0.00' : '');
t.total = t.valor !== '' ? (base + (Number(t.iva) || 0)).toFixed(2) : '';
t.estado = t.estado || 'Agendado';
    if ((t.estado === 'Concluído' || t.estado === 'Orçamento recusado') && !t.concluidoEm) t.concluidoEm = agora;
    if (t.estado !== 'Concluído' && t.estado !== 'Orçamento recusado') t.concluidoEm = '';
    t.material = (t.material === undefined || String(t.material).trim() === '') ? '' : String(Number(String(t.material).replace(',', '.')) || 0);
t.atualizado = agora;

if (linha > 0 && (t.data !== atual.data || t.hora !== atual.hora)) t.lembrado = '';
try { t.eventoId = sincronizarEvento_(t); } catch (e) { /* calendário falhou: guarda na mesma */ }

const row = [COLS.map(c => t[c] === undefined ? '' : t[c])];
if (linha > 0) sh.getRange(linha, 1, 1, COLS.length).setValues(row);
else sh.appendRow(row[0]);
try { financaAutoMovimento_(t); } catch (e) { /* finanças falhou: guarda o trabalho na mesma */ }
aviso = avisoPara_(u, linha < 0 ? null : atual, t);
return listar_();
} finally {
lock.releaseLock();
if (aviso) { try { notificar_(aviso[0], aviso[1], aviso[2]); } catch (e) {} }
}
}

/* Decide quem recebe notificação depois de guardar: [nomes[], titulo, corpo] */
function avisoPara_(u, antes, t) {
const dm = t.data ? t.data.slice(8, 10) + '/' + t.data.slice(5, 7) : '';
const onde = t.nome + (t.morada ? '\n' + t.morada : '');
const quem = nomes => nomes.filter(n => n !== u.nome); // não avisa quem fez a alteração
const equipa = tec => quem(tec ? tecs_(tec) : trabalhadores_()); // técnico escolhido, ou todos
if (u.papel === 'admin') {
if (!antes && t.estado !== 'Cancelado')
return [equipa(t.tecnico), '🔧 Novo trabalho: ' + (t.servico || 'Trabalho'), dm + ' às ' + t.hora + ' · ' + onde];
if (antes && t.estado === 'Cancelado' && antes.estado !== 'Cancelado')
return [equipa(t.tecnico || antes.tecnico), '❌ Trabalho cancelado', dm + ' às ' + t.hora + ' · ' + t.nome];
if (antes && t.estado !== 'Concluído' && (antes.data !== t.data || antes.hora !== t.hora || antes.morada !== t.morada || antes.tecnico !== t.tecnico))
return [equipa(t.tecnico), '📅 Trabalho alterado: ' + (t.servico || 'Trabalho'), dm + ' às ' + t.hora + ' · ' + onde];
}
const eur = v => Number(v).toFixed(2).replace('.', ',') + ' €';
if (antes && t.estado === 'Concluído' && antes.estado !== 'Concluído')
return [quem(admins_()), '✅ Concluído por ' + u.nome + ': ' + t.nome,
(t.valor === '' ? 'sem valor' : t.fatura === 'Sim'
? eur(t.valor) + ' + IVA ' + eur(t.iva) + ' = ' + eur(t.total) + ' · 🧾 fatura'
: eur(t.valor) + ' (sem fatura)') + (t.pagamento ? ' · ' + t.pagamento : '')];
if (antes && t.estado === 'Orçamento recusado' && antes.estado !== 'Orçamento recusado')
    return [quem(admins_()), '❌ Orçamento recusado: ' + t.nome, Number(t.valor) ? 'Taxa de deslocação ' + eur(t.valor) : 'Sem taxa de deslocação'];
  if (u.papel === 'tecnico' && t.estado === 'Orçamento dado' && (!antes || antes.estado !== 'Orçamento dado'))
return [quem(admins_()), '💬 Orçamento dado por ' + u.nome + ': ' + t.nome,
lerOrcs_(t.orcamento).map(o => (o.d ? o.d + ' ' : '') + eur(o.v)).join(' · ') || 'sem valor indicado'];
return null;
}

function apagar_(id) {
const sh = folha_();
const linha = linhaDe_(sh, id);
if (linha > 0) {
const evId = sh.getRange(linha, COLS.indexOf('eventoId') + 1).getDisplayValue();
if (evId) { try { CalendarApp.getEventById(evId).deleteEvent(); } catch (e) {} }
try { finRemoverPorJob_(id); } catch (e) {}
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
const titulo = (t.estado === 'Concluído' ? '✅ ' : t.estado === 'Orçamento dado' ? '💬 ' : '🔧 ') + (t.servico || 'Trabalho') + ' – ' + t.nome;
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

/* ============================================================
* FINANÇAS (serviço + casa) — folhas Contas / Movimentos / Bens / Config
* Quando um trabalho é concluído e pago, cria/atualiza automaticamente
* um movimento de Entrada ligado a esse trabalho (id "job-<id>") e
* ajusta o saldo da conta correspondente. Se o trabalho deixar de
* estar concluído, ficar "Por pagar" ou for apagado, o movimento é
* removido e o saldo é corrigido.
* ============================================================ */
const FOLHA_CONTAS = 'Contas';
const COLS_CONTAS = ['id', 'nome', 'moeda', 'saldo', 'ordem', 'pais', 'cdi', 'saldoData', 'atualizado'];
const FOLHA_MOV = 'Movimentos';
const COLS_MOV = ['id', 'data', 'area', 'tipo', 'categoria', 'descricao', 'valor', 'contaId', 'importado', 'origemJobId'];
const FOLHA_BENS = 'Bens';
const COLS_BENS = ['id', 'nome', 'valor', 'area', 'ordem'];
const FOLHA_CFG = 'Config';
const COLS_CFG = ['chave', 'eurbrl', 'taxaAnual', 'data', 'fonte'];
/* Para onde vai o dinheiro de um trabalho, conforme a forma de pagamento escolhida pelo técnico.
Ajuste aqui se quiser outra conta (os ids são os da folha Contas, coluna 'id'):
Multibanco e Transferência -> Santander; MB Way do Alex (925 375 475) -> Santander;
MB Way da Mariana (910 451 649) -> Novo Banco; Dinheiro em mãos -> dinheiro guardado em casa. */
const PAGAMENTO_CONTA = {
'Dinheiro': 'casa',
'MB Way – Alex (925 375 475)': 'santander',
'MB Way – Mariana (910 451 649)': 'novobanco',
'Multibanco': 'santander',
'Transferência': 'santander'
};

function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function checarFinancas_(podeFinancas) { if (!podeFinancas) throw new Error('Sem acesso às finanças'); }
function contaDoPagamento_(p) { return PAGAMENTO_CONTA[p] || 'santander'; }

function folhaContas_() {
const ss = SpreadsheetApp.getActive();
let sh = ss.getSheetByName(FOLHA_CONTAS);
if (!sh) { sh = ss.insertSheet(FOLHA_CONTAS); sh.getRange(1, 1, 1, COLS_CONTAS.length).setValues([COLS_CONTAS]).setFontWeight('bold'); sh.setFrozenRows(1); }
return sh;
}
function folhaMov_() {
const ss = SpreadsheetApp.getActive();
let sh = ss.getSheetByName(FOLHA_MOV);
if (!sh) { sh = ss.insertSheet(FOLHA_MOV); sh.getRange(1, 1, 1, COLS_MOV.length).setValues([COLS_MOV]).setFontWeight('bold'); sh.setFrozenRows(1); }
return sh;
}
function folhaBens_() {
const ss = SpreadsheetApp.getActive();
let sh = ss.getSheetByName(FOLHA_BENS);
if (!sh) { sh = ss.insertSheet(FOLHA_BENS); sh.getRange(1, 1, 1, COLS_BENS.length).setValues([COLS_BENS]).setFontWeight('bold'); sh.setFrozenRows(1); }
return sh;
}
function folhaCfg_() {
const ss = SpreadsheetApp.getActive();
let sh = ss.getSheetByName(FOLHA_CFG);
if (!sh) { sh = ss.insertSheet(FOLHA_CFG); sh.getRange(1, 1, 1, COLS_CFG.length).setValues([COLS_CFG]).setFontWeight('bold'); sh.setFrozenRows(1); }
return sh;
}

function listarContas_() {
const sh = folhaContas_();
const n = sh.getLastRow() - 1; if (n < 1) return [];
const vals = sh.getRange(2, 1, n, COLS_CONTAS.length).getDisplayValues();
return vals.filter(r => r[0]).map(r => {
const o = {}; COLS_CONTAS.forEach((c, i) => o[c] = r[i]);
o.saldo = Number(String(o.saldo).replace(',', '.')) || 0;
o.ordem = Number(o.ordem) || 0;
o.cdi = o.cdi === '' ? null : Number(o.cdi);
return o;
});
}
function listarMov_() {
const sh = folhaMov_();
const n = sh.getLastRow() - 1; if (n < 1) return [];
const vals = sh.getRange(2, 1, n, COLS_MOV.length).getDisplayValues();
return vals.filter(r => r[0]).map(r => {
const o = {}; COLS_MOV.forEach((c, i) => o[c] = r[i]);
o.valor = Number(String(o.valor).replace(',', '.')) || 0;
o.importado = o.importado === 'true' || o.importado === true;
return o;
}).sort((a, b) => b.data.localeCompare(a.data));
}
function listarBens_() {
const sh = folhaBens_();
const n = sh.getLastRow() - 1; if (n < 1) return [];
const vals = sh.getRange(2, 1, n, COLS_BENS.length).getDisplayValues();
return vals.filter(r => r[0]).map(r => {
const o = {}; COLS_BENS.forEach((c, i) => o[c] = r[i]);
o.valor = Number(String(o.valor).replace(',', '.')) || 0;
o.ordem = Number(o.ordem) || 0;
return o;
});
}
function linhaCfg_(sh, chave) {
const n = sh.getLastRow() - 1;
if (n > 0) { const ids = sh.getRange(2, 1, n, 1).getDisplayValues(); for (let i = 0; i < ids.length; i++) if (ids[i][0] === chave) return i + 2; }
sh.appendRow([chave]); return sh.getLastRow();
}
function lerCfg_(chave) {
const sh = folhaCfg_(); const linha = linhaCfg_(sh, chave);
const r = sh.getRange(linha, 1, 1, COLS_CFG.length).getDisplayValues()[0];
const o = {}; COLS_CFG.forEach((c, i) => o[c] = r[i]);
if (o.eurbrl !== '') o.eurbrl = Number(String(o.eurbrl).replace(',', '.'));
if (o.taxaAnual !== '') o.taxaAnual = Number(String(o.taxaAnual).replace(',', '.'));
return o;
}
function guardarCfg_(chave, campos) {
const sh = folhaCfg_(); const linha = linhaCfg_(sh, chave);
const atual = lerCfg_(chave);
const novo = Object.assign({}, atual, campos, { chave });
sh.getRange(linha, 1, 1, COLS_CFG.length).setValues([COLS_CFG.map(c => novo[c] === undefined ? '' : novo[c])]);
return novo;
}
function movLinhaDe_(sh, id) {
const n = sh.getLastRow() - 1; if (n < 1) return -1;
const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
return -1;
}
function bemLinhaDe_(sh, id) {
const n = sh.getLastRow() - 1; if (n < 1) return -1;
const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
return -1;
}
function contaLinhaDe_(sh, id) {
const n = sh.getLastRow() - 1; if (n < 1) return -1;
const ids = sh.getRange(2, 1, n, 1).getDisplayValues();
for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
return -1;
}
function contaAjustar_(contaId, delta) {
if (!contaId || !delta) return;
const sh = folhaContas_();
const linha = contaLinhaDe_(sh, contaId);
if (linha < 0) return;
const saldoCol = COLS_CONTAS.indexOf('saldo') + 1, atualCol = COLS_CONTAS.indexOf('atualizado') + 1, saldoDataCol = COLS_CONTAS.indexOf('saldoData') + 1;
const atual = Number(sh.getRange(linha, saldoCol).getValue()) || 0;
const hojeS = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
sh.getRange(linha, saldoCol).setValue(r2(atual + delta));
sh.getRange(linha, atualCol).setValue(hojeS);
sh.getRange(linha, saldoDataCol).setValue(hojeS);
}

function finTudo_() {
return { contas: listarContas_(), bens: listarBens_(), movimentos: listarMov_(), cambio: lerCfg_('cambio'), cdi: lerCfg_('cdi') };
}
function finContaSaldo_(d) {
const sh = folhaContas_();
const linha = contaLinhaDe_(sh, d.id);
if (linha < 0) throw new Error('Conta não encontrada');
const hojeS = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
sh.getRange(linha, COLS_CONTAS.indexOf('saldo') + 1).setValue(r2(Number(d.saldo) || 0));
sh.getRange(linha, COLS_CONTAS.indexOf('atualizado') + 1).setValue(hojeS);
sh.getRange(linha, COLS_CONTAS.indexOf('saldoData') + 1).setValue(hojeS);
return finTudo_();
}
function finMovNovo_(d) {
const valor = r2(Number(String(d.valor).replace(',', '.')) || 0);
if (!valor || !d.contaId) throw new Error('Falta o valor ou a conta');
const sh = folhaMov_();
const row = { id: 'm-' + Utilities.getUuid().slice(0, 8), data: d.data || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
area: d.area || 'Casa', tipo: d.tipo || 'Saída', categoria: d.categoria || '', descricao: d.descricao || '',
valor: String(valor), contaId: d.contaId, importado: 'false', origemJobId: '' };
sh.appendRow(COLS_MOV.map(c => row[c]));
contaAjustar_(d.contaId, d.tipo === 'Entrada' ? valor : -valor);
return finTudo_();
}
function finMovApagar_(id) {
const sh = folhaMov_();
const linha = movLinhaDe_(sh, id);
if (linha > 0) {
const r = sh.getRange(linha, 1, 1, COLS_MOV.length).getDisplayValues()[0];
const o = {}; COLS_MOV.forEach((c, i) => o[c] = r[i]);
contaAjustar_(o.contaId, o.tipo === 'Entrada' ? -Number(o.valor) : Number(o.valor));
sh.deleteRow(linha);
}
return finTudo_();
}
function finBemNovo_(d) {
const sh = folhaBens_();
const n = sh.getLastRow() - 1;
sh.appendRow(['b-' + Utilities.getUuid().slice(0, 8), d.nome, r2(Number(String(d.valor).replace(',', '.')) || 0), d.area || 'Serviço', n + 1]);
return finTudo_();
}
function finBemEditar_(d) {
const sh = folhaBens_();
const linha = bemLinhaDe_(sh, d.id);
if (linha < 0) throw new Error('Bem não encontrado');
sh.getRange(linha, COLS_BENS.indexOf('valor') + 1).setValue(r2(Number(String(d.valor).replace(',', '.')) || 0));
return finTudo_();
}
function finBemApagar_(id) {
const sh = folhaBens_();
const linha = bemLinhaDe_(sh, id);
if (linha > 0) sh.deleteRow(linha);
return finTudo_();
}
function finCambio_(d) {
guardarCfg_('cambio', { eurbrl: r2(Number(String(d.eurbrl).replace(',', '.')) || 0), data: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') });
return finTudo_();
}
function finCdi_(d) {
guardarCfg_('cdi', { taxaAnual: r2(Number(String(d.taxaAnual).replace(',', '.')) || 0), data: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), fonte: d.fonte || '' });
return finTudo_();
}

/* Liga a conclusão de um trabalho às finanças: cria/atualiza/remove o movimento "job-<id>" */
/* ---------- TRABALHO REPASSADO: tira a comissão do técnico e divide o resto a meio ---------- */
const COMISSAO_SRV = { 'Lucas': 0.10 };
function pascoaSrv_(y) { const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),mes=Math.floor((h+l-7*m+114)/31),dia=((h+l-7*m+114)%31)+1; return Date.UTC(y,mes-1,dia); }
function feriadoSrv_(ds) { const y = +ds.slice(0,4), p = pascoaSrv_(y), f = n => new Date(p + n*864e5).toISOString().slice(0,10);
  return ['01-01','04-25','05-01','06-10','08-15','10-05','11-01','12-01','12-08','12-25'].map(x => y+'-'+x).concat([f(-2), f(0), f(60)]).indexOf(ds) >= 0; }
function foraHorasSrv_(t) { if (!t.data) return false; const dow = new Date(t.data+'T12:00:00Z').getUTCDay(); if (dow===0 || dow===6 || feriadoSrv_(t.data)) return true; const hh = parseInt(String(t.hora||'12'),10); return hh < 8 || hh >= 17; }
function comissaoSrv_(t) { const tecs = tecs_(t.tecnico); let tot = 0;
  Object.keys(COMISSAO_SRV).forEach(n => { if (tecs.indexOf(n) >= 0 && foraHorasSrv_(t)) tot += Math.max(0, (Number(t.valor)||0) - (Number(t.material)||0)) * (t.estado === 'Orçamento recusado' ? 0.5 : COMISSAO_SRV[n]) /* taxa de deslocação: 50% */; });
  return Math.round(tot*100)/100; }
function parteEmpresa_(t) { const v = Number(t.valor) || 0; if (t.origem !== 'Trabalho repassado') return Math.round(v*100)/100; return Math.round((v - comissaoSrv_(t)) / 2 * 100) / 100; }

function financaAutoMovimento_(t) {
const deveExistir = t.estado === 'Concluído' && t.pagamento && t.pagamento !== 'Por pagar' && t.valor !== '' && Number(t.valor) > 0;
const sh = folhaMov_();
const linha = movLinhaDe_(sh, 'job-' + t.id);
let antigo = null;
if (linha > 0) {
const r = sh.getRange(linha, 1, 1, COLS_MOV.length).getDisplayValues()[0];
antigo = {}; COLS_MOV.forEach((c, i) => antigo[c] = r[i]);
}
if (!deveExistir) {
if (antigo) { contaAjustar_(antigo.contaId, -Number(antigo.valor || 0)); sh.deleteRow(linha); }
return;
}
const contaId = contaDoPagamento_(t.pagamento);
const valor = parteEmpresa_(t); // trabalho repassado: só a nossa parte
const novo = { id: 'job-' + t.id, data: t.data, area: 'Serviço', tipo: 'Entrada', categoria: t.servico || 'Desentupimentos', descricao: t.nome, valor: String(valor), contaId: contaId, importado: 'false', origemJobId: t.id };
if (antigo) {
const valorAntigo = Number(antigo.valor || 0);
if (antigo.contaId !== contaId) { contaAjustar_(antigo.contaId, -valorAntigo); contaAjustar_(contaId, valor); }
else if (valorAntigo !== valor) { contaAjustar_(contaId, valor - valorAntigo); }
sh.getRange(linha, 1, 1, COLS_MOV.length).setValues([COLS_MOV.map(c => novo[c])]);
} else {
contaAjustar_(contaId, valor);
sh.appendRow(COLS_MOV.map(c => novo[c]));
}
}
function finRemoverPorJob_(jobId) {
const sh = folhaMov_();
const linha = movLinhaDe_(sh, 'job-' + jobId);
if (linha > 0) {
const r = sh.getRange(linha, 1, 1, COLS_MOV.length).getDisplayValues()[0];
const o = {}; COLS_MOV.forEach((c, i) => o[c] = r[i]);
contaAjustar_(o.contaId, -Number(o.valor));
sh.deleteRow(linha);
}
}

/* Execute UMA VEZ à mão no editor, depois de "configurar", para trazer os dados que já
existiam no controlo financeiro antigo (contas, bens, movimentos, câmbio, CDI).
Não faz nada se a folha Contas já tiver dados, para nunca duplicar. */
function migrarFinancas_() {
const shC = folhaContas_();
if (shC.getLastRow() > 1) { Logger.log('Já há contas na folha — a migração já foi feita, não repete nada.'); return; }
const contas = [
['itau-cc', 'Itaú – conta corrente', 'BRL', 1939, 0.5, 'Brasil', '', '', '2026-09-30'],
['itau', 'Itaú – CDI (100% do CDI)', 'BRL', 124194, 1, 'Brasil', 100, '2026-09-30', '2026-09-30'],
['santander', 'Santander', 'EUR', 4095.4, 2, 'Portugal', '', '', '2026-09-30'],
['novobanco', 'Novo Banco', 'EUR', 1750, 3, 'Portugal', '', '', '2026-09-30'],
['casa', 'Dinheiro em casa', 'EUR', 6820, 4, 'Portugal', '', '', '2026-09-30']
];
contas.forEach(r => shC.appendRow(r));

const shB = folhaBens_();
const bens = [
['peugeot', 'Carrinha Peugeot', 6500, 'Serviço', 1],
['fiat-talento', 'Carrinha Fiat Talento', 7000, 'Serviço', 2],
['motor-alta-pressao', 'Motor de alta pressão', 3500, 'Serviço', 3],
['k60', 'K60', 2000, 'Serviço', 4],
['cobra22', 'Cobra 22', 500, 'Serviço', 5],
['camera', 'Câmara de vídeo', 100, 'Serviço', 6],
['bomba', 'Bomba', 150, 'Serviço', 7],
['casa', 'Coisas em casa', 1200, 'Casa', 8]
];
bens.forEach(r => shB.appendRow(r));

const shM = folhaMov_();
const movs = [
['nb-2026-09-30-desent-360', '2026-09-30', 'Serviço', 'Entrada', 'Desentupimentos', 'Trabalho de desentupimento', 360, 'novobanco', 'false', ''],
['santander-01', '2026-09-30', 'Casa', 'Saída', 'Outras despesas', 'MB Way para Thiago Santos Andrade', 248, 'santander', 'true', ''],
['santander-02', '2026-09-29', 'Casa', 'Saída', 'Lazer / restaurantes', "Vontade D'Alvorada Pedroso", 16.9, 'santander', 'true', ''],
['santander-03', '2026-09-29', 'Casa', 'Saída', 'Outras despesas', 'MB Way para nº …6708', 670, 'santander', 'true', ''],
['santander-04', '2026-09-29', 'Serviço', 'Entrada', 'Desentupimentos', 'Teya (pagamentos com cartão)', 1198.37, 'santander', 'true', ''],
['santander-05', '2026-09-29', 'Serviço', 'Entrada', 'Desentupimentos', 'MB Way de Marcelo Antonucci Araujo', 418.2, 'santander', 'true', ''],
['santander-06', '2026-09-29', 'Casa', 'Saída', 'Outras despesas', 'Pagamento Hipay', 43.05, 'santander', 'true', ''],
['santander-07', '2026-09-29', 'Casa', 'Saída', 'Outras despesas', 'Pagamento Hipay', 26.5, 'santander', 'true', ''],
['santander-08', '2026-09-28', 'Serviço', 'Entrada', 'Desentupimentos', 'Teya (pagamentos com cartão)', 318.08, 'santander', 'true', ''],
['santander-09', '2026-09-28', 'Casa', 'Saída', 'Outras despesas', 'Banco Credibom', 72.9, 'santander', 'true', ''],
['santander-10', '2026-09-28', 'Casa', 'Saída', 'Água', 'Águas de Gaia', 42.37, 'santander', 'true', ''],
['santander-11', '2026-09-28', 'Casa', 'Entrada', 'Outras entradas', 'Segurança Social', 637.9, 'santander', 'true', ''],
['santander-12', '2026-09-28', 'Casa', 'Saída', 'Outras despesas', 'MB Way para Fernanda Gabriella Jesus Macedo', 20, 'santander', 'true', ''],
['santander-13', '2026-09-25', 'Casa', 'Saída', 'Outras despesas', 'MB Way para Jessica Barbosa Siqueira Porto', 20, 'santander', 'true', ''],
['santander-14', '2026-09-25', 'Casa', 'Saída', 'Outras despesas', 'MB Way para Lucas Vinicios de Oliveira Quiaper', 38, 'santander', 'true', ''],
['santander-15', '2026-09-25', 'Casa', 'Saída', 'Outras despesas', 'Viver Pedroso', 39.9, 'santander', 'true', ''],
['santander-16', '2026-09-25', 'Casa', 'Saída', 'Educação', 'Associação de Pais da escola', 15, 'santander', 'true', ''],
['santander-17', '2026-09-25', 'Casa', 'Saída', 'Outras despesas', 'SquareTrade', 7.5, 'santander', 'true', ''],
['santander-18', '2026-09-24', 'Serviço', 'Entrada', 'Desentupimentos', 'MB Way de Daniel Filipe dos Santos Tavares', 20, 'santander', 'true', ''],
['santander-19', '2026-09-24', 'Casa', 'Saída', 'Lazer / restaurantes', 'Chimarrão Colombo', 1.8, 'santander', 'true', ''],
['santander-20', '2026-09-24', 'Casa', 'Saída', 'Lazer / restaurantes', 'Chimarrão Colombo', 63.2, 'santander', 'true', ''],
['santander-21', '2026-09-23', 'Casa', 'Saída', 'Supermercado', 'Continente', 0.1, 'santander', 'true', ''],
['santander-22', '2026-09-23', 'Casa', 'Saída', 'Supermercado', 'Continente', 13.55, 'santander', 'true', ''],
['santander-23', '2026-09-23', 'Casa', 'Saída', 'Outras despesas', 'MB Way para Thiago Santos Andrade', 152, 'santander', 'true', ''],
['santander-24', '2026-09-22', 'Serviço', 'Entrada', 'Desentupimentos', 'MB Way de Lucian Canito Firmino', 80, 'santander', 'true', ''],
['santander-25', '2026-09-22', 'Serviço', 'Entrada', 'Desentupimentos', 'MB Way de Hugo Daniel Santos Batista', 380, 'santander', 'true', ''],
['santander-26', '2026-09-22', 'Serviço', 'Saída', 'Portagens / estacionamento', 'Via Verde', 22.79, 'santander', 'true', ''],
['santander-27', '2026-09-22', 'Serviço', 'Saída', 'Google Ads / publicidade', 'Fraud Blocker', 86.43, 'santander', 'true', ''],
['santander-28', '2026-09-21', 'Casa', 'Saída', 'Supermercado', 'Intermarché', 95.03, 'santander', 'true', ''],
['santander-29', '2026-09-21', 'Serviço', 'Entrada', 'Desentupimentos', 'Transferência de Beatriz Vazzoler Poian Miller', 320, 'santander', 'true', ''],
['santander-30', '2026-09-21', 'Casa', 'Saída', 'Lazer / restaurantes', 'Café Zuca', 17, 'santander', 'true', ''],
['santander-31', '2026-09-21', 'Serviço', 'Saída', 'Telemóvel / internet', 'Carregamento Vodafone', 16, 'santander', 'true', ''],
['santander-32', '2026-09-21', 'Serviço', 'Entrada', 'Desentupimentos', 'Teya (pagamentos com cartão)', 268.38, 'santander', 'true', ''],
['santander-33', '2026-09-21', 'Serviço', 'Saída', 'Google Ads / publicidade', 'Google Ads', 350, 'santander', 'true', ''],
['santander-34', '2026-09-21', 'Casa', 'Saída', 'Outras despesas', 'Gu Yueguang', 24.9, 'santander', 'true', ''],
['santander-35', '2026-09-21', 'Serviço', 'Saída', 'Combustível', '675-Grijó', 98.3, 'santander', 'true', ''],
['santander-36', '2026-09-21', 'Casa', 'Saída', 'Lazer / restaurantes', "Vontade D'Alvorada", 33.4, 'santander', 'true', ''],
['santander-37', '2026-09-21', 'Casa', 'Saída', 'Outras despesas', 'MB Way para Lucas Vinicios de Oliveira Quiaper', 40, 'santander', 'true', ''],
['santander-38', '2026-09-21', 'Serviço', 'Saída', 'Material e equipamento', 'BCM Bricolage', 48, 'santander', 'true', '']
];
movs.forEach(r => shM.appendRow(r));

guardarCfg_('cambio', { eurbrl: 5.91, data: '2026-09-30' });
guardarCfg_('cdi', { taxaAnual: 13.65, data: '2026-09-30', fonte: 'numerando.com.br/cdi-hoje' });
Logger.log('Finanças migradas: ' + contas.length + ' contas, ' + bens.length + ' bens, ' + movs.length + ' movimentos.');
}

/* Execute uma vez à mão no editor para autorizar e criar a folha */
/* ---------- FOTOS (pasta no Google Drive do dono) ---------- */
function pastaFotos_() {
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('pastaFotos');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder('Agenda de Trabalhos - Fotos'); p.setProperty('pastaFotos', f.getId()); return f;
}
function fotosDe_(sh, linha) {
  try { return JSON.parse(sh.getRange(linha, COLS.indexOf('fotos') + 1).getDisplayValue() || '[]'); } catch (e) { return []; }
}
function podeVer_(u, sh, linha) {
  if (u.papel === 'admin') return true;
  const tec = sh.getRange(linha, COLS.indexOf('tecnico') + 1).getDisplayValue();
return !tec || tecs_(tec).includes(u.nome);
}
function addFoto_(d, u) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = folha_(); const linha = linhaDe_(sh, d.id);
    if (linha < 0) throw new Error('Trabalho não encontrado');
    if (!podeVer_(u, sh, linha)) throw new Error('Este trabalho não é seu');
    const b64 = String(d.foto || ''), b64t = String(d.mini || '');
    if (!b64 || b64.length > 4000000) throw new Error('Foto inválida ou demasiado grande');
    const fotos = fotosDe_(sh, linha);
    if (fotos.length >= 30) throw new Error('Máximo de 30 fotos por trabalho');
    const r = sh.getRange(linha, 1, 1, COLS.length).getDisplayValues()[0];
    const nome = (r[COLS.indexOf('data')] + ' ' + r[COLS.indexOf('nome')] + ' ' + (fotos.length + 1)).replace(/[\/:*?"<>|]/g, '');
    const pasta = pastaFotos_();
    const f = pasta.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nome + '.jpg'));
    const t = b64t ? pasta.createFile(Utilities.newBlob(Utilities.base64Decode(b64t), 'image/jpeg', nome + ' (mini).jpg')) : f;
    fotos.push({ id: f.getId(), t: t.getId(), por: u.nome, em: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
    sh.getRange(linha, COLS.indexOf('fotos') + 1).setValue(JSON.stringify(fotos));
    return listar_();
  } finally { lock.releaseLock(); }
}
/* Devolve a foto em base64 só a quem pode ver o trabalho (as fotos não ficam públicas) */
function verFoto_(d, u) {
  const sh = folha_(); const linha = linhaDe_(sh, d.id);
  if (linha < 0 || !podeVer_(u, sh, linha)) throw new Error('Sem acesso');
  const ok = fotosDe_(sh, linha).some(f => f.id === d.foto || f.t === d.foto);
  if (!ok) throw new Error('Foto não encontrada');
  return Utilities.base64Encode(DriveApp.getFileById(d.foto).getBlob().getBytes());
}
function apagarFoto_(d, u) {
  const sh = folha_(); const linha = linhaDe_(sh, d.id);
  if (linha < 0 || !podeVer_(u, sh, linha)) throw new Error('Sem acesso');
  const fotos = fotosDe_(sh, linha), f = fotos.find(x => x.id === d.foto);
  if (!f) throw new Error('Foto não encontrada');
  if (u.papel !== 'admin' && f.por !== u.nome) throw new Error('Só pode apagar as suas fotos');
  [f.id, f.t].forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }); // vai para o Lixo do Drive
  sh.getRange(linha, COLS.indexOf('fotos') + 1).setValue(JSON.stringify(fotos.filter(x => x !== f)));
  return listar_();
}

/* ---------- COMISSÕES (pagas ao domingo, semana de segunda a domingo) ---------- */
function folhaCom_() {
  const ss = SpreadsheetApp.getActive(); let sh = ss.getSheetByName('Comissoes');
  if (!sh) { sh = ss.insertSheet('Comissoes'); sh.appendRow(['semana', 'tecnico', 'valor', 'trabalhos', 'pagoEm', 'pagoPor']); sh.setFrozenRows(1); sh.getRange('A:F').setNumberFormat('@'); }
  return sh;
}
function comissoes_() {
  const sh = folhaCom_(); const n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, 6).getDisplayValues().map(r => ({ semana: r[0], tecnico: r[1], valor: Number(String(r[2]).replace(',', '.')) || 0, trabalhos: r[3], pagoEm: r[4], pagoPor: r[5] }));
}
function pagarComissao_(d, u) {
  const sem = String(d.semana || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(sem)) throw new Error('Semana inválida');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = folhaCom_(); const vals = sh.getDataRange().getDisplayValues();
    const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const row = [sem, String(d.tecnico || ''), String(Math.round((Number(d.valor) || 0) * 100) / 100), String(d.trabalhos || ''), agora, u.nome];
    for (let i = 1; i < vals.length; i++) if (vals[i][0] === sem && vals[i][1] === row[1]) { sh.getRange(i + 1, 1, 1, 6).setValues([row]); return comissoes_(); }
    sh.appendRow(row); return comissoes_();
  } finally { lock.releaseLock(); }
}
/* ---------- LEMBRETE DE FATURAS: 24h depois de concluído, 1 vez por dia, até carregar em "Fatura feita" ---------- */
/* ---------- SALÁRIOS: lembrete no dia do pagamento (e nos dias seguintes) até marcar como pago ---------- */
const SALARIOS = [{ nome: 'Lucas', valor: 920, dia: 16 }];
function lembrarSalario_() {
  const tz = Session.getScriptTimeZone(), agora = new Date(), F = f => Utilities.formatDate(agora, tz, f);
  const hh = Number(F('H')); if (hh < 9 || hh >= 21) return;
  const props = PropertiesService.getScriptProperties();
  SALARIOS.forEach(s => {
    if (Number(F('d')) < s.dia) return;
    const chave = F('yyyy-MM') + '-' + String(s.dia).padStart(2, '0');
    if (comissoes_().some(c => c.semana === chave && c.tecnico === s.nome + ' (salário)')) return;
    const k = 'sal_' + s.nome + '_' + F('yyyy-MM-dd'); if (props.getProperty(k)) return; props.setProperty(k, '1');
    notificar_(admins_(), '💶 Pagar salário do ' + s.nome + ': ' + s.valor.toFixed(2).replace('.', ',') + ' €', 'Dia ' + s.dia + '. Depois de pagar, marque como pago na app (Valores → Comissões).');
  });
}
function lembrarFaturas_() {
  try { lembrarSalario_(); } catch (e) {}
  const tz = Session.getScriptTimeZone(), hoje = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const hh = Number(Utilities.formatDate(new Date(), tz, 'H')); if (hh < 9 || hh >= 21) return;
  const sh = folha_(); const n = sh.getLastRow() - 1; if (n < 1) return;
  const vals = sh.getRange(2, 1, n, COLS.length).getDisplayValues(); const I = k => COLS.indexOf(k); const pend = [];
  vals.forEach((r, i) => {
    const est = r[I('estado')];
    if ((est === 'Concluído' || est === 'Orçamento recusado') && r[I('fatura')] === 'Sim' && !r[I('faturaFeita')] && r[I('faturaAviso')] !== hoje) {
      const c = r[I('concluidoEm')] || r[I('atualizado')];
      let t = 0; try { t = c ? Utilities.parseDate(c, tz, 'yyyy-MM-dd HH:mm').getTime() : 0; } catch (e) {}
      if (t && Date.now() - t >= 24 * 3600 * 1000) pend.push({ i: i, nome: r[I('nome')] });
    }
  });
  if (!pend.length) return;
  const dest = UTILIZADORES.some(x => x.nome === 'Mariana') ? ['Mariana'] : admins_();
  notificar_(dest, '🧾 ' + pend.length + (pend.length > 1 ? ' faturas por fazer' : ' fatura por fazer'), pend.map(x => x.nome).join(' · ') + ' — abra a app e carregue em "Fatura feita".');
  pend.forEach(x => sh.getRange(x.i + 2, I('faturaAviso') + 1).setValue(hoje));
}

function configurar() {
  folha_();
  folhaCom_();
  pastaFotos_();
folhaAvisos_();
CalendarApp.getDefaultCalendar();
UrlFetchApp.getRequest('https://www.google.com'); // pede autorização para enviar notificações
if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'lembretes'))
ScriptApp.newTrigger('lembretes').timeBased().everyMinutes(10).create();
Logger.log('Pronto! Agora faça Implementar > Nova implementação.');
}
