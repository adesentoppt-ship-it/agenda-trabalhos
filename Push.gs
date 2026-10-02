/**
 * NOTIFICAÇÕES PUSH (Web Push / VAPID) — envia avisos para os telemóveis
 * que carregaram em "Ativar notificações" na app.
 * A chave privada fica só aqui no Apps Script (nunca no GitHub).
 */
const VAPID_PRIVADA_HEX = 'COLE_AQUI_A_CHAVE_PRIVADA';
const VAPID_PUBLICA = 'BClgIdU7OVnICrWWfUoxW4s0x1Y7avuO_xoG6xPynUNUTxZI9ZCJ0vtsbRScN-LNf0rV_Tc5dEMpyPvimg8l5Go';
const VAPID_SUB = 'https://adesentoppt-ship-it.github.io/agenda-trabalhos/';
const FOLHA_AVISOS = 'Notificacoes';

/* ---------- curva P-256 (ECDSA) em BigInt ---------- */
const N0_ = BigInt(0), N1_ = BigInt(1), N2_ = BigInt(2), N3_ = BigInt(3), N4_ = BigInt(4), N8_ = BigInt(8); // números BigInt (o editor não aceita 0n)

const P256 = {
  p: BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'),
  a: BigInt('0xffffffff00000001000000000000000000000000fffffffffffffffffffffffc'),
  n: BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
  gx: BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
  gy: BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')
};
function pmod_(a, m) { const r = a % m; return r < N0_ ? r + m : r; }
function pinv_(a, m) {
  let [r0, r1] = [pmod_(a, m), m], [s0, s1] = [N1_, N0_];
  while (r1 !== N0_) { const q = r0 / r1; [r0, r1] = [r1, r0 - q * r1]; [s0, s1] = [s1, s0 - q * s1]; }
  return pmod_(s0, m);
}
// Pontos em coordenadas Jacobianas [X, Y, Z] (evita inversões a cada passo)
function jdbl_(P) {
  const p = P256.p; const [X, Y, Z] = P;
  if (Y === N0_ || Z === N0_) return [N0_, N1_, N0_];
  const YY = Y * Y % p, S = N4_ * X * YY % p, ZZ = Z * Z % p;
  const M = pmod_(N3_ * X * X + P256.a * ZZ % p * ZZ, p);
  const X3 = pmod_(M * M - N2_ * S, p);
  const Y3 = pmod_(M * (S - X3) - N8_ * YY * YY, p);
  const Z3 = N2_ * Y * Z % p;
  return [X3, Y3, Z3];
}
function jadd_(P, Q) {
  const p = P256.p;
  if (P[2] === N0_) return Q; if (Q[2] === N0_) return P;
  const Z1Z1 = P[2] * P[2] % p, Z2Z2 = Q[2] * Q[2] % p;
  const U1 = P[0] * Z2Z2 % p, U2 = Q[0] * Z1Z1 % p;
  const S1 = P[1] * Q[2] % p * Z2Z2 % p, S2 = Q[1] * P[2] % p * Z1Z1 % p;
  const H = pmod_(U2 - U1, p), R = pmod_(S2 - S1, p);
  if (H === N0_) return R === N0_ ? jdbl_(P) : [N0_, N1_, N0_];
  const HH = H * H % p, HHH = H * HH % p, V = U1 * HH % p;
  const X3 = pmod_(R * R - HHH - N2_ * V, p);
  const Y3 = pmod_(R * (V - X3) - S1 * HHH, p);
  const Z3 = P[2] * Q[2] % p * H % p;
  return [X3, Y3, Z3];
}
function jmul_(k) {
  let R = [N0_, N1_, N0_], Q = [P256.gx, P256.gy, N1_];
  while (k > N0_) { if (k & N1_) R = jadd_(R, Q); Q = jdbl_(Q); k >>= N1_; }
  const zi = pinv_(R[2], P256.p), zi2 = zi * zi % P256.p;
  return [R[0] * zi2 % P256.p, R[1] * zi2 % P256.p * zi % P256.p];
}
function bytesToBig_(b) { let h = '0x'; b.forEach(x => h += ((x & 255) + 256).toString(16).slice(1)); return BigInt(h === '0x' ? '0' : h); }
function bigToBytes32_(v) { const h = v.toString(16).padStart(64, '0'); const out = []; for (let i = 0; i < 64; i += 2) out.push(parseInt(h.substr(i, 2), 16)); return out; }
function sha256Bytes_(bytes) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes.map(x => x > 127 ? x - 256 : x)).map(x => x & 255); }
function b64url_(bytes) { return Utilities.base64EncodeWebSafe(bytes.map(x => x > 127 ? x - 256 : x)).replace(/=+$/, ''); }

function es256_(texto, dHex) {
  const d = BigInt('0x' + dHex.replace(/^0x/, ''));
  const msg = Utilities.newBlob(texto).getBytes().map(x => x & 255);
  const e = bytesToBig_(sha256Bytes_(msg));
  for (let tent = 0; tent < 10; tent++) {
    // k secreto e único: hash(chave privada + mensagem + aleatório)
    const semente = bigToBytes32_(d).concat(sha256Bytes_(msg), Utilities.newBlob(Utilities.getUuid() + Date.now() + Math.random()).getBytes().map(x => x & 255));
    const k = pmod_(bytesToBig_(sha256Bytes_(semente)), P256.n);
    if (k === N0_) continue;
    const r = pmod_(jmul_(k)[0], P256.n); if (r === N0_) continue;
    const s = pmod_(pinv_(k, P256.n) * (e + r * d), P256.n); if (s === N0_) continue;
    return bigToBytes32_(r).concat(bigToBytes32_(s));
  }
  throw new Error('Falha a assinar');
}

function jwtVapid_(audiencia) {
  const cache = CacheService.getScriptCache(), chave = 'jwt_' + audiencia;
  const c = cache.get(chave); if (c) return c;
  const enc = o => b64url_(Utilities.newBlob(JSON.stringify(o)).getBytes().map(x => x & 255));
  const corpo = enc({ typ: 'JWT', alg: 'ES256' }) + '.' + enc({ aud: audiencia, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: VAPID_SUB });
  const jwt = corpo + '.' + b64url_(es256_(corpo, VAPID_PRIVADA_HEX));
  cache.put(chave, jwt, 6 * 3600);
  return jwt;
}

/* ---------- subscrições (uma por telemóvel) ---------- */
function folhaAvisos_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FOLHA_AVISOS);
  if (!sh) { sh = ss.insertSheet(FOLHA_AVISOS); sh.appendRow(['endpoint', 'utilizador', 'criado']); sh.setFrozenRows(1); }
  return sh;
}
function subscrever_(d, nome) {
  const ep = String(d.endpoint || '');
  if (!/^https:\/\//.test(ep)) throw new Error('Subscrição inválida');
  const sh = folhaAvisos_(), vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === ep) { sh.getRange(i + 1, 2).setValue(nome); return { ok: true }; }
  sh.appendRow([ep, nome, new Date()]);
  return { ok: true };
}

/* Guarda a mensagem e "acorda" os telemóveis; o telemóvel vem buscar o texto (acao 'aviso') */
function notificar_(nomes, titulo, corpo) {
  if (!nomes || !nomes.length) return;
  const props = PropertiesService.getScriptProperties();
  nomes.forEach(n => props.setProperty('aviso_' + n, JSON.stringify({ titulo: titulo, corpo: corpo, t: Date.now() })));
  const sh = folhaAvisos_(), vals = sh.getDataRange().getValues(), apagar = [];
  for (let i = 1; i < vals.length; i++) {
    if (nomes.indexOf(vals[i][1]) < 0) continue;
    const ep = vals[i][0];
    try {
      const aud = ep.match(/^https:\/\/[^/]+/)[0];
      const r = UrlFetchApp.fetch(ep, {
        method: 'post', muteHttpExceptions: true, payload: '',
        headers: { TTL: '86400', Urgency: 'high', Authorization: 'vapid t=' + jwtVapid_(aud) + ', k=' + VAPID_PUBLICA }
      });
      const c = r.getResponseCode();
      if (c === 404 || c === 410) apagar.push(i + 1); // telemóvel desativou
    } catch (e) { /* ignora um telemóvel com erro */ }
  }
  apagar.reverse().forEach(l => sh.deleteRow(l));
}
function ultimoAviso_(nome) {
  const v = PropertiesService.getScriptProperties().getProperty('aviso_' + nome);
  return v ? JSON.parse(v) : null;
}

/* Corre de 10 em 10 min (acionador criado por configurar): avisa o técnico 1h antes */
function lembretes() {
  const tz = Session.getScriptTimeZone(), agora = Date.now();
  const sh = folha_(); const n = sh.getLastRow() - 1; if (n < 1) return;
  const vals = sh.getRange(2, 1, n, COLS.length).getDisplayValues();
  const iAv = COLS.indexOf('lembrado');
  vals.forEach((r, i) => {
    const o = {}; COLS.forEach((c, k) => o[c] = r[k]);
    if (o.estado !== 'Agendado' || o.lembrado === 'sim' || !o.data || !o.hora) return;
    const ini = Utilities.parseDate(o.data + ' ' + o.hora, tz, 'yyyy-MM-dd HH:mm').getTime();
    const falta = (ini - agora) / 60000;
    if (falta > 0 && falta <= 65) {
notificar_(o.tecnico ? tecs_(o.tecnico) : trabalhadores_(), '⏰ Daqui a ' + Math.round(falta) + ' min: ' + (o.servico || 'Trabalho'),
        o.hora + ' · ' + o.nome + (o.morada ? '\n' + o.morada : ''));
      sh.getRange(i + 2, iAv + 1).setValue('sim');
    }
  });
}
