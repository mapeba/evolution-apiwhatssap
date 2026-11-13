// Uso (después de `npm ci`, que aplica patches/ con patch-package): node scripts/prueba-parche-baileys-status.mjs
// Debe imprimir acks_status con VIVO1 y COLA1. Si sale vacío, el parche de Baileys no quedó puesto.
// Prueba del parche: ¿Baileys confirma un <status> de primer nivel? (en vivo y en la cola sin conexión)
import { WebSocketServer } from 'ws';
import makeWASocket, { initAuthCreds } from 'baileys';
const enviados = [];
const errores = [];
const logger = {
  level: 'debug',
  child() { return logger; },
  trace() {}, info() {}, warn() {},
  debug(obj, msg) { if (msg === 'sent ack') enviados.push(obj.sent); },
  error(obj, msg) { errores.push(msg || String(obj)); },
};
const wss = new WebSocketServer({ port: 0 });
await new Promise((r) => wss.on('listening', r));
const url = `ws://127.0.0.1:${wss.address().port}`;
const keys = {}; const store = { get: async (t, ids) => Object.fromEntries(ids.map((i) => [i, keys[`${t}-${i}`]])), set: async (d) => { for (const t in d) for (const i in d[t]) keys[`${t}-${i}`] = d[t][i]; } };
const sock = makeWASocket({ auth: { creds: initAuthCreds(), keys: store }, logger, waWebSocketUrl: new URL(url), connectTimeoutMs: 60000, printQRInTerminal: false, markOnlineOnConnect: false, fireInitQueries: false });
for (let i = 0; i < 50 && !sock.ws.isOpen; i++) await new Promise((r) => setTimeout(r, 100));
const status = (id, offline) => ({ tag: 'status', attrs: { from: '573001112233@s.whatsapp.net', id, type: 'media', t: '1763000000', ...(offline ? { offline: '1' } : {}) }, content: [] });
// 1) en vivo
sock.ws.emit('CB:status', status('VIVO1'));
// 2) cola sin conexión: primero un mensaje basura que hace fallar al procesador, luego un status
sock.ws.emit('CB:message', { tag: 'message', attrs: { from: 'roto', id: 'ROTO1', offline: '1' }, content: 'no-es-un-nodo' });
sock.ws.emit('CB:status', status('COLA1', true));
await new Promise((r) => setTimeout(r, 1500));
const acks = enviados.filter((s) => s.class === 'status').map((s) => `${s.id}:${s.type}`);
console.log(JSON.stringify({ ws_abierto: sock.ws.isOpen, acks_status: acks, otros_acks: enviados.filter((s) => s.class !== 'status').map((s) => s.class + ':' + s.id), errores: errores.slice(0, 4) }));
sock.end(undefined); wss.close(); process.exit(0);
