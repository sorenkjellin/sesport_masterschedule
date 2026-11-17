// A small FTP server for the tests, speaking what an FTPS server with explicit TLS speaks to
// basic-ftp: AUTH TLS on port 21, then login, EPSV data connections (TLS after PROT P), STOR,
// SIZE, RNFR/RNTO, DELE, CWD. Files are kept in memory.
//
//   const server = new FakeFtpServer({ users: { dmc: 'secret' }, dirs: ['/', '/incoming'] });
//   const port = await server.start();
//   ...
//   server.files.get('/incoming/clip.mp4')   // a Buffer
//   await server.stop();
'use strict';
const net = require('node:net');
const tls = require('node:tls');
const fs = require('node:fs');
const path = require('node:path');

const CERT = fs.readFileSync(path.join(__dirname, 'fixtures', 'cert.pem'));
const KEY = fs.readFileSync(path.join(__dirname, 'fixtures', 'key.pem'));

class FakeFtpServer {
  /**
   * options: users ({ name: password }), dirs (folders that exist), files ({ path: content }),
   * tls (false: refuses AUTH TLS), renameOverwrites (false: RNTO onto an existing file fails, as
   * on some servers), truncateUploads (stores only half of each upload).
   */
  constructor(options = {}) {
    this.users = options.users || { dmc: 'secret' };
    this.dirs = new Set(options.dirs || ['/']);
    this.files = new Map(Object.entries(options.files || {}).map(([p, c]) => [p, Buffer.from(c)]));
    this.tls = options.tls !== false;
    this.renameOverwrites = options.renameOverwrites !== false;
    this.truncateUploads = !!options.truncateUploads;
    this.secureContext = tls.createSecureContext({ cert: CERT, key: KEY });
    this.commands = []; // every command received, the password masked
    this.sockets = new Set();
  }

  start() {
    this.server = net.createServer((socket) => this.session(socket));
    return new Promise((resolve) => this.server.listen(0, '127.0.0.1', () => resolve(this.server.address().port)));
  }

  stop() {
    this.sockets.forEach((s) => s.destroy());
    return new Promise((resolve) => this.server.close(() => resolve()));
  }

  static get cert() { return CERT; }

  session(raw) {
    this.sockets.add(raw);
    raw.on('close', () => this.sockets.delete(raw));
    raw.on('error', () => {});
    const state = { socket: raw, cwd: '/', user: null, loggedIn: false, secure: false, prot: 'C', data: null, from: null };
    const reply = (text) => state.socket.write(text + '\r\n');
    let buffer = '';
    let queue = Promise.resolve();
    const onData = (chunk) => {
      buffer += chunk.toString('utf8');
      let i;
      while ((i = buffer.indexOf('\r\n')) >= 0) {
        const line = buffer.slice(0, i);
        buffer = buffer.slice(i + 2);
        queue = queue.then(() => this.command(line, state, reply, onData)).catch((e) => reply('451 ' + e.message));
      }
    };
    raw.on('data', onData);
    reply('220 Fake DMC FTP ready');
  }

  resolve(state, name) {
    return path.posix.resolve(state.cwd, name);
  }

  async command(line, state, reply, onData) {
    const space = line.indexOf(' ');
    const cmd = (space < 0 ? line : line.slice(0, space)).toUpperCase();
    const arg = space < 0 ? '' : line.slice(space + 1);
    this.commands.push(cmd === 'PASS' ? 'PASS ***' : line);

    const open = ['AUTH', 'USER', 'PASS', 'FEAT', 'QUIT', 'OPTS'];
    if (!state.loggedIn && open.indexOf(cmd) < 0) return reply('530 Please login');

    switch (cmd) {
      case 'AUTH': {
        if (!this.tls || arg.toUpperCase() !== 'TLS') return reply('502 TLS not available');
        reply('234 AUTH TLS ok');
        const raw = state.socket;
        raw.removeListener('data', onData);
        const secure = new tls.TLSSocket(raw, { isServer: true, secureContext: this.secureContext });
        secure.on('error', () => {});
        secure.on('data', onData);
        state.socket = secure;
        state.secure = true;
        return undefined;
      }
      case 'USER': state.user = arg; return reply('331 Password required');
      case 'PASS':
        if (state.user in this.users && this.users[state.user] === arg) {
          state.loggedIn = true;
          return reply('230 Logged in');
        }
        return reply('530 Login incorrect.');
      case 'FEAT': return reply('211-Features:\r\n UTF8\r\n SIZE\r\n EPSV\r\n211 End');
      case 'OPTS': case 'TYPE': case 'STRU': case 'PBSZ': return reply('200 OK');
      case 'PROT': state.prot = arg.toUpperCase(); return reply('200 PROT ' + state.prot);
      case 'PWD': return reply(`257 "${state.cwd}" is the current directory`);
      case 'CWD': {
        const dir = this.resolve(state, arg);
        if (!this.dirs.has(dir)) return reply('550 No such directory');
        state.cwd = dir;
        return reply('250 OK');
      }
      case 'EPSV': {
        const port = await this.openData(state);
        return reply(`229 Entering Extended Passive Mode (|||${port}|)`);
      }
      case 'STOR': return this.store(this.resolve(state, arg), state, reply);
      case 'SIZE': {
        const file = this.files.get(this.resolve(state, arg));
        return file ? reply('213 ' + file.length) : reply('550 No such file');
      }
      case 'DELE': {
        const p = this.resolve(state, arg);
        if (!this.files.has(p)) return reply('550 No such file');
        this.files.delete(p);
        return reply('250 Deleted');
      }
      case 'RNFR': {
        const p = this.resolve(state, arg);
        if (!this.files.has(p)) return reply('550 No such file');
        state.from = p;
        return reply('350 Ready for RNTO');
      }
      case 'RNTO': {
        const to = this.resolve(state, arg);
        if (!state.from) return reply('503 RNFR first');
        if (this.files.has(to) && !this.renameOverwrites) return reply('553 File exists');
        this.files.set(to, this.files.get(state.from));
        this.files.delete(state.from);
        state.from = null;
        return reply('250 Renamed');
      }
      case 'QUIT': reply('221 Bye'); state.socket.end(); return undefined;
      default: return reply('502 Not implemented');
    }
  }

  // Listens for the next data connection; state.data resolves to it (TLS after PROT P).
  openData(state) {
    return new Promise((resolve) => {
      const server = net.createServer();
      state.data = new Promise((resolveData) => {
        server.once('connection', (raw) => {
          this.sockets.add(raw);
          raw.on('error', () => {});
          server.close();
          if (state.prot !== 'P') return resolveData(raw);
          const secure = new tls.TLSSocket(raw, { isServer: true, secureContext: this.secureContext });
          secure.on('error', () => {});
          return resolveData(secure);
        });
      });
      server.listen(0, '127.0.0.1', () => resolve(server.address().port));
    });
  }

  async store(file, state, reply) {
    if (!this.dirs.has(path.posix.dirname(file))) return reply('553 No such directory');
    if (!state.data) return reply('425 Use EPSV first');
    const socket = await state.data;
    state.data = null;
    reply('150 Ready to receive');
    const chunks = [];
    await new Promise((resolve) => {
      socket.on('data', (c) => chunks.push(c));
      socket.on('end', resolve);
      socket.on('close', resolve);
    });
    let content = Buffer.concat(chunks);
    if (this.truncateUploads) content = content.subarray(0, Math.floor(content.length / 2));
    this.files.set(file, content);
    socket.end();
    return reply('226 Transfer complete');
  }
}

module.exports = { FakeFtpServer };
