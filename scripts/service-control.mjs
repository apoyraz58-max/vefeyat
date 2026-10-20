import { spawn } from 'node:child_process';
import { mkdir, readFile, unlink, writeFile } from 'node:fs/promises';
import { openSync, closeSync } from 'node:fs';
import { createServer } from 'node:net';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const local = resolve(root, '.local');
const services = {
  db: { command: process.execPath, args: [resolve(root, 'scripts/dev-db.mjs')], match: [resolve(root, 'scripts/dev-db.mjs')], port: 55432 },
  web: { command: resolve(root, '.venv/bin/python'), args: [resolve(root, 'manage.py'), 'runserver', '127.0.0.1:8000', '--noreload'], match: [resolve(root, 'manage.py'), 'runserver', '127.0.0.1:8000', '--noreload'], port: 8000 },
};
const delay = (ms) => new Promise((accept) => setTimeout(accept, ms));

async function ownedPid(name) {
  let pid;
  try { pid = Number((await readFile(resolve(local, `${name}.pid`), 'utf8')).trim()); }
  catch (error) { if (error.code === 'ENOENT') return null; throw error; }
  if (!Number.isSafeInteger(pid) || pid <= 1) throw new Error(`${name}.pid geçersiz; hiçbir süreç durdurulmadı.`);
  let args;
  try {
    process.kill(pid, 0);
    args = (await readFile(`/proc/${pid}/cmdline`, 'utf8')).split('\0');
    const status = await readFile(`/proc/${pid}/status`, 'utf8');
    if (/^State:\s+Z\b/m.test(status)) {
      await unlink(resolve(local, `${name}.pid`)).catch(() => {});
      return null;
    }
  } catch (error) {
    if (error.code === 'ESRCH' || error.code === 'ENOENT') {
      await unlink(resolve(local, `${name}.pid`)).catch(() => {});
      return null;
    }
    throw error;
  }
  if (!services[name].match.every((arg) => args.includes(arg))) {
    throw new Error(`${name}.pid farklı bir sürece ait. Güvenlik için işleme dokunulmadı.`);
  }
  return pid;
}

async function assertPortFree(port) {
  await new Promise((accept, reject) => {
    const probe = createServer();
    probe.once('error', () => reject(new Error(`127.0.0.1:${port} başka bir süreç tarafından kullanılıyor; mevcut süreç değiştirilmedi.`)));
    probe.listen({ host: '127.0.0.1', port, exclusive: true }, () => probe.close(accept));
  });
}

async function start(name) {
  if (await ownedPid(name)) { console.log(`${name}: mevcut geliştirme servisi kullanılıyor.`); return; }
  await assertPortFree(services[name].port);
  const log = openSync(resolve(local, `${name}.log`), 'a', 0o600);
  const child = spawn(services[name].command, services[name].args, { cwd: root, detached: true, stdio: ['ignore', log, log] });
  closeSync(log);
  await new Promise((accept, reject) => { child.once('spawn', accept); child.once('error', reject); });
  child.unref();
  // The database manager owns its PID file to prevent concurrent managers.
  if (name !== 'db') await writeFile(resolve(local, `${name}.pid`), String(child.pid), { mode: 0o600, flag: 'wx' });
  for (let attempt = 0; attempt < 30; attempt++) {
    await delay(100);
    if (await ownedPid(name)) { console.log(`${name}: başlatıldı.`); return; }
    try { process.kill(child.pid, 0); } catch { break; }
  }
  throw new Error(`${name} başlatılamadı; .local/${name}.log dosyasını inceleyin.`);
}

async function stop(name) {
  const pid = await ownedPid(name);
  if (!pid) return;
  process.kill(pid, 'SIGTERM');
  for (let attempt = 0; attempt < 100; attempt++) {
    await delay(100);
    if (!(await ownedPid(name))) { console.log(`${name}: durduruldu.`); return; }
  }
  throw new Error(`${name} henüz kapanmadı; zorla sonlandırılmadı.`);
}

process.umask(0o077);
await mkdir(local, { recursive: true, mode: 0o700 });
try {
  const action = process.argv[2];
  if (action === 'start-db') await start('db');
  else if (action === 'start-web') await start('web');
  else if (action === 'stop-web') await stop('web');
  else if (action === 'restart-web') { await stop('web'); await start('web'); }
  else if (action === 'stop') { await stop('web'); await stop('db'); }
  else throw new Error('Kullanım: service-control.mjs start-db|start-web|stop-web|restart-web|stop');
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
