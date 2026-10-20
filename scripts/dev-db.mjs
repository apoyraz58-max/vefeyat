import { randomBytes } from 'node:crypto';
import { mkdir, readFile, writeFile, chmod, unlink, open } from 'node:fs/promises';
import { createServer } from 'node:net';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import EmbeddedPostgres from 'embedded-postgres';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const local = resolve(root, '.local');
const pidPath = resolve(local, 'db.pid');
const configPath = resolve(local, 'dev-database.json');
let pg;
let stopping = false;
let ownsPid = false;

async function assertPortFree() {
  await new Promise((accept, reject) => {
    const probe = createServer();
    probe.once('error', () => reject(new Error('127.0.0.1:55432 başka bir süreç tarafından kullanılıyor.')));
    probe.listen({ host: '127.0.0.1', port: 55432, exclusive: true }, () => probe.close(accept));
  });
}

async function stop(exitCode = 0) {
  if (stopping) return;
  stopping = true;
  try {
    if (pg) await pg.stop();
  } catch {
    console.error('PostgreSQL kapatılırken hata oluştu; .local/db.log dosyasını inceleyin.');
    exitCode = 1;
  } finally {
    if (ownsPid) await unlink(pidPath).catch(() => {});
    process.exit(exitCode);
  }
}

async function main() {
  process.umask(0o077);
  await mkdir(local, { recursive: true, mode: 0o700 });
  let config;
  try {
    config = JSON.parse(await readFile(configPath, 'utf8'));
  } catch (error) {
    if (error.code !== 'ENOENT') throw new Error('Yerel veritabanı yapılandırması okunamadı.');
    config = { NAME: 'vefeyat', USER: 'vefeyat', PASSWORD: randomBytes(32).toString('hex'), HOST: '127.0.0.1', PORT: 55432 };
    await writeFile(configPath, `${JSON.stringify(config, null, 2)}\n`, { mode: 0o600, flag: 'wx' });
  }
  if (config.NAME !== 'vefeyat' || config.USER !== 'vefeyat' || config.HOST !== '127.0.0.1' || Number(config.PORT) !== 55432 || typeof config.PASSWORD !== 'string' || config.PASSWORD.length < 32) {
    throw new Error('Yerel veritabanı yapılandırması beklenen geliştirme ayarlarıyla uyuşmuyor.');
  }
  await chmod(configPath, 0o600);
  try {
    const handle = await open(pidPath, 'wx', 0o600);
    await handle.writeFile(String(process.pid));
    await handle.close();
    ownsPid = true;
  } catch (error) {
    if (error.code !== 'EEXIST') throw error;
    const existingPid = Number((await readFile(pidPath, 'utf8')).trim());
    if (!Number.isSafeInteger(existingPid) || existingPid <= 1) throw new Error('Geçersiz db.pid dosyası; güvenli biçimde başlatılamadı.');
    try {
      process.kill(existingPid, 0);
      throw new Error('db.pid canlı bir sürece ait; npm run stop ile mevcut geliştirme servislerini kapatın.');
    } catch (probeError) {
      if (probeError.code !== 'ESRCH') throw probeError;
    }
    await unlink(pidPath);
    const handle = await open(pidPath, 'wx', 0o600);
    await handle.writeFile(String(process.pid));
    await handle.close();
    ownsPid = true;
  }
  await assertPortFree();
  pg = new EmbeddedPostgres({
    databaseDir: resolve(local, 'postgres'),
    user: config.USER,
    password: config.PASSWORD,
    port: Number(config.PORT),
    authMethod: 'scram-sha-256',
    persistent: true,
    createPostgresUser: false,
    initdbFlags: ['--encoding=UTF8', '--locale=C.UTF-8'],
    postgresFlags: ['-c', 'listen_addresses=127.0.0.1', '-c', `unix_socket_directories=${local}`],
  });
  try {
    await readFile(resolve(local, 'postgres', 'PG_VERSION'), 'utf8');
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
    await pg.initialise();
  }
  await pg.start();
  const client = pg.getPgClient('postgres', config.HOST);
  try {
    await client.connect();
    const exists = await client.query('SELECT 1 FROM pg_database WHERE datname = $1', [config.NAME]);
    if (!exists.rowCount) await client.query('CREATE DATABASE vefeyat');
  } finally {
    await client.end();
  }
  console.log('Geliştirme PostgreSQL hazır: 127.0.0.1:55432 / vefeyat');
  // Keep the manager alive to forward termination signals to PostgreSQL.
  setInterval(() => {}, 60_000);
}

process.on('SIGTERM', () => void stop());
process.on('SIGINT', () => void stop());
main().catch(async (error) => {
  console.error(error?.message || 'PostgreSQL başlatılamadı; önceki günlük satırlarını inceleyin.');
  await stop(1);
});
