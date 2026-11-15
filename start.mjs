/**
 * Аудит СМР — запуск всего проекта одной командой.
 * Запуск: node start.mjs   или двойной клик по START.bat
 */
import http from 'node:http';
import path from 'node:path';
import vm from 'node:vm';
import { readFile, writeFile, mkdir, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { exec } from 'node:child_process';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const VENDOR = path.join(ROOT, 'vendor');

const LIBS = [
  { file: 'tailwindcss.js', min: 100000, urls: [
    'https://cdn.tailwindcss.com/3.4.17'
  ]},
  { file: 'xlsx.full.min.js', min: 400000, urls: [
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
  ]},
  { file: 'jszip.min.js', min: 50000, urls: [
    'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
    'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'
  ]}
];

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.webmanifest': 'application/manifest+json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

async function exists(p) { try { await stat(p); return true; } catch { return false; } }

function badCode(code, name) {
  if (/& &|functio n|ret urn|\bv ar\b/.test(code)) return 'файл повреждён (пробелы внутри команд)';
  try { new vm.Script(code, { filename: name }); return null; }
  catch (e) { return 'ошибка синтаксиса: ' + e.message; }
}

async function libOk(lib) {
  try {
    const code = await readFile(path.join(VENDOR, lib.file), 'utf8');
    return code.length >= lib.min && !badCode(code, lib.file);
  } catch { return false; }
}

async function download(url) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 40000);
  try {
    const res = await fetch(url, { signal: ctrl.signal, redirect: 'follow' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return await res.text();
  } finally { clearTimeout(timer); }
}

async function ensureLibs() {
  console.log('');
  console.log('=== Шаг 1. Библиотеки ===');
  await mkdir(VENDOR, { recursive: true });
  const failed = [];
  for (const lib of LIBS) {
    process.stdout.write('  ' + lib.file + ' ... ');
    if (await libOk(lib)) { console.log('уже на месте'); continue; }
    let ok = false, lastErr = '';
    for (const url of lib.urls) {
      try {
        const code = await download(url);
        if (code.length < lib.min) throw new Error('файл подозрительно маленький');
        const bad = badCode(code, lib.file);
        if (bad) throw new Error(bad);
        await writeFile(path.join(VENDOR, lib.file), code, 'utf8');
        console.log('скачан, ' + Math.round(code.length / 1024) + ' КБ');
        ok = true;
        break;
      } catch (e) { lastErr = e.message; }
    }
    if (!ok) { console.log('ОШИБКА — ' + lastErr); failed.push(lib); }
  }
  if (failed.length) {
    console.log('');
    console.log('Не удалось скачать: ' + failed.map(l => l.file).join(', '));
    console.log('Проверьте интернет и запустите снова.');
    console.log('Либо откройте эти ссылки в браузере, нажмите Ctrl+S');
    console.log('и сохраните файлы в папку vendor с указанными именами:');
    for (const l of failed) console.log('  ' + l.urls[0] + '  ->  vendor/' + l.file);
    return false;
  }
  console.log('  Все библиотеки на месте.');
  return true;
}

function startServer() {
  const server = http.createServer(async (req, res) => {
    try {
      if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); res.end(); return; }
      let urlPath = decodeURIComponent((req.url || '/').split('?')[0]);
      if (urlPath === '/') urlPath = '/index.html';
      let full = path.normalize(path.join(ROOT, urlPath));
      if (full !== ROOT && !full.startsWith(ROOT + path.sep)) { res.writeHead(403); res.end('Forbidden'); return; }

      let info = null;
      try { info = await stat(full); } catch {}
      if (info && info.isDirectory()) {
        full = path.join(full, 'index.html');
        try { info = await stat(full); } catch { info = null; }
      }
      if (!info) {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('404 — файл не найден: ' + urlPath);
        return;
      }
      const data = await readFile(full);
      res.writeHead(200, {
        'Content-Type': MIME[path.extname(full).toLowerCase()] || 'application/octet-stream',
        'Content-Length': data.length,
        'Cache-Control': 'no-cache'
      });
      res.end(req.method === 'HEAD' ? undefined : data);
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Ошибка сервера: ' + e.message);
    }
  });

  let port = 3000;
  server.on('error', e => {
    if (e.code === 'EADDRINUSE' && port < 3010) { port++; server.listen(port, '127.0.0.1'); }
    else { console.error('Не удалось запустить сервер: ' + e.message); process.exit(1); }
  });
  server.on('listening', () => {
    const url = 'http://localhost:' + port;
    console.log('');
    console.log('=== Шаг 2. Сервер запущен ===');
    console.log('');
    console.log('  Адрес: ' + url);
    console.log('  Браузер должен открыться сам.');
    console.log('');
    console.log('  НЕ ЗАКРЫВАЙТЕ это окно, пока работаете с приложением.');
    console.log('  Остановить: Ctrl+C или просто закрыть окно.');
    console.log('');
    const cmd = process.platform === 'win32' ? 'start "" "' + url + '"'
              : process.platform === 'darwin' ? 'open "' + url + '"'
              : 'xdg-open "' + url + '"';
    exec(cmd, () => {});
  });
  server.listen(port, '127.0.0.1');
}

(async () => {
  console.log('Папка проекта: ' + ROOT);
  if (!(await exists(path.join(ROOT, 'index.html')))) {
    console.log('');
    console.log('ОШИБКА: в этой папке нет файла index.html');
    console.log('Положите start.mjs и START.bat в ту же папку, где лежит index.html');
    process.exit(1);
  }
  const ok = await ensureLibs();
  if (!ok) process.exit(1);
  startServer();
})();