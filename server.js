/* 钟林 · 费用报销单系统 — 本地静态服务器（无依赖）
 * 用途：用 http 打开本目录，使随包 lib/ 引擎（OCR/PDF/Excel 等）可完全离线加载。
 * 启动：双击「启动.bat」即可。
 */
'use strict';
var http = require('http');
var fs = require('fs');
var path = require('path');
var cp = require('child_process');

var ROOT = __dirname;
var PORTS = [8080, 8081, 8082, 8083, 8084];
var PORT = PORTS[0];

var MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.wasm': 'application/wasm',
  '.wasm.js': 'text/javascript; charset=utf-8',
  '.gz': 'application/octet-stream',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.doc': 'application/msword',
  '.pdf': 'application/pdf',
  '.csv': 'text/csv; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8'
};

function safeJoin(urlPath) {
  // 防止路径穿越
  var p = decodeURIComponent(urlPath.split('?')[0]);
  if (p === '/' || p === '') p = '/index.html';
  var full = path.normalize(path.join(ROOT, p));
  if (full.indexOf(ROOT) !== 0) return null;
  return full;
}

function send(res, status, body, type, extra) {
  res.writeHead(status, Object.assign({
    'Content-Type': type || 'application/octet-stream',
    'Cache-Control': 'no-cache'
  }, extra || {}));
  res.end(body);
}

var server = http.createServer(function (req, res) {
  var full = safeJoin(req.url);
  if (!full) { send(res, 403, 'Forbidden', 'text/plain; charset=utf-8'); return; }
  fs.stat(full, function (err, st) {
    if (err || !st.isFile()) {
      // 兜底回 index.html
      var idx = path.join(ROOT, 'index.html');
      fs.readFile(idx, function (e2, b2) {
        if (e2) { send(res, 404, 'Not Found', 'text/plain; charset=utf-8'); }
        else { send(res, 200, b2, MIME['.html']); }
      });
      return;
    }
    fs.readFile(full, function (e3, buf) {
      if (e3) { send(res, 500, 'Read Error', 'text/plain; charset=utf-8'); return; }
      var ext = path.extname(full).toLowerCase();
      var type = MIME[ext] || 'application/octet-stream';
      // .gz 必须原样返回（Tesseract.js 自带解压），不要加 Content-Encoding
      send(res, 200, buf, type);
    });
  });
});

function tryListen(i) {
  if (i >= PORTS.length) {
    console.error('没有可用端口，请关闭占用 8080-8084 的程序后重试。');
    process.exit(1);
  }
  PORT = PORTS[i];
  server.listen(PORT, '127.0.0.1', function () {
    var url = 'http://localhost:' + PORT + '/index.html';
    console.log('报销单系统已启动：' + url);
    console.log('按 Ctrl+C 关闭服务。');
    openBrowser(url);
  });
  server.once('error', function () {
    tryListen(i + 1);
  });
}

function openBrowser(url) {
  try {
    if (process.platform === 'win32') {
      cp.exec('cmd /c start "" "' + url + '"');
    } else if (process.platform === 'darwin') {
      cp.exec('open "' + url + '"');
    } else {
      cp.exec('xdg-open "' + url + '"');
    }
  } catch (e) { /* 忽略：用户可手动打开链接 */ }
}

tryListen(0);
